"""
PDF (vision-based) Questionnaire Parser
=======================================
Parses questionnaires by converting the xlsx to PDF, rendering each page to a
PNG, and asking an Azure OpenAI vision model (via langchain) to extract the
question_id / question / answer triples. Read-only / locked PDFs are fine —
pypdfium2 renders them without modification.

Why a sliding window?
---------------------
Long questionnaires get split across many pages, and a single question (or its
answer) can land at the bottom of one page and continue at the top of the next.
To make sure every Q/A is seen in full at least once, we send the LLM
overlapping pairs of consecutive pages: [p1, p2], [p2, p3], [p3, p4], ...
Each Q/A then appears in at least one window where it is not cut. After all
windows are processed, results are deduped by question_id, preferring the most
complete (longest non-empty answer / longest question text) entry.

Inputs
------
A .xlsx file or a directory of .xlsx files. (A .pdf may also be passed
directly, in which case the xlsx→pdf step is skipped.)

Outputs
-------
For each input file, in ``--output-dir``:
    <name>.json       – list of {question_id, question, answer, source_pages}
    <name>.pdf        – the intermediate PDF (kept for inspection / debugging)
And a combined:
    <stem>.xlsx       – one row per question across all inputs

Environment
-----------
AZURE_OPENAI_API_KEY        – API key
AZURE_OPENAI_ENDPOINT       – e.g. https://my-resource.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT     – chat deployment name (must be a vision model,
                              e.g. gpt-4o / gpt-4o-mini / gpt-4.1)
AZURE_OPENAI_API_VERSION    – e.g. 2024-10-21 (default if unset)

Dependencies (add to pyproject.toml or install separately):
    pip install pypdfium2 langchain langchain-openai openpyxl pillow
And the system needs LibreOffice ("soffice") on PATH for the xlsx→pdf step.

Usage:
    python pdf_parse.py <file_or_dir> [--output-dir ./output] \
        [--dpi 200] [--window 2] [--stride 1] [--keep-pngs]
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import openpyxl
import pypdfium2 as pdfium
from PIL import Image
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import AzureChatOpenAI
from pydantic import BaseModel, Field

DEFAULT_API_VERSION = "2024-10-21"

# ─── Data model the LLM must return ─────────────────────────────────────────
class QA(BaseModel):
    question_id: str = Field(
        description=(
            "The question identifier exactly as printed (e.g. 'Q1', 'Q12', "
            "'1.a', '3.2'). If the page shows no explicit id but the row is "
            "clearly a question, use the visible numbering or a stable label."
        )
    )
    question: str = Field(
        description=(
            "The full question text, including any sub-prompts, italic hints, "
            "branch labels, or option labels that belong to it. Preserve "
            "newlines between sub-prompts."
        )
    )
    answer: str = Field(
        description=(
            "The respondent's answer text as written in the answer box / "
            "dropdown. Empty string if the answer box is blank. If there is "
            "both a side-dropdown (YES/NO) and a free-text answer, join them "
            "with ' - ' (e.g. 'YES - Acme Bank')."
        )
    )

class QAList(BaseModel):
    questions: list[QA] = Field(default_factory=list)

# ─── xlsx → pdf ─────────────────────────────────────────────────────────────
def xlsx_to_pdf(xlsx_path: Path, out_dir: Path) -> Path:
    """Convert an .xlsx to .pdf using headless LibreOffice. Returns the pdf path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        "soffice", "--headless", "--norestore", "--nolockcheck",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(xlsx_path),
    ]
    # LibreOffice can deadlock if another instance shares the user profile; give
    # each call its own throwaway profile directory.
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile:
        env = os.environ.copy()
        env["HOME"] = profile
        result = subprocess.run(
            cmd, env=env, capture_output=True, text=True, timeout=300,
        )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice failed for {xlsx_path.name}:\n"
            f"  stdout: {result.stdout.strip()}\n"
            f"  stderr: {result.stderr.strip()}"
        )
    pdf = out_dir / f"{xlsx_path.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"Expected {pdf} after LibreOffice convert, not found.")
    return pdf

# ─── pdf → PNG pages ────────────────────────────────────────────────────────
def render_pdf_pages(pdf_path: Path, dpi: int = 200) -> list[Image.Image]:
    """Render every page to a PIL Image. Works on read-only / locked PDFs."""
    scale = dpi / 72.0
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        return [
            pdf[i].render(scale=scale).to_pil().convert("RGB")
            for i in range(len(pdf))
        ]
    finally:
        pdf.close()

def image_to_data_url(img: Image.Image, max_side: int = 2000) -> str:
    """PNG-encode and base64-wrap a PIL image. Downscale if a side exceeds max_side."""
    w, h = img.size
    if max(w, h) > max_side:
        ratio = max_side / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

# ─── Azure OpenAI extraction ────────────────────────────────────────────────
SYSTEM_PROMPT = (
    "You extract questionnaire content from images of PDF pages. "
    "Each image is one page of a multi-page questionnaire. The user may send "
    "two consecutive pages at once so that questions spanning the page break "
    "are fully visible. "
    "Identify every question that has a printed question id (e.g. 'Q1', "
    "'Q12.a', '3.2'), and return its full question text and the respondent's "
    "answer. "
    "Rules:\n"
    "- Only return a question if its id is fully visible. If an id is on the "
    "  first image and its answer continues on the second image, return the "
    "  complete merged item. If an id starts at the very bottom and its body "
    "  is cut off (not visible on either image in this window), skip it — "
    "  another window will see it in full.\n"
    "- 'answer' must contain only what the respondent wrote (free text, "
    "  dropdown selection, ticked option). If the answer box is empty, use "
    "  an empty string.\n"
    "- If a question has both a side dropdown (YES/NO etc.) and a free-text "
    "  box, join them with ' - ' (e.g. 'YES - Acme Bank').\n"
    "- Preserve newlines inside the question text when sub-prompts / hints "
    "  appear on separate lines.\n"
    "- Do not invent ids, questions, or answers. Do not summarise.\n"
)

def build_llm(model_kwargs: dict | None = None) -> AzureChatOpenAI:
    endpoint   = os.environ.get("AZURE_OPENAI_ENDPOINT")
    api_key    = os.environ.get("AZURE_OPENAI_API_KEY")
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT")
    api_ver    = os.environ.get("AZURE_OPENAI_API_VERSION", DEFAULT_API_VERSION)
    missing = [n for n, v in [
        ("AZURE_OPENAI_ENDPOINT", endpoint),
        ("AZURE_OPENAI_API_KEY", api_key),
        ("AZURE_OPENAI_DEPLOYMENT", deployment),
    ] if not v]
    if missing:
        raise RuntimeError(
            "Missing required environment variables: " + ", ".join(missing)
        )
    return AzureChatOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        azure_deployment=deployment,
        api_version=api_ver,
        temperature=0,
        max_retries=3,
        **(model_kwargs or {}),
    )

def extract_from_window(
    llm: AzureChatOpenAI,
    page_images: list[Image.Image],
    page_numbers: list[int],
) -> list[dict]:
    """Run the vision model on a window of consecutive pages and return parsed
    items. Each returned dict carries `source_pages` so we know which window
    produced it."""
    user_content: list[dict] = [{
        "type": "text",
        "text": (
            f"This window contains page(s) {', '.join(map(str, page_numbers))} "
            f"of the questionnaire (in order). Extract every question per the "
            f"rules in the system message."
        ),
    }]
    for img in page_images:
        user_content.append({
            "type": "image_url",
            "image_url": {"url": image_to_data_url(img), "detail": "high"},
        })

    structured = llm.with_structured_output(QAList)
    result: QAList = structured.invoke([
        SystemMessage(content=SYSTEM_PROMPT),
        HumanMessage(content=user_content),
    ])
    return [
        {
            "question_id": q.question_id.strip(),
            "question":    q.question.strip(),
            "answer":      q.answer.strip(),
            "source_pages": list(page_numbers),
        }
        for q in result.questions
        if q.question_id.strip()
    ]

# ─── Sliding window + dedup ─────────────────────────────────────────────────
def iter_windows(n_pages: int, window: int, stride: int) -> Iterable[list[int]]:
    """1-indexed page-number windows. The final window is always anchored at the
    last page so it never gets dropped."""
    if n_pages == 0:
        return
    if n_pages <= window:
        yield list(range(1, n_pages + 1))
        return
    starts = list(range(1, n_pages - window + 2, stride))
    if starts[-1] != n_pages - window + 1:
        starts.append(n_pages - window + 1)
    for s in starts:
        yield list(range(s, s + window))

def _completeness(item: dict) -> tuple[int, int]:
    """Score used to pick the better duplicate. Prefer non-empty answers, then
    longer question text, then longer answer text."""
    has_answer = 1 if item["answer"] else 0
    return (has_answer, len(item["question"]) + len(item["answer"]))

def dedupe_by_id(items: list[dict]) -> list[dict]:
    best: dict[str, dict] = {}
    for it in items:
        qid = it["question_id"]
        if qid not in best:
            best[qid] = it
            continue
        if _completeness(it) > _completeness(best[qid]):
            merged_pages = sorted(set(best[qid]["source_pages"]) | set(it["source_pages"]))
            it = {**it, "source_pages": merged_pages}
            best[qid] = it
        else:
            best[qid]["source_pages"] = sorted(
                set(best[qid]["source_pages"]) | set(it["source_pages"])
            )
    # Natural sort by the numeric part of the id when possible, otherwise lexical.
    def sort_key(qid: str):
        import re
        nums = [int(n) for n in re.findall(r"\d+", qid)]
        return (nums or [10**9], qid)
    return [best[qid] for qid in sorted(best, key=sort_key)]

# ─── Orchestration ──────────────────────────────────────────────────────────
@dataclass
class ParseResult:
    file_name: str
    pdf_path: Path
    items: list[dict]

def parse_one(
    src: Path,
    out_dir: Path,
    llm: AzureChatOpenAI,
    dpi: int,
    window: int,
    stride: int,
    keep_pngs: bool,
) -> ParseResult:
    if src.suffix.lower() == ".pdf":
        pdf_path = src
    elif src.suffix.lower() in {".xlsx", ".xlsm"}:
        print(f"  → converting {src.name} to PDF via LibreOffice")
        pdf_path = xlsx_to_pdf(src, out_dir)
    else:
        raise ValueError(f"Unsupported input type: {src.suffix}")

    print(f"  → rendering {pdf_path.name} at {dpi} dpi")
    pages = render_pdf_pages(pdf_path, dpi=dpi)
    n = len(pages)
    print(f"    {n} page(s)")

    if keep_pngs:
        png_dir = out_dir / f"{src.stem}_pages"
        png_dir.mkdir(parents=True, exist_ok=True)
        for i, img in enumerate(pages, 1):
            img.save(png_dir / f"page_{i:03d}.png")

    all_items: list[dict] = []
    for win in iter_windows(n, window=window, stride=stride):
        print(f"    LLM call on pages {win}")
        win_imgs = [pages[p - 1] for p in win]
        try:
            items = extract_from_window(llm, win_imgs, win)
        except Exception as e:
            print(f"      ! failed on window {win}: {e}", file=sys.stderr)
            continue
        all_items.extend(items)

    merged = dedupe_by_id(all_items)
    print(f"    extracted {len(merged)} unique question(s)")
    return ParseResult(file_name=src.stem, pdf_path=pdf_path, items=merged)

# ─── Writers ────────────────────────────────────────────────────────────────
XLSX_FIELDS = ["file_name", "question_id", "question", "answer", "source_pages"]

def write_json(result: ParseResult, path: Path) -> None:
    path.write_text(
        json.dumps(
            {"file_name": result.file_name, "questions": result.items},
            indent=2, ensure_ascii=False,
        ),
        encoding="utf-8",
    )

def write_xlsx(results: list[ParseResult], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "questions"
    ws.append(XLSX_FIELDS)
    for r in results:
        for it in r.items:
            ws.append([
                r.file_name,
                it["question_id"],
                it["question"],
                it["answer"],
                ", ".join(map(str, it["source_pages"])),
            ])
    wb.save(path)

# ─── CLI ────────────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(
        description="PDF-vision questionnaire parser (Azure OpenAI via langchain).",
    )
    ap.add_argument("target", help="Path to .xlsx/.pdf file OR directory of them")
    ap.add_argument("--output-dir", default="./output",
                    help="Output directory (default: ./output)")
    ap.add_argument("--dpi", type=int, default=200,
                    help="Render DPI for PDF→PNG (default 200)")
    ap.add_argument("--window", type=int, default=2,
                    help="Pages per LLM call (default 2 = sliding pair)")
    ap.add_argument("--stride", type=int, default=1,
                    help="Step between window starts (default 1 = overlap)")
    ap.add_argument("--keep-pngs", action="store_true",
                    help="Also save the rendered page PNGs alongside the PDF")
    args = ap.parse_args()

    target  = Path(args.target)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if target.is_dir():
        files = sorted(
            p for p in target.iterdir()
            if p.suffix.lower() in {".xlsx", ".xlsm", ".pdf"}
            and not p.name.startswith("~$")
        )
    else:
        files = [target]
    if not files:
        print(f"No xlsx/pdf inputs found in {target}", file=sys.stderr)
        sys.exit(1)

    llm = build_llm()

    results: list[ParseResult] = []
    for fp in files:
        print(f"→ {fp.name}")
        try:
            res = parse_one(
                fp, out_dir, llm,
                dpi=args.dpi, window=args.window, stride=args.stride,
                keep_pngs=args.keep_pngs,
            )
        except Exception as e:
            print(f"  ✗ failed: {e}", file=sys.stderr)
            continue
        write_json(res, out_dir / f"{fp.stem}.json")
        results.append(res)

    if not results:
        print("No outputs produced.", file=sys.stderr)
        sys.exit(2)

    if target.is_dir():
        from datetime import date
        xlsx_path = out_dir / f"{target.name}_pdf_{date.today().isoformat()}.xlsx"
    else:
        xlsx_path = out_dir / f"{results[0].file_name}.xlsx"
    write_xlsx(results, xlsx_path)
    print(f"\n✓ Wrote {xlsx_path}")

if __name__ == "__main__":
    main()
