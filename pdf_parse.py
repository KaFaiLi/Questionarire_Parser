"""
PDF (vision-based) Questionnaire Parser
=======================================
Parses questionnaires by converting the xlsx to PDF (via Excel + pywin32),
rendering each page to a PNG, and asking an Azure OpenAI vision model (via
langchain) to extract the question_id / question / answer triples. Read-only /
locked PDFs are fine — pypdfium2 renders them without modification.

The xlsx→pdf step only keeps columns **B:AN** on every sheet (the print area is
restricted before export) so the irrelevant left margin (col A) and any noise
to the right of AN never reach the LLM.

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
    pip install pypdfium2 langchain langchain-openai openpyxl pillow pywin32
The xlsx→pdf step uses Excel COM automation via pywin32, so it requires
**Windows with Microsoft Excel installed**. PDF rendering and the LLM call
both work cross-platform.

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
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import openpyxl
import pypdfium2 as pdfium
from PIL import Image
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import AzureChatOpenAI
from pydantic import BaseModel, Field

# Helpers reused from the cell-based parser. The reconcile pass needs the same
# notion of "answer cell" (grey/yellow fill) and question-id pattern.
from parse import (
    ANSWER_FILLS,
    CONTENT_END_COL,
    CONTENT_START_COL,
    Q_PATTERN,
    build_merged_lookup,
    classify_fill,
    effective_value,
)

DEFAULT_API_VERSION = "2024-10-21"

# xlsx→pdf print-area limits (inclusive). Columns outside this range are
# excluded from every sheet before the PDF is exported.
PRINT_COL_START = "B"
PRINT_COL_END   = "AN"

# A column safely past PRINT_COL_END used as a measurement surface for the
# merged-cell row-height AutoFit pass. Nothing written here can reach the PDF.
SCRATCH_COL_LETTER = "AZ"

# Excel constant for ExportAsFixedFormat. (xlTypePDF = 0.)
XL_TYPE_PDF = 0

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
def _sheet_last_row(ws) -> int:
    """Last row that has any content anywhere on the sheet, via Excel's
    UsedRange. Falls back to 1 if the sheet is empty."""
    used = ws.UsedRange
    if used is None:
        return 1
    return max(1, used.Row + used.Rows.Count - 1)

def _is_blank_value(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())

def _autofit_merged_rows(ws, last_row: int) -> int:
    """Grow row heights so each merged wrap-text cell fits its full text.

    Excel does NOT auto-fit row height for merged cells (a well-known
    limitation), so answer boxes with long wrapped text get visually clipped
    in the exported PDF. We measure the required height for each merged
    wrap-text cell using a scratch column past the export band, and grow
    the anchor row accordingly. Returns the number of rows adjusted.
    """
    scratch_col = ws.Columns(SCRATCH_COL_LETTER)
    scratch_col_index = scratch_col.Column
    original_width = scratch_col.ColumnWidth
    # Pick a measurement row far past any real content so the AutoFit on
    # that row only sees our scratch cell, not real cells we'd grow by
    # accident.
    measurement_row = last_row + 10
    seen: set[str] = set()
    adjusted = 0

    try:
        for cell in ws.UsedRange:
            area = cell.MergeArea
            if area.Cells.Count == 1:
                continue
            addr = area.Address
            if addr in seen:
                continue
            seen.add(addr)

            anchor = area.Cells(1, 1)
            if not anchor.WrapText:
                continue
            if _is_blank_value(anchor.Value):
                continue

            first_col = area.Column
            n_cols = area.Columns.Count
            total_width = sum(
                ws.Columns(first_col + i).ColumnWidth for i in range(n_cols)
            )

            scratch_col.ColumnWidth = total_width
            scratch_cell = ws.Cells(measurement_row, scratch_col_index)
            scratch_cell.WrapText = True
            scratch_cell.Value = anchor.Value
            # Copy font properties so the measurement reflects the real cell.
            scratch_cell.Font.Name   = anchor.Font.Name
            scratch_cell.Font.Size   = anchor.Font.Size
            scratch_cell.Font.Bold   = anchor.Font.Bold
            scratch_cell.Font.Italic = anchor.Font.Italic
            scratch_cell.EntireRow.AutoFit()
            needed = scratch_cell.RowHeight

            anchor_row = ws.Rows(area.Row)
            if needed > anchor_row.RowHeight:
                anchor_row.RowHeight = needed
                adjusted += 1
    finally:
        scratch_col.Clear()
        scratch_col.ColumnWidth = original_width

    return adjusted

def xlsx_to_pdf(xlsx_path: Path, out_dir: Path) -> Path:
    """Convert an .xlsx to .pdf using Excel COM automation (pywin32).

    Every sheet's print area is restricted to ``{PRINT_COL_START}:{PRINT_COL_END}``
    (default B:AN), each sheet is scaled to one page wide, and every merged
    wrap-text cell has its row height grown so long answers don't get
    visually clipped in the PDF. Windows + Excel only.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "pywin32 (win32com) is required for the xlsx→pdf step. "
            "Install with `pip install pywin32` on Windows with Excel."
        ) from e

    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = (out_dir / f"{xlsx_path.stem}.pdf").resolve()
    src_path = xlsx_path.resolve()

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        # DispatchEx → dedicated Excel instance, won't share state with a user
        # session and is safe to Quit() unconditionally.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(src_path), ReadOnly=True, UpdateLinks=0)

        # Snapshot last row per sheet BEFORE any mutation so the scratch-column
        # writes inside _autofit_merged_rows can't inflate it.
        sheet_last_rows: dict[str, int] = {
            ws.Name: _sheet_last_row(ws) for ws in wb.Worksheets
        }

        for ws in wb.Worksheets:
            last_row = sheet_last_rows[ws.Name]
            n_adjusted = _autofit_merged_rows(ws, last_row=last_row)
            if n_adjusted:
                print(f"    auto-expanded {n_adjusted} merged row(s) on '{ws.Name}'")

            ws.PageSetup.PrintArea = (
                f"${PRINT_COL_START}$1:${PRINT_COL_END}${last_row}"
            )
            # Fit-to-width=1 so the B:AN band always lands on one page wide;
            # height is left free so long sheets paginate naturally.
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False

        wb.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path))

        if not pdf_path.exists():
            raise RuntimeError(
                f"Excel did not produce {pdf_path} (ExportAsFixedFormat succeeded "
                f"but the file is missing)."
            )
        return pdf_path
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        finally:
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()

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
            "question_id":   q.question_id.strip(),
            "question":      q.question.strip(),
            "answer":        q.answer.strip(),
            "source_pages":  list(page_numbers),
            "answer_source": "llm",
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

# ─── xlsx reconcile (truncation safety net) ────────────────────────────────
def _normalise(s: str) -> str:
    return " ".join(s.split()).casefold()

def _collect_xlsx_answers(xlsx_path: Path) -> dict[str, str]:
    """qid → full answer text, by scanning the source xlsx with openpyxl.

    Uses the same definition of "answer cell" as parse.py: a grey/yellow
    fill in the B:AO content band, belonging to the most recent Q-id seen.
    Multiple answer cells for the same Q-id are joined with newlines.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    answers: dict[str, str] = {}

    for ws in wb.worksheets:
        merged = build_merged_lookup(ws)
        current_qid: str | None = None
        seen_anchors: set = set()

        for row in range(1, ws.max_row + 1):
            for col in range(CONTENT_START_COL, CONTENT_END_COL + 1):
                anchor = merged.get((row, col), (row, col))
                if anchor in seen_anchors:
                    continue
                seen_anchors.add(anchor)

                cell = ws.cell(anchor[0], anchor[1])
                val = effective_value(ws, row, col, merged)
                text = "" if val is None else str(val).strip()

                m = Q_PATTERN.match(text) if text else None
                if m:
                    current_qid = m.group(0)
                    continue
                if current_qid is None:
                    continue
                if not text:
                    continue
                if classify_fill(cell) in ANSWER_FILLS:
                    existing = answers.get(current_qid, "")
                    answers[current_qid] = (
                        existing + "\n" + text if existing else text
                    )
    return answers

def reconcile_with_xlsx(
    items: list[dict], xlsx_path: Path
) -> tuple[list[dict], int]:
    """Replace LLM answers that look truncated with the full text from the
    source xlsx. Modifies items in place and also returns them. The second
    return value is the number of answers reconciled."""
    answers = _collect_xlsx_answers(xlsx_path)
    n_reconciled = 0

    for item in items:
        item.setdefault("answer_source", "llm")
        truth = answers.get(item["question_id"])
        if not truth:
            continue

        llm_ans = item["answer"]
        if not llm_ans:
            item["answer"] = truth
            item["answer_source"] = "xlsx_reconciled"
            n_reconciled += 1
            continue

        truth_n = _normalise(truth)
        llm_n   = _normalise(llm_ans)
        # Truncation signature: the LLM's answer is a prefix of the xlsx
        # cell's full text, and the xlsx has materially more content.
        if truth_n.startswith(llm_n) and len(truth_n) > len(llm_n) + 5:
            item["answer"] = truth
            item["answer_source"] = "xlsx_reconciled"
            n_reconciled += 1

    return items, n_reconciled

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
    xlsx_source: Path | None = None
    if src.suffix.lower() == ".pdf":
        pdf_path = src
    elif src.suffix.lower() in {".xlsx", ".xlsm"}:
        print(f"  → converting {src.name} to PDF via Excel COM")
        pdf_path = xlsx_to_pdf(src, out_dir)
        xlsx_source = src
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

    if xlsx_source is not None:
        try:
            merged, n_rec = reconcile_with_xlsx(merged, xlsx_source)
            if n_rec:
                print(f"    reconciled {n_rec} answer(s) from xlsx source text")
        except Exception as e:
            print(f"    ! xlsx reconcile failed: {e}", file=sys.stderr)

    return ParseResult(file_name=src.stem, pdf_path=pdf_path, items=merged)

# ─── Writers ────────────────────────────────────────────────────────────────
XLSX_FIELDS = [
    "file_name", "question_id", "question", "answer",
    "answer_source", "source_pages",
]

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
                it.get("answer_source", "llm"),
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
