"""Parse KYD questionnaire PDFs into the final JSON format with an LLM (Azure).

Sends the whole PDF (all pages rendered to images) to an Azure OpenAI chat model
via LangChain and extracts the nested questions[] -> sub_questions[] structure.

    python parse_questionnaires_llm.py --model nano  Demo/generated_questionnaires/11_questionnaire__*.pdf
    python parse_questionnaires_llm.py --model full  --dir Demo/generated_questionnaires -o output/llm_parsed

Config (endpoint / key / deployment) comes from config.yaml.
  --model nano -> gpt-4.1-nano (cheap, for testing)
  --model full -> gpt-4.1     (for actual processing)
"""
from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path

import yaml
from pydantic import BaseModel, Field

DEPLOYMENTS = {"nano": "gpt-4.1-nano", "full": "gpt-4.1"}


# --- output schema (matches the required final format) --------------------
class SubQuestion(BaseModel):
    option_label: str = Field("", description="Tick-box / option label, e.g. 'Settling distributor'. Empty if none.")
    prompt: str = Field("", description="Sub-prompt or instruction line attached to this answer, e.g. 'For HK, please specify license type'. Empty if none.")
    selection: str = Field("", description="Chosen value for an option, e.g. 'YES' or 'NO'. Empty if none.")
    answer: str = Field("", description="Free-text answer provided by the firm. Empty if blank.")


class Question(BaseModel):
    question_id: str = Field(..., description="Question id exactly as printed, e.g. 'Q1'.")
    question: str = Field("", description="Question heading text, verbatim. Empty if the row has only an id.")
    sub_questions: list[SubQuestion] = Field(default_factory=list)


class Questionnaire(BaseModel):
    questions: list[Question]


SYSTEM_PROMPT = """You extract a Know-Your-Distributor (KYD) questionnaire from page images into structured JSON.

The document is laid out in columns. Read it as a form:
- A question starts at a row carrying a question id like "Q1", "Q2" (left column) next to its heading text.
- Bold ALL-CAPS lines that are NOT next to a question id are SECTION HEADERS (e.g. "DISTRIBUTION CHANNEL"). They are NOT questions — ignore them, do not emit them.
- Under each question come its sub-questions. Each sub-question may have:
  * option_label: the label of a tick-box choice (e.g. "Settling distributor", "Other").
  * prompt: an instruction/sub-prompt line (e.g. "For HK, please specify license type", "Who is the custodian?", "a) Has your firm ...?").
  * selection: a YES/NO (or similar) value shown for an option, usually in the far-left column.
  * answer: the free-text response the firm wrote.

Grouping rules:
- A prompt line and the answer directly under/next to it belong to the SAME sub-question.
- An option_label and its YES/NO selection (and any answer beside it) belong to the SAME sub-question.
- If a question has only one free-text answer and no options/prompts, emit a single sub-question with just `answer` filled.
- A grey/example/hint line that paraphrases what to put (it ends in "etc.", lists examples, or restates the question) is a PROMPT, never the answer. The firm's actual typed value on the NEXT line is the `answer`. Do NOT merge the hint into the answer.
- The far-left YES/NO column belongs to the option/prompt line on the SAME row. Always capture it as `selection`, even though it sits far from the text.
- Preserve text VERBATIM, including odd spacing (e.g. double spaces) and capitalisation (e.g. ALL CAPS). Do not fix, normalise, translate, or summarise — copy character for character.
- Leave a field as "" when it does not apply. Never invent content.
- Distractor notes such as "External reviewer note: ..." or "Internal flag: ..." are NOT part of the questionnaire — ignore them.

Worked example. For a question printed like:

    Q2   Regulatory Authority
         SFC, Monetary Authority of Singapore etc.
         MAS

emit ONE sub-question:
    {"option_label":"", "prompt":"SFC, Monetary Authority of Singapore etc.", "selection":"", "answer":"MAS"}

For a capacity question printed like (YES/NO are in the far-left column):

    Q7   In which capacity will your firm act as distributor?
    YES  Settling distributor
         Who is the custodian?
    NO   Non-settling distributor
         Who is the custodian?            Citibank
         Is this PtP set up? ...

emit these sub-questions in order:
    {"option_label":"Settling distributor","prompt":"","selection":"YES","answer":""}
    {"option_label":"","prompt":"Who is the custodian?","selection":"","answer":""}
    {"option_label":"Non-settling distributor","prompt":"","selection":"NO","answer":""}
    {"option_label":"","prompt":"Who is the custodian?","selection":"","answer":"Citibank"}
    {"option_label":"","prompt":"Is this PtP set up? ...","selection":"","answer":""}

Emit every question id you find, in document order."""


def load_config(path: str = "config.yaml") -> dict:
    return yaml.safe_load(Path(path).read_text(encoding="utf-8"))


def make_llm(model: str, cfg: dict):
    from langchain_openai import AzureChatOpenAI
    az = cfg["azure"]
    return AzureChatOpenAI(
        azure_endpoint=az["endpoint"],
        azure_deployment=DEPLOYMENTS[model],
        api_version=az["api_version"],
        api_key=az["api_key"],
        temperature=0,
        max_tokens=4096,
    )


def pdf_to_image_blocks(pdf_path: Path, scale: float = 2.5) -> list[dict]:
    """Render every PDF page to a PNG and wrap as LangChain image content blocks."""
    import io

    import pypdfium2 as pdfium
    from PIL import Image, ImageChops
    blocks = []
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for page in pdf:
            bitmap = page.render(scale=scale)
            img = bitmap.to_pil().convert("RGB")
            # trim the large white margin so the form text fills the frame (vision reads small text better)
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bbox = ImageChops.difference(img, bg).getbbox()
            if bbox:
                pad = 20
                img = img.crop((max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                                min(img.width, bbox[2] + pad), min(img.height, bbox[3] + pad)))
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode()
            blocks.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"},
            })
    finally:
        pdf.close()
    return blocks


def parse_pdf(pdf_path: Path, llm, retries: int = 2) -> dict:
    from langchain_core.messages import HumanMessage, SystemMessage
    image_blocks = pdf_to_image_blocks(pdf_path)
    content = [{"type": "text", "text": "Extract this questionnaire to JSON."}, *image_blocks]
    structured = llm.with_structured_output(Questionnaire)
    msgs = [SystemMessage(content=SYSTEM_PROMPT), HumanMessage(content=content)]
    last = None
    for _ in range(retries + 1):
        try:
            result: Questionnaire = structured.invoke(msgs)
            return result.model_dump()
        except Exception as e:  # transient Azure errors / content-filter str responses
            last = e
    raise last


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdfs", nargs="*", help="PDF files to parse")
    ap.add_argument("--dir", help="parse every *.pdf in this directory")
    ap.add_argument("--model", choices=list(DEPLOYMENTS), default="nano")
    ap.add_argument("-o", "--out", default="output/llm_parsed", help="output dir for JSON")
    ap.add_argument("--workers", type=int, default=8, help="concurrent API requests")
    args = ap.parse_args()

    cfg = load_config()
    llm = make_llm(args.model, cfg)

    pdfs = [Path(p) for p in args.pdfs]
    if args.dir:
        pdfs += sorted(Path(args.dir).glob("*.pdf"))
    if not pdfs:
        ap.error("no PDFs given (pass paths or --dir)")

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    from concurrent.futures import ThreadPoolExecutor
    from tqdm import tqdm

    def work(pdf: Path) -> tuple[Path, str | None]:
        try:
            data = parse_pdf(pdf, llm)
        except Exception as e:
            return pdf, f"{type(e).__name__}: {str(e)[:120]}"
        data["file_name"] = pdf.stem
        (out / f"{pdf.stem}.json").write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        return pdf, None

    ok = fail = 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        for pdf, err in tqdm(ex.map(work, pdfs), total=len(pdfs), desc=f"parsing ({args.model})"):
            if err:
                fail += 1
                tqdm.write(f"{pdf.name}: FAILED ({err})")
            else:
                ok += 1
    print(f"done: {ok} ok, {fail} failed -> {out}")


if __name__ == "__main__":
    main()
