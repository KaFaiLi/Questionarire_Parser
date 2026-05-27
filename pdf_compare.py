"""
PDF-parser cross-questionnaire comparator
==========================================
Reads the xlsx produced by ``pdf_parse.py`` (the LLM/vision questionnaire
parser) and writes a single wide xlsx for side-by-side comparison, in the
**same layout as analyze.py**:

    Questionnaire | Sheet
        | Actual_Question Q1 | Q1 Answer | Q1 Match Percentage
        | Actual_Question Q2 | Q2 Answer | Q2 Match Percentage
        | ...

One row per (file_name, sheet). The Q-columns are canonical question groups
discovered by greedy fuzzy clustering (rapidfuzz token_set_ratio), so the same
question that carries a different question_id across questionnaires still lands
in one column.

Matching is whitespace/line-break insensitive: before comparison every question
is preprocessed by joining wrapped lines and collapsing repeated whitespace into
single spaces, so two questionnaires that differ only in formatting still match.

Input
-----
A single pdf_parse.py xlsx, or a directory of them. The 'questions' sheet
(columns: file_name, sheet, question_id, question, answer, ...) is read; a
single combined file holding several file_name/sheet questionnaires also works.

Usage
-----
    python pdf_compare.py ./output
    python pdf_compare.py parsed.xlsx -o comparison.xlsx --threshold 85

Requires: pandas, openpyxl, rapidfuzz
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz

_WS = re.compile(r"\s+")


# ── Shared match preprocessing ──────────────────────────────────────────────
def collapse_ws(text) -> str:
    """Join wrapped lines and collapse repeated whitespace into single spaces,
    then trim. This is the formatting-mismatch guard: questions that differ
    only by line breaks or spacing become identical strings."""
    if text is None:
        return ""
    return _WS.sub(" ", str(text)).strip()


def normalize_for_match(text: str) -> str:
    """Lowercased, whitespace/line-break-collapsed form used for fuzzy scoring."""
    return collapse_ws(text).lower()


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ""


# ── Input discovery + loading ───────────────────────────────────────────────
def discover_inputs(path: Path) -> list[Path]:
    if path.is_file():
        return [path]
    skip = ("_analysis", "_comparison")
    out: list[Path] = []
    for p in sorted(path.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        if any(p.stem.endswith(s) for s in skip):
            continue
        out.append(p)
    if not out:
        raise FileNotFoundError(
            f"No parsed *.xlsx found in {path.resolve()} "
            f"(excluding *_analysis.xlsx / *_comparison.xlsx)."
        )
    return out


def _read_sheet(p: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(p, engine="openpyxl")
    sheet = "questions" if "questions" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet)


def load_all(paths: list[Path]) -> pd.DataFrame:
    frames = []
    for p in paths:
        df = _read_sheet(p)
        if "question" not in df.columns:
            print(f"  ! {p.name}: no 'question' column (not pdf_parse output?), skipped.",
                  file=sys.stderr)
            continue
        df = df.copy()
        if "file_name" not in df.columns:
            df["file_name"] = p.stem
        if "sheet" not in df.columns:
            df["sheet"] = ""
        if "answer" not in df.columns:
            df["answer"] = ""
        frames.append(df)
    if not frames:
        raise ValueError("No readable pdf_parse questionnaires found.")
    df = pd.concat(frames, ignore_index=True)
    df = df[df["question"].apply(lambda v: not is_blank(v))].copy()
    if df.empty:
        raise ValueError("Inputs contained no question rows.")
    df["file_name"] = df["file_name"].fillna("").astype(str)
    df["sheet"] = df["sheet"].fillna("").astype(str)
    return df


# ── Clustering (greedy, seed-based, like analyze.py) ────────────────────────
def cluster_questions(df: pd.DataFrame, threshold: int) -> tuple[list[int], list[str]]:
    """Greedy clustering on the preprocessed question text. Returns (group_id
    per row, list of seed prompt keys)."""
    df = df.reset_index(drop=True)
    group_ids: list[int] = [-1] * len(df)
    seeds: list[str] = []

    for i, row in df.iterrows():
        key = normalize_for_match(str(row["question"]))
        best_g, best_score = -1, -1.0
        for g, seed in enumerate(seeds):
            score = fuzz.token_set_ratio(key, seed)
            if score > best_score:
                best_g, best_score = g, score
        if best_g >= 0 and best_score >= threshold:
            group_ids[i] = best_g
        else:
            seeds.append(key)
            group_ids[i] = len(seeds) - 1
    return group_ids, seeds


def pivot_to_wide(df: pd.DataFrame, group_ids: list[int], seeds: list[str]) -> pd.DataFrame:
    df = df.reset_index(drop=True).copy()
    df["_group"] = group_ids
    df["_match"] = [
        int(fuzz.token_set_ratio(normalize_for_match(str(df.at[i, "question"])),
                                 seeds[group_ids[i]]))
        for i in range(len(df))
    ]
    # display the actual question with wrapped lines joined (no info dropped)
    df["_display_q"] = [collapse_ws(df.at[i, "question"]) for i in range(len(df))]
    df["_answer"] = ["" if is_blank(df.at[i, "answer"]) else str(df.at[i, "answer"]).strip()
                     for i in range(len(df))]

    n_groups = len(seeds)
    questionnaires = (
        df[["file_name", "sheet"]]
        .drop_duplicates()
        .sort_values(["file_name", "sheet"], kind="stable")
        .values.tolist()
    )

    rows = []
    duplicates: list[str] = []
    for fname, sheet in questionnaires:
        sub = df[(df["file_name"] == fname) & (df["sheet"] == sheet)]
        row: dict[str, object] = {"Questionnaire": fname, "Sheet": sheet}
        for g in range(n_groups):
            hits = sub[sub["_group"] == g]
            if hits.empty:
                row[f"Actual_Question Q{g+1}"] = ""
                row[f"Q{g+1} Answer"] = ""
                row[f"Q{g+1} Match Percentage"] = ""
            else:
                if len(hits) > 1:
                    duplicates.append(
                        f"  - {fname} / {sheet} has {len(hits)} questions in group "
                        f"Q{g+1} (seed={seeds[g]!r}); keeping first."
                    )
                hit = hits.iloc[0]
                row[f"Actual_Question Q{g+1}"] = hit["_display_q"]
                row[f"Q{g+1} Answer"] = hit["_answer"]
                row[f"Q{g+1} Match Percentage"] = int(hit["_match"])
        rows.append(row)

    if duplicates:
        print("Warning: multiple questions from one questionnaire joined the same group:",
              file=sys.stderr)
        for d in duplicates:
            print(d, file=sys.stderr)
        print("Tune --threshold higher to split them.", file=sys.stderr)

    cols = ["Questionnaire", "Sheet"]
    for g in range(n_groups):
        cols += [f"Actual_Question Q{g+1}", f"Q{g+1} Answer", f"Q{g+1} Match Percentage"]
    return pd.DataFrame(rows, columns=cols)


def write_xlsx(wide: pd.DataFrame, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        wide.to_excel(writer, sheet_name="comparison", index=False)
        ws = writer.sheets["comparison"]
        widths = {"Questionnaire": 28, "Sheet": 22}
        for idx, col in enumerate(wide.columns, start=1):
            letter = ws.cell(row=1, column=idx).column_letter
            if col in widths:
                w = widths[col]
            elif col.startswith("Actual_Question "):
                w = 48
            elif col.endswith(" Answer"):
                w = 36
            elif col.endswith(" Match Percentage"):
                w = 10
            else:
                w = 18
            ws.column_dimensions[letter].width = w
        ws.freeze_panes = "C2"


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("input", nargs="?", default="./output",
                    help="pdf_parse.py xlsx file OR directory of them (default: ./output).")
    ap.add_argument("-o", "--output", default=None,
                    help="Output xlsx path (default: questionnaire_comparison.xlsx).")
    ap.add_argument("--threshold", type=int, default=85,
                    help="Fuzzy match cutoff 0-100 for joining a canonical group (default 85).")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"Input not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    if args.output:
        out_path = Path(args.output)
    elif in_path.is_file():
        out_path = in_path.parent / f"{in_path.stem}_comparison.xlsx"
    else:
        out_path = in_path / "questionnaire_comparison.xlsx"

    paths = discover_inputs(in_path)
    print(f"Loading {len(paths)} parsed file(s):")
    for p in paths:
        print(f"  - {p.name}")

    df = load_all(paths)
    print(f"Total questions: {len(df)}")

    group_ids, seeds = cluster_questions(df, args.threshold)
    print(f"Canonical groups discovered: {len(seeds)} (threshold={args.threshold})")

    wide = pivot_to_wide(df, group_ids, seeds)
    write_xlsx(wide, out_path)
    print(f"Wrote: {out_path.resolve()}  ({len(wide)} rows, {len(wide.columns)} cols)")


if __name__ == "__main__":
    main()
