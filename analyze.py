"""
Questionnaire Parsed-Output Analyzer
====================================
Reads a parsed-questionnaire xlsx (produced by parse.py) and writes an analysis
workbook with the original data plus per-grouping missingness tabs:

    raw_data            - the parsed rows, verbatim
    summary             - one-row high-level KPIs
    missing_by_file     - missing counts/% grouped by file
    missing_by_sheet    - grouped by (file, sheet)
    missing_by_section  - grouped by (file, sheet, section)

A sub-question is "missing" if EITHER `answer` OR `side_answer` is blank
(None, empty, or whitespace-only).

Usage:
    python analyze.py                          # auto-find newest *.xlsx in ./output
    python analyze.py <input.xlsx>             # explicit input
    python analyze.py <input.xlsx> -o <out>    # explicit output path

Requires: pandas, openpyxl
"""

from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd


def is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() == ""


def find_latest_parsed_xlsx(output_dir: Path) -> Path:
    candidates = [
        p for p in output_dir.glob("*.xlsx")
        if not p.stem.endswith("_analysis") and not p.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No parsed *.xlsx found in {output_dir.resolve()} "
            f"(excluding *_analysis.xlsx)."
        )
    return max(candidates, key=lambda p: p.stat().st_mtime)


def load_parsed(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df["_missing"] = df["answer"].apply(is_missing) | df["side_answer"].apply(is_missing)
    return df


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)
    missing = int(df["_missing"].sum())
    row = {
        "total_rows": total,
        "unique_files": df["file_name"].nunique(dropna=True),
        "unique_sheets": df[["file_name", "sheet"]].drop_duplicates().shape[0],
        "unique_sections": df[["file_name", "sheet", "section"]].drop_duplicates().shape[0],
        "unique_question_ids": df[["file_name", "sheet", "question_id"]].drop_duplicates().shape[0],
        "missing": missing,
        "present": total - missing,
        "missing_pct": round(100 * missing / total, 2) if total else 0.0,
    }
    return pd.DataFrame([row])


def build_missing_by_group(df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    grouped = df.groupby(group_cols, dropna=False).agg(
        rows=("_missing", "size"),
        missing=("_missing", "sum"),
    ).reset_index()
    grouped["missing"] = grouped["missing"].astype(int)
    grouped["present"] = grouped["rows"] - grouped["missing"]
    grouped["missing_pct"] = (100 * grouped["missing"] / grouped["rows"]).round(2)
    return grouped[group_cols + ["rows", "present", "missing", "missing_pct"]]


def write_report(df: pd.DataFrame, out_path: Path) -> None:
    raw = df.drop(columns=["_missing"])
    tabs = {
        "raw_data": raw,
        "summary": build_summary(df),
        "missing_by_file": build_missing_by_group(df, ["file_name"]),
        "missing_by_sheet": build_missing_by_group(df, ["file_name", "sheet"]),
        "missing_by_section": build_missing_by_group(df, ["file_name", "sheet", "section"]),
    }
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, frame in tabs.items():
            frame.to_excel(writer, sheet_name=name, index=False)


def main() -> None:
    ap = argparse.ArgumentParser(description="Analyze parsed-questionnaire xlsx output.")
    ap.add_argument("input", nargs="?",
                    help="Parsed xlsx to analyze. Defaults to newest *.xlsx in ./output.")
    ap.add_argument("-o", "--output",
                    help="Output xlsx path. Default: <input_stem>_analysis.xlsx next to input.")
    args = ap.parse_args()

    if args.input:
        in_path = Path(args.input)
        if not in_path.exists():
            print(f"Input not found: {in_path}", file=sys.stderr)
            sys.exit(1)
    else:
        in_path = find_latest_parsed_xlsx(Path("output"))
        print(f"Auto-selected newest parsed file: {in_path}")

    out_path = Path(args.output) if args.output else in_path.with_name(
        f"{in_path.stem}_analysis.xlsx"
    )

    df = load_parsed(in_path)
    write_report(df, out_path)
    print(f"Wrote analysis: {out_path.resolve()}  ({len(df)} rows)")


if __name__ == "__main__":
    main()
