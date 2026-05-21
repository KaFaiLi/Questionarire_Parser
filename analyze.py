"""
Cross-questionnaire answer comparator
=====================================
Reads every parsed-questionnaire *.xlsx in an input directory (output of
parse.py), fuzzy-groups equivalent sub-questions across files into canonical
groups, and writes a single wide xlsx for side-by-side comparison.

Output layout (one row per (Questionnaire, Sheet)):
    Questionnaire | Sheet
        | Actual_Question Q1 | Q1 Answer | Q1 Match Percentage
        | Actual_Question Q2 | Q2 Answer | Q2 Match Percentage
        | ...

The Q-columns are canonical question groups discovered by greedy fuzzy
clustering using fuzzywuzzy's token_set_ratio.

Usage:
    python analyze.py                          # input=./output, threshold=60
    python analyze.py <input_dir>
    python analyze.py <input_dir> -o <out.xlsx> --threshold 70

Requires: pandas, openpyxl, fuzzywuzzy
"""

from __future__ import annotations
import argparse
import re
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", message="Using slow pure-python SequenceMatcher")
from fuzzywuzzy import fuzz  # noqa: E402


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ""


def discover_inputs(input_dir: Path) -> list[Path]:
    skip_suffixes = ("_analysis", "_comparison")
    out: list[Path] = []
    for p in sorted(input_dir.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        if any(p.stem.endswith(s) for s in skip_suffixes):
            continue
        out.append(p)
    if not out:
        raise FileNotFoundError(
            f"No parsed *.xlsx found in {input_dir.resolve()} "
            f"(excluding *_analysis.xlsx and *_comparison.xlsx)."
        )
    return out


def load_all(paths: list[Path]) -> pd.DataFrame:
    frames = []
    for p in paths:
        df = pd.read_excel(p, sheet_name=0)
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)
    df = df[df["prompt"].apply(lambda v: not is_blank(v))].copy()
    df["sub_idx"] = pd.to_numeric(df["sub_idx"], errors="coerce").fillna(0).astype(int)
    return df


def normalize_for_match(text: str) -> str:
    """Lowercase, collapse whitespace, keep just the first line of the prompt
    (the hint text often follows on a newline and would inflate similarity)."""
    if not text:
        return ""
    first_line = text.split("\n", 1)[0]
    return re.sub(r"\s+", " ", first_line).strip().lower()


def split_match_keys(option_label, prompt) -> tuple[str, str]:
    """Return (option_label_key, prompt_key). Empty string means absent."""
    label = "" if is_blank(option_label) else normalize_for_match(str(option_label))
    pkey = normalize_for_match(str(prompt))
    return label, pkey


def composite_ratio(rec_label: str, rec_prompt: str,
                    seed_label: str, seed_prompt: str) -> int:
    """Combined match score. Option_label is a hard discriminator: two records
    with option_labels can only cluster if their labels are normalized-equal
    (semantically opposite branches like "Settling" vs "Non-settling" have
    high character-similarity but must stay separate)."""
    if rec_label and seed_label:
        if rec_label != seed_label:
            return 0
    elif rec_label or seed_label:
        return 0
    return fuzz.token_set_ratio(rec_prompt, seed_prompt)


def display_question(option_label, prompt) -> str:
    """The text shown in 'Actual_Question Q{n}' column."""
    prompt_first = str(prompt).split("\n", 1)[0].strip()
    if not is_blank(option_label):
        return f"{str(option_label).strip()} - {prompt_first}"
    return prompt_first


def merge_answer(side_answer, answer) -> str:
    s = "" if is_blank(side_answer) else str(side_answer).strip()
    a = "" if is_blank(answer)      else str(answer).strip()
    if s and a:
        return f"{s} - {a}"
    return s or a


def cluster_questions(df: pd.DataFrame, threshold: int
                      ) -> tuple[list[int], list[tuple[str, str]]]:
    """Greedy clustering. Returns (group_id per row, list of seed (label, prompt)
    keys). Two records cluster together only if their option_label keys are
    similar AND their prompt keys are similar (composite_ratio)."""
    df = df.reset_index(drop=True)
    group_ids: list[int] = [-1] * len(df)
    seeds: list[tuple[str, str]] = []

    for i, row in df.iterrows():
        label_key, prompt_key = split_match_keys(row.get("option_label"),
                                                 row.get("prompt", ""))
        best_g, best_score = -1, -1
        for g, (sl, sp) in enumerate(seeds):
            score = composite_ratio(label_key, prompt_key, sl, sp)
            if score > best_score:
                best_g, best_score = g, score
        if best_g >= 0 and best_score >= threshold:
            group_ids[i] = best_g
        else:
            seeds.append((label_key, prompt_key))
            group_ids[i] = len(seeds) - 1
    return group_ids, seeds


def pivot_to_wide(df: pd.DataFrame, group_ids: list[int],
                  seeds: list[tuple[str, str]]) -> pd.DataFrame:
    df = df.reset_index(drop=True).copy()
    df["_group"] = group_ids

    matches: list[int] = []
    for i in range(len(df)):
        label_key, prompt_key = split_match_keys(df.at[i, "option_label"],
                                                 df.at[i, "prompt"])
        sl, sp = seeds[group_ids[i]]
        matches.append(composite_ratio(label_key, prompt_key, sl, sp))
    df["_match"] = matches
    df["_display_q"] = [display_question(df.at[i, "option_label"], df.at[i, "prompt"])
                        for i in range(len(df))]
    df["_merged_answer"] = [merge_answer(df.at[i, "side_answer"], df.at[i, "answer"])
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
                        f"  - {fname} / {sheet} has {len(hits)} sub-questions "
                        f"in group Q{g+1} (seed={seeds[g][1]!r}); keeping first."
                    )
                hit = hits.iloc[0]
                row[f"Actual_Question Q{g+1}"] = hit["_display_q"]
                row[f"Q{g+1} Answer"] = hit["_merged_answer"]
                row[f"Q{g+1} Match Percentage"] = int(hit["_match"])
        rows.append(row)

    if duplicates:
        print("Warning: multiple sub-questions from one file joined the same group:",
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


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default="./output",
                    help="Directory containing parsed *.xlsx files (default: ./output).")
    ap.add_argument("-o", "--output", default=None,
                    help="Output xlsx path (default: <input>/questionnaire_comparison.xlsx).")
    ap.add_argument("--threshold", type=int, default=80,
                    help="Fuzzy match cutoff 0-100 for joining a canonical group (default 80).")
    args = ap.parse_args()

    in_dir = Path(args.input)
    if not in_dir.is_dir():
        print(f"Input directory not found: {in_dir}", file=sys.stderr)
        sys.exit(1)

    out_path = Path(args.output) if args.output else in_dir / "questionnaire_comparison.xlsx"

    paths = discover_inputs(in_dir)
    print(f"Loading {len(paths)} parsed file(s):")
    for p in paths:
        print(f"  - {p.name}")

    df = load_all(paths)
    print(f"Total sub-questions: {len(df)}")

    group_ids, seeds = cluster_questions(df, args.threshold)
    print(f"Canonical groups discovered: {len(seeds)} (threshold={args.threshold})")

    wide = pivot_to_wide(df, group_ids, seeds)
    write_xlsx(wide, out_path)
    print(f"Wrote: {out_path.resolve()}  ({len(wide)} rows, {len(wide.columns)} cols)")


if __name__ == "__main__":
    main()
