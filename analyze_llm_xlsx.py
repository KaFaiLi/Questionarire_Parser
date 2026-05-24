"""
LLM-parsed questionnaire analyser
=================================
Standalone analysis of the xlsx produced by the LLM questionnaire parsers
(``pdf_parse.py`` or ``parse.py``). It answers three questions across a set
of questionnaires:

  1. Missing-answer rates   - per questionnaire and per question
                              (e.g. "what fraction of questionnaires answered Q1").
  2. Question changes        - has the wording of the same question drifted,
                              and is the drift only formatting or a real edit?
  3. Question coverage       - which questions appear in some questionnaires
                              but are absent from others.

Three real-world hurdles are handled explicitly:

  * Question IDs are NOT trusted. "Q1" in one file and "Q3" in another may be
    the same question, so questions are matched across files by fuzzy text
    similarity, not by id.
  * The "same" question often differs slightly (whitespace, punctuation,
    capitalisation, leading numbering). Such cosmetic drift is reported as
    "formatting only" and kept separate from substantive wording changes.
  * Some questions need not be answered (conditional / "answer only if the
    previous answer is No"). These are detected heuristically and a blank
    answer on them is counted as "optional / skipped" rather than "missing".

Input
-----
Either a single parsed xlsx, or a directory of them. Both parser schemas are
auto-detected:
  * pdf_parse.py : columns file_name, sheet, question_id, question, answer, ...
  * parse.py     : columns file_name, sheet, question_id, prompt, answer,
                   option_label, side_answer, ...

Output
------
A multi-sheet xlsx report (default ``<input>_analysis.xlsx``) plus a console
summary. Sheets: overview, questionnaire_missing, question_answer_rate,
question_changes, coverage_matrix, question_presence, clusters (audit).

Usage
-----
    python analyze_llm_xlsx.py parsed.xlsx
    python analyze_llm_xlsx.py ./output -o report.xlsx --threshold 88
    python analyze_llm_xlsx.py ./output --id-by file
    python analyze_llm_xlsx.py parsed.xlsx --treat-as-blank "n/a,nil,-"

Requires: pandas, openpyxl, rapidfuzz
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz

# ── Schema column candidates ────────────────────────────────────────────────
# Each parser writes a slightly different schema; map both onto a common shape.
QUESTION_TEXT_COLS = ("question", "prompt")          # main question text
ANSWER_COL = "answer"
FILE_COL = "file_name"
SHEET_COL = "sheet"
QID_COL = "question_id"

# Heuristic markers that a question is conditional / optional to answer.
DEFAULT_CONDITIONAL_PATTERNS = [
    r"\bif\s+(?:yes|no|not|any|so|applicable|relevant|available)\b",
    r"\bif\s+the\s+(?:answer|above|previous|preceding)\b",
    r"\bif\s+applicable\b",
    r"\bwhere\s+applicable\b",
    r"\bonly\s+if\b",
    r"\b(?:answer|complete|fill|provide|specify)\b[^.?!]*\bif\b",
    r"\bskip\s+if\b",
    r"\bleave\s+blank\s+if\b",
    r"\b(?:otherwise|else)\s+(?:leave|skip|n/?a)\b",
    r"\(\s*if\s+[^)]*\)",                # "(if any)", "(if applicable)"
    r"\bplease\s+answer\b[^.?!]*\b(?:no|yes)\b",
]


# ── Blank / normalisation helpers ───────────────────────────────────────────
def is_blank(v, extra_blank: frozenset[str] = frozenset()) -> bool:
    """True for None, NaN, empty/whitespace, or a configured placeholder."""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    if s == "":
        return True
    return s.casefold() in extra_blank


_LEADING_NUM = re.compile(
    r"^\s*(?:q(?:uestion)?\s*[\-.]?\s*)?\d+[\).:\-]?\s*"  # Q1 / 1. / 1) / Q-1
    r"|^\s*[a-z][\).:]\s*",                                # a) b. (single letter)
    re.IGNORECASE,
)
_PUNCT = re.compile(r"[^\w\s]")
_WS = re.compile(r"\s+")


def norm_match(text: str) -> str:
    """Aggressive normalisation used for fuzzy clustering: drop leading
    numbering, lowercase, strip punctuation, collapse whitespace."""
    if not text:
        return ""
    t = str(text)
    prev = None
    # strip possibly-repeated leading numbering ("Q1 a) ...")
    while prev != t:
        prev = t
        t = _LEADING_NUM.sub("", t, count=1)
    t = _PUNCT.sub(" ", t)
    t = _WS.sub(" ", t).strip().casefold()
    return t


def norm_format(text: str) -> str:
    """Formatting-insensitive normalisation: two raw texts that map to the
    same value here differ only cosmetically (case / spacing / punctuation /
    leading numbering)."""
    return norm_match(text)


# ── Loading + schema normalisation ──────────────────────────────────────────
def discover_inputs(path: Path) -> list[Path]:
    """Resolve INPUT to a list of xlsx files, skipping our own/derived outputs."""
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
    """Read the 'questions' sheet if present, else the first sheet."""
    xls = pd.ExcelFile(p, engine="openpyxl")
    sheet = "questions" if "questions" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet)


def _question_text(row: pd.Series) -> str:
    """Build the question text from whichever schema this row came from.
    For parse.py, prepend the branch option_label so opposite branches of the
    same stem stay distinguishable (e.g. 'Settling' vs 'Non-settling')."""
    text = ""
    for col in QUESTION_TEXT_COLS:
        if col in row and not is_blank(row[col]):
            text = str(row[col]).strip()
            break
    label = row.get("option_label")
    if not is_blank(label):
        text = f"{str(label).strip()} - {text}" if text else str(label).strip()
    return text


def _answer_text(row: pd.Series) -> str:
    """Answer text, merging parse.py's side_answer (YES/NO dropdown) with the
    free-text answer when both exist."""
    ans = "" if ANSWER_COL not in row or is_blank(row[ANSWER_COL]) else str(row[ANSWER_COL]).strip()
    side = row.get("side_answer")
    side = "" if is_blank(side) else str(side).strip()
    if side and ans:
        return f"{side} - {ans}"
    return side or ans


def load_long(paths: list[Path], id_by: str, extra_blank: frozenset[str]) -> pd.DataFrame:
    """Load all inputs into one long frame with a common shape:
    questionnaire, file_name, sheet, question_id, question_text, answer, is_blank."""
    frames = []
    for p in paths:
        df = _read_sheet(p)
        if not any(c in df.columns for c in QUESTION_TEXT_COLS):
            print(f"  ! {p.name}: no question/prompt column, skipped.", file=sys.stderr)
            continue
        df = df.copy()
        if FILE_COL not in df.columns:
            df[FILE_COL] = p.stem
        if SHEET_COL not in df.columns:
            df[SHEET_COL] = ""
        if QID_COL not in df.columns:
            df[QID_COL] = ""
        frames.append(df)
    if not frames:
        raise ValueError("No readable parsed questionnaires found.")
    raw = pd.concat(frames, ignore_index=True)

    rows = []
    for _, r in raw.iterrows():
        qtext = _question_text(r)
        if is_blank(qtext):
            continue  # section headers / structural rows carry no question
        fname = "" if is_blank(r.get(FILE_COL)) else str(r[FILE_COL]).strip()
        sheet = "" if is_blank(r.get(SHEET_COL)) else str(r[SHEET_COL]).strip()
        if id_by == "file" or not sheet:
            questionnaire = fname
        else:
            questionnaire = f"{fname} | {sheet}" if fname else sheet
        ans = _answer_text(r)
        rows.append(
            {
                "questionnaire": questionnaire,
                "file_name": fname,
                "sheet": sheet,
                "question_id": "" if is_blank(r.get(QID_COL)) else str(r[QID_COL]).strip(),
                "question_text": qtext,
                "answer": ans,
                "is_blank": is_blank(ans, extra_blank),
            }
        )
    long = pd.DataFrame(rows)
    if long.empty:
        raise ValueError("Inputs contained no question rows.")
    return long


# ── Conditional-question detection ──────────────────────────────────────────
def build_conditional_matcher(extra_patterns: list[str]) -> re.Pattern:
    pats = DEFAULT_CONDITIONAL_PATTERNS + list(extra_patterns)
    return re.compile("|".join(f"(?:{p})" for p in pats), re.IGNORECASE)


def is_conditional(text: str, matcher: re.Pattern) -> bool:
    return bool(matcher.search(text or ""))


# ── Fuzzy clustering of equivalent questions across questionnaires ──────────
@dataclass
class Cluster:
    seed_norm: str                       # representative normalised text
    members: list[str] = field(default_factory=list)  # raw question texts


def cluster_questions(long: pd.DataFrame, threshold: int) -> dict[str, int]:
    """Greedy single-seed fuzzy clustering over *unique* question texts.

    Returns a mapping {raw_question_text -> cluster_id}. Unique texts are
    processed most-frequent-first so the dominant phrasing seeds each cluster
    and becomes its canonical representative."""
    counts = Counter(long["question_text"])
    # most common first, then longer text, then alphabet(stable + deterministic)
    ordered = sorted(counts, key=lambda t: (-counts[t], -len(t), t))

    clusters: list[Cluster] = []
    text_to_cluster: dict[str, int] = {}
    norm_cache: dict[str, str] = {}

    for raw in ordered:
        nm = norm_cache.setdefault(raw, norm_match(raw))
        best_idx, best_score = -1, -1.0
        for idx, cl in enumerate(clusters):
            score = fuzz.token_set_ratio(nm, cl.seed_norm)
            if score > best_score:
                best_idx, best_score = idx, score
        if best_idx >= 0 and best_score >= threshold:
            clusters[best_idx].members.append(raw)
            text_to_cluster[raw] = best_idx
        else:
            clusters.append(Cluster(seed_norm=nm, members=[raw]))
            text_to_cluster[raw] = len(clusters) - 1
    return text_to_cluster


def canonical_labels(long: pd.DataFrame) -> dict[int, str]:
    """Canonical display text per cluster = most frequent raw text. Ties are
    broken toward the cleanest-formatted variant (no double spaces, more
    capitalisation) so the 'official' wording wins over cosmetic noise."""
    def rank(text: str, freq: int) -> tuple:
        clean = text == " ".join(text.split())          # no doubled/edge spaces
        caps = sum(1 for c in text if c.isupper())
        return (freq, clean, caps, len(text))

    labels: dict[int, str] = {}
    for cid, grp in long.groupby("cluster"):
        c = Counter(grp["question_text"])
        labels[cid] = max(c, key=lambda t: rank(t, c[t]))
    return labels


# ── Analyses ────────────────────────────────────────────────────────────────
def per_questionnaire_missing(long: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for q, grp in long.groupby("questionnaire", sort=True):
        n = len(grp)
        n_blank = int(grp["is_blank"].sum())
        required = grp[~grp["conditional"]]
        n_req = len(required)
        n_req_blank = int(required["is_blank"].sum())
        rows.append(
            {
                "questionnaire": q,
                "file_name": grp["file_name"].iloc[0],
                "sheet": grp["sheet"].iloc[0],
                "n_questions": n,
                "n_answered": n - n_blank,
                "n_blank": n_blank,
                "answer_rate": round((n - n_blank) / n, 4) if n else 0.0,
                "n_required": n_req,
                "n_required_missing": n_req_blank,
                "required_answer_rate": round((n_req - n_req_blank) / n_req, 4) if n_req else None,
                "n_optional_blank": n_blank - n_req_blank,
            }
        )
    return pd.DataFrame(rows)


def per_question_answer_rate(long: pd.DataFrame, labels: dict[int, str],
                             n_questionnaires: int) -> pd.DataFrame:
    rows = []
    for cid, grp in long.groupby("cluster", sort=True):
        present = grp["questionnaire"].nunique()
        n_blank = int(grp["is_blank"].sum())
        n_inst = len(grp)
        ids = sorted({i for i in grp["question_id"] if i})
        rows.append(
            {
                "canonical_id": f"CQ{cid + 1}",
                "canonical_text": labels[cid],
                "is_conditional": bool(grp["conditional"].any()),
                "questionnaires_present": present,
                "questionnaires_missing": n_questionnaires - present,
                "instances": n_inst,
                "answered": n_inst - n_blank,
                "blank": n_blank,
                "answer_rate": round((n_inst - n_blank) / n_inst, 4) if n_inst else 0.0,
                "source_question_ids": ", ".join(ids),
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values(["answer_rate", "canonical_id"]).reset_index(drop=True)


def question_changes(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    """One row per (canonical question, distinct raw variant) when a cluster
    has more than one distinct raw text. Classifies the cluster's drift as
    'formatting only' (all variants share one normalised form) or
    'substantive' (wording genuinely differs)."""
    rows = []
    for cid, grp in long.groupby("cluster", sort=True):
        variants = grp["question_text"].value_counts()
        if len(variants) <= 1:
            continue
        norm_forms = {norm_format(v) for v in variants.index}
        drift = "formatting only" if len(norm_forms) == 1 else "substantive"
        for variant, _ in variants.items():
            users = sorted(grp.loc[grp["question_text"] == variant, "questionnaire"].unique())
            rows.append(
                {
                    "canonical_id": f"CQ{cid + 1}",
                    "Question": labels[cid],          # most repeated actual question
                    "drift_type": drift,
                    "n_variants": len(variants),
                    "variant_text": variant,
                    "used_by_count": len(users),
                    "used_by": "; ".join(users),
                }
            )
    return pd.DataFrame(
        rows,
        columns=["canonical_id", "Question", "drift_type", "n_variants",
                 "variant_text", "used_by_count", "used_by"],
    )


def coverage_matrix(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    """canonical question x questionnaire status grid.
    Cell = answered / blank / optional-blank / '' (absent)."""
    questionnaires = sorted(long["questionnaire"].unique())
    rows = []
    for cid, grp in long.groupby("cluster", sort=True):
        row = {"canonical_id": f"CQ{cid + 1}", "canonical_text": labels[cid]}
        for q in questionnaires:
            sub = grp[grp["questionnaire"] == q]
            if sub.empty:
                row[q] = ""  # absent
            else:
                rec = sub.iloc[0]
                if not rec["is_blank"]:
                    row[q] = "answered"
                elif rec["conditional"]:
                    row[q] = "blank (optional)"
                else:
                    row[q] = "blank (missing)"
        row["present_in"] = int(grp["questionnaire"].nunique())
        rows.append(row)
    cols = ["canonical_id", "canonical_text", *questionnaires, "present_in"]
    return pd.DataFrame(rows, columns=cols)


def question_presence(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    """Canonical questions that are NOT present in every questionnaire."""
    all_q = sorted(long["questionnaire"].unique())
    all_set = set(all_q)
    rows = []
    for cid, grp in long.groupby("cluster", sort=True):
        present = sorted(grp["questionnaire"].unique())
        if len(present) == len(all_q):
            continue  # present everywhere -> not an inconsistency
        missing = sorted(all_set - set(present))
        rows.append(
            {
                "canonical_id": f"CQ{cid + 1}",
                "canonical_text": labels[cid],
                "present_count": len(present),
                "missing_count": len(missing),
                "present_in": "; ".join(present),
                "missing_from": "; ".join(missing),
            }
        )
    return pd.DataFrame(
        rows,
        columns=["canonical_id", "canonical_text", "present_count",
                 "missing_count", "present_in", "missing_from"],
    ).sort_values(["missing_count", "canonical_id"]).reset_index(drop=True)


def clusters_audit(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    """Row-level audit so the fuzzy grouping can be eyeballed/tuned."""
    out = long.copy()
    out["canonical_id"] = out["cluster"].map(lambda c: f"CQ{c + 1}")
    out["canonical_text"] = out["cluster"].map(labels)
    out["match_to_canonical"] = [
        round(fuzz.token_set_ratio(norm_match(t), norm_match(labels[c])), 1)
        for t, c in zip(out["question_text"], out["cluster"])
    ]
    cols = ["canonical_id", "canonical_text", "match_to_canonical", "questionnaire",
            "file_name", "sheet", "question_id", "question_text", "answer",
            "is_blank", "conditional"]
    return out[cols].sort_values(["canonical_id", "questionnaire"]).reset_index(drop=True)


def build_overview(long: pd.DataFrame, qmiss: pd.DataFrame, qrate: pd.DataFrame,
                   changes: pd.DataFrame, presence: pd.DataFrame, threshold: int) -> pd.DataFrame:
    n_q = long["questionnaire"].nunique()
    n_inst = len(long)
    n_blank = int(long["is_blank"].sum())
    req = long[~long["conditional"]]
    n_req = len(req)
    n_req_blank = int(req["is_blank"].sum())
    fmt_only = changes[changes["drift_type"] == "formatting only"]["canonical_id"].nunique()
    substantive = changes[changes["drift_type"] == "substantive"]["canonical_id"].nunique()

    stats = [
        ("questionnaires", n_q),
        ("canonical questions", len(qrate)),
        ("question instances (rows)", n_inst),
        ("answered instances", n_inst - n_blank),
        ("blank instances", n_blank),
        ("overall answer rate", round((n_inst - n_blank) / n_inst, 4) if n_inst else 0.0),
        ("required answer rate (excl. conditional)",
         round((n_req - n_req_blank) / n_req, 4) if n_req else None),
        ("conditional questions (instances)", int(long["conditional"].sum())),
        ("required answers missing", n_req_blank),
        ("questions with changed wording (substantive)", substantive),
        ("questions with formatting-only drift", fmt_only),
        ("questions not present in every questionnaire", len(presence)),
        ("fuzzy match threshold", threshold),
        ("cell legend", "answered / blank (missing) / blank (optional) / '' = absent"),
    ]
    return pd.DataFrame(stats, columns=["metric", "value"])


# ── Excel writing ────────────────────────────────────────────────────────────
def _autosize(ws, df: pd.DataFrame, wide_cols: set[str] = frozenset(),
              max_w: int = 70) -> None:
    for idx, col in enumerate(df.columns, start=1):
        letter = ws.cell(row=1, column=idx).column_letter
        if str(col) in wide_cols or "text" in str(col).lower():
            ws.column_dimensions[letter].width = min(max_w, 60)
            continue
        sample = df[col].astype(str).head(200)
        width = max([len(str(col))] + [len(s) for s in sample]) + 2
        ws.column_dimensions[letter].width = min(max_w, max(10, width))


def write_report(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            _autosize(writer.sheets[safe], df,
                      wide_cols={"canonical_text", "question_text", "variant_text",
                                 "Question", "used_by", "present_in", "missing_from",
                                 "source_question_ids", "answer"})
            writer.sheets[safe].freeze_panes = "A2"


# ── CLI ───────────────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("input", help="Parsed questionnaire xlsx file OR a directory of them.")
    ap.add_argument("-o", "--output", default=None,
                    help="Report xlsx path (default: <input>_analysis.xlsx).")
    ap.add_argument("--threshold", type=int, default=85,
                    help="Fuzzy match cutoff 0-100 for grouping equivalent questions (default 85).")
    ap.add_argument("--id-by", choices=["file", "file+sheet"], default="file+sheet",
                    help="What counts as one questionnaire (default: file+sheet).")
    ap.add_argument("--conditional-regex", action="append", default=[],
                    help="Extra regex marking a question as conditional/optional (repeatable).")
    ap.add_argument("--treat-as-blank", default="",
                    help="Comma-separated placeholder answers to count as blank (e.g. 'n/a,nil,-').")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"Input not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    extra_blank = frozenset(
        s.strip().casefold() for s in args.treat_as_blank.split(",") if s.strip()
    )
    paths = discover_inputs(in_path)
    print(f"Loading {len(paths)} parsed file(s):")
    for p in paths:
        print(f"  - {p.name}")

    long = load_long(paths, id_by=args.id_by, extra_blank=extra_blank)
    matcher = build_conditional_matcher(args.conditional_regex)
    long["conditional"] = long["question_text"].map(lambda t: is_conditional(t, matcher))

    text_to_cluster = cluster_questions(long, args.threshold)
    long["cluster"] = long["question_text"].map(text_to_cluster)
    labels = canonical_labels(long)

    n_q = long["questionnaire"].nunique()
    print(f"Questionnaires: {n_q}   question instances: {len(long)}   "
          f"canonical questions: {len(labels)}   (threshold={args.threshold})")

    qmiss = per_questionnaire_missing(long)
    qrate = per_question_answer_rate(long, labels, n_q)
    changes = question_changes(long, labels)
    coverage = coverage_matrix(long, labels)
    presence = question_presence(long, labels)
    audit = clusters_audit(long, labels)
    overview = build_overview(long, qmiss, qrate, changes, presence, args.threshold)

    out_path = (
        Path(args.output)
        if args.output
        else (in_path.parent if in_path.is_file() else in_path)
        / f"{in_path.stem if in_path.is_file() else in_path.name}_analysis.xlsx"
    )
    write_report(
        out_path,
        {
            "overview": overview,
            "questionnaire_missing": qmiss,
            "question_answer_rate": qrate,
            "question_changes": changes,
            "coverage_matrix": coverage,
            "question_presence": presence,
            "clusters": audit,
        },
    )

    # ── Console digest ────────────────────────────────────────────────────
    print("\n=== Missing answers (per questionnaire) ===")
    for _, r in qmiss.iterrows():
        req_rate = "n/a" if r["required_answer_rate"] is None else f"{r['required_answer_rate']:.0%}"
        print(f"  {r['questionnaire']}: {r['n_answered']}/{r['n_questions']} answered "
              f"({r['answer_rate']:.0%}); required missing={r['n_required_missing']} "
              f"(required rate {req_rate})")

    low = qrate[qrate["answer_rate"] < 1.0].head(10)
    if not low.empty:
        print("\n=== Lowest question answer rates ===")
        for _, r in low.iterrows():
            cond = " [conditional]" if r["is_conditional"] else ""
            print(f"  {r['canonical_id']} ({r['answer_rate']:.0%}, "
                  f"{r['answered']}/{r['instances']}){cond}: {r['canonical_text'][:70]}")

    if not changes.empty:
        subs = changes[changes["drift_type"] == "substantive"]["canonical_id"].nunique()
        fmt = changes[changes["drift_type"] == "formatting only"]["canonical_id"].nunique()
        print(f"\n=== Question changes: {subs} substantive, {fmt} formatting-only ===")
        for cid in changes[changes["drift_type"] == "substantive"]["canonical_id"].unique()[:8]:
            print(f"  {cid} variants:")
            for _, r in changes[changes["canonical_id"] == cid].iterrows():
                print(f"     ({r['used_by_count']}x) {r['variant_text'][:80]}")

    if not presence.empty:
        print(f"\n=== Coverage gaps: {len(presence)} question(s) not in every questionnaire ===")
        for _, r in presence.head(10).iterrows():
            print(f"  {r['canonical_id']} in {r['present_count']}/{n_q}; "
                  f"missing from: {r['missing_from'][:60]} -- {r['canonical_text'][:50]}")

    print(f"\n[OK] Wrote report: {out_path.resolve()}")


if __name__ == "__main__":
    main()
