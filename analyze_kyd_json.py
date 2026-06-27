"""
KYD questionnaire reviewer
===========================
Reads a directory of "Know Your Distributor" questionnaires in the nested
final format and produces a multi-sheet Excel report that answers four goals:

  1. Empty answers     — which sub-questions are unanswered, split into
                         "missing" (a real gap) vs "optional/conditional"
                         (blank-because-not-applicable).
  2. Specific questions— questions present in only one questionnaire; broader
                         coverage gaps (present in N of M files).
  3. Question changes  — has wording drifted, and is it formatting-only or a
                         substantive edit?
  4. Answer outliers   — within the same question, answers that are a clear
                         minority (e.g. 9 firms say YES, 1 says NO).

Input format (one .json file per distributor firm)::

    {
      "file_name": "acme_hk",
      "questions": [
        {
          "question_id": "Q1",
          "question": "Firm Full Legal Name and Address",
          "sub_questions": [
            {"option_label": "", "prompt": "", "selection": "", "answer": "..."}
          ]
        }
      ]
    }

Question-text derivation (per sub-question):
  option_label alone   — when present; identifies the branch without inflating
                         similarity across branches that share the same prompt.
  prompt               — the actual question when no branch label is present.
  parent question      — fallback when neither option_label nor prompt exist.

Answer derivation:
  selection | answer   — merged; a YES/NO dropdown with no free text still
                         counts as answered.

Limitation handled: question_id is NOT trusted across files. Questions are
matched by fuzzy text similarity (token_sort_ratio, which is stricter than
token_set_ratio for clean, precise question strings).

Output sheets
-------------
  overview             — headline KPIs.
  comparison           — wide grid: one row per questionnaire, three columns
                         per canonical question (text / answer / match %).
  fact_responses       — tidy long grid, one row per (canonical question ×
                         questionnaire) including absent combinations.
  dim_question         — per canonical question: answer rate, drift, coverage.
  dim_questionnaire    — per questionnaire: answered / missing counts.
  question_changes     — one row per (question, distinct wording variant).
  answer_outliers      — one row per flagged minority answer.

Usage
-----
    python analyze_kyd_json.py Demo/kyd_examples/
    python analyze_kyd_json.py Demo/kyd_examples/ -o output/kyd_analysis.xlsx
    python analyze_kyd_json.py Demo/kyd_examples/ --threshold 88
    python analyze_kyd_json.py Demo/kyd_examples/ --treat-as-blank "n/a,nil,-"

Multiple folders (each analysed independently — one report per folder)::

    python analyze_kyd_json.py 2024Q4/ 2025Q1/ 2025Q2/ --out-dir output/

Requires: pandas, openpyxl, rapidfuzz
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz


# ── Long-path handling ─────────────────────────────────────────────────────────
# On Windows a path over ~260 chars raises "file not found" unless it is absolute,
# backslash-separated and prefixed with the extended-length marker \\?\ (UNC paths
# use \\?\UNC\). These helpers apply that transparently and are no-ops elsewhere,
# so every read/write below is safe for deeply-nested inputs and outputs.
def longpath(p) -> str:
    """Return a filesystem string for *p* that is safe past the OS max-path limit."""
    s = os.fspath(p)
    if os.name != "nt" or s.startswith("\\\\?\\"):
        return s
    ap = os.path.abspath(s)
    if ap.startswith("\\\\"):              # UNC: \\server\share -> \\?\UNC\server\share
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


def path_exists(p) -> bool:
    return os.path.exists(longpath(p))


def read_text(p, encoding: str = "utf-8") -> str:
    with open(longpath(p), "r", encoding=encoding) as fh:
        return fh.read()


def write_text(p, data: str, encoding: str = "utf-8") -> None:
    ensure_parent(p)
    with open(longpath(p), "w", encoding=encoding) as fh:
        fh.write(data)


def ensure_parent(p) -> None:
    parent = Path(p).parent
    if str(parent):
        os.makedirs(longpath(parent), exist_ok=True)


def iter_json(path: Path) -> list[Path]:
    """List *.json under a directory, tolerant of long directory paths."""
    out = []
    with os.scandir(longpath(path)) as it:
        for e in it:
            if e.is_file() and e.name.lower().endswith(".json") and not e.name.startswith("~$"):
                out.append(path / e.name)
    return sorted(out)

# ── Status values ─────────────────────────────────────────────────────────────
STATUS_ANSWERED = "answered"
STATUS_MISSING  = "blank (missing)"
STATUS_OPTIONAL = "blank (optional)"
STATUS_ABSENT   = "absent"

# Patterns that mark a question as conditional / optional to answer.
_CONDITIONAL_PATTERNS = [
    r"\bif\s+(?:yes|no|not|any|so|applicable|relevant|available|non-?settling)\b",
    r"\bif\s+the\s+(?:answer|above|previous|preceding)\b",
    r"\bif\s+applicable\b",
    r"\bwhere\s+applicable\b",
    r"\bonly\s+if\b",
    r"\bskip\s+if\b",
    r"\bleave\s+blank\s+if\b",
    r"\(\s*(?:only\s+)?if\s+[^)]*\)",
]


# ── Text helpers ───────────────────────────────────────────────────────────────
def is_blank(v, extra: frozenset[str] = frozenset()) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.casefold() in extra


_LEADING_NUM = re.compile(
    r"^\s*(?:q(?:uestion)?\s*[\-.]?\s*)?\d+[\).:\-]?\s*"
    r"|^\s*[a-z][\).:]\s*",
    re.IGNORECASE,
)
_PUNCT = re.compile(r"[^\w\s]")
_WS    = re.compile(r"\s+")


def _norm(text: str) -> str:
    """Normalise for fuzzy matching: strip leading numbering, lowercase,
    remove punctuation, collapse whitespace."""
    t = str(text or "")
    prev = None
    while prev != t:
        prev = t
        t = _LEADING_NUM.sub("", t, count=1)
    return _WS.sub(" ", _PUNCT.sub(" ", t)).strip().casefold()


def _norm_fmt(text: str) -> str:
    """Same as _norm — two texts equal here differ only cosmetically."""
    return _norm(text)


def _build_conditional(extra: list[str]) -> re.Pattern:
    pats = _CONDITIONAL_PATTERNS + list(extra)
    return re.compile("|".join(f"(?:{p})" for p in pats), re.IGNORECASE)


# ── JSON flattening ────────────────────────────────────────────────────────────
def _question_text(parent: str, option_label: str, prompt: str) -> str:
    """Derive the matchable question string for one sub-question.

    Priority: option_label (alone) > prompt > parent question.
    option_label alone avoids inflating similarity between branches that share
    the same follow-up prompt (e.g. 'Settling' and 'Non-settling' both ask
    'Who is the custodian?')."""
    opt = (option_label or "").strip()
    prt = (prompt or "").strip()
    return opt or prt or (parent or "").strip()


def _answer_text(selection: str, answer: str) -> str:
    """Merge a dropdown selection with a free-text answer into one string."""
    sel = (selection or "").strip()
    ans = (answer or "").strip()
    if sel and ans:
        return f"{sel} | {ans}"
    return sel or ans


def _discover(path: Path) -> list[Path]:
    if path.is_file() or (not path.is_dir() and path_exists(path) and path.suffix):
        return [path]
    files = iter_json(path)
    if not files:
        raise FileNotFoundError(f"No *.json questionnaires found in {path}")
    return files


def load_long(paths: list[Path], extra_blank: frozenset[str]) -> pd.DataFrame:
    """Flatten every sub-question of every file into one row.

    Columns: file_name, question_id, question_text, answer, is_blank, src_order.
    """
    rows = []
    for p in paths:
        data = json.loads(read_text(p))
        fname = str(data.get("file_name") or p.stem).strip() or p.stem
        order = 0
        for q in data.get("questions", []):
            qid    = str(q.get("question_id") or "").strip()
            parent = str(q.get("question") or "").strip()
            for sub in (q.get("sub_questions") or [{}]):
                opt    = sub.get("option_label", "") or ""
                prompt = sub.get("prompt", "")       or ""
                sel    = sub.get("selection", "")    or ""
                ans    = sub.get("answer", "")       or ""
                qtext  = _question_text(parent, opt, prompt)
                if not qtext:
                    continue
                answer = _answer_text(sel, ans)
                rows.append({
                    "file_name":     fname,
                    "question_id":   qid,
                    "question_text": qtext,
                    "answer":        answer,
                    "is_blank":      is_blank(answer, extra_blank),
                    "src_order":     order,
                })
                order += 1

    if not rows:
        raise ValueError("No question rows found in the input file(s).")
    long = pd.DataFrame(rows)
    # "questionnaire" is the identity key used throughout the analysis.
    long["questionnaire"] = long["file_name"]
    return long


# ── Fuzzy question clustering ──────────────────────────────────────────────────
@dataclass
class _Cluster:
    seed_norm: str
    members: list[str] = field(default_factory=list)


def _cluster_questions(long: pd.DataFrame, threshold: int) -> dict[str, int]:
    """Greedy single-seed fuzzy clustering with token_sort_ratio.

    token_sort (not token_set) is deliberate: token_set scores a subset of
    tokens as a near-perfect match, which wrongly merges distinct sub-questions
    that share a long common stem. token_sort still groups genuine wording
    drift (e.g. 'Address' vs 'Registered Address')."""
    counts  = Counter(long["question_text"])
    ordered = sorted(counts, key=lambda t: (-counts[t], -len(t), t))
    clusters: list[_Cluster] = []
    mapping:  dict[str, int] = {}
    for raw in ordered:
        nm = _norm(raw)
        best_idx, best_score = -1, -1.0
        for idx, cl in enumerate(clusters):
            s = fuzz.token_sort_ratio(nm, cl.seed_norm)
            if s > best_score:
                best_idx, best_score = idx, s
        if best_idx >= 0 and best_score >= threshold:
            clusters[best_idx].members.append(raw)
            mapping[raw] = best_idx
        else:
            clusters.append(_Cluster(seed_norm=nm, members=[raw]))
            mapping[raw] = len(clusters) - 1
    return mapping


def _canonical_labels(long: pd.DataFrame) -> dict[int, str]:
    def rank(text: str, freq: int) -> tuple:
        clean = text == " ".join(text.split())
        caps  = sum(1 for c in text if c.isupper())
        return (freq, clean, caps, len(text))
    labels: dict[int, str] = {}
    for cid, grp in long.groupby("cluster"):
        c = Counter(grp["question_text"])
        labels[cid] = max(c, key=lambda t: rank(t, c[t]))
    return labels


def _order_clusters(long: pd.DataFrame) -> dict[int, int]:
    mean_pos = long.groupby("cluster")["src_order"].mean()
    ordered  = sorted(mean_pos.index, key=lambda c: (mean_pos[c], c))
    return {cid: rank for rank, cid in enumerate(ordered)}


def _enrich(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    rank = _order_clusters(long)
    long = long.copy()
    long["rank"]               = long["cluster"].map(rank)
    long["canonical_id"]       = long["rank"].map(lambda r: f"CQ{r + 1}")
    long["canonical_question"] = long["cluster"].map(labels)
    norm_label = {cid: _norm(t) for cid, t in labels.items()}
    long["match_pct"] = [
        round(fuzz.token_sort_ratio(_norm(t), norm_label[c]))
        for t, c in zip(long["question_text"], long["cluster"])
    ]
    return long


def _is_conditional(text: str, matcher: re.Pattern) -> bool:
    return bool(matcher.search(text or ""))


def _compute_optional(long: pd.DataFrame, blank_frac: float,
                      min_present: int) -> dict[int, bool]:
    optional: dict[int, bool] = {}
    for cid, grp in long.groupby("cluster"):
        text_cond    = bool(grp["conditional"].any())
        per_q_blank  = grp.groupby("questionnaire")["is_blank"].min()
        present      = len(per_q_blank)
        blank_rate   = int(per_q_blank.sum()) / present if present else 0.0
        data_cond    = present >= min_present and blank_rate >= blank_frac
        optional[cid] = text_cond or data_cond
    return optional


def _drift_type(grp: pd.DataFrame) -> tuple[str, int]:
    variants = grp["question_text"].unique()
    if len(variants) <= 1:
        return "none", len(variants)
    norm_forms = {_norm_fmt(v) for v in variants}
    return ("formatting only" if len(norm_forms) == 1 else "substantive"), len(variants)


# ── Report tables ──────────────────────────────────────────────────────────────
def build_fact_responses(long: pd.DataFrame, optional: dict[int, bool]) -> pd.DataFrame:
    """One row per (canonical question × questionnaire), including absent combos."""
    questionnaires = sorted(long["questionnaire"].unique())
    clusters = sorted(long["cluster"].unique(),
                      key=lambda c: long.loc[long["cluster"] == c, "rank"].iloc[0])
    rows = []
    for cid in clusters:
        grp   = long[long["cluster"] == cid]
        cqid  = grp["canonical_id"].iloc[0]
        ctext = grp["canonical_question"].iloc[0]
        is_opt = optional[cid]
        for q in questionnaires:
            sub = grp[grp["questionnaire"] == q]
            if sub.empty:
                rows.append(dict(questionnaire=q, canonical_id=cqid,
                                 canonical_question=ctext, source_question_id="",
                                 raw_question_text="", answer="",
                                 status=STATUS_ABSENT, is_optional=is_opt,
                                 is_present=False, is_answered=False,
                                 match_pct=float("nan")))
                continue
            answered = not bool(sub["is_blank"].all())
            pool = sub[~sub["is_blank"]] if answered else sub
            rep  = pool.sort_values(["match_pct", "src_order"],
                                    ascending=[False, True]).iloc[0]
            status = (STATUS_ANSWERED if answered
                      else STATUS_OPTIONAL if is_opt else STATUS_MISSING)
            rows.append(dict(questionnaire=q, canonical_id=cqid,
                             canonical_question=ctext,
                             source_question_id=rep["question_id"],
                             raw_question_text=rep["question_text"],
                             answer=rep["answer"], status=status,
                             is_optional=is_opt, is_present=True,
                             is_answered=answered,
                             match_pct=float(rep["match_pct"])))
    return pd.DataFrame(rows)


def build_dim_question(long: pd.DataFrame, optional: dict[int, bool],
                       n_q: int) -> pd.DataFrame:
    rows = []
    for cid, grp in long.groupby("cluster"):
        per_q_blank = grp.groupby("questionnaire")["is_blank"].min()
        present     = len(per_q_blank)
        n_blank     = int(per_q_blank.sum())
        n_answered  = present - n_blank
        drift, n_var = _drift_type(grp)
        ids = sorted({i for i in grp["question_id"] if i})
        specific = sorted(per_q_blank.index) if present == 1 else []
        rows.append(dict(
            canonical_id              = grp["canonical_id"].iloc[0],
            canonical_question        = grp["canonical_question"].iloc[0],
            is_optional               = optional[cid],
            drift_type                = drift,
            n_variants                = n_var,
            n_present                 = present,
            n_missing_from            = n_q - present,
            specific_to               = ", ".join(specific),
            n_answered                = n_answered,
            n_blank                   = n_blank,
            answer_rate               = round(n_answered / present, 4) if present else 0.0,
            source_question_ids       = ", ".join(ids),
        ))
    df = pd.DataFrame(rows)
    df["_n"] = df["canonical_id"].str[2:].astype(int)
    return df.sort_values("_n").drop(columns="_n").reset_index(drop=True)


def build_dim_questionnaire(fact: pd.DataFrame) -> pd.DataFrame:
    present = fact[fact["is_present"]]
    rows = []
    for q, grp in present.groupby("questionnaire", sort=True):
        n           = len(grp)
        n_answered  = int(grp["is_answered"].sum())
        req         = grp[~grp["is_optional"]]
        n_req       = len(req)
        n_req_miss  = int((~req["is_answered"]).sum())
        rows.append(dict(
            questionnaire          = q,
            n_questions            = n,
            n_answered             = n_answered,
            n_blank                = n - n_answered,
            answer_rate            = round(n_answered / n, 4) if n else 0.0,
            n_required             = n_req,
            n_required_missing     = n_req_miss,
            required_answer_rate   = round((n_req - n_req_miss) / n_req, 4) if n_req else None,
            n_optional_blank       = (n - n_answered) - n_req_miss,
        ))
    return pd.DataFrame(rows)


def build_question_changes(long: pd.DataFrame, labels: dict[int, str]) -> pd.DataFrame:
    rows = []
    for cid, grp in long.groupby("cluster"):
        variants = grp["question_text"].value_counts()
        if len(variants) <= 1:
            continue
        drift, n_var = _drift_type(grp)
        norm_label   = _norm(labels[cid])
        for variant in variants.index:
            users = sorted(grp.loc[grp["question_text"] == variant,
                                   "questionnaire"].unique())
            rows.append(dict(
                canonical_id      = grp["canonical_id"].iloc[0],
                canonical_question= labels[cid],
                drift_type        = drift,
                n_variants        = n_var,
                variant_text      = variant,
                variant_match_pct = round(fuzz.token_sort_ratio(_norm(variant), norm_label)),
                used_by_count     = len(users),
                used_by           = "; ".join(users),
            ))
    df = pd.DataFrame(rows, columns=["canonical_id", "canonical_question",
                                     "drift_type", "n_variants", "variant_text",
                                     "variant_match_pct", "used_by_count", "used_by"])
    if df.empty:
        return df
    df["_n"] = df["canonical_id"].str[2:].astype(int)
    return (df.sort_values(["_n", "used_by_count"], ascending=[True, False])
              .drop(columns="_n").reset_index(drop=True))


def build_comparison(fact: pd.DataFrame) -> pd.DataFrame:
    qorder = (fact[["questionnaire"]].drop_duplicates()
                                     .sort_values("questionnaire"))
    canon  = (fact[["canonical_id"]].drop_duplicates()
                                    .assign(_n=lambda d: d["canonical_id"].str[2:].astype(int))
                                    .sort_values("_n")["canonical_id"].tolist())
    lookup = {(r.questionnaire, r.canonical_id): r for r in fact.itertuples()}
    rows = []
    for qr in qorder.itertuples():
        row: dict = {"Questionnaire": qr.questionnaire}
        for cid in canon:
            n   = int(cid[2:])
            rec = lookup.get((qr.questionnaire, cid))
            if rec is None or not rec.is_present:
                row[f"Actual_Question Q{n}"] = ""
                row[f"Q{n} Answer"]           = ""
                row[f"Q{n} Match %"]          = ""
            else:
                row[f"Actual_Question Q{n}"] = rec.raw_question_text
                row[f"Q{n} Answer"]           = rec.answer
                row[f"Q{n} Match %"]          = int(rec.match_pct)
        rows.append(row)
    cols = ["Questionnaire"]
    for cid in canon:
        n = int(cid[2:])
        cols += [f"Actual_Question Q{n}", f"Q{n} Answer", f"Q{n} Match %"]
    return pd.DataFrame(rows, columns=cols)


# ── Goal 4: fuzzy answer-outlier detection ─────────────────────────────────────
def _cluster_answers(values: list[str], threshold: int) -> list[int]:
    """Greedy answer clustering; uses token_set_ratio (lenient) to avoid
    flooding the outlier report with spurious free-text differences."""
    counts  = Counter(values)
    ordered = sorted(counts, key=lambda t: (-counts[t], -len(t), t))
    seeds: list[str] = []
    v2g:   dict[str, int] = {}
    for raw in ordered:
        nm = _norm(raw)
        best_idx, best_score = -1, -1.0
        for idx, seed in enumerate(seeds):
            s = fuzz.token_set_ratio(nm, seed)
            if s > best_score:
                best_idx, best_score = idx, s
        if best_idx >= 0 and best_score >= threshold:
            v2g[raw] = best_idx
        else:
            seeds.append(nm)
            v2g[raw] = len(seeds) - 1
    return [v2g[v] for v in values]


def build_answer_outliers(long: pd.DataFrame, labels: dict[int, str],
                          answer_threshold: int, min_answers: int,
                          outlier_frac: float, majority_frac: float) -> pd.DataFrame:
    rows = []
    for cid, grp in long.groupby("cluster"):
        answered = grp[~grp["is_blank"]]
        if answered.empty:
            continue
        per_q = (answered.sort_values("src_order")
                         .groupby("questionnaire", as_index=False).first())
        n = len(per_q)
        if n < min_answers:
            continue
        values = list(per_q["answer"])
        groups = _cluster_answers(values, answer_threshold)
        per_q  = per_q.assign(_grp=groups)
        sizes  = Counter(groups)
        maj_grp, maj_size = max(sizes.items(), key=lambda kv: kv[1])
        maj_frac_val = maj_size / n
        if maj_frac_val < majority_frac:
            continue
        maj_value = Counter(
            per_q.loc[per_q["_grp"] == maj_grp, "answer"]).most_common(1)[0][0]
        outlier_max = max(1, math.floor(n * outlier_frac))
        for g, size in sizes.items():
            if g == maj_grp or size > outlier_max:
                continue
            conf = ("high"   if size == 1 and maj_frac_val >= 0.8 else
                    "medium" if maj_frac_val >= 0.6 else "low")
            for _, r in per_q[per_q["_grp"] == g].iterrows():
                rows.append(dict(
                    canonical_id        = grp["canonical_id"].iloc[0],
                    canonical_question  = labels[cid],
                    questionnaire       = r["questionnaire"],
                    outlier_answer      = r["answer"],
                    majority_answer     = maj_value,
                    agree_with_majority = f"{maj_size}/{n}",
                    minority_size       = size,
                    consensus_pct       = round(maj_frac_val * 100),
                    confidence          = conf,
                ))
    df = pd.DataFrame(rows, columns=["canonical_id", "canonical_question",
                                     "questionnaire", "outlier_answer",
                                     "majority_answer", "agree_with_majority",
                                     "minority_size", "consensus_pct", "confidence"])
    if df.empty:
        return df
    df["_n"] = df["canonical_id"].str[2:].astype(int)
    return (df.sort_values(["_n", "questionnaire"])
              .drop(columns="_n").reset_index(drop=True))


def build_overview(fact: pd.DataFrame, dimq: pd.DataFrame,
                   changes: pd.DataFrame, outliers: pd.DataFrame,
                   n_q: int, threshold: int, answer_threshold: int) -> pd.DataFrame:
    present      = fact[fact["is_present"]]
    n_present    = len(present)
    n_answered   = int(present["is_answered"].sum())
    req          = present[~present["is_optional"]]
    n_req        = len(req)
    n_req_ans    = int(req["is_answered"].sum())
    substantive  = (int((changes["drift_type"] == "substantive")
                         .groupby(changes["canonical_id"]).any().sum())
                    if not changes.empty else 0)
    fmt_only     = (int((changes["drift_type"] == "formatting only")
                         .groupby(changes["canonical_id"]).any().sum())
                    if not changes.empty else 0)
    stats = [
        ("questionnaires",                              n_q),
        ("canonical questions",                         fact["canonical_id"].nunique()),
        ("present (question × questionnaire)",          n_present),
        ("answered",                                    n_answered),
        ("blank (present)",                             n_present - n_answered),
        ("overall answer rate",                         round(n_answered / n_present, 4) if n_present else 0.0),
        ("required answer rate (excl. optional)",       round(n_req_ans / n_req, 4) if n_req else None),
        ("required answers MISSING",                    n_req - n_req_ans),
        ("questions specific to one questionnaire",     int((dimq["specific_to"] != "").sum())),
        ("questions not in every questionnaire",        int((dimq["n_missing_from"] > 0).sum())),
        ("questions with substantive wording change",   substantive),
        ("questions with formatting-only drift",        fmt_only),
        ("answer outliers flagged",                     0 if outliers.empty else len(outliers)),
        ("question match threshold",                    threshold),
        ("answer match threshold",                      answer_threshold),
        ("status legend",
         f"{STATUS_ANSWERED} / {STATUS_MISSING} / {STATUS_OPTIONAL} / {STATUS_ABSENT}"),
    ]
    return pd.DataFrame(stats, columns=["metric", "value"])


# ── Excel output ───────────────────────────────────────────────────────────────
_WIDE = frozenset({"canonical_question", "raw_question_text", "variant_text",
                   "answer", "outlier_answer", "majority_answer", "used_by",
                   "source_question_ids"})


def _autosize(ws, df: pd.DataFrame, max_w: int = 60) -> None:
    for idx, col in enumerate(df.columns, start=1):
        letter = ws.cell(row=1, column=idx).column_letter
        if str(col) in _WIDE or str(col).startswith("Actual_Question ") \
                or str(col).endswith(" Answer"):
            ws.column_dimensions[letter].width = 50
        else:
            w = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 2
            ws.column_dimensions[letter].width = min(max_w, max(10, w))


def write_report(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    ensure_parent(path)
    with pd.ExcelWriter(longpath(path), engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            (df if not df.empty else pd.DataFrame({"(none)": []})).to_excel(
                writer, sheet_name=safe, index=False)
            if not df.empty:
                _autosize(writer.sheets[safe], df)
            writer.sheets[safe].freeze_panes = "A2"


# ── CLI ────────────────────────────────────────────────────────────────────────
def _resolve_out(in_path: Path, output: str | None, out_dir: str | None,
                 default_name: str) -> Path:
    """Where one input's report is written.
    --out-dir wins (named <input>_<default>); else -o (single input); else the
    report lands inside the input folder (or its parent for a single file)."""
    if out_dir:
        stem = in_path.name if in_path.is_dir() else in_path.stem
        return Path(out_dir) / f"{stem}_{default_name}"
    if output:
        return Path(output)
    base = in_path.parent if in_path.is_file() else in_path
    return base / default_name


def _plan_outputs(inputs: list[Path], output: str | None, out_dir: str | None,
                  default_name: str) -> list[tuple[Path, Path]]:
    """One distinct output path per input. Folder basenames can collide (e.g. two
    inputs both named 'ground_truth'); disambiguate with the parent name, then a
    numeric suffix, so no report silently overwrites another."""
    plan: list[tuple[Path, Path]] = []
    used: set[str] = set()
    for in_path in inputs:
        out = _resolve_out(in_path, output, out_dir, default_name)
        if str(out) in used:
            stem = in_path.name if in_path.is_dir() else in_path.stem
            parent = in_path.parent.name
            cand = out.with_name(f"{parent}_{stem}_{default_name}") if parent else out
            i = 2
            while str(cand) in used:
                cand = out.with_name(f"{stem}_{i}_{default_name}")
                i += 1
            out = cand
        used.add(str(out))
        plan.append((in_path, out))
    return plan


def analyze_one(in_path: Path, args, out_path: Path) -> None:
    """Run the full review on one input (file or folder) and write its report.

    Each input is analysed independently — the consensus baseline is computed
    per folder, so separate cohorts are never pooled."""
    extra_blank = frozenset(
        s.strip().casefold() for s in args.treat_as_blank.split(",") if s.strip())
    paths = _discover(in_path)
    print(f"\n=== {in_path} ===")
    print(f"Loading {len(paths)} questionnaire(s):")
    for p in paths:
        print(f"  - {p.name}")

    long    = load_long(paths, extra_blank)
    matcher = _build_conditional(args.conditional_regex)
    long["conditional"] = long["question_text"].map(lambda t: _is_conditional(t, matcher))

    mapping = _cluster_questions(long, args.threshold)
    long["cluster"] = long["question_text"].map(mapping)
    labels  = _canonical_labels(long)
    long    = _enrich(long, labels)
    optional = _compute_optional(long, args.optional_blank_frac,
                                 args.optional_min_present)

    n_q = long["questionnaire"].nunique()
    print(f"Questionnaires: {n_q}   sub-questions: {len(long)}   "
          f"canonical questions: {long['cluster'].nunique()}   "
          f"(threshold={args.threshold})")

    fact        = build_fact_responses(long, optional)
    dimq        = build_dim_question(long, optional, n_q)
    dimn        = build_dim_questionnaire(fact)
    changes     = build_question_changes(long, labels)
    comparison  = build_comparison(fact)
    outliers    = build_answer_outliers(long, labels, args.answer_threshold,
                                        args.outlier_min_answers, args.outlier_frac,
                                        args.outlier_majority_frac)
    overview    = build_overview(fact, dimq, changes, outliers, n_q,
                                 args.threshold, args.answer_threshold)

    write_report(out_path, {
        "overview":            overview,
        "comparison":          comparison,
        "fact_responses":      fact,
        "dim_question":        dimq,
        "dim_questionnaire":   dimn,
        "question_changes":    changes,
        "answer_outliers":     outliers,
    })

    # ── Console digest ─────────────────────────────────────────────────────
    print("\n--- Goal 1: empty answers ---")
    miss = fact[fact["is_present"] & ~fact["is_answered"] & ~fact["is_optional"]]
    opt  = fact[fact["is_present"] & ~fact["is_answered"] &  fact["is_optional"]]
    print(f"  {len(miss)} required answer(s) MISSING   "
          f"{len(opt)} optional blank(s)")
    for _, r in miss.iterrows():
        print(f"    [MISSING] {r['questionnaire']}  {r['canonical_id']}: "
              f"{r['canonical_question'][:55]}")

    print("\n--- Goal 2: questionnaire-specific / coverage gaps ---")
    spec = dimq[dimq["specific_to"] != ""]
    gaps = dimq[(dimq["n_missing_from"] > 0) & (dimq["specific_to"] == "")]
    if spec.empty and gaps.empty:
        print("  none.")
    for _, r in spec.iterrows():
        print(f"    [SPECIFIC] {r['canonical_id']} only in {r['specific_to']}: "
              f"{r['canonical_question'][:55]}")
    for _, r in gaps.iterrows():
        print(f"    [GAP]      {r['canonical_id']} in {r['n_present']}/{n_q}: "
              f"{r['canonical_question'][:55]}")

    print("\n--- Goal 3: question wording changes ---")
    if changes.empty:
        print("  none.")
    for cid in changes["canonical_id"].unique():
        sub = changes[changes["canonical_id"] == cid]
        print(f"    {cid} [{sub['drift_type'].iloc[0]}]:")
        for _, r in sub.iterrows():
            print(f"      ({r['used_by_count']}×) {r['variant_text'][:65]}")

    print("\n--- Goal 4: answer outliers ---")
    if outliers.empty:
        print("  none.")
    for _, r in outliers.iterrows():
        print(f"    [{r['confidence']}] {r['canonical_id']} {r['questionnaire']}: "
              f"'{r['outlier_answer'][:35]}' vs majority '{r['majority_answer'][:35]}' "
              f"({r['agree_with_majority']} agree)")

    print(f"[OK] Wrote {out_path.resolve()}")


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+",
                    help="One or more KYD .json files OR directories of them. "
                         "Each is analysed independently (one report per input).")
    ap.add_argument("-o", "--output", default=None,
                    help="Report xlsx path. Single input only; for many inputs use --out-dir.")
    ap.add_argument("--out-dir", default=None,
                    help="Write every report into this directory as <input>_kyd_analysis.xlsx.")
    ap.add_argument("--threshold", type=int, default=85,
                    help="Fuzzy cutoff 0-100 for grouping equivalent questions (default 85).")
    ap.add_argument("--answer-threshold", type=int, default=80,
                    help="Fuzzy cutoff 0-100 for grouping equivalent answers (default 80).")
    ap.add_argument("--treat-as-blank", default="",
                    help="Comma-separated placeholders to treat as blank (e.g. 'n/a,nil,-').")
    ap.add_argument("--conditional-regex", action="append", default=[],
                    help="Extra regex marking a question as optional/conditional (repeatable).")
    ap.add_argument("--optional-blank-frac", type=float, default=0.6,
                    help="Flag optional if blank in >= this fraction of files (default 0.6).")
    ap.add_argument("--optional-min-present", type=int, default=2,
                    help="Min files before blank-pattern rule applies (default 2).")
    ap.add_argument("--outlier-min-answers", type=int, default=4,
                    help="Min answered questionnaires before outlier check runs (default 4).")
    ap.add_argument("--outlier-frac", type=float, default=0.25,
                    help="An answer group is a minority if its share <= this (default 0.25).")
    ap.add_argument("--outlier-majority-frac", type=float, default=0.6,
                    help="Require a majority group >= this share to flag (default 0.6).")
    args = ap.parse_args()

    inputs = [Path(p) for p in args.inputs]
    missing = [p for p in inputs if not path_exists(p)]
    if missing:
        for p in missing:
            print(f"Input not found: {p}", file=sys.stderr)
        sys.exit(1)
    if args.output and len(inputs) > 1:
        ap.error("-o/--output takes a single input; use --out-dir for multiple inputs.")
    if args.out_dir:
        Path(args.out_dir).mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    for in_path, out_path in _plan_outputs(inputs, args.output, args.out_dir,
                                           "kyd_analysis.xlsx"):
        try:
            analyze_one(in_path, args, out_path)
            written.append(out_path)
        except (FileNotFoundError, ValueError) as e:
            print(f"  SKIP {in_path}: {e}", file=sys.stderr)

    print(f"\n[DONE] {len(written)}/{len(inputs)} report(s) written.")
    for w in written:
        print(f"  - {w.resolve()}")
    if not written:
        sys.exit(1)


if __name__ == "__main__":
    main()
