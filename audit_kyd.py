"""
KYD audit — question & answer drift checks
==========================================
An audit-framed layer over ``analyze_kyd_json``. It reuses that module's
offline, deterministic engine (rapidfuzz; population-consensus canonical
wording) and turns the drift signals into a single **findings register** an
auditor can sign off, plus two checks the base analyzer does not surface
cleanly: missing/extra questions and configurable red-flag answer rules.

Drift is measured against **population consensus** — the canonical wording /
answer is the most common variant across the submitted questionnaires; whoever
deviates is flagged. No approved master template is required.

Checks (each emits rows into the findings register)
---------------------------------------------------
  QDRIFT-SUB  substantive question-wording drift (real edit vs consensus)   [medium]
  QDRIFT-FMT  formatting-only question drift (spacing / caps only)          [low]
  QMISSING    a consensus/core question is absent from a firm's form        [medium]
  QEXTRA      a question appears in only <= --extra-max firm(s)             [low]
  ADRIFT-OUT  a minority answer (e.g. 9x YES / 1x NO, or divergent text)    [from confidence]
  REDFLAG     an answer matches a configured regulatory red-flag rule       [from rule]

Input: a directory of (or one) KYD questionnaire JSON in the nested final
format (questions[] -> sub_questions[] with option_label / prompt / selection
/ answer), one file per distributor firm — same format ``analyze_kyd_json``
reads.

Output: ``audit_findings.xlsx`` with sheets
  summary          counts by severity / check / questionnaire.
  findings         the consolidated register (sorted by severity).
  question_drift   every wording variant per canonical question.
  coverage         per canonical question: presence, drift, answer rate.
  answer_outliers  minority answers with consensus %.
  redflags         red-flag rule hits with the rule that fired.

Usage
-----
    python audit_kyd.py Demo/kyd_examples/
    python audit_kyd.py Demo/kyd_examples/ -o output/audit_findings.xlsx
    python audit_kyd.py path/to/firms/ --core-frac 0.6 --extra-max 1
    python audit_kyd.py path/to/firms/ --rules my_rules.yaml
    python audit_kyd.py path/to/firms/ --use-embeddings   # optional, needs config.yaml

Requires: pandas, openpyxl, rapidfuzz, pyyaml  (all already project deps).
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

import analyze_kyd_json as akj

# ── Severity model ─────────────────────────────────────────────────────────────
SEV_RANK = {"high": 0, "medium": 1, "low": 2, "info": 3}
_CONFIDENCE_TO_SEV = {"high": "high", "medium": "medium", "low": "low"}

FINDING_COLS = [
    "finding_id", "check_id", "check", "severity", "questionnaire",
    "canonical_id", "canonical_question", "source_question_id",
    "observation", "evidence",
]

# Red-flag rules embedded as a fallback so the tool runs with zero config even
# if audit_rules.yaml is missing. Keep in sync with audit_rules.yaml.
_DEFAULT_RULES = [
    {"id": "RF-AML", "description": "No AML / KYC programme in place", "severity": "high",
     "question_pattern": r"\b(aml|anti[- ]?money|money\s+laundering|kyc|know\s+your\s+customer|cdd|customer\s+due\s+diligence)\b",
     "answer_pattern": r"^\s*(no|none|n/?a|nil|not\s+in\s+place|we\s+do\s+not|no\s+(policy|programme|program))\b"},
    {"id": "RF-SANCTIONS", "description": "Sanctions / PEP exposure identified", "severity": "high",
     "question_pattern": r"\b(sanction|sanctioned|ofac|embargo|politically\s+exposed|\bpep\b)\b",
     "answer_pattern": r"\b(yes|exposed|exposure|present|identified|some)\b"},
    {"id": "RF-LICENCE", "description": "Regulatory licence / authorisation missing, expired or pending", "severity": "high",
     "question_pattern": r"\b(licen[cs]e|licens|authoris|authoriz|regulat(?:ed|or|ory)|registration|registered\s+with)\b",
     "answer_pattern": r"\b(no|none|not\s+(licen|authoris|registered)|without|expired|revoked|suspended|pending|lapsed|in\s+progress)\b"},
    {"id": "RF-ADVERSE-MEDIA", "description": "Adverse media / litigation / enforcement action disclosed", "severity": "high",
     "question_pattern": r"\b(adverse\s+media|litigation|investigation|enforcement|regulatory\s+action|disciplinary|sanctioned\s+by|breach|fine[ds]?)\b",
     "answer_pattern": r"\b(yes|ongoing|pending|present|disclosed|several|multiple)\b"},
    {"id": "RF-BENEFICIAL-OWNER", "description": "Beneficial / ultimate ownership not disclosed", "severity": "high",
     "question_pattern": r"\b(beneficial\s+owner|ultimate\s+owner|ownership\s+structure|\bubo\b)\b",
     "answer_pattern": r"\b(no|undisclosed|not\s+disclosed|unable|unwilling|decline|declined|unknown|withheld)\b"},
    {"id": "RF-CONTROL-NO", "description": "Negative attestation on a policy / control / compliance question", "severity": "medium",
     "question_pattern": r"\b(policy|procedure|control|compliance|programme|program|framework|attest|confirm|do\s+you|does\s+your)\b",
     "answer_pattern": r"^\s*no\b"},
    {"id": "RF-AUDITED-FINANCIALS", "description": "Financial statements not audited / unavailable", "severity": "medium",
     "question_pattern": r"\b(audited|financial\s+statement|annual\s+report|accounts|financials)\b",
     "answer_pattern": r"\b(no|none|not\s+audited|un-?audited|n/?a|unavailable)\b"},
]


# ── Rule loading / engine ──────────────────────────────────────────────────────
def load_rules(path: str | None) -> list[dict]:
    """Load red-flag rules from YAML; fall back to audit_rules.yaml next to this
    script, then to the embedded defaults. Each rule's patterns are compiled."""
    import yaml
    raw = None
    candidates = [path] if path else []
    candidates.append(str(Path(__file__).with_name("audit_rules.yaml")))
    for cand in candidates:
        if cand and Path(cand).exists():
            raw = yaml.safe_load(Path(cand).read_text(encoding="utf-8"))
            break
    if not raw:
        raw = _DEFAULT_RULES
    rules = []
    for r in raw:
        if not r.get("answer_pattern"):
            continue
        rules.append({
            "id": r["id"],
            "description": r.get("description", r["id"]),
            "severity": r.get("severity", "medium"),
            "q_re": re.compile(r["question_pattern"], re.IGNORECASE) if r.get("question_pattern") else None,
            "a_re": re.compile(r["answer_pattern"], re.IGNORECASE),
        })
    return rules


def apply_redflag_rules(long: pd.DataFrame, rules: list[dict]) -> pd.DataFrame:
    """Evaluate every rule against each answered sub-question. One row per
    (questionnaire, rule, canonical question) hit."""
    cols = ["rule_id", "description", "severity", "questionnaire", "canonical_id",
            "canonical_question", "source_question_id", "raw_question_text", "answer"]
    seen, rows = set(), []
    answered = long[~long["is_blank"]]
    for r in answered.itertuples():
        q_can, q_raw, ans = r.canonical_question, r.question_text, r.answer
        for rule in rules:
            if rule["q_re"] and not (rule["q_re"].search(q_can or "") or rule["q_re"].search(q_raw or "")):
                continue
            if not rule["a_re"].search(ans or ""):
                continue
            key = (r.questionnaire, rule["id"], r.canonical_id)
            if key in seen:
                continue
            seen.add(key)
            rows.append(dict(rule_id=rule["id"], description=rule["description"],
                             severity=rule["severity"], questionnaire=r.questionnaire,
                             canonical_id=r.canonical_id, canonical_question=q_can,
                             source_question_id=r.question_id, raw_question_text=q_raw,
                             answer=ans))
    return pd.DataFrame(rows, columns=cols)


# ── Finding builders (signal -> uniform register rows) ──────────────────────────
def _qid_lookup(long: pd.DataFrame) -> dict[tuple[str, str], str]:
    """(questionnaire, raw question text) -> a source question_id, for evidence."""
    lk: dict[tuple[str, str], str] = {}
    for r in long.itertuples():
        lk.setdefault((r.questionnaire, r.question_text), r.question_id)
    return lk


def findings_question_drift(changes: pd.DataFrame, labels: dict[int, str],
                            long: pd.DataFrame) -> list[dict]:
    """Per non-canonical wording variant, one finding per firm using it.
    Classify per-variant: same text after normalisation as canonical ->
    formatting-only (low); otherwise substantive (medium)."""
    if changes.empty:
        return []
    qid = _qid_lookup(long)
    out = []
    for r in changes.itertuples():
        canon = r.canonical_question
        if r.variant_text == canon:
            continue  # the canonical wording itself is not a finding
        fmt_only = akj._norm(r.variant_text) == akj._norm(canon)
        check_id = "QDRIFT-FMT" if fmt_only else "QDRIFT-SUB"
        check = ("Question wording drift (formatting only)" if fmt_only
                 else "Question wording drift (substantive)")
        severity = "low" if fmt_only else "medium"
        for firm in [f.strip() for f in str(r.used_by).split(";") if f.strip()]:
            out.append(dict(
                check_id=check_id, check=check, severity=severity, questionnaire=firm,
                canonical_id=r.canonical_id, canonical_question=canon,
                source_question_id=qid.get((firm, r.variant_text), ""),
                observation=f"Wording differs from consensus ({r.variant_match_pct}% match).",
                evidence=f"used: {r.variant_text!r}  |  consensus: {canon!r}"))
    return out


def findings_coverage(fact: pd.DataFrame, dimq: pd.DataFrame, long: pd.DataFrame,
                      n_q: int, core_frac: float, extra_max: int) -> list[dict]:
    """QMISSING — a core question (present in >= core_frac of firms) absent from a
    firm. QEXTRA — a question present in only <= extra_max firm(s)."""
    out = []
    core_ids = set(dimq.loc[dimq["n_present"] >= max(2, math.ceil(core_frac * n_q)),
                            "canonical_id"])
    opt = dict(zip(dimq["canonical_id"], dimq["is_optional"]))
    # QMISSING: core question, firm where it is structurally absent
    miss = fact[(fact["canonical_id"].isin(core_ids)) & (fact["status"] == akj.STATUS_ABSENT)]
    for r in miss.itertuples():
        is_opt = bool(opt.get(r.canonical_id, False))
        out.append(dict(
            check_id="QMISSING", check="Missing question (absent from form)",
            severity="low" if is_opt else "medium", questionnaire=r.questionnaire,
            canonical_id=r.canonical_id, canonical_question=r.canonical_question,
            source_question_id="",
            observation=f"Consensus question absent from this questionnaire"
                        f"{' (question is optional)' if is_opt else ''}.",
            evidence=f"present in others; missing here"))
    # QEXTRA: question carried by only a handful of firms
    extra = dimq[dimq["n_present"] <= extra_max]
    for r in extra.itertuples():
        firms = sorted(long.loc[long["canonical_id"] == r.canonical_id, "questionnaire"].unique())
        for firm in firms:
            out.append(dict(
                check_id="QEXTRA", check="Extra / questionnaire-specific question",
                severity="low", questionnaire=firm, canonical_id=r.canonical_id,
                canonical_question=r.canonical_question, source_question_id="",
                observation=f"Question appears in only {r.n_present}/{n_q} questionnaire(s).",
                evidence=f"carried by: {', '.join(firms)}"))
    return out


def findings_answer_outliers(outliers: pd.DataFrame) -> list[dict]:
    if outliers.empty:
        return []
    out = []
    for r in outliers.itertuples():
        out.append(dict(
            check_id="ADRIFT-OUT", check="Answer outlier (minority response)",
            severity=_CONFIDENCE_TO_SEV.get(r.confidence, "low"),
            questionnaire=r.questionnaire, canonical_id=r.canonical_id,
            canonical_question=r.canonical_question, source_question_id="",
            observation=f"Answer is a {r.confidence}-confidence minority "
                        f"({r.consensus_pct}% consensus, {r.agree_with_majority} agree).",
            evidence=f"answer: {r.outlier_answer!r}  |  majority: {r.majority_answer!r}"))
    return out


def findings_redflags(redflags: pd.DataFrame) -> list[dict]:
    if redflags.empty:
        return []
    out = []
    for r in redflags.itertuples():
        out.append(dict(
            check_id="REDFLAG", check=f"Red flag: {r.description}",
            severity=r.severity, questionnaire=r.questionnaire,
            canonical_id=r.canonical_id, canonical_question=r.canonical_question,
            source_question_id=r.source_question_id,
            observation=f"Answer triggered rule {r.rule_id}.",
            evidence=f"answer: {r.answer!r}"))
    return out


def build_findings(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows, columns=[c for c in FINDING_COLS if c != "finding_id"])
    if df.empty:
        return pd.DataFrame(columns=FINDING_COLS)
    df["_rank"] = df["severity"].map(lambda s: SEV_RANK.get(s, 9))
    df = (df.sort_values(["_rank", "check_id", "questionnaire", "canonical_id"])
            .drop(columns="_rank").reset_index(drop=True))
    df.insert(0, "finding_id", [f"AUD-{i:04d}" for i in range(1, len(df) + 1)])
    return df[FINDING_COLS]


def build_summary(findings: pd.DataFrame, n_q: int) -> pd.DataFrame:
    rows = [dict(group="total", key="findings", n_findings=len(findings)),
            dict(group="total", key="questionnaires", n_findings=n_q)]
    if not findings.empty:
        for sev in sorted(findings["severity"].unique(), key=lambda s: SEV_RANK.get(s, 9)):
            rows.append(dict(group="severity", key=sev,
                             n_findings=int((findings["severity"] == sev).sum())))
        for chk in sorted(findings["check_id"].unique()):
            rows.append(dict(group="check", key=chk,
                             n_findings=int((findings["check_id"] == chk).sum())))
        for q, n in findings["questionnaire"].value_counts().items():
            rows.append(dict(group="questionnaire", key=q, n_findings=int(n)))
    return pd.DataFrame(rows, columns=["group", "key", "n_findings"])


# ── Optional embeddings clustering (best-effort) ───────────────────────────────
def _embedding_clusters(long: pd.DataFrame, cluster_threshold: float | None) -> dict[str, int]:
    """Replace fuzzy clustering with analyze_drift's semantic pass. Needs
    config.yaml + Azure keys. Returns raw question_text -> cluster id."""
    import analyze_drift as ad
    cfg = ad.load_config()
    ct = cluster_threshold if cluster_threshold is not None else cfg["drift"]["cluster_threshold"]
    texts = sorted(long["question_text"].unique())
    units = [{"file_name": "", "question_id": "", "sub_idx": None,
              "level": "question", "text": t} for t in texts]
    embeddings = ad.make_embeddings(cfg)
    vecs = ad.embed_texts(texts, embeddings, workers=8,
                          cache_path=Path("output/.audit_vec_cache.json"))
    assign, _ = ad.assign_clusters(units, vecs, ct)
    return {t: assign[("question", t)] for t in texts}


# ── Pipeline ───────────────────────────────────────────────────────────────────
def run_audit(paths, *, threshold=85, answer_threshold=80, core_frac=0.6,
              extra_max=1, treat_as_blank="", conditional_regex=None,
              rules_path=None, use_embeddings=False,
              embed_cluster_threshold=None) -> dict[str, pd.DataFrame]:
    """End-to-end: load -> cluster/canonical -> reuse base tables -> findings.
    Returns the sheet dict ready for ``akj.write_report``."""
    extra_blank = frozenset(s.strip().casefold()
                            for s in treat_as_blank.split(",") if s.strip())
    long = akj.load_long(paths, extra_blank)
    matcher = akj._build_conditional(conditional_regex or [])
    long["conditional"] = long["question_text"].map(lambda t: akj._is_conditional(t, matcher))

    if use_embeddings:
        try:
            mapping = _embedding_clusters(long, embed_cluster_threshold)
            print("clustering: embeddings (analyze_drift)")
        except Exception as exc:  # noqa: BLE001 — fall back to offline, never crash the audit
            print(f"WARN: embeddings unavailable ({exc}); falling back to fuzzy.", file=sys.stderr)
            mapping = akj._cluster_questions(long, threshold)
    else:
        mapping = akj._cluster_questions(long, threshold)
    long["cluster"] = long["question_text"].map(mapping)

    labels = akj._canonical_labels(long)
    long = akj._enrich(long, labels)
    optional = akj._compute_optional(long, 0.6, 2)
    n_q = long["questionnaire"].nunique()

    # reused base tables
    fact = akj.build_fact_responses(long, optional)
    dimq = akj.build_dim_question(long, optional, n_q)
    changes = akj.build_question_changes(long, labels)
    outliers = akj.build_answer_outliers(long, labels, answer_threshold,
                                         min_answers=4, outlier_frac=0.25,
                                         majority_frac=0.6)
    redflags = apply_redflag_rules(long, load_rules(rules_path))

    # findings register
    rows = []
    rows += findings_question_drift(changes, labels, long)
    rows += findings_coverage(fact, dimq, long, n_q, core_frac, extra_max)
    rows += findings_answer_outliers(outliers)
    rows += findings_redflags(redflags)
    findings = build_findings(rows)
    summary = build_summary(findings, n_q)

    coverage = dimq[["canonical_id", "canonical_question", "is_optional", "drift_type",
                     "n_present", "n_missing_from", "specific_to", "answer_rate"]]
    return {
        "summary": summary,
        "findings": findings,
        "question_drift": changes,
        "coverage": coverage,
        "answer_outliers": outliers,
        "redflags": redflags,
    }


def _digest(sheets: dict[str, pd.DataFrame]) -> None:
    findings = sheets["findings"]
    print(f"\n=== Audit findings: {len(findings)} ===")
    if findings.empty:
        print("  none.")
        return
    for sev in sorted(findings["severity"].unique(), key=lambda s: SEV_RANK.get(s, 9)):
        sub = findings[findings["severity"] == sev]
        print(f"\n  [{sev.upper()}] {len(sub)}")
        for r in sub.itertuples():
            print(f"    {r.finding_id} {r.check_id:11s} {r.questionnaire:24.24s} "
                  f"{r.canonical_id:>5s}: {r.observation}")


# ── CLI ─────────────────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="A KYD .json file OR a directory of them.")
    ap.add_argument("-o", "--output", default=None,
                    help="Report xlsx path (default: <input>/audit_findings.xlsx).")
    ap.add_argument("--threshold", type=int, default=85,
                    help="Fuzzy cutoff 0-100 for grouping equivalent questions (default 85).")
    ap.add_argument("--answer-threshold", type=int, default=80,
                    help="Fuzzy cutoff 0-100 for grouping equivalent answers (default 80).")
    ap.add_argument("--core-frac", type=float, default=0.6,
                    help="A question is 'core' if present in >= this fraction of firms (default 0.6).")
    ap.add_argument("--extra-max", type=int, default=1,
                    help="Flag a question as extra/specific if present in <= this many firms (default 1).")
    ap.add_argument("--rules", default=None,
                    help="Red-flag rules YAML (default: audit_rules.yaml beside this script).")
    ap.add_argument("--treat-as-blank", default="",
                    help="Comma-separated placeholders to treat as blank (e.g. 'n/a,nil,-').")
    ap.add_argument("--conditional-regex", action="append", default=[],
                    help="Extra regex marking a question as optional/conditional (repeatable).")
    ap.add_argument("--use-embeddings", action="store_true",
                    help="Use analyze_drift semantic clustering (needs config.yaml); offline fallback otherwise.")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"Input not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    paths = akj._discover(in_path)
    print(f"Loading {len(paths)} questionnaire(s):")
    for p in paths:
        print(f"  - {p.name}")

    sheets = run_audit(
        paths, threshold=args.threshold, answer_threshold=args.answer_threshold,
        core_frac=args.core_frac, extra_max=args.extra_max,
        treat_as_blank=args.treat_as_blank, conditional_regex=args.conditional_regex,
        rules_path=args.rules, use_embeddings=args.use_embeddings)

    out_path = (Path(args.output) if args.output
                else (in_path.parent if in_path.is_file() else in_path) / "audit_findings.xlsx")
    akj.write_report(out_path, sheets)

    _digest(sheets)
    print(f"\n[OK] Wrote {out_path.resolve()}")


if __name__ == "__main__":
    main()
