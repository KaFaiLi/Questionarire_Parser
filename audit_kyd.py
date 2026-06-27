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

Multiple folders (each audited independently — one report per folder)::

    python audit_kyd.py 2024Q4/ 2025Q1/ 2025Q2/ --out-dir output/

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


# ── Visual HTML dashboard ───────────────────────────────────────────────────────
SEV_COLOR = {"high": "#c62828", "medium": "#ef6c00", "low": "#2e7d32", "info": "#607d8b"}
_CHECK_LABEL = {
    "QDRIFT-SUB": "Question drift (substantive)",
    "QDRIFT-FMT": "Question drift (formatting)",
    "QMISSING": "Missing question",
    "QEXTRA": "Extra / specific question",
    "ADRIFT-OUT": "Answer outlier",
    "REDFLAG": "Red flag",
}


def write_html_report(sheets: dict[str, pd.DataFrame], path: Path, *,
                      title: str = "KYD Audit — Findings") -> None:
    """Self-contained HTML dashboard: KPI cards, a firm × check severity heatmap,
    a per-firm severity bar, and a colour-coded findings table."""
    import html as _html
    import plotly.graph_objects as go

    f = sheets["findings"]
    n_q = int(sheets["summary"].loc[sheets["summary"]["key"] == "questionnaires",
                                    "n_findings"].iloc[0]) if not sheets["summary"].empty else 0

    def card(label, value, color="#1f2937"):
        return (f'<div class="card"><div class="card-val" style="color:{color}">{value}</div>'
                f'<div class="card-lbl">{label}</div></div>')

    sev_counts = {s: int((f["severity"] == s).sum()) for s in SEV_RANK}
    n_firms_flagged = f["questionnaire"].nunique() if not f.empty else 0
    cards = "".join([
        card("Total findings", len(f)),
        card("High", sev_counts["high"], SEV_COLOR["high"]),
        card("Medium", sev_counts["medium"], SEV_COLOR["medium"]),
        card("Low", sev_counts["low"], SEV_COLOR["low"]),
        card("Questionnaires flagged", f"{n_firms_flagged}/{n_q}"),
        card("Red flags", int((f["check_id"] == "REDFLAG").sum()) if not f.empty else 0,
             SEV_COLOR["high"]),
    ])

    charts_html = ""
    if not f.empty:
        firms = sorted(f["questionnaire"].unique())
        checks = [c for c in _CHECK_LABEL if c in set(f["check_id"])]
        # heatmap: worst severity per (firm, check); z = severity rank inverted (3=high)
        z, hover = [], []
        for firm in firms:
            zr, hr = [], []
            for c in checks:
                sub = f[(f["questionnaire"] == firm) & (f["check_id"] == c)]
                if sub.empty:
                    zr.append(None); hr.append("")
                else:
                    worst = min(SEV_RANK[s] for s in sub["severity"])
                    zr.append(3 - worst)
                    hr.append(f"<b>{firm}</b><br>{_CHECK_LABEL[c]}<br>"
                              f"{len(sub)} finding(s), worst: "
                              f"{['high','medium','low','info'][worst]}")
            z.append(zr); hover.append(hr)
        heat = go.Figure(go.Heatmap(
            z=z, x=[_CHECK_LABEL[c] for c in checks], y=firms,
            customdata=hover, hovertemplate="%{customdata}<extra></extra>",
            zmin=0, zmax=3, xgap=2, ygap=2,
            colorscale=[[0.0, "#607d8b"], [0.33, "#607d8b"], [0.34, "#2e7d32"],
                        [0.66, "#2e7d32"], [0.67, "#ef6c00"], [0.83, "#ef6c00"],
                        [0.84, "#c62828"], [1.0, "#c62828"]],
            colorbar={"tickvals": [0.4, 1.3, 2.1, 2.8],
                      "ticktext": ["info", "low", "medium", "high"],
                      "title": {"text": "severity"}, "thickness": 14, "len": 0.6}))
        heat.update_layout(
            title="Findings by questionnaire × check (worst severity)",
            xaxis={"side": "top", "tickangle": -30}, yaxis={"autorange": "reversed"},
            height=120 + 30 * len(firms), margin={"l": 160, "r": 40, "t": 120, "b": 20},
            plot_bgcolor="white", font={"family": "system-ui,sans-serif", "size": 12})

        # stacked bar: count per firm by severity
        bar = go.Figure()
        for s in ["high", "medium", "low", "info"]:
            counts = [int(((f["questionnaire"] == firm) & (f["severity"] == s)).sum())
                      for firm in firms]
            if sum(counts):
                bar.add_bar(name=s, y=firms, x=counts, orientation="h",
                            marker_color=SEV_COLOR[s])
        bar.update_layout(
            title="Findings per questionnaire by severity", barmode="stack",
            yaxis={"autorange": "reversed"}, xaxis={"title": "findings"},
            height=120 + 30 * len(firms), margin={"l": 160, "r": 40, "t": 60, "b": 40},
            plot_bgcolor="white", legend={"orientation": "h", "y": 1.08},
            font={"family": "system-ui,sans-serif", "size": 12})

        # embed plotly.js inline (once) so the report is fully self-contained / offline.
        charts_html = (
            f'<div class="chart">{heat.to_html(full_html=False, include_plotlyjs=True)}</div>'
            f'<div class="chart">{bar.to_html(full_html=False, include_plotlyjs=False)}</div>')

    # findings table
    if f.empty:
        table_html = '<p class="empty">No findings — every questionnaire matches consensus.</p>'
    else:
        trs = []
        for r in f.itertuples():
            sev = r.severity
            badge = (f'<span class="badge" style="background:{SEV_COLOR.get(sev, "#777")}">'
                     f'{sev.upper()}</span>')
            trs.append(
                "<tr>"
                f'<td class="mono">{_html.escape(r.finding_id)}</td>'
                f"<td>{badge}</td>"
                f'<td class="mono">{_html.escape(r.check_id)}</td>'
                f"<td>{_html.escape(r.questionnaire)}</td>"
                f'<td class="mono">{_html.escape(str(r.canonical_id))}</td>'
                f'<td class="wrap">{_html.escape(str(r.canonical_question))}</td>'
                f'<td class="wrap">{_html.escape(str(r.observation))}</td>'
                f'<td class="wrap small">{_html.escape(str(r.evidence))}</td>'
                "</tr>")
        table_html = (
            '<table><thead><tr>'
            '<th>ID</th><th>Severity</th><th>Check</th><th>Questionnaire</th>'
            '<th>Q</th><th>Canonical question</th><th>Observation</th><th>Evidence</th>'
            "</tr></thead><tbody>" + "".join(trs) + "</tbody></table>")

    doc = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_html.escape(title)}</title>
<style>
  :root {{ --bg:#f4f5f7; --fg:#1f2937; --line:#e5e7eb; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--fg);
         font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif; }}
  header {{ background:linear-gradient(120deg,#0f2740,#1d4e6f); color:#fff;
            padding:28px 32px; }}
  header h1 {{ margin:0; font-size:22px; }}
  header p {{ margin:6px 0 0; opacity:.8; font-size:13px; }}
  .wrap-main {{ max-width:1200px; margin:0 auto; padding:24px 32px 64px; }}
  .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
            gap:14px; margin:-44px 0 28px; }}
  .card {{ background:#fff; border:1px solid var(--line); border-radius:12px;
           padding:18px 16px; box-shadow:0 1px 3px rgba(0,0,0,.06); }}
  .card-val {{ font-size:30px; font-weight:700; line-height:1; }}
  .card-lbl {{ margin-top:8px; font-size:12px; color:#6b7280; text-transform:uppercase;
               letter-spacing:.04em; }}
  .chart {{ background:#fff; border:1px solid var(--line); border-radius:12px;
            padding:8px 12px; margin-bottom:22px; overflow-x:auto; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; border:1px solid var(--line);
           border-radius:12px; overflow:hidden; font-size:13px; }}
  thead th {{ background:#0f2740; color:#fff; text-align:left; padding:10px 12px;
              font-weight:600; position:sticky; top:0; }}
  tbody td {{ padding:9px 12px; border-top:1px solid var(--line); vertical-align:top; }}
  tbody tr:nth-child(even) {{ background:#fafbfc; }}
  .badge {{ color:#fff; padding:2px 9px; border-radius:999px; font-size:11px; font-weight:700; }}
  .mono {{ font-family:ui-monospace,SFMono-Regular,Menlo,monospace; white-space:nowrap; }}
  .wrap {{ max-width:280px; }} .small {{ color:#6b7280; font-size:12px; }}
  .empty {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:40px;
            text-align:center; color:#6b7280; }}
  h2 {{ font-size:15px; margin:26px 0 12px; color:#374151; }}
</style></head><body>
<header>
  <h1>{_html.escape(title)}</h1>
  <p>Drift measured against population consensus · {len(f)} finding(s) across {n_q} questionnaire(s)</p>
</header>
<div class="wrap-main">
  <div class="cards">{cards}</div>
  {charts_html}
  <h2>Findings register</h2>
  {table_html}
</div></body></html>"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(doc, encoding="utf-8")


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
def audit_one(in_path: Path, args, out_path: Path) -> None:
    """Audit one input (file or folder) and write its xlsx + HTML dashboard.

    Each input is audited independently — the consensus baseline is computed
    per folder, so separate cohorts are never pooled."""
    paths = akj._discover(in_path)
    print(f"\n=== {in_path} ===")
    print(f"Loading {len(paths)} questionnaire(s):")
    for p in paths:
        print(f"  - {p.name}")

    sheets = run_audit(
        paths, threshold=args.threshold, answer_threshold=args.answer_threshold,
        core_frac=args.core_frac, extra_max=args.extra_max,
        treat_as_blank=args.treat_as_blank, conditional_regex=args.conditional_regex,
        rules_path=args.rules, use_embeddings=args.use_embeddings)

    akj.write_report(out_path, sheets)
    html_path = None
    if not args.no_html:
        html_path = out_path.with_suffix(".html")
        write_html_report(sheets, html_path, title=f"KYD Audit — {in_path.name or in_path.stem}")

    _digest(sheets)
    print(f"[OK] Wrote {out_path.resolve()}")
    if html_path:
        print(f"[OK] Wrote {html_path.resolve()}")


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+",
                    help="One or more KYD .json files OR directories of them. "
                         "Each is audited independently (one report per input).")
    ap.add_argument("-o", "--output", default=None,
                    help="Report xlsx path. Single input only; for many inputs use --out-dir.")
    ap.add_argument("--out-dir", default=None,
                    help="Write every report into this directory as <input>_audit_findings.xlsx.")
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
    ap.add_argument("--no-html", action="store_true",
                    help="Skip the visual HTML dashboard (xlsx only).")
    args = ap.parse_args()

    inputs = [Path(p) for p in args.inputs]
    missing = [p for p in inputs if not p.exists()]
    if missing:
        for p in missing:
            print(f"Input not found: {p}", file=sys.stderr)
        sys.exit(1)
    if args.output and len(inputs) > 1:
        ap.error("-o/--output takes a single input; use --out-dir for multiple inputs.")
    if args.out_dir:
        Path(args.out_dir).mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    for in_path, out_path in akj._plan_outputs(inputs, args.output, args.out_dir,
                                               "audit_findings.xlsx"):
        try:
            audit_one(in_path, args, out_path)
            written.append(out_path)
        except (FileNotFoundError, ValueError) as e:
            print(f"  SKIP {in_path}: {e}", file=sys.stderr)

    print(f"\n[DONE] {len(written)}/{len(inputs)} audit(s) written.")
    for w in written:
        print(f"  - {w.resolve()}")
    if not written:
        sys.exit(1)


if __name__ == "__main__":
    main()
