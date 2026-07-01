"""Opt-in LLM slot review: flag peer-disagreement and intrinsic-risk answer
outliers the deterministic vote in analyze_drift misses. Pure heuristics +
one bounded LLM call per suspicious slot. See
docs/superpowers/specs/2026-07-01-llm-answer-check-design.md.
"""
from __future__ import annotations

import json
import re
import sys
import statistics

_NA = {"", "n/a", "na", "n.a.", "nil", "none", "-", "not applicable"}
_NEG = re.compile(r"^\s*(no|not|never|none)\b", re.I)
_HEDGE = re.compile(r"\b(maybe|we believe|to our knowledge|not sure|possibly|might|unsure|tbc|tbd)\b", re.I)


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def risk_flags(slots, rules, canon) -> dict:
    """Per-answer deterministic risk labels. Key (slot_id, file_name, response)."""
    out: dict = {}
    for sid, recs in slots.items():
        lengths = [len(r["response"]) for r in recs]
        med = statistics.median(lengths) if lengths else 0
        qtext = canon.get(sid, "")
        for r in recs:
            resp = r["response"]
            labels = []
            if _norm(resp) in _NA:
                labels.append("na")
            if _NEG.match(resp):
                labels.append("negation")
            if _HEDGE.search(resp):
                labels.append("hedge")
            # short only when the slot's typical answer is substantial (avoid all-short slots)
            if med >= 20 and len(resp) < 0.4 * med:
                labels.append("short")
            for rule in rules:
                if rule.get("q_re") and not rule["q_re"].search(qtext) and not rule["q_re"].search(r["anchor_text"]):
                    continue
                if rule["a_re"].search(resp):
                    labels.append(rule["id"])
            if labels:
                out[(sid, r["file_name"], resp)] = labels
    return out


LLM_COLS = ["slot_id", "canonical_question", "file_name", "response",
            "llm_is_outlier", "category", "severity", "rationale",
            "det_votes", "det_methods", "risk_flags"]

_SEV_ORDER = {"low": 0, "med": 1, "high": 2}
_SEV_SYNONYMS = {"critical": "high", "crit": "high", "medium": "med", "moderate": "med", "minor": "low"}
_CATEGORIES = {"peer", "evasive", "gap", "contradiction", "none"}


def _norm_verdict(cat, sev):
    """Normalize LLM-returned category/severity to their enums so downstream
    filtering never sees an off-enum value. Off-enum severity fails safe to
    "high" (surface rather than silently drop); off-enum category falls back
    to "none"."""
    sev = _SEV_SYNONYMS.get(str(sev).strip().lower(), str(sev).strip().lower())
    if sev not in _SEV_ORDER:
        sev = "high"
    cat = str(cat).strip().lower()
    if cat not in _CATEGORIES:
        cat = "none"
    return cat, sev

_PROMPT = """You audit KYD questionnaire answers. For ONE question you are given every firm's free-text answer.
Flag an answer as an outlier if EITHER:
  - it diverges in substance from the other answers (peer-disagreement), OR
  - it is intrinsically concerning regardless of peers: evasive, refuses, admits a compliance gap,
    or contradicts another answer here.
Do NOT flag mere paraphrases that mean the same thing as the majority.

Question: {question}

Answers (JSON): {answers}

Return ONLY valid JSON, no prose:
{{"answers":[{{"file":<file>,"is_outlier":<bool>,"category":"peer|evasive|gap|contradiction|none","severity":"high|med|low","rationale":<one short line>}}]}}"""


def suspicious_slots(det_outliers, rmap) -> set:
    """A slot is suspicious if the vote flagged >=1 answer OR a risk heuristic did."""
    ids = {o["slot_id"] for o in det_outliers}
    ids |= {key[0] for key in rmap}
    return ids


def llm_review(slots, susp_ids, canon, chat, llm_cfg) -> list:
    """One LLM call per suspicious slot; parse per-answer verdicts. A slot whose
    call or JSON fails is skipped with a warning (best-effort, no crash)."""
    cap = llm_cfg.get("max_answers_per_slot", 40)
    out = []
    for sid in sorted(susp_ids):
        recs = slots.get(sid, [])
        if not recs:
            continue
        # ponytail: cap slot size to bound tokens; first `cap` answers. Add sampling only if a real slot exceeds it.
        sample = recs[:cap]
        payload = [{"file": r["file_name"], "answer": r["response"]} for r in sample]
        prompt = _PROMPT.format(question=canon.get(sid, ""), answers=json.dumps(payload, ensure_ascii=False))
        try:
            content = chat.invoke(prompt).content
            data = json.loads(content)
            by_file = {r["file_name"]: r["response"] for r in sample}
            for a in data["answers"]:
                fn = a["file"]
                if fn not in by_file:
                    continue
                cat, sev = _norm_verdict(a.get("category", "none"), a.get("severity", "low"))
                out.append({
                    "slot_id": sid, "file_name": fn, "response": by_file[fn],
                    "is_outlier": bool(a.get("is_outlier")),
                    "category": cat,
                    "severity": sev,
                    "rationale": a.get("rationale", ""),
                })
        except Exception as exc:  # noqa: BLE001 - best-effort per slot
            print(f"WARN: llm-review slot {sid} skipped ({exc})", file=sys.stderr)
    return out


def build_llm_review_df(verdicts, det_outliers, rmap, canon, llm_cfg):
    import pandas as pd
    floor = _SEV_ORDER.get(llm_cfg.get("min_severity", "low"), 0)
    det_by = {(o["slot_id"], o["file_name"]): o for o in det_outliers}
    rows = []
    for v in verdicts:
        if not v["is_outlier"] or _SEV_ORDER.get(v["severity"], 0) < floor:
            continue
        d = det_by.get((v["slot_id"], v["file_name"]), {})
        labels = rmap.get((v["slot_id"], v["file_name"], v["response"]), [])
        rows.append({
            "slot_id": v["slot_id"],
            "canonical_question": canon.get(v["slot_id"], ""),
            "file_name": v["file_name"],
            "response": v["response"],
            "llm_is_outlier": True,
            "category": v["category"],
            "severity": v["severity"],
            "rationale": v["rationale"],
            "det_votes": d.get("votes", 0),
            "det_methods": d.get("methods_fired", ""),
            "risk_flags": ",".join(labels),
        })
    return pd.DataFrame(rows, columns=LLM_COLS)


def make_chat(cfg, llm_cfg):
    """Azure chat client for the review pass (temperature 0). Not unit-tested."""
    import analyze_kyd_json as akj
    from langchain_openai import AzureChatOpenAI
    az = cfg["azure"]
    return AzureChatOpenAI(
        azure_endpoint=az["endpoint"],
        azure_deployment=llm_cfg.get("deployment", az["deployment"]),
        api_version=az["api_version"],
        api_key=akj.azure_api_key(az),
        temperature=0,
        max_tokens=2048,
    )
