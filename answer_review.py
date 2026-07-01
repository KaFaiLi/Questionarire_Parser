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
from collections import Counter

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
