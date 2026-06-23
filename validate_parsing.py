"""Score LLM-parsed questionnaires against generator ground truth.

    python validate_parsing.py --pred output/llm_parsed --truth Demo/generated_questionnaires/ground_truth

Matches files by the leading index (e.g. "11_questionnaire..."), aligns
questions by question_id, and reports field-level accuracy. Reports both EXACT
and whitespace-normalised text matching (the generator plants doubled-space /
case drift on purpose, so exact is the strict bar).
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

FIELDS = ("option_label", "prompt", "selection", "answer")


def _idx(name: str) -> str:
    m = re.match(r"(\d+)", Path(name).stem)
    return m.group(1) if m else Path(name).stem


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _load(d: Path) -> dict[str, dict]:
    return {_idx(p.name): json.loads(p.read_text(encoding="utf-8")) for p in d.glob("*.json")}


def _subq_tuples(q: dict, normed: bool):
    out = []
    for s in q.get("sub_questions", []):
        vals = tuple((_norm(s.get(f, "")) if normed else (s.get(f, "") or "")) for f in FIELDS)
        out.append(vals)
    return out


def score_file(pred: dict, truth: dict) -> dict:
    tq = {q["question_id"]: q for q in truth["questions"]}
    pq = {q["question_id"]: q for q in pred.get("questions", [])}
    res = {
        "gt_questions": len(tq),
        "found_ids": len(set(tq) & set(pq)),
        "missing_ids": sorted(set(tq) - set(pq)),
        "extra_ids": sorted(set(pq) - set(tq)),
        "qtext_exact": 0, "qtext_norm": 0,
        "subq_fields_total": 0, "subq_fields_exact": 0, "subq_fields_norm": 0,
        "subq_count_match": 0,
    }
    for qid, t in tq.items():
        p = pq.get(qid)
        if not p:
            res["subq_fields_total"] += len(t.get("sub_questions", [])) * len(FIELDS)
            continue
        if (p.get("question", "") or "") == (t.get("question", "") or ""):
            res["qtext_exact"] += 1
        if _norm(p.get("question", "")) == _norm(t.get("question", "")):
            res["qtext_norm"] += 1
        ts, ps = _subq_tuples(t, False), _subq_tuples(p, False)
        tn, pn = _subq_tuples(t, True), _subq_tuples(p, True)
        if len(ts) == len(ps):
            res["subq_count_match"] += 1
        for i, tup in enumerate(ts):
            for j, f in enumerate(FIELDS):
                res["subq_fields_total"] += 1
                if i < len(ps) and ps[i][j] == tup[j]:
                    res["subq_fields_exact"] += 1
                if i < len(pn) and pn[i][j] == tn[i][j]:
                    res["subq_fields_norm"] += 1
    return res


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--pred", default="output/llm_parsed")
    ap.add_argument("--truth", default="Demo/generated_questionnaires/ground_truth")
    args = ap.parse_args()

    preds, truths = _load(Path(args.pred)), _load(Path(args.truth))
    common = sorted(set(preds) & set(truths), key=int)
    if not common:
        ap.error("no matching files between pred and truth")

    agg = {"gt_questions": 0, "found_ids": 0, "qtext_exact": 0, "qtext_norm": 0,
           "subq_fields_total": 0, "subq_fields_exact": 0, "subq_fields_norm": 0,
           "subq_count_match": 0, "files": 0, "missing": 0}
    print(f"{'file':>5} {'ids':>7} {'qtext_x':>8} {'subq#':>6} {'fld_x':>7} {'fld_n':>7}  missing")
    for k in common:
        r = score_file(preds[k], truths[k])
        agg["files"] += 1
        for f in ("gt_questions", "found_ids", "qtext_exact", "qtext_norm",
                  "subq_fields_total", "subq_fields_exact", "subq_fields_norm", "subq_count_match"):
            agg[f] += r[f]
        agg["missing"] += len(r["missing_ids"])
        fx = r["subq_fields_exact"] / r["subq_fields_total"] if r["subq_fields_total"] else 1
        fn = r["subq_fields_norm"] / r["subq_fields_total"] if r["subq_fields_total"] else 1
        print(f"{k:>5} {r['found_ids']:>3}/{r['gt_questions']:<3} "
              f"{r['qtext_exact']:>3}/{r['gt_questions']:<3} "
              f"{r['subq_count_match']:>3}/{r['gt_questions']:<2} "
              f"{fx:>6.0%} {fn:>6.0%}  {','.join(r['missing_ids']) or '-'}")

    n = agg["gt_questions"] or 1
    ft = agg["subq_fields_total"] or 1
    print("\n=== AGGREGATE ===")
    print(f"files                : {agg['files']}")
    print(f"question recall      : {agg['found_ids']}/{agg['gt_questions']}  ({agg['found_ids']/n:.1%})")
    print(f"question-text exact  : {agg['qtext_exact']/n:.1%}   norm: {agg['qtext_norm']/n:.1%}")
    print(f"sub-q count match    : {agg['subq_count_match']/n:.1%}")
    print(f"sub-q field exact    : {agg['subq_fields_exact']/ft:.1%}   norm: {agg['subq_fields_norm']/ft:.1%}")


if __name__ == "__main__":
    main()
