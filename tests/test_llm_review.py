# tests/test_llm_review.py
import json, sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_review as ar

def _rec(fn, resp, sid=0): return {"file_name": fn, "question_id": "Q1",
                                   "anchor_level": "prompt", "anchor_text": "x",
                                   "response": resp, "slot_id": sid}

CANON = {0: "Describe your AML programme."}
LLM_CFG = {"deployment": "gpt-4.1-nano", "min_severity": "low", "max_answers_per_slot": 40}

class FakeChat:
    def __init__(self, content): self._content = content
    def invoke(self, prompt):
        return type("M", (), {"content": self._content})

def test_gate_includes_risk_only_slot():
    det = []  # vote flagged nothing
    rmap = {(0, "f0", "N/A"): ["na"]}
    assert ar.suspicious_slots(det, rmap) == {0}

def test_gate_includes_vote_slot():
    det = [{"slot_id": 7, "file_name": "f9"}]
    assert ar.suspicious_slots(det, {}) == {7}

def test_llm_review_parses_and_filters():
    slots = {0: [_rec("f0", "N/A"), _rec("f1", "Full programme reviewed annually.")]}
    content = json.dumps({"answers": [
        {"file": "f0", "is_outlier": True, "category": "evasive", "severity": "high", "rationale": "no substance"},
        {"file": "f1", "is_outlier": False, "category": "none", "severity": "low", "rationale": "ok"},
    ]})
    verdicts = ar.llm_review(slots, {0}, CANON, FakeChat(content), LLM_CFG)
    assert len(verdicts) == 2
    v0 = [v for v in verdicts if v["file_name"] == "f0"][0]
    assert v0["is_outlier"] and v0["category"] == "evasive"

def test_malformed_json_skips_slot_without_crash():
    slots = {0: [_rec("f0", "N/A")]}
    verdicts = ar.llm_review(slots, {0}, CANON, FakeChat("not json"), LLM_CFG)
    assert verdicts == []

def test_build_df_only_keeps_outliers_meeting_min_severity():
    verdicts = [
        {"slot_id": 0, "file_name": "f0", "response": "N/A", "is_outlier": True,
         "category": "evasive", "severity": "high", "rationale": "no substance"},
        {"slot_id": 0, "file_name": "f1", "response": "ok", "is_outlier": False,
         "category": "none", "severity": "low", "rationale": "fine"},
    ]
    det = [{"slot_id": 0, "file_name": "f0", "votes": 1, "methods_fired": "centroid"}]
    rmap = {(0, "f0", "N/A"): ["na"]}
    df = ar.build_llm_review_df(verdicts, det, rmap, CANON, {"min_severity": "low"})
    assert list(df["file_name"]) == ["f0"]
    assert df.iloc[0]["det_methods"] == "centroid"
    assert df.iloc[0]["risk_flags"] == "na"
    assert df.iloc[0]["canonical_question"] == "Describe your AML programme."

if __name__ == "__main__":
    test_gate_includes_risk_only_slot()
    test_gate_includes_vote_slot()
    test_llm_review_parses_and_filters()
    test_malformed_json_skips_slot_without_crash()
    test_build_df_only_keeps_outliers_meeting_min_severity()
    print("OK")
