# tests/test_llm_review_wiring.py
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a
import answer_review as ar

def test_run_llm_stage_returns_df(monkeypatch=None):
    slots = {0: [{"file_name": "f0", "response": "N/A", "anchor_text": "x", "question_id": "Q1"}]}
    canon = {0: "AML?"}
    det = []
    # stub the pieces so no network is needed
    fake_verdicts = [{"slot_id": 0, "file_name": "f0", "response": "N/A", "is_outlier": True,
                      "category": "evasive", "severity": "high", "rationale": "empty"}]
    ar.make_chat = lambda cfg, lc: object()
    ar.risk_flags = lambda s, r, c: {(0, "f0", "N/A"): ["na"]}
    ar.llm_review = lambda s, ids, c, chat, lc: fake_verdicts
    df = a.run_llm_review_stage(slots, det, canon, cfg={"azure": {}}, llm_cfg=a.DEFAULT_LLM_CFG)
    assert list(df["file_name"]) == ["f0"]
    assert df.iloc[0]["category"] == "evasive"

if __name__ == "__main__":
    test_run_llm_stage_returns_df()
    print("OK")
