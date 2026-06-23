# tests/test_detect.py
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a

CFG = {"min_samples": 5, "minority_frac": 0.2, "answer_threshold": 0.84, "z_k": 1.5, "min_votes": 2}

def _rec(fn, resp): return {"file_name": fn, "question_id": "Q9", "response": resp, "slot_id": 0}

def test_lone_no_among_yes_is_outlier():
    slots = {0: [_rec(f"f{i}", "Yes") for i in range(9)] + [_rec("f9", "No")]}
    vecs = {"Yes": [1.0, 0.0], "No": [-1.0, 0.0]}   # opposite -> far on all 3 signals
    out = a.detect_outliers(slots, vecs, CFG)
    assert len(out) == 1 and out[0]["file_name"] == "f9"
    assert out[0]["votes"] == 3

def test_even_split_flags_nothing():
    # MAS/SFC/HKMA roughly even -> no dominant majority -> freq & cluster silent -> < 2 votes
    vals = ["MAS", "SFC", "HKMA"] * 3
    slots = {0: [_rec(f"f{i}", v) for i, v in enumerate(vals)]}
    vecs = {"MAS": [1.0, 0.0, 0.0], "SFC": [0.0, 1.0, 0.0], "HKMA": [0.0, 0.0, 1.0]}
    out = a.detect_outliers(slots, vecs, CFG)
    assert out == []

def test_small_slot_skipped():
    slots = {0: [_rec("f0", "Yes"), _rec("f1", "No")]}
    vecs = {"Yes": [1.0, 0.0], "No": [-1.0, 0.0]}
    assert a.detect_outliers(slots, vecs, CFG) == []
    print("OK")

if __name__ == "__main__":
    test_lone_no_among_yes_is_outlier()
    test_even_split_flags_nothing()
    test_small_slot_skipped()
