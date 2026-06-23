# tests/test_flag.py
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a

def test_flag_counts_and_threshold():
    outliers = [
        {"file_name": "f1", "question_id": "Q2"},
        {"file_name": "f1", "question_id": "Q9"},
        {"file_name": "f2", "question_id": "Q5"},
    ]
    rows = a.flag_questionnaires(outliers, all_files=["f1", "f2", "f3"], n=2)
    by = {r["file_name"]: r for r in rows}
    assert by["f1"]["n_outliers"] == 2 and by["f1"]["flagged"] is True
    assert by["f1"]["questions"] == "Q2,Q9"
    assert by["f2"]["n_outliers"] == 1 and by["f2"]["flagged"] is False
    assert by["f3"]["n_outliers"] == 0 and by["f3"]["flagged"] is False
    assert rows[0]["file_name"] == "f1"   # sorted desc
    print("OK")

if __name__ == "__main__":
    test_flag_counts_and_threshold()
