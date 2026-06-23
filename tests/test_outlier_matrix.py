# tests/test_outlier_matrix.py
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd, analyze_drift as a

def test_builds_html():
    slots = {0: [{"file_name": "f1", "question_id": "Q9", "response": "Yes"},
                 {"file_name": "f2", "question_id": "Q9", "response": "No"}]}
    canon = {0: "Distribute directly?"}
    df_out = pd.DataFrame([{"slot_id": 0, "question_id": "Q9", "file_name": "f2",
                            "response": "No", "votes": 3, "methods_fired": "freq|cluster|centroid",
                            "freq_share": 0.5, "cluster_share": 0.5, "centroid_z": -1.0}])
    df_flags = pd.DataFrame([{"file_name": "f1", "n_outliers": 0, "flagged": False, "questions": ""},
                             {"file_name": "f2", "n_outliers": 1, "flagged": False, "questions": "Q9"}])
    p = Path("tests/_om.html")
    a.make_outlier_matrix(df_out, df_flags, slots, canon, p, {"multi_outlier_n": 2})
    h = p.read_text(encoding="utf-8")
    assert "heatmap" in h and "Q9" in h
    p.unlink()
    print("OK")

if __name__ == "__main__":
    test_builds_html()
