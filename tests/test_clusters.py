import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a

def test_assign_clusters_groups_and_picks_canonical():
    units = [
        {"file_name": "f1", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Firm Name"},
        {"file_name": "f2", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Firm Name"},
        {"file_name": "f3", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Company Name"},
    ]
    # identical vectors for the two "Firm Name", a near one for "Company Name"
    vecs = {"Firm Name": [1.0, 0.0], "Company Name": [0.99, 0.14]}  # cos ~0.99
    assign, canon = a.assign_clusters(units, vecs, ct=0.84)
    cid = assign[("question", "Firm Name")]
    assert assign[("question", "Company Name")] == cid           # merged
    assert canon[cid] == "Firm Name"                              # most frequent is canonical

def test_build_rows_uses_assignment():
    units = [
        {"file_name": "f1", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Firm Name"},
        # second "Firm Name" occurrence so its count (2) beats "Company Name" (1) under the
        # canonical tie-break (-count, -len, alpha) -- otherwise the longer "Company Name"
        # string would win the tie at count=1 each, which is not what this test is checking.
        {"file_name": "f3", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Firm Name"},
        {"file_name": "f2", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Company Name"},
    ]
    vecs = {"Firm Name": [1.0, 0.0], "Company Name": [0.99, 0.14]}
    assign, canon = a.assign_clusters(units, vecs, ct=0.84)
    rows = a.build_rows(units, vecs, assign, canon, drift_threshold=0.90)
    canon_row = next(r for r in rows if r["is_canonical"])
    assert canon_row["variant_text"] == "Firm Name"
    drift_row = next(r for r in rows if not r["is_canonical"])
    assert drift_row["is_drift"] is False                         # sim 0.99 >= 0.90 -> not drift
    print("OK")

if __name__ == "__main__":
    test_assign_clusters_groups_and_picks_canonical()
    test_build_rows_uses_assignment()
