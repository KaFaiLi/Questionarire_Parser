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

def _split_fixture():
    # two clusters that SHOULD be one question: cos 0.80 < ct 0.84 so they split
    units = [
        {"file_name": "f1", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Source of wealth?"},
        {"file_name": "f2", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Source of wealth?"},
        {"file_name": "f3", "question_id": "Q1", "sub_idx": None, "level": "question", "text": "Wealth origin?"},
    ]
    vecs = {"Source of wealth?": [1.0, 0.0], "Wealth origin?": [0.8, 0.6]}  # cos 0.80
    return units, vecs

def test_near_miss_flags_split_pair():
    units, vecs = _split_fixture()
    assign, canon = a.assign_clusters(units, vecs, ct=0.84)
    assert assign[("question", "Source of wealth?")] != assign[("question", "Wealth origin?")]  # split
    nm = a.near_miss_pairs(units, assign, canon, vecs, ct=0.84, floor=0.75)
    assert len(nm) == 1 and abs(nm[0]["cosine"] - 0.80) < 1e-6
    assert {nm[0]["canon_a"], nm[0]["canon_b"]} == {"Source of wealth?", "Wealth origin?"}
    # nothing in-band when floor above the pair's cosine
    assert a.near_miss_pairs(units, assign, canon, vecs, ct=0.84, floor=0.85) == []

def test_merge_map_collapses_split():
    units, vecs = _split_fixture()
    assign, canon = a.assign_clusters(units, vecs, ct=0.84)
    assign, canon = a.apply_merge_map(units, assign, canon,
                                      [["Source of wealth?", "Wealth origin?"]])
    cid = assign[("question", "Source of wealth?")]
    assert assign[("question", "Wealth origin?")] == cid           # now one cluster
    assert canon[cid] == "Source of wealth?"                        # most frequent wins
    assert len(set(assign.values())) == 1

if __name__ == "__main__":
    test_assign_clusters_groups_and_picks_canonical()
    test_build_rows_uses_assignment()
    test_near_miss_flags_split_pair()
    test_merge_map_collapses_split()
