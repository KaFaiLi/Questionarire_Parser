# tests/test_e2e_outliers.py
# Fixture: 6 copies of real questionnaires 11-16. Planted outliers (no real-file edits):
#   - file 11 Q5 first answer set to "No" (others "Yes"; anchor = shared prompt).
#   - file 11 Q8 sub0 selection set to "NO" (others "YES"; anchor = shared option_label).
# Both planted on questions with a STABLE anchor identical across files, so they share one
# slot even under the orthogonal stub vectors below (question-text anchors drift and would
# scatter under the stub; real embeddings merge them — see the live run flagging file 17).
# Expected: file 11 flagged (>=2 outliers) on questions Q5 and Q8.
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a

FIX = Path("tests/fixtures/outliers")

def _stub_vectors(texts):
    """Deterministic embedding: one orthogonal axis per normalized value. Identical text ->
    identical vector; different values -> orthogonal (far). Exercises voting without the API."""
    keys = sorted({a._norm(t) for t in texts})
    idx = {k: i for i, k in enumerate(keys)}
    dim = len(keys)
    vecs = {}
    for t in texts:
        v = [0.0] * dim
        v[idx[a._norm(t)]] = 1.0
        vecs[t] = v
    return vecs

def test_planted_outliers_flag_file_11():
    units = a.extract_units(FIX)
    answers = a.extract_answers(FIX)
    vecs = _stub_vectors([u["text"] for u in units] + [r["response"] for r in answers])
    assign, canon = a.assign_clusters(units, vecs, ct=0.84)
    slots = a.group_by_slot(answers, assign)
    cfg = {"min_samples": 5, "minority_frac": 0.2, "answer_threshold": 0.84, "z_k": 1.5, "min_votes": 2}
    outliers = a.detect_outliers(slots, vecs, cfg)
    files = sorted({r["file_name"] for r in answers})
    flags = a.flag_questionnaires(outliers, files, n=2)
    by = {r["file_name"]: r for r in flags}
    f11 = next(k for k in by if k.startswith("11"))
    assert by[f11]["flagged"] is True, by[f11]
    assert "Q5" in by[f11]["questions"] and "Q8" in by[f11]["questions"], by[f11]
    print("OK", by[f11])

if __name__ == "__main__":
    test_planted_outliers_flag_file_11()
