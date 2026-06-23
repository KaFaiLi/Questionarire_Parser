# tests/test_answers.py
import sys, json; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a

def _write(tmp, name, questions):
    p = tmp / name
    p.write_text(json.dumps({"questions": questions, "file_name": name[:-5]}), encoding="utf-8")

def test_extract_and_slot(tmp_path=Path("tests/_tmp_answers")):
    tmp_path.mkdir(parents=True, exist_ok=True)
    q = [{"question_id": "Q9", "question": "Distribute directly?",
          "sub_questions": [{"option_label": "", "prompt": "", "selection": "YES", "answer": "Yes"}]}]
    _write(tmp_path, "01_x.json", q)
    _write(tmp_path, "02_x.json", q)
    recs = a.extract_answers(tmp_path)
    assert len(recs) == 2
    assert recs[0]["anchor_level"] == "question" and recs[0]["anchor_text"] == "Distribute directly?"
    assert recs[0]["response"] == "Yes"
    # slot binding: both anchor texts cluster together
    vecs = {"Distribute directly?": [1.0, 0.0]}
    units = [{"file_name": r["file_name"], "question_id": "Q9", "sub_idx": None,
              "level": "question", "text": "Distribute directly?"} for r in recs]
    assign, _ = a.assign_clusters(units, vecs, ct=0.84)
    slots = a.group_by_slot(recs, assign)
    assert len(slots) == 1 and len(next(iter(slots.values()))) == 2
    import shutil; shutil.rmtree(tmp_path)
    print("OK")

if __name__ == "__main__":
    test_extract_and_slot()
