# tests/test_combine.py
# --combine pools every input folder into ONE population: a single shared
# consensus baseline, one report, each questionnaire id namespaced as
# '<source>/<file>' so files keep their folder of origin and never clash.
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_kyd_json as akj
import audit_kyd as au


def _firm(fid, *, reg="Securities and Futures Commission", aml_ok=True):
    qs = [
        {"question_id": "Q1", "question": "Firm Full Legal Name and Address",
         "sub_questions": [{"option_label": "", "prompt": "", "selection": "", "answer": f"{fid} Ltd"}]},
        {"question_id": "Q2", "question": "Primary Regulatory Authority",
         "sub_questions": [{"option_label": "", "prompt": "", "selection": "", "answer": reg}]},
        {"question_id": "Q3", "question": "Describe your AML / KYC programme",
         "sub_questions": [{"option_label": "", "prompt": "", "selection": "",
                            "answer": "Full AML programme reviewed annually"
                            if aml_ok else "No - we do not maintain an AML policy"}]},
    ]
    return {"file_name": fid, "questions": qs}


def _make_folder(parent: Path, name: str, fids, **kw):
    d = parent / name
    d.mkdir(parents=True)
    for i, fid in enumerate(fids):
        firm = _firm(fid, aml_ok=(i != 0) if kw.get("plant_aml") else True)
        (d / f"{fid}.json").write_text(json.dumps(firm), encoding="utf-8")
    return d


def test_unique_sources_disambiguates_colliding_basenames():
    with tempfile.TemporaryDirectory() as tmp:
        t = Path(tmp)
        a = _make_folder(t / "A", "ground_truth", ["f1", "f2"])
        b = _make_folder(t / "B", "ground_truth", ["f3", "f4"])
        labeled = akj.unique_sources([a, b])
        labels = [lbl for _, lbl in labeled]
        assert len(set(labels)) == 2, labels          # no collision
        assert "ground_truth" in labels[0]
        print("OK unique_sources:", labels)


def test_pooled_load_and_audit():
    with tempfile.TemporaryDirectory() as tmp:
        t = Path(tmp)
        a = _make_folder(t / "A", "ground_truth", ["f1", "f2", "f3"])
        b = _make_folder(t / "B", "ground_truth", ["f4", "f5", "f6"], plant_aml=True)
        labeled = akj.unique_sources([a, b])

        long = akj.load_long_many(labeled, frozenset())
        assert "source" in long.columns
        assert long["questionnaire"].str.contains("/").all()
        assert long["questionnaire"].nunique() == 6           # pooled population
        assert long["source"].nunique() == 2

        sheets = au.run_audit_combined(labeled, treat_as_blank="")
        f = sheets["findings"]
        # the planted "No AML policy" firm is flagged, and is traceable to its folder
        rf = f[f["check_id"] == "REDFLAG"]
        assert not rf.empty
        assert rf["questionnaire"].str.contains("/").all()
        assert {s.split("/")[0] for s in f["questionnaire"]} <= set(l for _, l in labeled)
        print("OK pooled audit: %d findings across %d questionnaires"
              % (len(f), long["questionnaire"].nunique()))


if __name__ == "__main__":
    test_unique_sources_disambiguates_colliding_basenames()
    test_pooled_load_and_audit()
