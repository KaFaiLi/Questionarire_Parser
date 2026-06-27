# tests/test_audit.py
# Deterministic, offline (no network, no config.yaml). Builds 6 synthetic firms
# in a temp dir with one planted issue each and asserts the matching audit check
# fires. *.json is gitignored, so fixtures are generated at runtime.
#
# Planted:
#   F2  substantive Q1 wording  ("...and Registered Address")      -> QDRIFT-SUB
#   F3  formatting-only Q1 wording (double spaces)                 -> QDRIFT-FMT
#   F4  drops the core "Primary Regulatory Authority" question     -> QMISSING
#   F1  carries a unique "MAS Full Bank condition" question        -> QEXTRA
#   F5  answers NO to direct-distribution (others YES)             -> ADRIFT-OUT
#   F6  "No - we do not maintain an AML policy"                    -> REDFLAG RF-AML
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import audit_kyd as au


def _q(qid, text, *, answer="", selection=""):
    return {"question_id": qid, "question": text,
            "sub_questions": [{"option_label": "", "prompt": "", "selection": selection, "answer": answer}]}


def _firms():
    name = "Firm Full Legal Name and Address"
    reg = "Primary Regulatory Authority"
    direct = "Will the firm distribute directly to retail investors?"
    aml = "Describe your AML / KYC programme"
    aml_ok = "We maintain a full AML and KYC programme reviewed annually by compliance."

    def base(fid, *, q1_text=name, reg_ok=True, direct_sel="YES", aml_ans=aml_ok, extra=False):
        qs = [_q("Q1", q1_text, answer=f"{fid} Holdings Limited")]
        if reg_ok:
            qs.append(_q("Q2", reg, answer="Securities and Futures Commission"))
        qs.append(_q("Q3", direct, selection=direct_sel))
        qs.append(_q("Q4", aml, answer=aml_ans))
        if extra:
            qs.append(_q("Q9", "MAS Full Bank licensing condition applicable to this firm",
                         answer="Yes, condition noted"))
        return {"file_name": fid, "questions": qs}

    return [
        base("F1_apex", extra=True),
        base("F2_bea", q1_text="Firm Full Legal Name and Registered Address"),
        base("F3_crest", q1_text="Firm  Full  Legal  Name  and  Address"),
        base("F4_delta", reg_ok=False),
        base("F5_echo", direct_sel="NO"),
        base("F6_foxtrot", aml_ans="No - we do not maintain an AML policy"),
    ]


def _run(tmp):
    d = Path(tmp)
    for firm in _firms():
        (d / f"{firm['file_name']}.json").write_text(
            json.dumps(firm, indent=2), encoding="utf-8")
    paths = au.akj._discover(d)
    return au.run_audit(paths)


def test_audit_checks_fire():
    with tempfile.TemporaryDirectory() as tmp:
        sheets = _run(tmp)
    f = sheets["findings"]
    assert not f.empty, "expected findings"

    def hits(check_id, firm):
        return f[(f["check_id"] == check_id) & (f["questionnaire"] == firm)]

    assert len(hits("QDRIFT-SUB", "F2_bea")) >= 1, f["check_id"].tolist()
    assert len(hits("QDRIFT-FMT", "F3_crest")) >= 1, f["check_id"].tolist()
    assert len(hits("QMISSING", "F4_delta")) >= 1, f.to_dict("records")
    assert len(hits("QEXTRA", "F1_apex")) >= 1, f.to_dict("records")
    assert len(hits("ADRIFT-OUT", "F5_echo")) >= 1, sheets["answer_outliers"].to_dict("records")

    rf = sheets["redflags"]
    aml = rf[(rf["rule_id"] == "RF-AML") & (rf["questionnaire"] == "F6_foxtrot")]
    assert len(aml) >= 1, rf.to_dict("records")
    assert len(f[(f["check_id"] == "REDFLAG") & (f["questionnaire"] == "F6_foxtrot")]) >= 1

    # register is well-formed: severities valid, ids unique & ordered by severity
    assert set(f["severity"]).issubset(set(au.SEV_RANK)), set(f["severity"])
    assert f["finding_id"].is_unique
    ranks = [au.SEV_RANK[s] for s in f["severity"]]
    assert ranks == sorted(ranks), "findings must be sorted by severity"
    print("OK", {c: int((f["check_id"] == c).sum())
                 for c in sorted(f["check_id"].unique())})


def test_redflag_rule_loading():
    rules = au.load_rules(None)
    ids = {r["id"] for r in rules}
    assert {"RF-AML", "RF-SANCTIONS", "RF-LICENCE"}.issubset(ids), ids
    for r in rules:
        assert r["a_re"] is not None and r["severity"] in au.SEV_RANK
    print("OK rules:", sorted(ids))


if __name__ == "__main__":
    test_audit_checks_fire()
    test_redflag_rule_loading()
