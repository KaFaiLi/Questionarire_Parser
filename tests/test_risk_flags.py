# tests/test_risk_flags.py
import re, sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_review as ar

def _rec(fn, resp): return {"file_name": fn, "question_id": "Q1",
                            "anchor_level": "prompt", "anchor_text": "x",
                            "response": resp, "slot_id": 0}

# a single rule: question mentions AML, answer starts with "no"
RULES = [{"id": "RF-AML", "description": "no AML", "severity": "high",
          "q_re": re.compile(r"aml", re.I), "a_re": re.compile(r"^\s*no\b", re.I)}]
CANON = {0: "Do you have an AML programme?"}

def test_na_and_negation_and_hedge_trip():
    slots = {0: [_rec("f0", "N/A"), _rec("f1", "no, we do not"),
                 _rec("f2", "maybe, to our knowledge"),
                 _rec("f3", "We maintain a full AML programme reviewed annually.")]}
    flags = ar.risk_flags(slots, RULES, CANON)
    assert "na" in flags[(0, "f0", "N/A")]
    assert "negation" in flags[(0, "f1", "no, we do not")]
    assert "RF-AML" in flags[(0, "f1", "no, we do not")]   # rule also fires
    assert "hedge" in flags[(0, "f2", "maybe, to our knowledge")]
    assert (0, "f3", "We maintain a full AML programme reviewed annually.") not in flags

def test_short_answer_relative_to_slot():
    long = "We maintain a comprehensive documented programme reviewed annually by compliance."
    slots = {0: [_rec(f"f{i}", long) for i in range(4)] + [_rec("f4", "ok")]}
    flags = ar.risk_flags(slots, [], {0: "desc?"})
    assert "short" in flags[(0, "f4", "ok")]
    assert (0, "f0", long) not in flags

if __name__ == "__main__":
    test_na_and_negation_and_hedge_trip()
    test_short_answer_relative_to_slot()
    print("OK")
