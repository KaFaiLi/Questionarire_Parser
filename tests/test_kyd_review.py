"""Planted scenarios in Demo/kyd_examples must all be detected (see its README)."""
from pathlib import Path

import kyd_review as kr

# difflib path: deterministic, no Azure needed (embeddings covered separately)
REPORT = kr.analyze(kr.load_files([(p, None) for p in sorted(Path("Demo/kyd_examples").glob("*.json"))]))


def _findings(kind=None, file=None):
    return [f for f in REPORT["findings"]
            if (kind is None or f["kind"] == kind) and (file is None or file in f["file"])]


def _cluster_of(f):
    return REPORT["clusters"][f["cluster"]]["title"]


def test_ids_not_trusted_capacity_question_merges_across_hk_sg():
    c = next(c for c in REPORT["clusters"] if "capacity" in c["title"])
    assert c["coverage"] == 10 and {"Q10", "Q12"} <= set(c["ids"].values())


def test_questionnaire_specific_mas_question():
    fs = _findings("questionnaire-specific")
    assert len(fs) == 1 and fs[0]["file"] == "02_pacific_trust_sg"


def test_substantive_vs_formatting_wording():
    assert any("Registered Address" in f["message"] for f in _findings("wording variant"))
    assert _findings("formatting variant", "06_summit")   # extra spaces in Q1
    assert _findings("formatting variant", "09_apex")     # REGULATORY  AUTHORITY


def test_missing_required_answers():
    assert any("conomic data" in _cluster_of(f) for f in _findings("missing answer", "04_crescent"))
    assert any("Marketing" in _cluster_of(f) for f in _findings("missing answer", "06_summit"))


def test_conditional_blanks_not_flagged():
    assert not any("custodian" in _cluster_of(f) for f in _findings("missing answer"))


def test_answer_outliers():
    assert _findings("selection outlier", "07_harbourview")   # 9x YES / 1x NO
    assert _findings("free-text outlier", "05_northwind")     # divergent marketing answer
    assert _findings("numeric outlier", "03_meridian")        # 45 issuers vs 8-12


def test_declared_na_flagged():
    fs = _findings("declared N/A")
    assert len(fs) == 1 and fs[0]["file"] == "08_silkroad_sg"
    assert "conomic data" in _cluster_of(fs[0])


def test_yes_no_synonym_not_an_outlier():
    answer_kinds = {"selection outlier", "free-text outlier", "numeric outlier",
                    "selection/answer mismatch", "declared N/A", "missing answer"}
    assert not [f for f in _findings(file="10_lighthouse") if f["kind"] in answer_kinds]
    # and canonicalization must not hide the real 9-vs-1 NO outlier
    assert _findings("free-text outlier", "07_harbourview")


def test_heatmap_matrix_shape():
    assert len(REPORT["cells"]) == 10
    assert all(len(v) == len(REPORT["clusters"]) for v in REPORT["cells"].values())


def test_qvar_wording_grid():
    # question x firm wording-divergence grid: codes 0 same / 1 formatting / 2 substantive / -1 absent
    qv = REPORT["qvar"]
    assert len(qv) == len(REPORT["clusters"])
    assert all(len(row) == len(REPORT["files"]) for row in qv)
    codes = {c for row in qv for c in row}
    assert codes <= {-1, 0, 1, 2}
    # a firm marked substantive (2) for a cluster must have a real wording finding there
    warn = {(f["file"], f["cluster"]) for f in REPORT["findings"] if f["kind"] == "wording variant"}
    for ci, row in enumerate(qv):
        for fi, code in enumerate(row):
            if code == 2:
                assert (REPORT["files"][fi], ci) in warn


def test_xlsx_export(tmp_path):
    from openpyxl import load_workbook
    p = tmp_path / "r.xlsx"
    kr.render_xlsx(REPORT, p)
    wb = load_workbook(p)
    assert [ws.title for ws in wb] == ["Heatmap", "Question variance", "Findings", "Questions", "Answers"]
    assert wb["Findings"].max_row == len(REPORT["findings"]) + 1
    assert wb["Heatmap"].max_column == len(REPORT["clusters"]) + 1
    # one variance row per canonical question, worst-first
    assert wb["Question variance"].max_row == len(REPORT["clusters"]) + 1
    scores = [wb["Question variance"].cell(row=r, column=8).value
              for r in range(2, wb["Question variance"].max_row + 1)]
    assert scores == sorted(scores, reverse=True)
