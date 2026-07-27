"""Focused-review behaviour for flat, long-format questionnaires."""

import questionnaire_review as qr


def _question(question: str, answer: str = "complete") -> dict:
    return {
        "question_id": "Q1",
        "question": question,
        "section": "General",
        "answer": answer,
        "sub_questions": [],
    }


def _file(index: int, questions: list[dict]) -> dict:
    return {"name": f"f{index}", "folder": "", "questions": questions}


def _cluster(report: dict, title: str) -> tuple[int, dict]:
    return next((i, c) for i, c in enumerate(report["clusters"])
                if c["title"] == title)


def test_tolerant_matching_profile_is_shared_with_kyd_review():
    matching = qr.question_matching_settings({"question_matching": {"mode": "tolerant"}})

    assert matching["mode"] == "tolerant"
    assert matching["embedding_threshold"] == 0.78
    assert matching["slot_similarity"] == 0.60


def test_focused_review_only_marks_broad_consensus_failures():
    files = []
    for i in range(10):
        questions = [_question("Common question")]
        if i < 5:
            questions.append(_question("Optional question"))
        if i == 0:
            questions.append(_question("Questionnaire-only question"))
        files.append(_file(i, questions))

    report = qr.analyze(files, False, None, None, None, 0.75)
    partial_index, partial = _cluster(report, "Optional question")

    assert not partial["expected"]
    assert not [f for f in report["findings"] if f["kind"] == "question missing"]
    assert report["cells"]["f5"][partial_index] == -1

    specific = [f for f in report["findings"] if f["kind"] == "questionnaire-specific"]
    assert len(specific) == 1
    assert specific[0]["file"] == "f0"
    assert specific[0]["level"] == qr.NOTE


def test_answer_consensus_thresholds_are_configurable():
    files = [
        _file(i, [_question("Completeness", "" if i == 0 else "complete")])
        for i in range(6)
    ]

    default = qr.analyze(files, False, None, None, None, 0.75)
    assert not [f for f in default["findings"] if f["kind"] == "missing answer"]

    overridden = qr.analyze(
        files, False, None, None, None, 0.75,
        {"required_answer_share": 0.8},
    )
    missing = [f for f in overridden["findings"] if f["kind"] == "missing answer"]
    assert len(missing) == 1
    assert missing[0]["file"] == "f0"
