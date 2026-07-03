# KYD Questionnaire Parser & Review

Two entry points:

1. **Parse** — `parse_questionnaires_llm.py`: PDF questionnaires → JSON
   (`questions[] -> sub_questions[]` with `option_label` / `prompt` /
   `selection` / `answer`) via an Azure OpenAI vision model.

   ```
   python parse_questionnaires_llm.py --model full --dir Demo/generated_questionnaires -o output/llm_parsed
   ```

2. **Review** — `kyd_review.py` (stdlib only): cross-questionnaire audit of the
   parsed JSON. Matches the same question across files by text similarity
   (question ids are not trusted — the same question may carry different ids),
   classifies wording drift (formatting-only vs substantive), flags
   questionnaire-specific and missing questions, and hunts answer outliers
   (categorical minorities, numeric outliers via median/MAD, divergent
   free text, blank required answers, YES/NO vs written-answer mismatches).

   ```
   python kyd_review.py --dir Demo/kyd_examples -o output/kyd_review.html
   ```

   Output is one self-contained HTML report (severity tiles, a files ×
   questions heat map, click-through drill-down: per-cell findings, all
   wording variants, every firm's answer side by side) plus a matching
   `.xlsx` next to it (Heatmap / Findings / Questions / Answers sheets)
   for auditors who work in Excel. Hand both to the auditor.

Support files: `gen_kyd_questionnaires.py` regenerates the demo source
documents; `Demo/kyd_examples/` is a hand-checked JSON set with planted
findings (see its README); `tests/test_kyd_review.py` asserts every planted
finding is detected (`python -m pytest tests/`).
