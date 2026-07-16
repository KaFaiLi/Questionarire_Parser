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

   The review defaults are deliberately focused on broad-consensus failures.
   Tune them in the ignored local `config.yaml` under `review`:

   ```yaml
   review:
     specific_question_share: 0.15   # informational only
     common_question_share: 0.85     # absence becomes missing only at this coverage
     wording_min_coverage: 0.80
     wording_variant_max_share: 0.25
     required_answer_share: 0.85
     na_substantive_share: 0.85
   ```

   For OCR/parser-heavy questionnaires, use tolerant question matching. It
   preserves the full signature (including sub-question prompts) so records
   without top-level question text still match:

   ```yaml
   question_matching:
     mode: tolerant
     embedding_threshold: 0.78
     slot_similarity: 0.60
     parent_presence:
       allow_top_level_evidence: true
       top_level_similarity: 0.78
       allow_subquestion_evidence: true
       slot_similarity: 0.78
       minimum_distinctive_slot_matches: 1
       uniqueness_margin: 0.05
       ignore_subquestion_wording_drift: true
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
