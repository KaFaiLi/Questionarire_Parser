# LLM slot-review answer check

**Date:** 2026-07-01
**Status:** Approved (design)

## Problem

The KYD pipeline flags *outlier answers* by peer-disagreement. Three deterministic
methods vote per slot (`analyze_drift.py::detect_outliers`): frequency, embedding-cluster
minority, centroid z-score. All three measure the same thing — does this answer differ
from peer answers to the same canonical question.

Free text breaks this. When every firm writes a unique sentence:

- no majority exists → `maj_exists` / `cluster_dom` go false → freq & cluster methods abstain;
- only centroid-z fires, and it flags mere paraphrases.

Two outlier kinds are therefore missed:

1. **Peer-disagreement** phrased uniquely (substantively different, but no exact/near majority to diverge from).
2. **Intrinsic risk** — an answer concerning on its own regardless of peers: evasive,
   `N/A`/refusal, admits a compliance gap, or contradicts another answer in the same form.
   The regex rules in `audit_rules.yaml` catch only a slice, and only when peers *disagree*.

## Goal

Add an opt-in LLM review that catches both kinds, with bounded token cost and graceful
degradation. The existing deterministic output stays unchanged and fully reproducible; the
LLM layer is purely additive.

## Decisions (locked during brainstorming)

- **Two-stage:** deterministic layer builds a shortlist → LLM (nano) judges only the shortlist.
- **Slot-level review (approach C):** the LLM sees a whole slot at once (all answers for that
  canonical question), not one answer in isolation — gives real peer context and enables
  contradiction detection.
- **Gate = suspicious slots only:** a slot is sent to the LLM iff ≥1 of its answers already
  tripped the peer-vote **or** a risk heuristic. Cheapest bounded option.
  - *Ceiling:* a contradiction in a slot where nothing tripped any heuristic is missed.
    Acceptable v1; upgrade path is a looser gate (skip-unanimous or all-slots).

## Architecture

New stage in `analyze_drift.py`, in the run-one path (currently ~L537–567, between
`detect_outliers` and report write), enabled by a new `--llm-review` flag. Three steps:

1. **Risk heuristics** — `risk_flags(slots, rules)` → per-answer deterministic risk hits.
   Reuses `audit_rules.yaml` regexes plus cheap heuristics:
   - `n/a` / `nil` / `none` / blank-ish,
   - leading negation,
   - hedge tokens (`maybe`, `we believe`, `to our knowledge`, `not sure`, …),
   - length below a slot-relative percentile (suspiciously short vs peers).
2. **Gate** — a slot is *suspicious* if ≥1 answer appears in `detect_outliers` output **or**
   in `risk_flags`. Only suspicious slots proceed. This is what plugs the everyone-evasive
   hole: peers agree, vote stays silent, but the risk heuristic still selects the slot.
3. **LLM slot review** — `llm_review(suspicious_slots, canon, cfg)` → one `AzureChatOpenAI`
   call per slot, `temperature=0` (reuse the client pattern from
   `parse_questionnaires_llm.py::make_llm`).

## Components

### `risk_flags(slots, rules) -> dict[(sid, file_name), list[str]]`
Pure, no LLM, fully testable. Returns the list of heuristic/rule labels each answer tripped
(empty ⇒ clean). Drives the gate and is also surfaced as evidence.

### `llm_review(slots_subset, canon, cfg) -> list[verdict]`
Per slot, sends:

```
{ "canonical_question": <canon[sid]>,
  "answers": [ { "file": <file_name>, "answer": <response> }, ... ] }
```

Structured JSON back, one entry per answer:

```
{ "file": <file_name>,
  "is_outlier": <bool>,
  "category": "peer" | "evasive" | "gap" | "contradiction" | "none",
  "severity": "high" | "med" | "low",
  "rationale": <one line> }
```

Seeing every answer in the slot gives peer context and lets the model detect
contradictions across firms. `ponytail:` comment caps slot size via
`llm_review.max_answers_per_slot`; sampling is only added if a real slot exceeds the cap.

### Merge
LLM `is_outlier` is authoritative for inclusion. Deterministic evidence
(`votes`, `methods_fired`, `freq_share`, `cluster_share`, `centroid_z`, `risk_flags`) rides
along per row as supporting columns.

## Data flow

```
answer_recs → slots → { detect_outliers , risk_flags }
                            │            │
                            └──── suspicious-slot gate ────► llm_review
                                                                │
                        verdicts ⨝ deterministic evidence ──────┘
                                        │
                                     report
```

## Output

- **New sheet `llm_answer_review`:** `slot_id, canonical_question, file_name, answer,
  llm_is_outlier, category, severity, rationale, det_votes, det_methods, risk_flags`.
- Existing `answer_outliers` sheet **unchanged** (deterministic, reproducible).
- LLM sheet is additive — nothing is removed or altered in existing sheets/HTML.

## Error handling

- No `--llm-review` flag → LLM stage skipped. The `llm_answer_review` sheet is still
  written but empty (columns only), so every run yields a consistent workbook shape.
  Existing sheets and HTML are unchanged.
- LLM unavailable or a call throws → best-effort per slot: warn to stderr, skip that slot,
  keep all deterministic output. Mirrors the existing `--use-embeddings` degrade pattern.
  No data loss.
- Malformed LLM JSON for a slot → skip that slot with a warning; never crash the run.

## API key from environment (folded-in scope)

`config.yaml` is gitignored, but the committed local copy hardcodes a live Azure key. Read
the key from the environment instead, at both call sites:

- `analyze_drift.py:53`
- `parse_questionnaires_llm.py:104`

Change: `api_key = os.environ.get("AZURE_OPENAI_API_KEY") or az.get("api_key")`, and raise a
clear error if neither is set. Env wins; the `config.yaml` value stays as a local fallback so
existing setups keep working. Rotate the currently-committed key separately (operational, not
code).

## Config (`config.yaml`, new `llm_review:` block)

```yaml
llm_review:
  deployment: gpt-4.1-nano       # reuse azure: creds
  min_severity: low              # lowest severity surfaced in the sheet
  max_answers_per_slot: 40       # cap; ponytail — sample only if exceeded
```

## Testing

Follow existing test style in `tests/` (plain asserts, no new framework/fixtures).

- `risk_flags`: n/a / negation / hedge / short answers trip; normal prose does not.
- Gate: a slot with a heuristic hit but **no** peer-vote is still selected (everyone-evasive case).
- `llm_review`: inject a **fake client** returning canned JSON → verify merge, malformed-JSON
  skip, and LLM-down fallback (deterministic output preserved).
- One runnable `assert`-based self-check per new non-trivial function.

## Scope cuts (YAGNI)

- No verdict caching, no cross-run diffing.
- No severity-weighted questionnaire flagging v1 (deterministic `multi_outlier_n` stays as-is).
- No looser gate v1 (suspicious-only); upgrade path documented above.
