# LLM Slot-Review Answer Check Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an opt-in LLM pass that reviews suspicious answer slots for peer-disagreement and intrinsic-risk outliers the deterministic vote misses, and read the Azure API key from the environment.

**Architecture:** A new focused module `answer_review.py` holds all new logic (risk heuristics, suspicious-slot gate, LLM slot review, result merge). `analyze_drift.py` gains a thin wiring block in `_drift_report` behind a `--llm-review` flag; the existing deterministic sheets are untouched. The LLM sees a whole suspicious slot at once (all answers for one canonical question) for peer context and contradiction detection.

**Tech Stack:** Python 3, pandas, `langchain_openai.AzureChatOpenAI`, PyYAML, rapidfuzz (existing). Reuses `analyze_drift.py` (slots/canon/vecs) and `audit_kyd.load_rules`.

## Global Constraints

- **Config file `config.yaml` is gitignored** — never commit it. Config may lack the new `llm_review:` block; code must supply defaults.
- **Azure API key comes from env var `AZURE_OPENAI_API_KEY`**, falling back to `config.yaml`'s `azure.api_key`. Raise a clear error if neither is set.
- **LLM temperature = 0** (reproducibility), matching `parse_questionnaires_llm.py::make_llm`.
- **Graceful degrade:** any LLM failure warns to stderr and preserves all deterministic output. No `--llm-review` flag ⇒ behavior identical to today.
- **Test style:** plain functions, `sys.path.insert(0, parents[1])`, a `if __name__ == "__main__":` runner ending in `print("OK")`. Run with `uv run python tests/<file>.py`. No pytest fixtures.
- **Slot record shape** (from `analyze_drift.extract_answers` + `group_by_slot`): `{"file_name","question_id","sub_idx","anchor_level","anchor_text","response","slot_id"}`.
- **`canon`** is `dict[int, str]`: cluster/slot id → canonical question text.
- **`detect_outliers` output rows** have keys: `slot_id, question_id, file_name, response, votes, methods_fired, freq_share, cluster_share, centroid_z`.

---

### Task 1: Read Azure API key from the environment

**Files:**
- Modify: `analyze_kyd_json.py` (add helper near the other small utils, ~after `read_text`)
- Modify: `analyze_drift.py:46-54` (`make_embeddings`)
- Modify: `parse_questionnaires_llm.py:97-107` (`make_llm`)
- Test: `tests/test_azure_key.py`

**Interfaces:**
- Produces: `analyze_kyd_json.azure_api_key(az: dict) -> str` — returns `os.environ["AZURE_OPENAI_API_KEY"]` if set, else `az.get("api_key")`; raises `SystemExit` with a clear message if neither.

- [ ] **Step 1: Write the failing test**

Create `tests/test_azure_key.py`:

```python
# tests/test_azure_key.py
import os, sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_kyd_json as akj

def test_env_wins_over_config():
    os.environ["AZURE_OPENAI_API_KEY"] = "env-key"
    try:
        assert akj.azure_api_key({"api_key": "cfg-key"}) == "env-key"
    finally:
        del os.environ["AZURE_OPENAI_API_KEY"]

def test_falls_back_to_config():
    os.environ.pop("AZURE_OPENAI_API_KEY", None)
    assert akj.azure_api_key({"api_key": "cfg-key"}) == "cfg-key"

def test_raises_when_missing():
    os.environ.pop("AZURE_OPENAI_API_KEY", None)
    try:
        akj.azure_api_key({})
        assert False, "expected SystemExit"
    except SystemExit:
        pass

if __name__ == "__main__":
    test_env_wins_over_config()
    test_falls_back_to_config()
    test_raises_when_missing()
    print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run python tests/test_azure_key.py`
Expected: FAIL — `AttributeError: module 'analyze_kyd_json' has no attribute 'azure_api_key'`

- [ ] **Step 3: Add the helper**

In `analyze_kyd_json.py`, ensure `import os` is present at the top (add if missing), then add after `read_text`:

```python
def azure_api_key(az: dict) -> str:
    """Prefer the AZURE_OPENAI_API_KEY env var; fall back to config.yaml's azure.api_key."""
    key = os.environ.get("AZURE_OPENAI_API_KEY") or az.get("api_key")
    if not key:
        raise SystemExit("Azure key missing: set AZURE_OPENAI_API_KEY or azure.api_key in config.yaml")
    return key
```

- [ ] **Step 4: Wire both call sites**

In `analyze_drift.py::make_embeddings`, replace `api_key=az["api_key"],` with:

```python
        api_key=akj.azure_api_key(az),
```

In `parse_questionnaires_llm.py::make_llm`, add `import analyze_kyd_json as akj` at the top of the file if absent, then replace `api_key=az["api_key"],` with:

```python
        api_key=akj.azure_api_key(az),
```

- [ ] **Step 5: Run test to verify it passes**

Run: `uv run python tests/test_azure_key.py`
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add analyze_kyd_json.py analyze_drift.py parse_questionnaires_llm.py tests/test_azure_key.py
git commit -m "feat: read Azure API key from AZURE_OPENAI_API_KEY env var"
```

---

### Task 2: Risk heuristics (`risk_flags`)

**Files:**
- Create: `answer_review.py`
- Test: `tests/test_risk_flags.py`

**Interfaces:**
- Consumes: slots dict `{sid: [record,...]}`, `canon: dict[int,str]`, `rules` from `audit_kyd.load_rules(None)` (each rule: `{"id","description","severity","q_re" (compiled|None),"a_re" (compiled)}`).
- Produces: `answer_review.risk_flags(slots, rules, canon) -> dict[tuple[int,str,str], list[str]]` — key `(slot_id, file_name, response)`, value list of labels tripped (rule ids + heuristic names). Answers with no hit are absent from the dict.

- [ ] **Step 1: Write the failing test**

Create `tests/test_risk_flags.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run python tests/test_risk_flags.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'answer_review'`

- [ ] **Step 3: Create `answer_review.py` with `risk_flags`**

```python
"""Opt-in LLM slot review: flag peer-disagreement and intrinsic-risk answer
outliers the deterministic vote in analyze_drift misses. Pure heuristics +
one bounded LLM call per suspicious slot. See
docs/superpowers/specs/2026-07-01-llm-answer-check-design.md.
"""
from __future__ import annotations

import json
import re
import sys
import statistics
from collections import Counter

_NA = {"", "n/a", "na", "n.a.", "nil", "none", "-", "not applicable"}
_NEG = re.compile(r"^\s*(no|not|never|none)\b", re.I)
_HEDGE = re.compile(r"\b(maybe|we believe|to our knowledge|not sure|possibly|might|unsure|tbc|tbd)\b", re.I)


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def risk_flags(slots, rules, canon) -> dict:
    """Per-answer deterministic risk labels. Key (slot_id, file_name, response)."""
    out: dict = {}
    for sid, recs in slots.items():
        lengths = [len(r["response"]) for r in recs]
        med = statistics.median(lengths) if lengths else 0
        qtext = canon.get(sid, "")
        for r in recs:
            resp = r["response"]
            labels = []
            if _norm(resp) in _NA:
                labels.append("na")
            if _NEG.match(resp):
                labels.append("negation")
            if _HEDGE.search(resp):
                labels.append("hedge")
            # short only when the slot's typical answer is substantial (avoid all-short slots)
            if med >= 20 and len(resp) < 0.4 * med:
                labels.append("short")
            for rule in rules:
                if rule.get("q_re") and not rule["q_re"].search(qtext) and not rule["q_re"].search(r["anchor_text"]):
                    continue
                if rule["a_re"].search(resp):
                    labels.append(rule["id"])
            if labels:
                out[(sid, r["file_name"], resp)] = labels
    return out
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run python tests/test_risk_flags.py`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add answer_review.py tests/test_risk_flags.py
git commit -m "feat: risk_flags heuristics for answer review"
```

---

### Task 3: Suspicious-slot gate, LLM review, and result merge

**Files:**
- Modify: `answer_review.py`
- Test: `tests/test_llm_review.py`

**Interfaces:**
- Consumes: `detect_outliers` output list, `risk_flags` dict, slots, `canon`, a `chat` object exposing `.invoke(prompt: str)` returning an object with a `.content` str, and `llm_cfg: {"deployment","min_severity","max_answers_per_slot"}`.
- Produces:
  - `suspicious_slots(det_outliers, rmap) -> set[int]`
  - `llm_review(slots, susp_ids, canon, chat, llm_cfg) -> list[dict]` — verdict rows `{slot_id, file_name, response, is_outlier, category, severity, rationale}`.
  - `build_llm_review_df(verdicts, det_outliers, rmap, canon, llm_cfg) -> pandas.DataFrame` with columns `LLM_COLS` (below).
  - `make_chat(cfg, llm_cfg)` — real `AzureChatOpenAI` client (not unit-tested; used only in wiring).

- [ ] **Step 1: Write the failing test**

Create `tests/test_llm_review.py`:

```python
# tests/test_llm_review.py
import json, sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_review as ar

def _rec(fn, resp, sid=0): return {"file_name": fn, "question_id": "Q1",
                                   "anchor_level": "prompt", "anchor_text": "x",
                                   "response": resp, "slot_id": sid}

CANON = {0: "Describe your AML programme."}
LLM_CFG = {"deployment": "gpt-4.1-nano", "min_severity": "low", "max_answers_per_slot": 40}

class FakeChat:
    def __init__(self, content): self._content = content
    def invoke(self, prompt):
        return type("M", (), {"content": self._content})

def test_gate_includes_risk_only_slot():
    det = []  # vote flagged nothing
    rmap = {(0, "f0", "N/A"): ["na"]}
    assert ar.suspicious_slots(det, rmap) == {0}

def test_gate_includes_vote_slot():
    det = [{"slot_id": 7, "file_name": "f9"}]
    assert ar.suspicious_slots(det, {}) == {7}

def test_llm_review_parses_and_filters():
    slots = {0: [_rec("f0", "N/A"), _rec("f1", "Full programme reviewed annually.")]}
    content = json.dumps({"answers": [
        {"file": "f0", "is_outlier": True, "category": "evasive", "severity": "high", "rationale": "no substance"},
        {"file": "f1", "is_outlier": False, "category": "none", "severity": "low", "rationale": "ok"},
    ]})
    verdicts = ar.llm_review(slots, {0}, CANON, FakeChat(content), LLM_CFG)
    assert len(verdicts) == 2
    v0 = [v for v in verdicts if v["file_name"] == "f0"][0]
    assert v0["is_outlier"] and v0["category"] == "evasive"

def test_malformed_json_skips_slot_without_crash():
    slots = {0: [_rec("f0", "N/A")]}
    verdicts = ar.llm_review(slots, {0}, CANON, FakeChat("not json"), LLM_CFG)
    assert verdicts == []

def test_build_df_only_keeps_outliers_meeting_min_severity():
    verdicts = [
        {"slot_id": 0, "file_name": "f0", "response": "N/A", "is_outlier": True,
         "category": "evasive", "severity": "high", "rationale": "no substance"},
        {"slot_id": 0, "file_name": "f1", "response": "ok", "is_outlier": False,
         "category": "none", "severity": "low", "rationale": "fine"},
    ]
    det = [{"slot_id": 0, "file_name": "f0", "votes": 1, "methods_fired": "centroid"}]
    rmap = {(0, "f0", "N/A"): ["na"]}
    df = ar.build_llm_review_df(verdicts, det, rmap, CANON, {"min_severity": "low"})
    assert list(df["file_name"]) == ["f0"]
    assert df.iloc[0]["det_methods"] == "centroid"
    assert df.iloc[0]["risk_flags"] == "na"
    assert df.iloc[0]["canonical_question"] == "Describe your AML programme."

if __name__ == "__main__":
    test_gate_includes_risk_only_slot()
    test_gate_includes_vote_slot()
    test_llm_review_parses_and_filters()
    test_malformed_json_skips_slot_without_crash()
    test_build_df_only_keeps_outliers_meeting_min_severity()
    print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run python tests/test_llm_review.py`
Expected: FAIL — `AttributeError: module 'answer_review' has no attribute 'suspicious_slots'`

- [ ] **Step 3: Add gate, LLM review, merge, and client to `answer_review.py`**

Append to `answer_review.py`:

```python
LLM_COLS = ["slot_id", "canonical_question", "file_name", "response",
            "llm_is_outlier", "category", "severity", "rationale",
            "det_votes", "det_methods", "risk_flags"]

_SEV_ORDER = {"low": 0, "med": 1, "high": 2}

_PROMPT = """You audit KYD questionnaire answers. For ONE question you are given every firm's free-text answer.
Flag an answer as an outlier if EITHER:
  - it diverges in substance from the other answers (peer-disagreement), OR
  - it is intrinsically concerning regardless of peers: evasive, refuses, admits a compliance gap,
    or contradicts another answer here.
Do NOT flag mere paraphrases that mean the same thing as the majority.

Question: {question}

Answers (JSON): {answers}

Return ONLY valid JSON, no prose:
{{"answers":[{{"file":<file>,"is_outlier":<bool>,"category":"peer|evasive|gap|contradiction|none","severity":"high|med|low","rationale":<one short line>}}]}}"""


def suspicious_slots(det_outliers, rmap) -> set:
    """A slot is suspicious if the vote flagged >=1 answer OR a risk heuristic did."""
    ids = {o["slot_id"] for o in det_outliers}
    ids |= {key[0] for key in rmap}
    return ids


def llm_review(slots, susp_ids, canon, chat, llm_cfg) -> list:
    """One LLM call per suspicious slot; parse per-answer verdicts. A slot whose
    call or JSON fails is skipped with a warning (best-effort, no crash)."""
    cap = llm_cfg.get("max_answers_per_slot", 40)
    out = []
    for sid in sorted(susp_ids):
        recs = slots.get(sid, [])
        if not recs:
            continue
        # ponytail: cap slot size to bound tokens; first `cap` answers. Add sampling only if a real slot exceeds it.
        sample = recs[:cap]
        payload = [{"file": r["file_name"], "answer": r["response"]} for r in sample]
        prompt = _PROMPT.format(question=canon.get(sid, ""), answers=json.dumps(payload, ensure_ascii=False))
        try:
            content = chat.invoke(prompt).content
            data = json.loads(content)
            by_file = {r["file_name"]: r["response"] for r in sample}
            for a in data["answers"]:
                fn = a["file"]
                if fn not in by_file:
                    continue
                out.append({
                    "slot_id": sid, "file_name": fn, "response": by_file[fn],
                    "is_outlier": bool(a.get("is_outlier")),
                    "category": a.get("category", "none"),
                    "severity": a.get("severity", "low"),
                    "rationale": a.get("rationale", ""),
                })
        except Exception as exc:  # noqa: BLE001 - best-effort per slot
            print(f"WARN: llm-review slot {sid} skipped ({exc})", file=sys.stderr)
    return out


def build_llm_review_df(verdicts, det_outliers, rmap, canon, llm_cfg):
    import pandas as pd
    floor = _SEV_ORDER.get(llm_cfg.get("min_severity", "low"), 0)
    det_by = {(o["slot_id"], o["file_name"]): o for o in det_outliers}
    rows = []
    for v in verdicts:
        if not v["is_outlier"] or _SEV_ORDER.get(v["severity"], 0) < floor:
            continue
        d = det_by.get((v["slot_id"], v["file_name"]), {})
        labels = rmap.get((v["slot_id"], v["file_name"], v["response"]), [])
        rows.append({
            "slot_id": v["slot_id"],
            "canonical_question": canon.get(v["slot_id"], ""),
            "file_name": v["file_name"],
            "response": v["response"],
            "llm_is_outlier": True,
            "category": v["category"],
            "severity": v["severity"],
            "rationale": v["rationale"],
            "det_votes": d.get("votes", 0),
            "det_methods": d.get("methods_fired", ""),
            "risk_flags": ",".join(labels),
        })
    return pd.DataFrame(rows, columns=LLM_COLS)


def make_chat(cfg, llm_cfg):
    """Azure chat client for the review pass (temperature 0). Not unit-tested."""
    import analyze_kyd_json as akj
    from langchain_openai import AzureChatOpenAI
    az = cfg["azure"]
    return AzureChatOpenAI(
        azure_endpoint=az["endpoint"],
        azure_deployment=llm_cfg.get("deployment", az["deployment"]),
        api_version=az["api_version"],
        api_key=akj.azure_api_key(az),
        temperature=0,
        max_tokens=2048,
    )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run python tests/test_llm_review.py`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add answer_review.py tests/test_llm_review.py
git commit -m "feat: suspicious-slot gate + LLM slot review + merge"
```

---

### Task 4: Wire `--llm-review` into the drift pipeline

**Files:**
- Modify: `analyze_drift.py` (`_drift_report`, `analyze_drift_one`, `analyze_drift_combined`, `main`)
- Modify: `config.yaml` (add `llm_review:` block — local file, not committed)
- Test: `tests/test_llm_review_wiring.py`

**Interfaces:**
- Consumes: `answer_review.risk_flags/suspicious_slots/llm_review/build_llm_review_df/make_chat`, `audit_kyd.load_rules`.
- Produces: new Excel sheet `llm_answer_review`; new CLI flag `--llm-review`; `llm_cfg` threaded through the report functions (`None` ⇒ stage skipped).
- `DEFAULT_LLM_CFG = {"deployment": "gpt-4.1-nano", "min_severity": "low", "max_answers_per_slot": 40}` defined in `analyze_drift.py`.

- [ ] **Step 1: Write the failing test**

Create `tests/test_llm_review_wiring.py` (tests the seam function without real Azure by monkeypatching `answer_review`):

```python
# tests/test_llm_review_wiring.py
import sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_drift as a
import answer_review as ar

def test_run_llm_stage_returns_df(monkeypatch=None):
    slots = {0: [{"file_name": "f0", "response": "N/A", "anchor_text": "x", "question_id": "Q1"}]}
    canon = {0: "AML?"}
    det = []
    # stub the pieces so no network is needed
    fake_verdicts = [{"slot_id": 0, "file_name": "f0", "response": "N/A", "is_outlier": True,
                      "category": "evasive", "severity": "high", "rationale": "empty"}]
    ar.make_chat = lambda cfg, lc: object()
    ar.risk_flags = lambda s, r, c: {(0, "f0", "N/A"): ["na"]}
    ar.llm_review = lambda s, ids, c, chat, lc: fake_verdicts
    df = a.run_llm_review_stage(slots, det, canon, cfg={"azure": {}}, llm_cfg=a.DEFAULT_LLM_CFG)
    assert list(df["file_name"]) == ["f0"]
    assert df.iloc[0]["category"] == "evasive"

if __name__ == "__main__":
    test_run_llm_stage_returns_df()
    print("OK")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run python tests/test_llm_review_wiring.py`
Expected: FAIL — `AttributeError: module 'analyze_drift' has no attribute 'run_llm_review_stage'`

- [ ] **Step 3: Add the seam function + defaults to `analyze_drift.py`**

Add near the top of `analyze_drift.py` (after imports):

```python
import answer_review

DEFAULT_LLM_CFG = {"deployment": "gpt-4.1-nano", "min_severity": "low", "max_answers_per_slot": 40}
```

Add this function above `_drift_report`:

```python
def run_llm_review_stage(slots, det_outliers, canon, *, cfg, llm_cfg):
    """Build the llm_answer_review DataFrame. Best-effort: any failure warns and
    returns an empty frame so deterministic output is never lost."""
    import pandas as pd
    import audit_kyd  # lazy: audit_kyd imports analyze_drift, avoid circular top-level import
    try:
        chat = answer_review.make_chat(cfg, llm_cfg)
        rules = audit_kyd.load_rules(None)
        rmap = answer_review.risk_flags(slots, rules, canon)
        susp = answer_review.suspicious_slots(det_outliers, rmap)
        verdicts = answer_review.llm_review(slots, susp, canon, chat, llm_cfg)
        return answer_review.build_llm_review_df(verdicts, det_outliers, rmap, canon, llm_cfg)
    except Exception as exc:  # noqa: BLE001
        print(f"WARN: llm-review unavailable ({exc}); skipping.", file=sys.stderr)
        return pd.DataFrame(columns=answer_review.LLM_COLS)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run python tests/test_llm_review_wiring.py`
Expected: `OK`

- [ ] **Step 5: Thread `llm_cfg` and write the sheet**

In `analyze_drift.py`, add `llm_cfg=None` to the keyword params of `analyze_drift_one`, `analyze_drift_combined`, and `_drift_report`, and pass it through the two calls to `_drift_report`.

In `_drift_report`, after the `df_flags = pd.DataFrame(flags, ...)` line and before the `NM_COLS` block, add:

```python
    df_llm = (run_llm_review_stage(slots, outliers, canon, cfg=cfg, llm_cfg=llm_cfg)
              if llm_cfg is not None else
              pd.DataFrame(columns=answer_review.LLM_COLS))
```

Inside the `with pd.ExcelWriter(...) as xl:` block, after the `suspected_merges` sheet write, add:

```python
        df_llm.to_excel(xl, sheet_name="llm_answer_review", index=False)
```

After the existing `outliers:` print line, add:

```python
    if llm_cfg is not None:
        print(f"llm-review: {len(df_llm)} flagged answers (min_severity={llm_cfg['min_severity']})")
```

- [ ] **Step 6: Add the CLI flag and build `llm_cfg` in `main`**

In `main`, alongside the other `ap.add_argument` calls, add:

```python
    ap.add_argument("--llm-review", action="store_true",
                    help="LLM pass over suspicious slots for peer/intrinsic-risk outliers (needs config.yaml azure creds)")
```

Where `ocfg` is assembled from config, add:

```python
    llm_cfg = {**DEFAULT_LLM_CFG, **cfg.get("llm_review", {})} if args.llm_review else None
```

Pass `llm_cfg=llm_cfg` into each `analyze_drift_one(...)` and `analyze_drift_combined(...)` call in `main`.

- [ ] **Step 7: Add the config block (local file)**

Append to `config.yaml` (gitignored — local only):

```yaml
llm_review:
  deployment: gpt-4.1-nano
  min_severity: low
  max_answers_per_slot: 40
```

- [ ] **Step 8: Verify no regression + flag off is identical**

Run: `uv run python tests/test_detect.py && uv run python tests/test_llm_review.py && uv run python tests/test_risk_flags.py && uv run python tests/test_llm_review_wiring.py`
Expected: each prints `OK`

Run (smoke, real embeddings, LLM stage OFF — must behave exactly as before):
`uv run python analyze_drift.py --in Demo/ -o output/_smoke.xlsx`
Expected: runs to completion, prints the usual cluster/outlier summary, no `llm-review:` line.

- [ ] **Step 9: Commit**

```bash
git add analyze_drift.py tests/test_llm_review_wiring.py
git commit -m "feat: --llm-review flag wires LLM answer review into drift report"
```

---

## Self-Review Notes

- **Spec coverage:** risk heuristics (T2) ✓; suspicious-slot gate (T3) ✓; slot-level LLM review with structured verdict (T3) ✓; merge + additive `llm_answer_review` sheet, existing sheets untouched (T3/T4) ✓; graceful degrade + malformed-JSON skip + flag-off identical (T3/T4) ✓; env-var key both sites (T1) ✓; config `llm_review:` block with code defaults (T4) ✓; testing with fake client (T3) ✓; YAGNI cuts (no caching/diffing/looser gate) — not built ✓.
- **Category enum** consistent everywhere: `peer|evasive|gap|contradiction|none`. **Severity** enum `high|med|low` consistent (`_SEV_ORDER`).
- **Function names** consistent across tasks: `azure_api_key`, `risk_flags`, `suspicious_slots`, `llm_review`, `build_llm_review_df`, `make_chat`, `run_llm_review_stage`, `LLM_COLS`, `DEFAULT_LLM_CFG`.
- **Note for implementer:** the exact insertion lines in `_drift_report`/`main` may shift by a few lines from those cited; anchor on the named neighbouring statements, not line numbers.
