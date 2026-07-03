"""Cross-questionnaire review of parsed KYD JSON files.

Reads every parsed questionnaire (the ``questions[] -> sub_questions[]`` format
produced by parse_questionnaires_llm.py), matches the *same* question across
files by text similarity (question ids are NOT trusted — the same question can
be Q10 in one file and Q12 in another), then checks:

Questions
  - wording drift per canonical question: formatting-only vs substantive
  - questions present in only a few files (questionnaire-specific)
  - common questions missing from a file

Answers (per canonical question, per aligned sub-question slot)
  - required answers left blank (blank-in-a-minority => missing;
    blank-in-the-majority => conditional branch, not a defect)
  - categorical outliers (e.g. 9x YES / 1x NO)
  - numeric outliers (robust median/MAD)
  - free-text minority answers diverging from the common wording
  - YES/NO selection contradicting the written answer

Output: one self-contained HTML report (heat map of files x questions with
click-through drill-down) plus a console summary.

    python kyd_review.py --dir Demo/kyd_examples -o output/kyd_review.html

Stdlib only.
"""
from __future__ import annotations

import argparse
import difflib
import html
import json
import re
import statistics
from collections import Counter
from pathlib import Path

# severity levels
OK, NOTE, WARN, MISSING, OUTLIER = 0, 1, 2, 3, 4
LEVEL_NAMES = {OK: "ok", NOTE: "formatting", WARN: "wording / specific",
               MISSING: "missing", OUTLIER: "answer outlier"}

SIM_QUESTION = 0.60   # min similarity to treat two questions as the same
SIM_SLOT = 0.70       # min similarity to align two sub-question slots
SIM_TEXT_GROUP = 0.75 # min similarity to group two free-text answers
SPECIFIC_SHARE = 0.25 # present in <= this share of files => questionnaire-specific
MINORITY_SHARE = 0.25 # answer value held by <= this share => outlier candidate
REQUIRED_SHARE = 0.70 # answered in >= this share => blanks are "missing"


def norm(s: str) -> str:
    """Case/whitespace/punctuation-insensitive form for comparisons."""
    return re.sub(r"[^a-z0-9 ]+", "", re.sub(r"\s+", " ", (s or "").lower())).strip()


# "N/A", "n.a.", "NA", "Not applicable", "Nil", "None" — a declared non-answer,
# distinct from a blank: the firm asserts the question does not apply.
NA_VALUES = {"na", "n a", "not applicable", "nil", "none"}
YESNO = {"y": "YES", "yes": "YES", "n": "NO", "no": "NO"}


def answer_class(s: str) -> str:
    """'blank' | 'na' (declared not-applicable) | 'text' (substantive)."""
    t = norm(s)
    return "blank" if not t else ("na" if t in NA_VALUES else "text")


def canon(s: str) -> str:
    """Comparison value: yes/no synonyms collapsed (Y == yes == YES)."""
    t = norm(s)
    return YESNO.get(t, t)


def ratio(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()


def q_signature(q: dict) -> str:
    """Text that identifies a question: heading + prompts + option labels."""
    parts = [q.get("question", "")]
    for s in q.get("sub_questions", []):
        parts += [s.get("option_label", ""), s.get("prompt", "")]
    return " ".join(p for p in parts if p)


def load_files(paths: list[Path]) -> list[dict]:
    out = []
    for p in sorted(paths):
        d = json.loads(p.read_text(encoding="utf-8"))
        out.append({"name": d.get("file_name") or p.stem, "questions": d.get("questions", [])})
    return out


# --- question matching -----------------------------------------------------

def cluster_questions(files: list[dict]) -> list[dict]:
    """Greedy clustering of questions across files by signature similarity.

    ponytail: O(files x questions x clusters) difflib scan — fine for hundreds
    of files; add a token-overlap prefilter if it ever gets slow.
    """
    clusters: list[dict] = []
    for fi, f in enumerate(files):
        for pos, q in enumerate(f["questions"]):
            sig = norm(q_signature(q))
            best, best_r = None, 0.0
            for c in clusters:
                if f["name"] in c["members"]:
                    continue  # one question per file per cluster
                r = ratio(sig, c["sig"])
                if r > best_r:
                    best, best_r = c, r
            if best is not None and best_r >= SIM_QUESTION:
                best["members"][f["name"]] = q
                best["positions"].append(pos)
            else:
                clusters.append({"sig": sig, "members": {f["name"]: q}, "positions": [pos]})
    clusters.sort(key=lambda c: statistics.median(c["positions"]))
    return clusters


def cluster_title(c: dict) -> str:
    texts = [q["question"] for q in c["members"].values() if q.get("question")]
    if not texts:
        texts = [s["prompt"] for q in c["members"].values()
                 for s in q.get("sub_questions", []) if s.get("prompt")]
    return Counter(texts).most_common(1)[0][0] if texts else "(untitled)"


# --- checks ------------------------------------------------------------------

def check_variants(c: dict, findings: list, ci: int) -> list[dict]:
    """Classify wording variants of one canonical question."""
    by_raw: dict[str, list[str]] = {}
    for fname, q in c["members"].items():
        by_raw.setdefault(q_signature(q), []).append(fname)
    variants = [{"text": raw, "files": fs} for raw, fs in by_raw.items()]
    if len(variants) == 1:
        variants[0]["kind"] = "majority"
        return variants
    majority = max(variants, key=lambda v: len(v["files"]))
    for v in variants:
        if v is majority:
            v["kind"] = "majority"
        elif norm(v["text"]) == norm(majority["text"]):
            v["kind"] = "formatting"
            for fn in v["files"]:
                findings.append(dict(file=fn, cluster=ci, level=NOTE, kind="formatting variant",
                                     message="Question wording differs only in spacing/case from the common version."))
        else:
            v["kind"] = "substantive"
            for fn in v["files"]:
                findings.append(dict(file=fn, cluster=ci, level=WARN, kind="wording variant",
                                     message=f"Question wording differs substantively from the common version: “{v['text'][:120]}”"))
    return variants


def _enum_marker(s: str) -> str:
    """Leading enumeration marker like 'a)', '(ii)', '3.' — '' if none."""
    m = re.match(r"\(?([a-z0-9]{1,3})[\).]", s.strip().lower())
    return m.group(1) if m else ""


def align_slots(c: dict) -> list[dict]:
    """Align sub-questions across files by option label + prompt similarity.

    Slots with different enumeration markers (a) vs b)) are never merged —
    similar wording there means sibling sub-questions, not the same one.
    """
    slots: list[dict] = []
    for fname, q in c["members"].items():
        for s in q.get("sub_questions", []):
            raw = (s.get("option_label", "") + " " + s.get("prompt", "")).strip()
            key, mark = norm(raw), _enum_marker(raw)
            hit = next((sl for sl in slots if sl["mark"] == mark
                        and (sl["key"] == key or ratio(key, sl["key"]) >= SIM_SLOT)), None)
            if hit is None:
                label = (s.get("option_label") or s.get("prompt") or "answer").strip()
                hit = {"key": key, "mark": mark, "label": label, "entries": []}
                slots.append(hit)
            hit["entries"].append({"file": fname, "selection": s.get("selection", ""),
                                   "answer": s.get("answer", ""), "flags": []})
    return slots


def _num(s: str) -> float | None:
    m = re.fullmatch(r"[~\s]*([-+]?[\d,]+(?:\.\d+)?)\s*%?", s.strip())
    return float(m.group(1).replace(",", "")) if m else None


def _flag(entry: dict, findings: list, ci: int, level: int, kind: str, message: str) -> None:
    entry["flags"].append(kind)
    findings.append(dict(file=entry["file"], cluster=ci, level=level, kind=kind, message=message))


def check_answers(slot: dict, findings: list, ci: int) -> None:
    entries = slot["entries"]
    n = len(entries)
    if n < 3:
        return

    # YES/NO selection vs written answer contradiction
    for e in entries:
        sel = YESNO.get(norm(e["selection"]), "")
        first = YESNO.get(norm(e["answer"]).split(" ")[0] if e["answer"].strip() else "", "")
        if sel and first and first != sel:
            _flag(e, findings, ci, OUTLIER, "selection/answer mismatch",
                  f"Selection is {e['selection']} but the written answer says “{e['answer'][:80]}”.")

    # categorical minority on the selection column (yes/no synonyms collapsed)
    sels = [e for e in entries if e["selection"].strip()]
    if len(sels) >= 5:
        counts = Counter(canon(e["selection"]) for e in sels)
        if len(counts) <= 4:
            top = counts.most_common(1)[0][1]
            for val, cnt in counts.items():
                if cnt / len(sels) <= MINORITY_SHARE and cnt < top:
                    for e in sels:
                        if canon(e["selection"]) == val:
                            _flag(e, findings, ci, OUTLIER, "selection outlier",
                                  f"Selection “{e['selection']}” given by {cnt}/{len(sels)} firms; the rest answered differently.")

    # blanks vs declared N/A vs substantive answers
    answered = [e for e in entries if answer_class(e["answer"]) != "blank"]
    filled = [e for e in entries if answer_class(e["answer"]) == "text"]
    if answered and len(answered) / n >= REQUIRED_SHARE:
        for e in entries:
            if answer_class(e["answer"]) == "blank" and not e["selection"].strip():
                _flag(e, findings, ci, MISSING, "missing answer",
                      f"Answer left blank while {len(answered)}/{n} firms answered.")
    # declared N/A while peers answered substantively => the assertion needs challenging
    if answered and len(filled) / len(answered) >= REQUIRED_SHARE:
        for e in answered:
            if answer_class(e["answer"]) == "na":
                _flag(e, findings, ci, MISSING, "declared N/A",
                      f"Firm answered “{e['answer']}” while {len(filled)}/{len(answered)} peers gave a substantive answer — the not-applicable claim should be challenged.")
    if len(filled) < 5:
        return

    nums = [(_num(e["answer"]), e) for e in filled]
    if sum(1 for v, _ in nums if v is not None) / len(filled) >= 0.8:
        vals = [v for v, _ in nums if v is not None]
        med = statistics.median(vals)
        mad = statistics.median([abs(v - med) for v in vals])
        for v, e in nums:
            if v is None:
                continue
            z = 0.6745 * abs(v - med) / mad if mad else (abs(v - med) / (abs(med) or 1))
            if z > 3.5:
                _flag(e, findings, ci, OUTLIER, "numeric outlier",
                      f"Value {e['answer']} is far from the peer median of {med:g}.")
        return

    # free text: greedy similarity groups, minority group => outlier
    groups: list[list[dict]] = []
    for e in filled:
        t = canon(e["answer"])
        g = next((g for g in groups if ratio(t, canon(g[0]["answer"])) >= SIM_TEXT_GROUP), None)
        (g.append(e) if g is not None else groups.append([e]))
    top = max(len(g) for g in groups)
    if top / len(filled) >= 0.6:
        for g in groups:
            if len(g) / len(filled) <= MINORITY_SHARE and len(g) < top:
                for e in g:
                    _flag(e, findings, ci, OUTLIER, "free-text outlier",
                          f"Answer diverges from the wording used by {top}/{len(filled)} firms: “{e['answer'][:100]}”")


# --- analysis ---------------------------------------------------------------

def analyze(files: list[dict]) -> dict:
    n_files = len(files)
    clusters = cluster_questions(files)
    findings: list[dict] = []
    out_clusters = []
    for ci, c in enumerate(clusters):
        coverage = len(c["members"])
        specific = coverage / n_files <= SPECIFIC_SHARE
        if specific:
            for fn in c["members"]:
                findings.append(dict(file=fn, cluster=ci, level=WARN, kind="questionnaire-specific",
                                     message=f"Question appears in only {coverage}/{n_files} questionnaires."))
        else:
            for f in files:
                if f["name"] not in c["members"]:
                    findings.append(dict(file=f["name"], cluster=ci, level=MISSING, kind="question missing",
                                         message=f"Question present in {coverage}/{n_files} questionnaires but absent here."))
        variants = check_variants(c, findings, ci)
        slots = align_slots(c)
        for slot in slots:
            check_answers(slot, findings, ci)
        out_clusters.append({
            "title": cluster_title(c),
            "ids": {fn: q["question_id"] for fn, q in c["members"].items()},
            "coverage": coverage, "specific": specific,
            "variants": variants,
            "slots": [{"label": s["label"], "entries": s["entries"]} for s in slots],
        })

    fnames = [f["name"] for f in files]
    cells = {fn: [] for fn in fnames}
    per_cell = {(g["file"], g["cluster"]): [] for g in findings}
    for g in findings:
        per_cell[(g["file"], g["cluster"])].append(g)
    for fn in fnames:
        for ci, c in enumerate(out_clusters):
            if fn not in c["ids"] and c["specific"]:
                cells[fn].append(-1)  # not expected here
            else:
                cells[fn].append(max((g["level"] for g in per_cell.get((fn, ci), [])), default=OK))
    findings.sort(key=lambda g: (-g["level"], g["cluster"], g["file"]))
    return {"files": fnames, "clusters": out_clusters, "cells": cells, "findings": findings}


# --- report -----------------------------------------------------------------

HTML_TEMPLATE = """<!doctype html>
<meta charset="utf-8">
<title>KYD questionnaire review</title>
<style>
:root {
  --surface:#fcfcfb; --page:#f9f9f7; --ink:#0b0b0b; --ink2:#52514e; --muted:#898781;
  --grid:#e1e0d9; --border:rgba(11,11,11,.10);
  --ok:#0ca30c; --note:#9ec5f4; --warn:#fab219; --serious:#ec835a; --critical:#d03b3b;
}
@media (prefers-color-scheme: dark) { :root {
  --surface:#1a1a19; --page:#0d0d0d; --ink:#ffffff; --ink2:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --border:rgba(255,255,255,.10); --note:#1c5cab;
}}
body { margin:0; padding:24px; background:var(--page); color:var(--ink);
       font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif; }
h1 { font-size:20px; margin:0 0 4px; } h2 { font-size:16px; margin:28px 0 8px; }
.sub { color:var(--ink2); margin:0 0 16px; }
.tiles { display:flex; gap:12px; flex-wrap:wrap; margin:16px 0; }
.tile { background:var(--surface); border:1px solid var(--border); border-radius:8px;
        padding:10px 16px; min-width:110px; }
.tile b { display:block; font-size:24px; }
.tile span { color:var(--ink2); font-size:12px; }
.legend { display:flex; gap:16px; flex-wrap:wrap; margin:8px 0 12px; color:var(--ink2); font-size:12px; }
.legend i { display:inline-block; width:14px; height:14px; border-radius:3px; vertical-align:-2px;
            margin-right:5px; border:1px solid var(--border); font-style:normal; text-align:center;
            font-size:10px; line-height:14px; color:#0b0b0b; }
.wrap { overflow-x:auto; background:var(--surface); border:1px solid var(--border); border-radius:8px; padding:8px; }
table.heat { border-collapse:separate; border-spacing:2px; }
table.heat th { font-weight:500; color:var(--ink2); font-size:12px; }
table.heat thead th { height:130px; vertical-align:bottom; padding:0; }
table.heat thead th > div { transform:rotate(-45deg); transform-origin:bottom left;
   width:24px; white-space:nowrap; text-align:left; cursor:pointer; }
table.heat tbody th { text-align:right; padding-right:8px; white-space:nowrap; }
td.cell { width:26px; height:22px; border-radius:4px; text-align:center; cursor:pointer;
          font-size:12px; color:#0b0b0b; border:1px solid var(--border); }
td.cell.sel { outline:2px solid var(--ink); }
td.l0 { background:color-mix(in srgb, var(--ok) 18%, var(--surface)); }
td.l1 { background:var(--note); } td.l2 { background:var(--warn); }
td.l3 { background:var(--serious); } td.l4 { background:var(--critical); color:#fff; }
td.na { background:repeating-linear-gradient(45deg, var(--surface), var(--surface) 3px, var(--grid) 3px, var(--grid) 5px); cursor:default; }
#detail { background:var(--surface); border:1px solid var(--border); border-radius:8px;
          padding:16px; margin-top:16px; }
#detail table { border-collapse:collapse; width:100%; margin:8px 0 16px; }
#detail th, #detail td { text-align:left; padding:4px 10px 4px 0; border-bottom:1px solid var(--grid);
                         vertical-align:top; font-size:13px; }
#detail th { color:var(--muted); font-weight:500; }
tr.hot td { background:color-mix(in srgb, var(--critical) 14%, transparent); }
tr.miss td { background:color-mix(in srgb, var(--serious) 14%, transparent); }
.chip { display:inline-block; padding:1px 8px; border-radius:10px; font-size:11px; color:#0b0b0b;
        border:1px solid var(--border); margin-right:6px; }
.chip.l1{background:var(--note);} .chip.l2{background:var(--warn);}
.chip.l3{background:var(--serious);} .chip.l4{background:var(--critical);color:#fff;}
.muted { color:var(--muted); }
#findings li { margin:4px 0; } #findings b { cursor:pointer; text-decoration:underline; }
</style>
<h1>KYD questionnaire review</h1>
<p class="sub">Questions matched across files by text similarity (ids not trusted). Click a cell or a column header to drill down.</p>
<div class="tiles" id="tiles"></div>
<div class="legend" id="legend"></div>
<div class="wrap"><table class="heat" id="heat"></table></div>
<div id="detail"><span class="muted">Click a heat-map cell (one firm, one question) or a column header (all firms) for details.</span></div>
<h2>All findings</h2>
<ol id="findings"></ol>
<script>
const R = __DATA__;
const GLYPH = {0:"", 1:"\\u2248", 2:"\\u25B2", 3:"!", 4:"\\u2715"};
const NAME = {0:"no issue", 1:"formatting-only wording", 2:"substantive wording / questionnaire-specific",
              3:"missing (question or required answer)", 4:"answer outlier"};
const esc = s => s.replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

// tiles
const counts = {1:0,2:0,3:0,4:0};
R.findings.forEach(f => counts[f.level] !== undefined && counts[f.level]++);
document.getElementById("tiles").innerHTML =
  `<div class="tile"><b>${R.files.length}</b><span>questionnaires</span></div>` +
  `<div class="tile"><b>${R.clusters.length}</b><span>canonical questions</span></div>` +
  [4,3,2,1].map(l => `<div class="tile"><b>${counts[l]}</b><span>${NAME[l]}</span></div>`).join("");

document.getElementById("legend").innerHTML =
  [0,1,2,3,4].map(l => `<span><i class="l${l} cell" style="background:var(--${["ok","note","warn","serious","critical"][l]})">${GLYPH[l]}</i>${NAME[l]}</span>`).join("") +
  `<span><i style="background:repeating-linear-gradient(45deg,var(--surface),var(--surface) 3px,var(--grid) 3px,var(--grid) 5px)"></i>question not expected in this file</span>`;

// heat map
const heat = document.getElementById("heat");
let h = "<thead><tr><th></th>" + R.clusters.map((c,i) =>
  `<th><div onclick="showCluster(${i})" title="${esc(c.title)}">${esc(c.title.slice(0,28))}${c.title.length>28?"\\u2026":""}</div></th>`).join("") + "</tr></thead><tbody>";
R.files.forEach(fn => {
  h += `<tr><th>${esc(fn)}</th>` + R.cells[fn].map((lv,ci) => {
    if (lv < 0) return `<td class="cell na" title="not expected"></td>`;
    const t = `${fn} \\u00d7 ${R.clusters[ci].title.slice(0,60)}: ${NAME[lv]}`;
    return `<td class="cell l${lv}" id="c-${fn}-${ci}" title="${esc(t)}" onclick="showCell('${esc(fn)}',${ci})">${GLYPH[lv]}</td>`;
  }).join("") + "</tr>";
});
heat.innerHTML = h + "</tbody>";

function findingsFor(ci, fn) {
  return R.findings.filter(f => f.cluster === ci && (!fn || f.file === fn));
}
function chips(fs) {
  return fs.length ? fs.map(f => `<div><span class="chip l${f.level}">${f.kind}</span>${!f.__one?`<b>${esc(f.file)}</b> \\u2014 `:""}${esc(f.message)}</div>`).join("")
                   : `<span class="muted">No findings.</span>`;
}
function clusterDetail(ci, focusFile) {
  const c = R.clusters[ci];
  const flagged = new Set(findingsFor(ci).map(f => f.file));
  let s = `<h2 style="margin-top:0">${esc(c.title)}</h2>
    <p class="muted">Present in ${c.coverage}/${R.files.length} questionnaires${c.specific?" (questionnaire-specific)":""}.
    Ids used: ${esc([...new Set(Object.values(c.ids))].join(", "))}</p>`;
  s += `<h3>Question wording variants</h3><table><tr><th>Wording (question + prompts)</th><th>Type</th><th>Files</th></tr>`;
  c.variants.forEach(v => {
    s += `<tr><td>${esc(v.text)}</td><td>${v.kind}</td><td>${esc(v.files.join(", "))}</td></tr>`;
  });
  s += `</table>`;
  c.slots.forEach(sl => {
    s += `<h3>${esc(sl.label)}</h3><table><tr><th>File</th><th>Selection</th><th>Answer</th><th>Flags</th></tr>`;
    sl.entries.forEach(e => {
      const cls = e.flags.some(k => k.includes("missing")) ? "miss" : (e.flags.length ? "hot" : "");
      const focus = e.file === focusFile ? " style='font-weight:700'" : "";
      s += `<tr class="${cls}"><td${focus}>${esc(e.file)}</td><td>${esc(e.selection)}</td><td>${esc(e.answer)||"<span class=muted>(blank)</span>"}</td><td>${esc(e.flags.join(", "))}</td></tr>`;
    });
    s += `</table>`;
  });
  return s;
}
function showCluster(ci) {
  select(null, ci);
  document.getElementById("detail").innerHTML =
    chips(findingsFor(ci)) + "<hr style='border:none;border-top:1px solid var(--grid)'>" + clusterDetail(ci);
}
function showCell(fn, ci) {
  select(fn, ci);
  const fs = findingsFor(ci, fn).map(f => ({...f, __one:true}));
  document.getElementById("detail").innerHTML =
    `<p style="margin-top:0"><b>${esc(fn)}</b> \\u00d7 <b>${esc(R.clusters[ci].title)}</b></p>` +
    chips(fs) + "<hr style='border:none;border-top:1px solid var(--grid)'>" + clusterDetail(ci, fn);
}
function select(fn, ci) {
  document.querySelectorAll("td.sel").forEach(e => e.classList.remove("sel"));
  if (fn) { const el = document.getElementById(`c-${fn}-${ci}`); if (el) el.classList.add("sel"); }
  document.getElementById("detail").scrollIntoView({behavior:"smooth", block:"nearest"});
}
document.getElementById("findings").innerHTML = R.findings.map(f =>
  `<li><span class="chip l${f.level}">${f.kind}</span><b onclick="showCell('${esc(f.file)}',${f.cluster})">${esc(f.file)}</b>
   \\u00d7 ${esc(R.clusters[f.cluster].title.slice(0,60))} \\u2014 ${esc(f.message)}</li>`).join("");
</script>
"""


def render_html(report: dict) -> str:
    return HTML_TEMPLATE.replace("__DATA__", json.dumps(report, ensure_ascii=False))


XLSX_FILLS = {OK: "DFF2DF", NOTE: "9EC5F4", WARN: "FAB219", MISSING: "EC835A", OUTLIER: "D03B3B"}


def render_xlsx(report: dict, path: Path) -> None:
    """Same content as the HTML, flat, for auditors who work in Excel."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    bold = Font(bold=True)

    def sheet(title, headers, rows, widths, fills=None):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for c in ws[1]:
            c.font = bold
        for ri, row in enumerate(rows, start=2):
            ws.append(row)
            for ci, fill in (fills(ri - 2) if fills else {}).items():
                ws.cell(row=ri, column=ci).fill = PatternFill("solid", start_color=fill)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "B2"
        return ws

    clusters = report["clusters"]
    titles = [c["title"] for c in clusters]

    ws = sheet("Heatmap", ["File"] + titles,
               [[fn] + [("n/a" if lv < 0 else LEVEL_NAMES[lv]) for lv in report["cells"][fn]]
                for fn in report["files"]],
               [22] + [18] * len(titles),
               fills=lambda r: {ci + 2: XLSX_FILLS[lv]
                                for ci, lv in enumerate(report["cells"][report["files"][r]]) if lv >= 0})
    for c in ws[1][1:]:
        c.alignment = Alignment(wrap_text=True, vertical="top")

    sheet("Findings", ["Severity", "Type", "File", "Question", "Detail"],
          [[LEVEL_NAMES[f["level"]], f["kind"], f["file"], titles[f["cluster"]], f["message"]]
           for f in report["findings"]],
          [16, 22, 22, 45, 90],
          fills=lambda r: {1: XLSX_FILLS[report["findings"][r]["level"]]})

    qrows = [(c, v) for c in clusters for v in c["variants"]]
    sheet("Questions", ["Question", "Coverage", "Specific?", "Variant type", "Wording (question + prompts)", "Files"],
          [[c["title"], f"{c['coverage']}/{len(report['files'])}", "yes" if c["specific"] else "",
            v["kind"], v["text"], ", ".join(sorted(v["files"]))] for c, v in qrows],
          [45, 10, 9, 14, 70, 40])

    arows = [(c, sl, e) for c in clusters for sl in c["slots"] for e in sl["entries"]]
    sheet("Answers", ["Question", "Sub-question", "File", "Id", "Selection", "Answer", "Flags"],
          [[c["title"], sl["label"], e["file"], c["ids"].get(e["file"], ""),
            e["selection"], e["answer"], ", ".join(e["flags"])] for c, sl, e in arows],
          [45, 40, 22, 6, 22, 70, 30])

    del wb["Sheet"]
    wb.save(path)


def main() -> None:
    import sys
    if hasattr(sys.stdout, "reconfigure"):  # Windows consoles default to a legacy codepage
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("jsons", nargs="*", help="parsed questionnaire JSON files")
    ap.add_argument("--dir", help="review every *.json in this directory")
    ap.add_argument("-o", "--out", default="output/kyd_review.html")
    args = ap.parse_args()

    paths = [Path(p) for p in args.jsons]
    if args.dir:
        paths += sorted(Path(args.dir).glob("*.json"))
    if len(paths) < 2:
        ap.error("need at least 2 JSON files (pass paths or --dir)")

    report = analyze(load_files(paths))
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render_html(report), encoding="utf-8")
    xlsx = out.with_suffix(".xlsx")
    render_xlsx(report, xlsx)

    counts = Counter(g["level"] for g in report["findings"])
    print(f"{len(report['files'])} questionnaires, {len(report['clusters'])} canonical questions")
    for lv in (OUTLIER, MISSING, WARN, NOTE):
        print(f"  {LEVEL_NAMES[lv]:>18}: {counts.get(lv, 0)}")
    for g in report["findings"]:
        if g["level"] >= MISSING:
            print(f"  [{LEVEL_NAMES[g['level']]}] {g['file']} x {report['clusters'][g['cluster']]['title'][:50]}: {g['message']}")
    print(f"report -> {out} and {xlsx}")


if __name__ == "__main__":
    main()
