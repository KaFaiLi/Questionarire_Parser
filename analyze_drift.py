"""Cluster questionnaire wordings and flag question drift across parsed JSONs.

Embeds every template string (question text, sub-question prompt, option_label) with
Azure embeddings via LangChain, clusters semantically-equal variants, derives a canonical
wording per cluster (most frequent), and flags drift. Output: one Excel tab, one row per
(file, variant) so you can see which file carries which wording.

    uv run python analyze_drift.py --in output/llm_parsed_full2 -o output/drift_analysis.xlsx

Multiple folders (each analysed independently — one report per folder)::

    uv run python analyze_drift.py --in 2024Q4/ 2025Q1/ 2025Q2/ --out-dir output/

Pool every folder into ONE combined report (single shared baseline; file ids
prefixed with their source folder as '<source>/<file>')::

    uv run python analyze_drift.py --in 2024Q4/ 2025Q1/ 2025Q2/ --combine -o output/combined.xlsx

Config (endpoint / key / embedding deployment / thresholds) comes from config.yaml.
Tune --cluster-threshold (merges wordings into one logical question) and --drift-threshold
(flags a variant as real drift) by reviewing the Excel.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import yaml

import analyze_kyd_json as akj  # reuse _plan_outputs for consistent multi-input output naming
import answer_review

FIELDS_TEXT = {"prompt": "prompt", "option_label": "option"}  # sub-question field -> level

DEFAULT_LLM_CFG = {"deployment": "gpt-4.1-nano", "min_severity": "low", "max_answers_per_slot": 40}


def load_config(path: str = "config.yaml") -> dict:
    return yaml.safe_load(akj.read_text(path))


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def make_embeddings(cfg: dict):
    from langchain_openai import AzureOpenAIEmbeddings
    az = cfg["azure"]
    return AzureOpenAIEmbeddings(
        azure_endpoint=az["endpoint"],
        azure_deployment=az["embedding_deployment"],
        api_version=az["api_version"],
        api_key=akj.azure_api_key(az),
    )


def extract_units(in_dir: Path) -> list[dict]:
    """One unit per non-empty template string (question / prompt / option_label)."""
    units = []
    for p in akj.iter_json(Path(in_dir)):
        data = json.loads(akj.read_text(p))
        fname = data.get("file_name", p.stem)
        for q in data.get("questions", []):
            qid = q.get("question_id", "")
            if (q.get("question") or "").strip():
                units.append({"file_name": fname, "question_id": qid, "sub_idx": None,
                              "level": "question", "text": q["question"]})
            for i, s in enumerate(q.get("sub_questions", [])):
                for field, level in FIELDS_TEXT.items():
                    if (s.get(field) or "").strip():
                        units.append({"file_name": fname, "question_id": qid, "sub_idx": i,
                                      "level": level, "text": s[field]})
    return units


def extract_answers(in_dir) -> list[dict]:
    """One record per answered sub-question. anchor = option_label, else prompt, else question
    text (raw, matching extract_units so the cluster lookup hits). response = answer or selection."""
    recs = []
    for p in akj.iter_json(Path(in_dir)):
        d = json.loads(akj.read_text(p))
        fn = d.get("file_name", p.stem)
        for q in d.get("questions", []):
            qid = q.get("question_id", "")
            qtext = q.get("question", "")
            for i, s in enumerate(q.get("sub_questions", [])):
                resp = (s.get("answer") or "").strip() or (s.get("selection") or "").strip()
                if not resp:
                    continue
                if (s.get("option_label") or "").strip():
                    alevel, atext = "option", s["option_label"]
                elif (s.get("prompt") or "").strip():
                    alevel, atext = "prompt", s["prompt"]
                elif qtext.strip():
                    alevel, atext = "question", qtext
                else:
                    continue  # nothing to anchor on
                recs.append({"file_name": fn, "question_id": qid, "sub_idx": i,
                             "anchor_level": alevel, "anchor_text": atext, "response": resp})
    return recs


def group_by_slot(answer_recs, assign) -> dict:
    slots = {}
    for r in answer_recs:
        key = (r["anchor_level"], r["anchor_text"])
        if key not in assign:
            continue
        r["slot_id"] = assign[key]
        slots.setdefault(r["slot_id"], []).append(r)
    return slots


def detect_outliers(slots, vecs, cfg) -> list[dict]:
    """Per slot, vote across freq / minority-cluster / centroid-z. Flag if votes >= min_votes."""
    import numpy as np
    from collections import Counter

    out = []
    for sid, recs in slots.items():
        n = len(recs)
        if n < cfg["min_samples"]:
            continue
        responses = [r["response"] for r in recs]
        # --- freq (on normalized exact value) ---
        normed = [_norm(x) for x in responses]
        fc = Counter(normed)
        maj_exists = (fc.most_common(1)[0][1] / n) >= 0.5
        # --- minority cluster (embeddings) ---
        # local maps each DISTINCT response -> cluster id; weight by occurrence for true sizes
        local = cluster(responses, vecs, cfg["answer_threshold"])
        csize = Counter(local[x] for x in responses)
        cluster_dom = (max(csize.values()) / n) >= 0.5
        # --- centroid z-score ---
        M = np.array([vecs[x] for x in responses], dtype=float)
        M /= (np.linalg.norm(M, axis=1, keepdims=True) + 1e-12)
        cen = M.mean(axis=0)
        cen /= (np.linalg.norm(cen) + 1e-12)
        sims = M @ cen
        sd = float(sims.std())
        z = (sims - sims.mean()) / sd if sd > 1e-9 else np.zeros(n)
        for i, r in enumerate(recs):
            # strict '<': a share exactly == minority_frac (e.g. 2/10 at 0.2) is NOT a minority.
            # Raise minority_frac if you want such borderline splits flagged.
            f = maj_exists and (fc[normed[i]] / n < cfg["minority_frac"])
            cl = cluster_dom and (csize[local[responses[i]]] / n < cfg["minority_frac"])
            ce = bool(z[i] < -cfg["z_k"])
            votes = int(f) + int(cl) + int(ce)
            if votes >= cfg["min_votes"]:
                fired = "|".join(m for m, v in (("freq", f), ("cluster", cl), ("centroid", ce)) if v)
                out.append({
                    "slot_id": sid, "question_id": r["question_id"], "file_name": r["file_name"],
                    "response": r["response"], "votes": votes, "methods_fired": fired,
                    "freq_share": round(fc[normed[i]] / n, 3),
                    "cluster_share": round(csize[local[responses[i]]] / n, 3),
                    "centroid_z": round(float(z[i]), 3),
                })
    return out


def flag_questionnaires(outliers, all_files, n) -> list[dict]:
    from collections import defaultdict
    byf = defaultdict(list)
    for o in outliers:
        byf[o["file_name"]].append(o["question_id"])
    rows = []
    for fn in all_files:
        qs = sorted(set(byf.get(fn, [])))
        rows.append({"file_name": fn, "n_outliers": len(byf.get(fn, [])),
                     "flagged": len(byf.get(fn, [])) >= n, "questions": ",".join(qs)})
    return sorted(rows, key=lambda r: -r["n_outliers"])


def embed_texts(texts: list[str], embeddings, workers: int, cache_path: Path) -> dict[str, list[float]]:
    """Embed each distinct text once, multithreaded. Cached to cache_path (text->vector JSON)
    so repeat runs are reproducible and only new strings hit the API."""
    from concurrent.futures import ThreadPoolExecutor
    from tqdm import tqdm

    vecs: dict[str, list[float]] = {}
    if akj.path_exists(cache_path):
        vecs = json.loads(akj.read_text(cache_path))
    uniq = sorted(set(texts))
    todo = [t for t in uniq if t not in vecs]
    if todo:
        batch = 64
        batches = [todo[i:i + batch] for i in range(0, len(todo), batch)]
        with ThreadPoolExecutor(max_workers=workers) as ex:
            for b, res in tqdm(zip(batches, ex.map(embeddings.embed_documents, batches)),
                               total=len(batches), desc="embedding"):
                vecs.update(zip(b, res))
        akj.write_text(cache_path, json.dumps(vecs))
    print(f"embeddings: {len(todo)} new, {len(uniq) - len(todo)} cached  ({cache_path})")
    return {t: vecs[t] for t in uniq}


def cluster(texts: list[str], vecs: dict[str, list[float]], threshold: float) -> dict[str, int]:
    """Single-linkage union-find on cosine sim. Returns text -> cluster id (per call)."""
    import numpy as np

    uniq = sorted(set(texts))
    if not uniq:
        return {}
    M = np.array([vecs[t] for t in uniq], dtype=float)
    M /= (np.linalg.norm(M, axis=1, keepdims=True) + 1e-12)
    sim = M @ M.T

    parent = list(range(len(uniq)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    # ponytail: numpy union-find single-linkage, small N. Swap to sklearn AgglomerativeClustering if chaining becomes a problem.
    for i in range(len(uniq)):
        for j in range(i + 1, len(uniq)):
            if sim[i, j] >= threshold:
                parent[find(i)] = find(j)
    return {t: find(i) for i, t in enumerate(uniq)}


def cos(a, b) -> float:
    import numpy as np
    a, b = np.array(a, dtype=float), np.array(b, dtype=float)
    return float(a @ b / ((np.linalg.norm(a) * np.linalg.norm(b)) + 1e-12))


def assign_clusters(units: list[dict], vecs: dict, ct: float):
    """One clustering pass. Returns (assign, canon):
    assign: (level, text) -> global cluster id; canon: cid -> canonical (most-frequent) text."""
    from collections import Counter
    assign, canon = {}, {}
    next_cid = 0
    for level in sorted({u["level"] for u in units}):
        lvl_units = [u for u in units if u["level"] == level]
        local = cluster([u["text"] for u in lvl_units], vecs, ct)
        groups = {}
        for u in lvl_units:
            groups.setdefault(local[u["text"]], []).append(u)
        for _, gunits in sorted(groups.items()):
            cid = next_cid
            next_cid += 1
            counts = Counter(u["text"] for u in gunits)
            canon[cid] = sorted(counts, key=lambda t: (-counts[t], -len(t), t))[0]
            for u in gunits:
                assign[(level, u["text"])] = cid
    return assign, canon


def near_miss_pairs(units: list[dict], assign: dict, canon: dict, vecs: dict,
                    ct: float, floor: float) -> list[dict]:
    """Same-level cluster pairs whose canonicals sit just below ct (in [floor, ct)).
    Surfaces logical questions the hard threshold wrongly split into two clusters so a
    human can force a merge via --merge-map. Sorted closest-to-merging first."""
    from collections import Counter
    meta = {}  # cid -> {level, files:set, n}
    for u in units:
        cid = assign[(u["level"], u["text"])]
        m = meta.setdefault(cid, {"level": u["level"], "files": set(), "n": 0})
        m["files"].add(u["file_name"]); m["n"] += 1
    cids = sorted(meta)
    rows = []
    for a_i in range(len(cids)):
        for b_i in range(a_i + 1, len(cids)):
            ca, cb = cids[a_i], cids[b_i]
            if meta[ca]["level"] != meta[cb]["level"]:
                continue
            c = cos(vecs[canon[ca]], vecs[canon[cb]])
            if floor <= c < ct:
                rows.append({"level": meta[ca]["level"], "cosine": round(c, 4),
                             "cluster_a": ca, "cluster_b": cb,
                             "canon_a": canon[ca], "canon_b": canon[cb],
                             "files_a": len(meta[ca]["files"]), "files_b": len(meta[cb]["files"]),
                             "n_a": meta[ca]["n"], "n_b": meta[cb]["n"]})
    return sorted(rows, key=lambda r: -r["cosine"])


def apply_merge_map(units, assign, canon, merges: list[list[str]]):
    """Force-union clusters listed in the merge-map (each entry = a group of canonical
    texts that mean the same question). Keyed on text, not cluster id, because ids
    renumber every run. Returns fresh (assign, canon) with groups collapsed and each
    merged cluster's canonical recomputed as its most-frequent wording."""
    from collections import Counter
    parent = {}
    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]; x = parent[x]
        return x
    def union(a, b):
        parent[find(a)] = find(b)
    # map each merge-group text to every cid it appears under, then union them
    text_cids = {}
    for (level, text), cid in assign.items():
        text_cids.setdefault(text, set()).add(cid)
    for group in merges:
        cids = [c for t in group for c in text_cids.get(t, ())]
        for c in cids[1:]:
            union(cids[0], c)
    new_assign = {k: find(v) for k, v in assign.items()}
    groups = {}
    for u in units:
        groups.setdefault(new_assign[(u["level"], u["text"])], []).append(u)
    new_canon = {}
    for cid, gunits in groups.items():
        counts = Counter(u["text"] for u in gunits)
        new_canon[cid] = sorted(counts, key=lambda t: (-counts[t], -len(t), t))[0]
    return new_assign, new_canon


def build_rows(units: list[dict], vecs: dict, assign: dict, canon: dict, drift_threshold: float) -> list[dict]:
    from collections import Counter
    by_cid = {}
    for u in units:
        by_cid.setdefault(assign[(u["level"], u["text"])], []).append(u)
    rows = []
    for cid, gunits in sorted(by_cid.items()):
        canonical = canon[cid]
        counts = Counter(u["text"] for u in gunits)
        dom_qid = Counter(u["question_id"] for u in gunits).most_common(1)[0][0]
        level = gunits[0]["level"]
        for u in gunits:
            is_canon = u["text"] == canonical
            sim_c = 1.0 if is_canon else cos(vecs[u["text"]], vecs[canonical])
            rows.append({
                "cluster_id": cid, "level": level, "question_id": dom_qid,
                "file_name": u["file_name"], "variant_text": u["text"],
                "is_canonical": is_canon,
                "is_drift": (not is_canon) and sim_c < drift_threshold,
                "occurrences": counts[u["text"]], "cosine_to_canonical": round(sim_c, 4),
            })
    return rows


def _variant_meta(df):
    """Per cluster, order distinct variants canonical-first and tag status + a Vn id."""
    import pandas as pd
    meta = {}  # (cluster_id, text) -> {"vid","status","occ","sim","files"}
    for cid, sub in df.groupby("cluster_id"):
        order = (sub.drop_duplicates("variant_text")
                    .sort_values(["is_canonical", "occurrences"], ascending=[False, False]))
        for vid, (_, r) in enumerate(order.iterrows()):
            status = "canon" if r["is_canonical"] else ("drift" if r["is_drift"] else "near")
            files = sorted(sub[sub["variant_text"] == r["variant_text"]]["file_name"])
            meta[(cid, r["variant_text"])] = {
                "vid": vid, "status": status, "occ": int(r["occurrences"]),
                "sim": float(r["cosine_to_canonical"]), "files": files}
    return meta


def make_matrix(df, path, ct, dt):
    """Plotly heatmap: files (rows) x question clusters (cols), colored by drift status.
    Cell text = Vn variant id; hover = full wording + sim + status."""
    import plotly.graph_objects as go

    meta = _variant_meta(df)
    STATUS_Z = {"canon": 0, "near": 1, "drift": 2}

    qdf = df[df["level"] == "question"]
    cols = (qdf.drop_duplicates("cluster_id")
               .sort_values(["question_id", "cluster_id"])[["cluster_id", "question_id"]]
               .values.tolist())
    files = sorted(df["file_name"].unique())
    xlabels = [f"{q} <span style='color:#999'>(c{cid})</span>" for cid, q in cols]

    z, text, hover = [], [], []
    for f in files:
        zr, tr, hr = [], [], []
        for cid, q in cols:
            r = qdf[(qdf["cluster_id"] == cid) & (qdf["file_name"] == f)]
            if r.empty:
                zr.append(None); tr.append(""); hr.append("")
                continue
            txt = r.iloc[0]["variant_text"]
            m = meta[(cid, txt)]
            zr.append(STATUS_Z[m["status"]])
            tr.append(f"V{m['vid']}")
            hr.append(f"<b>{f}</b> · {q}<br>V{m['vid']} ({m['status']}, sim {m['sim']:.2f})"
                      f"<br>{txt}")
        z.append(zr); text.append(tr); hover.append(hr)

    # discrete green/amber/red over z in {0,1,2}
    colorscale = [[0.0, "#2e7d32"], [0.33, "#2e7d32"],
                  [0.34, "#f9a825"], [0.66, "#f9a825"],
                  [0.67, "#c62828"], [1.0, "#c62828"]]
    # in-cell Vn text gets unreadable past ~40 rows; rely on hover at scale
    show_text = len(files) <= 40 and len(cols) <= 40
    fig = go.Figure(go.Heatmap(
        z=z, x=xlabels, y=files,
        text=(text if show_text else None),
        texttemplate=("%{text}" if show_text else None),
        customdata=hover, hovertemplate="%{customdata}<extra></extra>",
        zmin=0, zmax=2, colorscale=colorscale,
        xgap=1, ygap=1, textfont={"size": 10, "color": "#fff"},
        colorbar={"title": {"text": "status", "side": "right"},
                  "tickvals": [0.33, 1.0, 1.67], "ticktext": ["canonical", "minor", "drift"],
                  "thickness": 16, "len": 0.4, "y": 1, "yanchor": "top"}))

    n_drift = int(df["is_drift"].sum())
    # fixed cell size -> page scrolls at scale (100 files x 30+ questions) instead of squashing
    cell_w, cell_h = 46, 22
    width = 260 + cell_w * len(cols)
    height = 200 + cell_h * len(files)
    fig.update_layout(
        title={"text": (f"<b>Question drift</b> — {len(files)} files · "
                        f"{df['cluster_id'].nunique()} clusters · {n_drift} drift occurrences"
                        f"<br><sub>cell = wording each file used · hover for full text · "
                        f"cluster_threshold={ct} drift_threshold={dt}</sub>"),
               "x": 0, "xanchor": "left", "y": 0.985, "yanchor": "top", "font": {"size": 16}},
        xaxis={"side": "top", "tickangle": -45, "tickfont": {"size": 11},
               "showgrid": False, "ticks": "", "constrain": "domain"},
        yaxis={"autorange": "reversed", "tickfont": {"size": 11},
               "showgrid": False, "ticks": ""},
        width=width, height=height,
        margin={"l": 200, "r": 80, "t": 170, "b": 20},
        plot_bgcolor="white", font={"family": "system-ui,sans-serif"})
    fig.write_html(akj.longpath(path), include_plotlyjs="cdn",
                   full_html=True, default_width=f"{width}px", default_height=f"{height}px")


def make_outlier_matrix(df_out, df_flags, slots, canon, path, ocfg):
    """Heatmap: files (rows) x answer-slots (cols). green=answered, red=outlier, blank=missing."""
    import plotly.graph_objects as go
    from collections import Counter

    slot_ids = sorted(slots)
    dom_qid = {sid: Counter(r["question_id"] for r in slots[sid]).most_common(1)[0][0]
               for sid in slot_ids}
    files = sorted(df_flags["file_name"])
    flagged = set(df_flags[df_flags["flagged"]]["file_name"])
    outset = {(r.file_name, r.slot_id) for r in df_out.itertuples()}
    resp = {(r["file_name"], sid): r["response"]
            for sid in slot_ids for r in slots[sid]}

    xlabels = [f"{dom_qid[sid]} <span style='color:#999'>(s{sid})</span>" for sid in slot_ids]
    z, hover = [], []
    for f in files:
        zr, hr = [], []
        for sid in slot_ids:
            if (f, sid) not in resp:
                zr.append(None); hr.append("")
            elif (f, sid) in outset:
                zr.append(2); hr.append(f"<b>{f}</b> · {dom_qid[sid]}<br>OUTLIER: {resp[(f, sid)]}")
            else:
                zr.append(0); hr.append(f"{f} · {dom_qid[sid]}<br>{resp[(f, sid)]}")
        z.append(zr); hover.append(hr)

    ylabels = [("⚑ " + f) if f in flagged else f for f in files]
    colorscale = [[0.0, "#2e7d32"], [0.5, "#2e7d32"], [0.5, "#c62828"], [1.0, "#c62828"]]
    fig = go.Figure(go.Heatmap(
        z=z, x=xlabels, y=ylabels, customdata=hover,
        hovertemplate="%{customdata}<extra></extra>",
        zmin=0, zmax=2, colorscale=colorscale, showscale=False, xgap=1, ygap=1))
    cell_w, cell_h = 60, 22
    width = 260 + cell_w * len(slot_ids)
    height = 200 + cell_h * len(files)
    n_flag = len(flagged)
    fig.update_layout(
        title={"text": (f"<b>Answer outliers</b> — {n_flag} flagged questionnaire(s) "
                        f"(≥{ocfg['multi_outlier_n']} outliers)<br>"
                        f"<sub>red = outlier answer · ⚑ = flagged file · hover for value</sub>"),
               "x": 0, "xanchor": "left", "y": 0.985, "yanchor": "top", "font": {"size": 16}},
        xaxis={"side": "top", "tickangle": -45, "tickfont": {"size": 11}, "showgrid": False, "ticks": ""},
        yaxis={"autorange": "reversed", "tickfont": {"size": 11}, "showgrid": False, "ticks": ""},
        width=width, height=height, margin={"l": 200, "r": 80, "t": 170, "b": 20},
        plot_bgcolor="white", font={"family": "system-ui,sans-serif"})
    fig.write_html(akj.longpath(path), include_plotlyjs="cdn",
                   full_html=True, default_width=f"{width}px", default_height=f"{height}px")


def extract_units_many(labeled: list[tuple[Path, str]]) -> list[dict]:
    """Pool units across folders; namespace file_name as '<source>/<file>'."""
    units = []
    for in_dir, label in labeled:
        for u in extract_units(Path(in_dir)):
            u["file_name"] = f"{label}/{u['file_name']}"
            units.append(u)
    return units


def extract_answers_many(labeled: list[tuple[Path, str]]) -> list[dict]:
    recs = []
    for in_dir, label in labeled:
        for r in extract_answers(Path(in_dir)):
            r["file_name"] = f"{label}/{r['file_name']}"
            recs.append(r)
    return recs


def analyze_drift_one(in_dir: Path, out: Path, *, cfg, ct, dt, ocfg, nm_floor, merges,
                      embeddings, cache_path: Path, workers: int, llm_cfg=None) -> bool:
    """Per-folder run: extract this folder only, then build the report."""
    units = extract_units(in_dir)
    answer_recs = extract_answers(in_dir)
    return _drift_report(units, answer_recs, out, label=str(in_dir), cfg=cfg, ct=ct, dt=dt,
                         ocfg=ocfg, nm_floor=nm_floor, merges=merges,
                         embeddings=embeddings, cache_path=cache_path, workers=workers,
                         llm_cfg=llm_cfg)


def analyze_drift_combined(labeled: list[tuple[Path, str]], out: Path, *, cfg, ct, dt,
                           ocfg, nm_floor, merges, embeddings, cache_path: Path, workers: int,
                           llm_cfg=None) -> bool:
    """Pool every folder into ONE population (source-tagged file ids) and build a
    single combined report."""
    units = extract_units_many(labeled)
    answer_recs = extract_answers_many(labeled)
    return _drift_report(units, answer_recs, out,
                         label=f"combined ({len(labeled)} folders)", cfg=cfg, ct=ct, dt=dt,
                         ocfg=ocfg, nm_floor=nm_floor, merges=merges,
                         embeddings=embeddings, cache_path=cache_path, workers=workers,
                         llm_cfg=llm_cfg)


def run_llm_review_stage(slots, det_outliers, canon, *, cfg, llm_cfg):
    """Build the llm_answer_review DataFrame. Returns an empty frame when the
    feature is off (llm_cfg is None). Best-effort otherwise: any failure warns
    and returns an empty frame so deterministic output is never lost."""
    import pandas as pd
    if llm_cfg is None:
        return pd.DataFrame(columns=answer_review.LLM_COLS)
    import audit_kyd  # lazy: audit_kyd imports analyze_drift, avoid circular top-level import
    try:
        chat = answer_review.make_chat(cfg, llm_cfg)
        rules = audit_kyd.load_rules(None)
        rmap = answer_review.risk_flags(slots, rules, canon)
        susp = answer_review.suspicious_slots(det_outliers, rmap)
        verdicts = answer_review.llm_review(slots, susp, canon, chat, llm_cfg)
        return answer_review.build_llm_review_df(verdicts, det_outliers, rmap, canon, llm_cfg)
    except (Exception, SystemExit) as exc:  # noqa: BLE001 - SystemExit: azure_api_key raises this on missing key
        print(f"WARN: llm-review unavailable ({exc}); skipping.", file=sys.stderr)
        return pd.DataFrame(columns=answer_review.LLM_COLS)


def _drift_report(units: list[dict], answer_recs: list[dict], out: Path, *, label, cfg,
                  ct, dt, ocfg, nm_floor, merges, embeddings, cache_path: Path, workers: int,
                  llm_cfg=None) -> bool:
    """Embed -> cluster -> drift rows + outliers -> xlsx + two HTML matrices,
    for units/answers from one folder or several pooled together. The
    consensus canonical wording is computed over whatever is passed in."""
    import pandas as pd

    if not units:
        print(f"  SKIP {label}: no template strings found", file=sys.stderr)
        return False
    print(f"\n=== {label} ===")
    print(f"{len(units)} units from {len(set(u['file_name'] for u in units))} files; "
          f"{len(set(u['text'] for u in units))} distinct strings")

    all_texts = [u["text"] for u in units] + [r["response"] for r in answer_recs]
    vecs = embed_texts(all_texts, embeddings, workers, cache_path)

    assign, canon = assign_clusters(units, vecs, ct)
    near_miss = near_miss_pairs(units, assign, canon, vecs, ct, nm_floor)
    if merges:
        assign, canon = apply_merge_map(units, assign, canon, merges)
    rows = build_rows(units, vecs, assign, canon, dt)

    slots = group_by_slot(answer_recs, assign)
    outliers = detect_outliers(slots, vecs, ocfg)
    all_files = sorted({r["file_name"] for r in answer_recs})
    flags = flag_questionnaires(outliers, all_files, ocfg["multi_outlier_n"])

    df = pd.DataFrame(rows).sort_values(
        ["cluster_id", "is_canonical", "file_name"], ascending=[True, False, True])
    akj.ensure_parent(out)

    OUT_COLS = ["slot_id", "question_id", "file_name", "response", "votes",
                "methods_fired", "freq_share", "cluster_share", "centroid_z"]
    df_out = pd.DataFrame(outliers, columns=OUT_COLS).sort_values(
        ["file_name", "question_id"]) if outliers else pd.DataFrame(columns=OUT_COLS)
    df_flags = pd.DataFrame(flags, columns=["file_name", "n_outliers", "flagged", "questions"])

    df_llm = run_llm_review_stage(slots, outliers, canon, cfg=cfg, llm_cfg=llm_cfg)

    NM_COLS = ["level", "cosine", "cluster_a", "cluster_b", "canon_a", "canon_b",
               "files_a", "files_b", "n_a", "n_b"]
    df_nm = pd.DataFrame(near_miss, columns=NM_COLS)

    with pd.ExcelWriter(akj.longpath(out)) as xl:
        df.to_excel(xl, sheet_name="drift", index=False)
        df_out.to_excel(xl, sheet_name="answer_outliers", index=False)
        df_flags.to_excel(xl, sheet_name="questionnaire_flags", index=False)
        df_nm.to_excel(xl, sheet_name="suspected_merges", index=False)
        df_llm.to_excel(xl, sheet_name="llm_answer_review", index=False)

    html_out = out.with_suffix(".html")
    # derive the outlier-matrix name from this report's stem so multiple folders
    # never collide on a hardcoded filename.
    outlier_html = out.with_name(f"{out.stem}_outliers.html")
    make_matrix(df, html_out, ct, dt)
    make_outlier_matrix(df_out, df_flags, slots, canon, outlier_html, ocfg)

    n_clusters = df["cluster_id"].nunique()
    n_drift = int(df["is_drift"].sum())
    print(f"clusters: {n_clusters}  |  drift rows: {n_drift}  "
          f"(cluster_threshold={ct}, drift_threshold={dt})")
    print(f"suspected_merges: {len(near_miss)} same-level pairs in [{nm_floor}, {ct}) "
          f"-- review the sheet, add real splits to --merge-map")
    print(f"outliers: {len(outliers)} answers  |  flagged questionnaires: "
          f"{int(df_flags['flagged'].sum())}/{len(df_flags)} (min_votes={ocfg['min_votes']}, "
          f"multi_outlier_n={ocfg['multi_outlier_n']})")
    if llm_cfg is not None:
        print(f"llm-review: {len(df_llm)} flagged answers (min_severity={llm_cfg['min_severity']})")
    print(f"wrote {out}\nwrote {html_out}\nwrote {outlier_html}")
    return True


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--in", dest="in_dirs", nargs="+", default=["output/llm_parsed_full2"],
                    help="one or more dirs of parsed *.json (each analysed independently)")
    ap.add_argument("-o", "--out", default=None,
                    help="output Excel file. Single input only; for many inputs use --out-dir.")
    ap.add_argument("--out-dir", default=None,
                    help="write every report into this dir as <input>_drift_analysis.xlsx.")
    ap.add_argument("--combine", action="store_true",
                    help="pool ALL inputs into one population and write a single combined "
                         "report (file ids prefixed with their source folder).")
    ap.add_argument("--workers", type=int, default=8, help="concurrent embedding requests")
    ap.add_argument("--cache", default="output/.vec_cache.json", help="vector cache (text->embedding)")
    ap.add_argument("--cluster-threshold", type=float, help="override config drift.cluster_threshold")
    ap.add_argument("--drift-threshold", type=float, help="override config drift.drift_threshold")
    ap.add_argument("--min-votes", type=int, help="override outliers.min_votes")
    ap.add_argument("--multi-outlier-n", type=int, help="override outliers.multi_outlier_n")
    ap.add_argument("--answer-threshold", type=float, help="override outliers.answer_threshold")
    ap.add_argument("--near-miss-floor", type=float, help="override drift.near_miss_floor")
    ap.add_argument("--merge-map", default=None,
                    help="YAML with a `merges:` list of canonical-text groups to force-union")
    ap.add_argument("--llm-review", action="store_true",
                    help="LLM pass over suspicious slots for peer/intrinsic-risk outliers (needs config.yaml azure creds)")
    args = ap.parse_args()

    cfg = load_config()
    ct = args.cluster_threshold if args.cluster_threshold is not None else cfg["drift"]["cluster_threshold"]
    dt = args.drift_threshold if args.drift_threshold is not None else cfg["drift"]["drift_threshold"]
    nm_floor = (args.near_miss_floor if args.near_miss_floor is not None
                else cfg["drift"].get("near_miss_floor", 0.75))
    merges = yaml.safe_load(akj.read_text(args.merge_map)).get("merges", []) if args.merge_map else []
    ocfg = dict(cfg["outliers"])
    if args.min_votes is not None: ocfg["min_votes"] = args.min_votes
    if args.multi_outlier_n is not None: ocfg["multi_outlier_n"] = args.multi_outlier_n
    if args.answer_threshold is not None: ocfg["answer_threshold"] = args.answer_threshold
    llm_cfg = {**DEFAULT_LLM_CFG, **cfg.get("llm_review", {})} if args.llm_review else None

    inputs = [Path(p) for p in args.in_dirs]
    missing = [p for p in inputs if not akj.path_exists(p)]
    if missing:
        for p in missing:
            print(f"Input not found: {p}", file=sys.stderr)
        sys.exit(1)
    if args.out and len(inputs) > 1 and not args.combine:
        ap.error("-o/--out takes a single input; use --out-dir, or --combine for one pooled report.")
    if args.out_dir:
        Path(args.out_dir).mkdir(parents=True, exist_ok=True)

    embeddings = make_embeddings(cfg)
    cache_path = Path(args.cache)

    if args.combine:
        labeled = akj.unique_sources(inputs)
        print(f"\n=== combined: pooling {len(inputs)} input(s) "
              f"[{', '.join(l for _, l in labeled)}] ===")
        out = (Path(args.out) if args.out
               else (Path(args.out_dir) if args.out_dir else Path("."))
               / "combined_drift_analysis.xlsx")
        ok = analyze_drift_combined(labeled, out, cfg=cfg, ct=ct, dt=dt, ocfg=ocfg,
                                    nm_floor=nm_floor, merges=merges,
                                    embeddings=embeddings, cache_path=cache_path,
                                    workers=args.workers, llm_cfg=llm_cfg)
        print(f"\n[DONE] {int(ok)} combined report written.")
        if ok:
            print(f"  - {out.resolve()}")
        sys.exit(0 if ok else 1)

    written: list[Path] = []
    for in_dir, out in akj._plan_outputs(inputs, args.out, args.out_dir, "drift_analysis.xlsx"):
        if analyze_drift_one(in_dir, out, cfg=cfg, ct=ct, dt=dt, ocfg=ocfg,
                             nm_floor=nm_floor, merges=merges,
                             embeddings=embeddings, cache_path=cache_path, workers=args.workers,
                             llm_cfg=llm_cfg):
            written.append(out)

    print(f"\n[DONE] {len(written)}/{len(inputs)} report(s) written.")
    for w in written:
        print(f"  - {w.resolve()}")
    if not written:
        sys.exit(1)


if __name__ == "__main__":
    main()
