# KYD example dataset

10 fictional "Know Your Distributor" (KYD) questionnaires for structured-product
distributors, in the nested **final format**
(`questions[] -> sub_questions[]` with `option_label` / `prompt` / `selection` /
`answer`). One JSON file per firm.

These files are **generated** by `../../gen_kyd_examples.py` — edit the scenario
table there and re-run, don't hand-edit the JSON. They exist to exercise every
review goal in `../../analyze_kyd_json.py`.

## What each file contributes

| File | Firm | Planted scenario(s) |
|------|------|---------------------|
| `01_bea_hk.json` | The Bank of East Asia (HK) | baseline; Q1 "…and Address" wording |
| `02_pacific_trust_sg.json` | Pacific Trust Securities (SG) | **questionnaire-specific** Q5 "MAS Full Bank condition" (in no other file); Q1 "…and Registered Address" (substantive variant) |
| `03_meridian_hk.json` | Meridian Capital Markets (HK) | **numeric answer outlier** — Q9a issuer panel = 45 vs 8–12 elsewhere |
| `04_crescent_sg.json` | Crescent Wealth Partners (SG) | **empty answer (missing)** — Q4 economic data left blank |
| `05_northwind_hk.json` | Northwind Securities (HK) | **free-text answer outlier** — Q20 marketing answer diverges from the other 9 |
| `06_summit_sg.json` | Summit Asset Management (SG) | **empty answer (missing)** — Q20 marketing left blank; Q1 "Firm  Full  Legal…" (formatting-only variant) |
| `07_harbourview_hk.json` | Harbour View Investments (HK) | **YES/NO answer outlier** — Q21 = NO while the other 9 = YES |
| `08_silkroad_sg.json` | Silk Road Brokerage (SG) | **declared N/A** — Q4 economic data answered "N/A" while peers answer substantively |
| `09_apex_hk.json` | Apex Financial (HK) | **formatting-only** drift — Q2 "REGULATORY  AUTHORITY" vs the other 9 "Regulatory Authority" |
| `10_lighthouse_sg.json` | Lighthouse Capital (SG) | **yes/no synonym, NOT an outlier** — Q21 answer "Y" must canonicalize to YES |

## Goal → where to look

1. **Empty answers** — Firm 04 (Q4) and Firm 06 (Q20) leave *required* answers
   blank → flagged **missing**. The "Non-settling distributor / Who is the
   custodian?" branch and the "If non-settling…(only if applicable)" question are
   blank for nearly every firm → flagged **optional/conditional**, not a defect.
2. **Questionnaire-specific questions** — Firm 02's "MAS Full Bank condition"
   question is present in 1/10 files → flagged specific to that questionnaire.
3. **Question consistency / changes** —
   * **Q1** firm-name has a **substantive** edit ("…and Registered Address")
     plus a cosmetic extra-spaces variant; the canonical question is classified
     *substantive* because a real edit is present.
   * **Q2** "Regulatory Authority" has a **formatting-only** variant (Firm 09's
     "REGULATORY  AUTHORITY") — classified *formatting only*.
   * **Q3** authorised-activities also surfaces as a wording difference: the
     HK files prompt "For HK…" and the SG files "For Singapore…" — a real,
     jurisdiction-driven phrasing difference the reviewer should see.
4. **Answer outliers** — Q21 direct-distribution is **9× YES / 1× NO** (Firm 07);
   Q20 marketing has a **minority free-text cluster** (Firm 05).

## Intentional limitation: question IDs are not trusted

The same capacity/custodian question is `Q10`/`Q11` in HK files and `Q12`/`Q13`
in SG files. The analyzer must match questions by **text similarity**, not by id.
