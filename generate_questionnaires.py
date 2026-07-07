"""Generate KYC questionnaire Excel files, one per entity.

Each file has a single sheet "Questionnaire_Long" with columns:
    Source_file | Section | question_id | question | answer

A shared question bank drives every file. Per entity we override the answers
and lightly reword a few questions, so files differ in both.
"""

from openpyxl import Workbook

COLUMNS = ["Source_file", "Section", "question_id", "question", "answer"]

# Base question bank: (section, question_id, question)
BANK = [
    ("1. Entity Identification", "KYC_001",
     "What is the full legal name of the entity and any trading/DBA names?"),
    ("1. Entity Identification", "KYC_002",
     "What is the country of incorporation and legal registration number?"),
    ("1. Entity Identification", "KYC_003",
     "What is the registered office address and primary principal place of business?"),
    ("2. Ownership & Control", "KYC_004",
     "Are there any individual Ultimate Beneficial Owners (UBOs) holding direct or "
     "indirect shares/voting rights of 25% or more?"),
    ("2. Ownership & Control", "KYC_005",
     "Please list the full names and nationalities of all current executive directors."),
    ("3. Business Nature & Activity", "KYC_006",
     "What is the primary industry sector and nature of core business operations?"),
    ("3. Business Nature & Activity", "KYC_007",
     "Which geographic regions or countries does the company primarily operate in or "
     "generate revenue from?"),
    ("4. Financial & Transaction Profile", "KYC_008",
     "What is the expected annual turnover and the projected monthly volume of "
     "transactions through this account?"),
    ("4. Financial & Transaction Profile", "KYC_009",
     "What are the primary source of funds and source of wealth for the entity's operations?"),
    ("4. Financial & Transaction Profile", "KYC_010",
     "Will the account be used for cross-border wire transfers, and if so, to which "
     "high-frequency counterparty locations?"),
    ("5. AML & Sanctions Risk", "KYC_011",
     "Does the entity, its subsidiaries, or any major stakeholder operate in or conduct "
     "business with high-risk or sanctioned jurisdictions?"),
    ("5. AML & Sanctions Risk", "KYC_012",
     "Are any UBOs, directors, or authorized signatories classified as a Politically "
     "Exposed Person (PEP) or closely associated with one?"),
]

# Per entity: source_file, answers keyed by question_id, and optional question
# reword-s (slight wording differences) keyed by question_id.
ENTITIES = [
    {
        "source_file": "kyc_core_questionnaire_v1.0",
        "answers": {
            "KYC_001": "Global Logistics Solutions Ltd. (DBA: GLS Express)",
            "KYC_002": "Hong Kong (CR No. 2948102)",
            "KYC_003": "Suite 1402, 14/F, Central Plaza, 18 Harbour Road, Wanchai, Hong Kong",
            "KYC_004": "Yes, Mr. David Chen holds 45% direct ownership. Remaining 55% held by "
                       "institutional investors below the disclosure threshold.",
            "KYC_005": "1. David Chen (Hong Kong SAR), 2. Sarah Jenkins (United Kingdom), "
                       "3. Kenji Tanaka (Japan)",
            "KYC_006": "Cross-border supply chain management and freight forwarding services.",
            "KYC_007": "East Asia (60%), Southeast Asia (30%), Western Europe (10%).",
            "KYC_008": "Expected annual turnover: USD 15,000,000; Projected monthly account "
                       "volume: USD 1,250,000 across ~50 transactions.",
            "KYC_009": "Primary Source of Funds: Commercial revenues from freight service "
                       "contracts. Source of Wealth: Retained corporate earnings accumulated "
                       "over 8 years of operation.",
            "KYC_010": "Yes, recurring cross-border transfers to port agents and carriers in "
                       "Singapore, Japan, Thailand, and mainland China.",
            "KYC_011": "No. Internal policies strictly prohibit direct or indirect transactions "
                       "involving countries currently under UN, OFAC, or EU sanctions.",
            "KYC_012": "No PEP relationships have been identified among current management, "
                       "board members, or beneficial owners.",
        },
        "questions": {},
    },
    {
        "source_file": "kyc_core_questionnaire_v1.1",
        "answers": {
            "KYC_001": "Meridian Capital Advisors Pte. Ltd. (DBA: Meridian Wealth)",
            "KYC_002": "Singapore (UEN 201734567K)",
            "KYC_003": "9 Raffles Place, #18-01 Republic Plaza, Singapore 048619",
            "KYC_004": "Yes, Ms. Priya Nair holds 30% direct ownership and Mr. Alan Wong holds "
                       "28% indirectly via Meridian Holdings.",
            "KYC_005": "1. Priya Nair (Singapore), 2. Alan Wong (Malaysia), "
                       "3. Elena Petrova (Cyprus)",
            "KYC_006": "Independent asset management and investment advisory services.",
            "KYC_007": "Southeast Asia (50%), Middle East (35%), Western Europe (15%).",
            "KYC_008": "Expected annual turnover: USD 42,000,000; Projected monthly account "
                       "volume: USD 3,500,000 across ~120 transactions.",
            "KYC_009": "Primary Source of Funds: Advisory and management fees. Source of Wealth: "
                       "Founder equity and reinvested profits over 12 years.",
            "KYC_010": "Yes, frequent transfers to custodian banks and fund administrators in "
                       "Luxembourg, Switzerland, and the United Arab Emirates.",
            "KYC_011": "Limited. The firm advises one client with operations in a medium-risk "
                       "jurisdiction; enhanced due diligence is applied.",
            "KYC_012": "Yes. One director, Elena Petrova, is a close associate of a foreign "
                       "PEP; enhanced monitoring is in place.",
        },
        # Slight reword-s for this version.
        "questions": {
            "KYC_005": "Please provide the full names and nationalities of all currently "
                       "serving executive directors.",
            "KYC_012": "Are any beneficial owners, directors, or authorized signatories a "
                       "Politically Exposed Person (PEP) or a known close associate of one?",
        },
    },
    {
        "source_file": "kyc_core_questionnaire_v2.0",
        "answers": {
            "KYC_001": "Andes Mining & Metals S.A. (DBA: AMM Group)",
            "KYC_002": "Peru (RUC 20512345678)",
            "KYC_003": "Av. Javier Prado Este 4200, Santiago de Surco, Lima 15038, Peru",
            "KYC_004": "Yes, the Rojas family trust holds 62% collectively; no single individual "
                       "exceeds 25% except Mr. Carlos Rojas at 33%.",
            "KYC_005": "1. Carlos Rojas (Peru), 2. Maria Fernanda Silva (Brazil), "
                       "3. Thomas Muller (Germany)",
            "KYC_006": "Extraction, processing, and export of copper and zinc concentrates.",
            "KYC_007": "South America (70%), East Asia (20%), North America (10%).",
            "KYC_008": "Expected annual turnover: USD 210,000,000; Projected monthly account "
                       "volume: USD 18,000,000 across ~40 large transactions.",
            "KYC_009": "Primary Source of Funds: Mineral export sales. Source of Wealth: "
                       "Long-standing family mining concessions held since 1978.",
            "KYC_010": "Yes, regular settlements to metal traders and smelters in China, "
                       "South Korea, and Germany.",
            "KYC_011": "No. All counterparties are screened; no dealings with sanctioned "
                       "jurisdictions or restricted parties.",
            "KYC_012": "No PEPs identified; however one director previously held a regional "
                       "trade advisory post, disclosed for transparency.",
        },
        "questions": {
            "KYC_006": "What is the principal industry sector and the nature of the entity's "
                       "core operating activities?",
            "KYC_008": "State the expected annual turnover and the anticipated monthly "
                       "transaction volume passing through this account.",
        },
    },
]


def build_rows(entity):
    rows = [COLUMNS]
    for section, qid, base_q in BANK:
        question = entity["questions"].get(qid, base_q)
        answer = entity["answers"].get(qid, "")
        rows.append([entity["source_file"], section, qid, question, answer])
    return rows


def write_file(entity):
    wb = Workbook()
    ws = wb.active
    ws.title = "Questionnaire_Long"
    for row in build_rows(entity):
        ws.append(row)
    path = entity["source_file"] + ".xlsx"
    wb.save(path)
    return path


def main():
    for entity in ENTITIES:
        print("wrote", write_file(entity))


if __name__ == "__main__":
    # Self-check: every file has the sheet, header, and one row per question.
    for e in ENTITIES:
        rows = build_rows(e)
        assert rows[0] == COLUMNS
        assert len(rows) == len(BANK) + 1
        assert all(len(r) == len(COLUMNS) for r in rows)
    main()
