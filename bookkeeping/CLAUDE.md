# CLAUDE.md — Finance Project

This file is read automatically by Claude Code at the start of every session.
Do not delete it. Keep it updated when the project changes significantly.

---

## Who & What

Personal finance automation for Chris (Den Haag).
Two bank accounts at ABN AMRO — one personal, one for company (ZZP/freelance).
Goal: replace manual Excel kasboek with a fully automated system.

---

## Project Structure

```
finance/
├── CLAUDE.md                          ← you are here
├── README.md                          ← human-readable project overview
└── bookkeeping/
    ├── process.py                     ← main script, run this each month
    ├── rules.xlsx                     ← categorisation rules (keyword → category)
    ├── input/                         ← drop .TAB files here (gitignored)
    ├── output/                        ← Excel output lands here (gitignored)
    └── InfoFiles/
        └── Jaaroverzichten_kasboekken.xlsx  ← old manual kasboek (reference/inspiration)
```

---

## How the Script Works

1. Reads all `.TAB` files from `input/` (deduplicates on Windows case-insensitive FS)
2. Validates format (must be ABN AMRO Excel TXT export, 8 tab-separated columns)
3. Optionally filters to a single year with `--year YYYY`
4. Extracts merchant name from raw description using regex
5. Matches merchant against `rules.xlsx` → assigns category
6. Detects internal transfers between own accounts (betaalrekening ↔ spaarrekening) → tags as `Interne Overboeking`, excluded from all totals
7. Outputs Excel to `output/` with 4 sheets

### Input Format (.TAB file from ABN AMRO)

Tab-separated, no header, 8 columns:
| # | Field | Example |
|---|-------|---------|
| 0 | Account number | 536542171 |
| 1 | Currency | EUR |
| 2 | Date (YYYYMMDD) | 20260321 |
| 3 | Balance before | 211,90 |
| 4 | Balance after | 192,48 |
| 5 | Value date | 20260321 |
| 6 | Amount (comma decimal) | -19,42 |
| 7 | Description (raw) | BEA, Apple Pay   Albert Heijn 1870... |

Two description formats:
- BEA (pin): `BEA, Apple Pay   MERCHANT NAME,PAS555   NR:xxx`
- SEPA: `/TRTP/SEPA.../NAME/Counterparty Name/REMI/...`

### Run

```bash
cd bookkeeping
python process.py                  # process all TAB files
python process.py --year 2025      # filter to one year
python process.py --create-rules   # regenerate rules.xlsx from scratch
python process.py --migrate-rules  # rename old category names to current system
```

On Windows, prefix with `python -X utf8` to avoid terminal encoding errors.

Dependencies: `pip install pandas openpyxl`

---

## Category System

Inspired by the old kasboek in InfoFiles/. Two blocks: INKOMEN and UITGAVEN.

### INKOMEN
- Salaris
- DUO / Studiefinanciering
- Zorgtoeslag
- Familie & Giften
- Inkomsten Overig

### VASTE LASTEN
- Huur & Wonen
- Zorgverzekering
- Telefoon & Internet
- Bankkosten
- Abonnementen

### DAGELIJKSE UITGAVEN
- Boodschappen
- Eten & Drinken
- OV & Reizen
- Sport & Fitness
- Online Winkelen
- Kleding
- Gezondheid
- Tabak
- Cultuur & Entertainment
- Studie

### OVERIG
- Sparen
- Diversen

**Important:** Always use these exact category names. rules.xlsx and process.py must stay in sync.

---

## Excel Output Structure

The output mimics the old kasboek style:

- **Transacties** — all transactions, date-sorted, with auto-filter; Interne Overboeking rows visible but neutral
- **Maand Overzicht** — rows = categories, columns = months + Totaal + Gemiddelde
  - Structured in blocks: INKOMEN / VASTE LASTEN / DAGELIJKS / OVERIG / SAMENVATTING (netto)
- **Jaar Samenvatting** — full-year totals per category + `% van inkomen` column
- **Onbekend** — grouped by merchant, with count + total, empty "Jouw categorie" column

When `--year YYYY` is used, the output file is named `boekhouding_YYYY.xlsx` and sheet names include the year.

---

## Coding Conventions

- Python 3, no external deps beyond `pandas` and `openpyxl`
- Dutch variable names for UI-facing strings, English for code
- Dutch error messages (users are Dutch)
- Always validate input files before processing — fail loudly with clear messages
- Keep all styling in helper functions (`hfill`, `hfont`, `money`, etc.)
- Never hardcode paths — always derive from `BASE_DIR`
- Test after every change: `python process.py`

---

## Current state (as of April 2026)

### What works
- Full kasboek-style Excel output with 4 colour-coded sheets
- `--year YYYY` flag for single-year processing
- Dual-account support: internal transfers between betaalrekening (536542171) and spaarrekening (844835730) are detected and excluded from netto
- 35 categorisation rules covering the main recurring merchants
- Terminal summary per month: inkomen | vaste lasten | dagelijks | netto

### Work in progress
- **2025 full-year rules gap-fill**: first run gives ~49% coverage (637/1298 tx). The "Onbekend" sheet in `output/boekhouding_2025.xlsx` has the grouped unknowns. Next session: open that sheet, add missing keywords to `rules.xlsx`, rerun until ≥90%.
- Known gaps: salary/income keywords probably don't match 2025 merchant names; some recurring costs unmatched.

### Next session checklist
1. Open `output/boekhouding_2025.xlsx` → sheet "Onbekend"
2. Add missing rules to `rules.xlsx` for top unknown merchants
3. Rerun `python -X utf8 process.py --year 2025` until ≥90% coverage
4. Commit final `rules.xlsx` and output notes

---

## Backlog (planned future features)

### Phase 2 — Company account (ZZP)
- [ ] Second ABN AMRO account processed separately from personal
- [ ] Company-specific categories: BTW, Zakelijke kosten, Facturen, etc.
- [ ] Separate output sheet/file for company vs personal

### Phase 3 — Receipt/invoice scanning (OCR)
- [ ] Upload photo of a receipt or invoice (jpg/png/pdf)
- [ ] Extract: vendor, amount, date, BTW, payment method using Claude vision
- [ ] Match scanned receipts against bank transactions automatically
- [ ] Useful for: ZZP facturen, zakelijke bonnetjes, BTW administratie

### Phase 4 — Dashboard
- [ ] Simple web UI or Streamlit dashboard
- [ ] Monthly spending chart per category
- [ ] Savings goals tracker (spaarpotjes like old kasboek)
- [ ] Year-over-year comparison

### Phase 5 — Automation
- [ ] Watch input/ folder and auto-process on new file
- [ ] Email/notification summary after processing

---

## Reference Files

- `InfoFiles/Jaaroverzichten_kasboekken.xlsx` — old manual kasboek 2023
  - Sheet "2023 betaalrekening": monthly income/expense breakdown with transaction log
  - Sheet "sparen 2023": savings goals (Apparaten, Noodfonds, SpaarRing, Kleding, Trips, etc.)
  - Use this as visual/structural inspiration only — data is from 2023

---

## When Starting a New Claude Code Session

1. Read this file (CLAUDE.md) — especially "Current state" above
2. Read `bookkeeping/process.py`
3. Check if there are new `.TAB` files in `bookkeeping/input/`
4. Run `python -X utf8 bookkeeping/process.py --year 2025` to verify current state
5. Ask what to work on — most likely: rules gap-fill for 2025, or starting on Phase 2 (company account)
