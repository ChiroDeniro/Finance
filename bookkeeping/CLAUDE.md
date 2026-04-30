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
├── README_PROMPTS.md                  ← bewezen prompts per taak
└── bookkeeping/
    ├── process.py                     ← hoofdscript — bankbestanden verwerken
    ├── receipt_scanner.py             ← (gepland) bonnetjes OCR en matching
    ├── tax_report.py                  ← (gepland) belastingaangifte overzicht
    ├── annual_report.py               ← (gepland) jaarverslag en prognose
    ├── rules.xlsx                     ← categorisatieregels (keyword → categorie)
    ├── input/                         ← drop .TAB bestanden hier (gitignored)
    ├── receipts/                      ← drop bonnetjes/facturen hier (gitignored)
    ├── output/                        ← alle Excel outputs (gitignored)
    └── InfoFiles/
        └── Jaaroverzichten_kasboekken.xlsx  ← oud kasboek 2023 (referentie)
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
- Huur
- Inclusief Huur
- Zorgverzekering
- Telefoon & Internet
- Bankkosten
- Abonnementen
- Sport & Fitness
- Onderhoud

### DAGELIJKSE UITGAVEN
- Boodschappen
- Eten & Drinken
- Uitgaan
- OV & Reizen
- Kleding
- Kapper
- Gezondheid
- WbW
- Cultuur & Entertainment
- Studie
- Dagelijks Overig

### OVERIG
- Sparen

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

## Ontwerpbeslissingen

Keuzes die niet uit de code zelf blijken — bewaar dit voor toekomstige sessies.

| Keuze | Reden |
|-------|-------|
| Alleen `pandas` + `openpyxl` | Geen zware deps, geen API-sleutels, werkt offline |
| Keyword-matching i.p.v. ML | Regels zijn zichtbaar in rules.xlsx, deterministisch, makkelijk te debuggen |
| Interne overboeking op bedrag+datum | Omschrijving varieert; alleen bedrag+datum is betrouwbaar |
| Output naar Drive-map als die bestaat | Automatische sync naar Google Drive zonder extra stap |
| `--year YYYY` filter | Verwerkt één jaar zonder andere data aan te raken; geeft schone Excel per jaar |
| Stijl in module-niveau constanten | Één plek voor kleuren/fonts — geen magic values door de code heen |
| Prompt design rule | One task per prompt, explicit file list, explicit exclusions, one test command, one commit. Prevents freezing. |
| File size limit | Any file Claude Code reads regularly: max ~150 lines. Over 300 lines and touched every session → split it. |
| README vs CLAUDE.md | README = for humans who've never seen the project. CLAUDE.md = for Claude Code starting fresh. Never mix audiences. |

---

## Current state (as of April 2026)

### What works
- Full kasboek-style Excel output with 4 colour-coded sheets
- `--year YYYY` flag for single-year processing
- Dual-account support: internal transfers between betaalrekening (536542171) and spaarrekening (844835730) are detected and excluded from netto
- 35 categorisation rules covering the main recurring merchants
- Terminal summary per month: inkomen | vaste lasten | dagelijks | netto
- Output auto-saves to `C:\Users\chris\Documents\Finance\Kasboek\` if that Drive-synced folder exists

### Work in progress
- **2025 full-year rules gap-fill**: current run gives **56% coverage (725/1298 tx)**, 573 uncategorised remain.
- **Inkomen = €0 for Jan–Sep 2025** — salary/income SEPA rules not matching. Top priority: add rules for the merchants below.
- **process.py has uncommitted category changes** (249-line diff) — category names diverge from what's documented in CLAUDE.md. Commit or revert before next session.

### Top unknowns to fix (open `Onbekend 2025` sheet, add keywords to rules.xlsx)
| Priority | Merchant | Count | Amount | Likely category |
|----------|----------|-------|--------|-----------------|
| 1 | J.P.Q. Hoyng — Huur SJ 109 | 12 | -€5,851 | Huur & Wonen |
| 2 | D. van Rhijn e | 7 | +€5,200 | Salaris / Inkomsten Overig |
| 3 | Optimal Solutions | 3 | +€4,200 | Salaris |
| 4 | Dhr. C. van Rhijn | 1 | +€1,800 | Salaris / Inkomsten Overig |
| 5 | Christiaan Van Rhijn — WDR leeg | 1 | +€1,705 | Inkomsten Overig |
| 6 | RHIJN D VAN en | 1 | -€1,050 | Diversen |
| 7 | Revolut (eCom Apple Pay) | 3 | -€400 | Dagelijks Overig |

### Next session checklist
1. Decide: commit or revert uncommitted process.py changes (category rename)
2. Open `boekhouding_2025.xlsx` → sheet `Onbekend 2025`
3. Add missing SEPA NAME keywords to `rules.xlsx` — start with rows 1–5 in table above
4. Rerun `python -X utf8 process.py --year 2025` until ≥90% coverage
5. Commit `rules.xlsx` and updated output notes

---

## Backlog

See `BACKLOG.md` in the repo root.
Do NOT read it unless explicitly asked to.
Current focus: see "Now" section in BACKLOG.md.

---

## Architecture decisions pending

These must be decided before building phase 2+ modules.
Decide in the session that starts that phase — not before.

**Storage layer (decide before receipt_scanner.py)**
Current: Excel is both database and output format.
Option A: keep as-is — simple, no new dependencies
Option B: SQLite as database, Excel as export only
  Pros: cross-year queries, dashboard-ready, faster lookups
  Cons: adds complexity, requires migration of existing output
Trigger to choose B: if you need to query across years or
build a dashboard that updates without rerunning process.py

**Context window budget per session**
No single file read by Claude Code should exceed ~150 lines.
Current violations: process.py (~1000 lines) → refactor first.
Rule: if a file is over 300 lines and touched every session, split it.

---

## Hoofddoelen

1. Bankbestanden inlezen ✅ → process.py
2. Bonnetjes inlezen → receipt_scanner.py (planned)
3. Belastingaangifte → tax_report.py (planned)
4. Jaarverslag + prognose → annual_report.py (planned)
5. Dashboard (planned)

---

## Sessie starten

1. Lees dit bestand — vooral "Current state"
2. Plak een prompt uit `README_PROMPTS.md` en pas aan
3. Commit na elke werkende feature

---

## Reference Files

- `InfoFiles/Jaaroverzichten_kasboekken.xlsx` — oud kasboek 2023, structuur/kleur als referentie
