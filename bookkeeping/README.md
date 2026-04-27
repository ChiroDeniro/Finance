# Finance — Personal Bookkeeping System

Personal finance automation for ABN AMRO bank accounts. Replaces manual Excel kasboek.

---

## Project

Personal finance automation for Chris (Den Haag).
Two ABN AMRO accounts — betaalrekening (personal) and spaarrekening. Replaces a manual Excel kasboek that broke down when unknowns weren't tracked and year summaries drifted from reality.

Design principles: local first, one command to run, transparent Excel output with visible formulas, incremental modules built in phases.

---

## Quick start

```bash
cd bookkeeping
pip install pandas openpyxl   # first time only

python process.py              # process all TAB files in input/
python process.py --year 2025  # process only 2025 transactions
```

Drop `.TAB` files in `bookkeeping/input/` first (gitignored). Output lands in `bookkeeping/output/`.

> **Windows terminal note:** use `python -X utf8 process.py` if you get encoding errors.

---

## What it produces

One Excel file with 4 sheets:

| Sheet | Contents |
|-------|----------|
| Transacties | All transactions, date-sorted, auto-filter, colour-coded by type |
| Maand Overzicht | Income/expense by category, columns = months, totals + average |
| Jaar Samenvatting | Full-year totals per category + % van inkomen |
| Onbekend | Unknown merchants grouped, sorted by amount — fill in categories here |

---

## Input format

ABN AMRO → Internetbankieren → Bij/Afschriften → Exporteren → **Excel (TXT)**

Saves as a `.TAB` file. Tab-separated, 8 columns, no header:

| Col | Field | Example |
|-----|-------|---------|
| 0 | Account number | 536542171 |
| 1 | Currency | EUR |
| 2 | Date (YYYYMMDD) | 20250321 |
| 3 | Balance before | 211,90 |
| 4 | Balance after | 192,48 |
| 5 | Value date | 20250321 |
| 6 | Amount (comma decimal) | -19,42 |
| 7 | Description | BEA, Apple Pay   Albert Heijn 1870... |

Multiple accounts or files in one run are fine — duplicates are removed automatically.

**Two accounts (betaalrekening + spaarrekening):** transfers between them are detected and tagged as *Interne Overboeking* — they appear in Transacties but are excluded from all totals and netto calculations.

---

## Categories

Defined in `rules.xlsx` (Sheet: **Rules**). Rules use case-insensitive keyword matching on the merchant name. First match wins — put specific rules above general ones.

### INKOMEN
`Salaris` · `DUO / Studiefinanciering` · `Zorgtoeslag` · `Familie & Giften` · `Inkomsten Overig`

### VASTE LASTEN
`Huur & Wonen` · `Zorgverzekering` · `Telefoon & Internet` · `Bankkosten` · `Abonnementen`

### DAGELIJKSE UITGAVEN
`Boodschappen` · `Eten & Drinken` · `OV & Reizen` · `Sport & Fitness` · `Online Winkelen` · `Kleding` · `Gezondheid` · `Tabak` · `Cultuur & Entertainment` · `Studie`

### OVERIG
`Sparen` · `Diversen`

To fix an unknown: add a row to `rules.xlsx` → Sheet Rules → keyword + category. Re-run the script.

---

## Other flags

```bash
python process.py --create-rules   # regenerate rules.xlsx from scratch (overwrites!)
python process.py --migrate-rules  # rename old category names to new system
```

---

## Folder structure

```
finance/
├── CLAUDE.md
├── README.md
├── README_PROMPTS.md          ← reusable prompt templates
├── BACKLOG.md                 ← ideas, planned features, work in progress
└── bookkeeping/
    ├── process.py             ← main entry point
    ├── loader.py              ← TAB reading + validation
    ├── categoriser.py         ← rules matching
    ├── excel_output.py        ← Excel writing + styling
    ├── config.py              ← constants + paths
    ├── receipt_scanner.py     ← (planned) receipt OCR
    ├── tax_report.py          ← (planned) tax overview
    ├── annual_report.py       ← (planned) year report
    ├── rules.xlsx
    ├── input/
    ├── receipts/              ← drop receipts/invoices here
    └── output/
```

---

## Roadmap

| # | Module | Status |
|---|--------|--------|
| 1 | process.py — bank file import + categorisation | ✅ done |
| 2 | receipt_scanner.py — receipt OCR + matching | planned |
| 3 | tax_report.py — BTW + IB aangifte export | planned |
| 4 | annual_report.py — year report + forecast | planned |
| 5 | Dashboard — visualisation + net worth | planned |

---

## Owner

Chris — Den Haag
