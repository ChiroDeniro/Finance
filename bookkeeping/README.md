# Finance — Personal Bookkeeping System

Personal finance automation for Chris (Den Haag). Replaces a manual Excel kasboek.
Supports ABN AMRO (betaal + spaar), Revolut, and Knab zakelijke rekening.

---

## Quick start

```bash
cd bookkeeping
pip install pandas openpyxl   # first time only

python -X utf8 process.py --year 2026   # one year
python -X utf8 process.py               # all years + Knab
```

Drop exports in the right input subfolder (see below) before running.

---

## Input

| Folder | Source | Format |
|--------|--------|--------|
| `input/ABN spaar en betaal/` | ABN AMRO betaal + spaarrekening | `.TAB` — Excel (TXT) export |
| `input/Revolut/` | Revolut | `.CSV` — Dutch export (Status=VOLTOOID) |
| `input/ODS/` | Knab zakelijke rekening | `.CSV` — Knab export |

All folders are gitignored. Duplicates are removed automatically.

### ABN AMRO export
Internetbankieren → Bij/Afschriften → Exporteren → **Excel (TXT)** → save as `.TAB`

### Revolut export
App → Profile → Statements → Export → **CSV** (Dutch)

---

## Output

| File | Contents |
|------|----------|
| `output/boekhouding_YYYY.xlsx` | ABN + Revolut per jaar |
| `output/boekhouding_alles.xlsx` | All years combined (master) |
| `output/boekhouding_knab.xlsx` | Knab zakelijke rekening |

Sheets in `boekhouding_YYYY.xlsx`:

| Sheet | Contents |
|-------|----------|
| Transacties | All transactions, date-sorted, auto-filter, Bron column (ABN/Revolut) |
| Maand Overzicht | Income/expense by category × month, SUMIFS formulas stay live |
| Onbekende Transacties | Unmatched transactions (Dagelijks Overig) |
| Controle | Balance check |
| Saldo Overzicht | Opening balance per month (betaal + spaar) |

Output auto-saves to `C:\Users\chris\Documents\Finance\Kasboek\` if that folder exists.

---

## Categories

Defined in `rules.xlsx` (Sheet: **Rules**). Keyword matching on merchant name — first match wins.

### INKOMEN
`Salaris` · `DUO / Studiefinanciering` · `Zorgtoeslag` · `Familie & Giften` · `ZZP Opname` · `ZZP Inkomen` · `Inkomsten Overig`

### VASTE LASTEN
`Huur` · `Inclusief Huur` · `Zorgverzekering` · `Telefoon & Internet` · `Bankkosten` · `Abonnementen` · `Sport & Fitness` · `Onderhoud` · `Belasting`

### DAGELIJKSE UITGAVEN
`Boodschappen` · `Eten & Drinken` · `Uitgaan` · `OV & Reizen` · `Kleding` · `Kapper` · `Gezondheid` · `WbW` · `Cultuur & Entertainment` · `Studie` · `Dagelijks Overig`

### OVERIG
`Sparen` · `Beleggen` · `Zakelijke Kosten`

Unknowns land in **Dagelijks Overig**. To fix: add a row to `rules.xlsx` → re-run.

---

## Internal transfers

Transfers between own accounts are tagged **Interne Overboeking** and excluded from all totals:
- ABN betaal ↔ spaar: detected by matching account numbers in description
- ABN → Revolut: matched via Revolut IBAN in description (`revonl22` in `config.json`)
- Revolut → own bank: Revolut `Geld toevoegen` type is always internal

---

## Flags

```bash
python -X utf8 process.py --year 2025      # filter to one year
python -X utf8 process.py --knab-only      # only process Knab
python -X utf8 process.py --create-rules   # regenerate rules.xlsx from scratch (overwrites!)
python -X utf8 process.py --migrate-rules  # rename old category names
```

---

## Folder structure

```
bookkeeping/
├── process.py             ← main entry point
├── loader.py              ← ABN AMRO TAB reader
├── loader_revolut.py      ← Revolut CSV reader
├── loader_knab.py         ← Knab CSV reader
├── categoriser.py         ← rules matching
├── excel_output.py        ← workbook factory
├── sheet_transactions.py  ← Transacties sheet
├── sheet_overview.py      ← Maand Overzicht
├── sheet_controle.py      ← Controle sheet
├── sheet_knab.py          ← Knab sheets
├── config.py              ← constants + paths
├── config.json            ← account labels, IBAN rules
├── rules.xlsx             ← categorisation rules (edit this)
├── input/
│   ├── ABN spaar en betaal/   ← drop .TAB files here
│   ├── Revolut/               ← drop Revolut .CSV files here
│   └── ODS/                   ← drop Knab .CSV files here
├── receipts/              ← drop receipts/invoices here
└── output/                ← all Excel outputs (gitignored)
```

---

## Owner

Chris — Den Haag
