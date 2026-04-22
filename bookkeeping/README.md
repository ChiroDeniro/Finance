# Finance — Personal Bookkeeping System

Personal finance automation for ABN AMRO bank accounts.

## What this does

Downloads from ABN AMRO → Python script → Clean Excel with:
- All transactions auto-categorised
- Monthly overview per category (income vs expenses)
- Flagged unknowns to review

## Folder structure

```
finance/
├── bookkeeping/
│   ├── process.py        # Main script — run this each month
│   ├── rules.xlsx        # Categorisation rules (edit this to add/fix categories)
│   ├── input/            # Drop your .TAB file here (gitignored)
│   └── output/           # Excel output appears here (gitignored)
└── README.md
```

## How to run

```bash
cd bookkeeping
python process.py
```

First time setup:
```bash
pip install pandas openpyxl
```

## Input format

ABN AMRO export → **Excel (TXT)** format → saves as `.TAB` file.
Download via: Internetbankieren → Bij/Afschriften → Exporteren → Excel (TXT)

The file is tab-separated with 8 columns (no header):
| Col | Field | Example |
|-----|-------|---------|
| 0 | Account number | 536542171 |
| 1 | Currency | EUR |
| 2 | Date | 20260321 |
| 3 | Balance before | 211,90 |
| 4 | Balance after | 192,48 |
| 5 | Value date | 20260321 |
| 6 | Amount | -19,42 |
| 7 | Description | BEA, Apple Pay   Albert Heijn 1870... |

Multiple accounts can be exported into one file or separate files — script handles both.

## Categories

Defined in `rules.xlsx` → Sheet: **Rules**

| Category | What goes in it |
|----------|----------------|
| Inkomsten | Salary, DUO, family transfers in |
| Huur & Wonen | Rent, service costs |
| Boodschappen | Albert Heijn, Jumbo, Dirk |
| Eten & Drinken | Restaurants, cafés, takeaway |
| OV & Reizen | NS, bus, train, travel |
| Sport & Fitness | Sportcity, David Lloyd, gym |
| Verzekeringen | Zorgverzekering, UNIVE |
| Telefoon & Internet | Youfone, Odido, Tele2 |
| Bankkosten | ABN AMRO fees |
| Online Winkelen | Bol.com, Media Markt |
| Tabak | Tabakshop, GoGo |
| Cultuur & Entertainment | Concerts, museums, events |
| Sparen | Transfers to savings account |
| Diversen | Everything else |

Rules use **keyword matching** (case-insensitive, partial match). First match wins — put specific rules above general ones.

## History / context

- Previously tracked manually in Excel kasboek (see `Copy_of_Jaaroverzichten_kasboekken.xlsx`)
- Old kasboek had categories: BS, PV, wbw, Uit, OV, Studie, Huur, Telefoon, Zorgverzekering, Gym
- Spaarpotjes tracked separately: Apparaten, Noodfonds, SpaarRing, Kleding, Trips etc.
- This script replaces the manual work — same categories, fully automated

## Known issues / TODO

- [ ] 15 transactions still uncategorised from first run — need rules added
- [ ] Spaarpotjes (savings goals) not yet integrated
- [ ] Second bank account not yet tested
- [ ] Would be nice: auto-feedback loop (fill in unknowns → auto-adds to rules)

## Owner

Chris — Den Haag
