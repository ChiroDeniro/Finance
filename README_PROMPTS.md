# README_PROMPTS.md — Bewezen prompts per taak

Gebruik deze prompts als startpunt voor elke Claude Code sessie.
Pas de placeholders (`[...]`) aan naar de actuele situatie.

---

## Backlog beheren

#### Add an idea
```
Read ONLY: BACKLOG.md
Add this to the "Later" section:
[IDEA — 1-2 lines max]
Commit: "backlog: [short name] added"
```

#### Start a backlog item
```
Read ONLY: BACKLOG.md, bookkeeping/CLAUDE.md
Move this item from "Next" to "Now":
[ITEM]
Create branch if needed: git checkout -b [feature/name]
Commit: "backlog: [item] moved to now"
```

#### Close a completed item
```
Read ONLY: BACKLOG.md
Mark this item as done and move it to a new
"## Done" section at the bottom with completion date:
[ITEM]
Commit: "backlog: [item] done"
```

---

## process.py — Bankbestanden verwerken

### Onbekende transacties categoriseren
```
Open output/boekhouding_[JAAR].xlsx, sheet "Onbekend".
Voeg voor de top-[N] onbekende merchants een regel toe aan rules.xlsx
met de juiste categorie uit het categoriesysteem in CLAUDE.md.
Herstart daarna met: python -X utf8 process.py --year [JAAR]
Doel: minimaal 90% coverage.
```

### Nieuw jaar verwerken
```
Er staan nieuwe .TAB bestanden in bookkeeping/input/ voor jaar [JAAR].
Verwerk ze met: python -X utf8 process.py --year [JAAR]
Controleer de output in output/boekhouding_[JAAR].xlsx.
Rapporteer: hoeveel transacties, welk % gecategoriseerd, top onbekenden.
```

### Nieuwe categorisatieregel toevoegen
```
Voeg een regel toe aan rules.xlsx:
- Keyword: "[MERCHANT_NAAM]"
- Categorie: "[CATEGORIE]"
Zorg dat de categorie exact overeenkomt met de lijst in CLAUDE.md.
Test daarna met python -X utf8 process.py --year [JAAR].
```

---

## receipt_scanner.py — Bonnetjes OCR

### Eerste opzet (nog niet gebouwd)
```
Maak bookkeeping/receipt_scanner.py.
Doel: scan alle bestanden in bookkeeping/receipts/ (jpg/png/pdf),
gebruik Claude Vision API om per bonnetje te extraheren:
datum, bedrag, BTW-bedrag, leverancier, betaalmethode.
Output: Excel met alle bonnetjes en een kolom "Gematch met transactie" (ja/nee).
Gebruik de Anthropic SDK met claude-sonnet-4-6.
```

### Bonnetje matchen met banktransactie
```
In receipt_scanner.py: voeg matching toe op datum (±1 dag) + bedrag (exact).
Lees de transacties uit output/boekhouding_[JAAR].xlsx sheet "Transacties".
Markeer gematchte transacties in beide outputs.
```

---

## tax_report.py — Belastingaangifte

### BTW kwartaaloverzicht (nog niet gebouwd)
```
Maak bookkeeping/tax_report.py.
Doel: lees output/boekhouding_[JAAR].xlsx en genereer een BTW-overzicht per kwartaal.
Splits per BTW-tarief: 0%, 9%, 21%.
Input: kolom "BTW" in de bonnetjes-output van receipt_scanner.py.
Output: Excel met per kwartaal: omzet, BTW te betalen, aftrekbare kosten.
```

---

## annual_report.py — Jaarverslag en prognose

### Jaarverslag genereren (nog niet gebouwd)
```
Maak bookkeeping/annual_report.py.
Doel: volledig jaarverslag privé + zakelijk gecombineerd.
Lees output/boekhouding_[JAAR].xlsx.
Output:
- Blad 1: Totaaloverzicht per categorie (privé vs zakelijk naast elkaar)
- Blad 2: 3-maands prognose op basis van gemiddelde maanduitgaven
- Blad 3: Spaardoelen tracker (saldo per doel, maandelijkse aangroei)
Gebruik dezelfde stijl (kleuren, fonts) als process.py.
```

---

## Algemeen — sessie starten

```
Lees CLAUDE.md voor de huidige projectstatus.
Wat werkt al, wat is work in progress, en wat staat als volgende stap?
Stel voor wat we vandaag kunnen aanpakken op basis van de Hoofddoelen.
```
