# Backlog — Finance Project

## Now — in progress
- [ ] Refactor process.py into modules (branch: refactor/split-process-modules)
- [ ] Fix unknowns in 2025 monthly overview + controle sheet

## Next — ready to pick up
- [ ] Auto-detect account numbers from TAB into config.json
- [ ] New 2026 category structure (ODS, BS, PV, wbw, Huur, Inclusief etc.)
- [ ] Saldo betaalrekening auto-read from TAB (balance_before first tx)
- [ ] Saldo spaarrekening auto-read from TAB with fallback if no mutations

## Later — ideas + wishes
- [ ] Receipt OCR via Claude Vision API (receipt_scanner.py)
- [ ] BTW kwartaaloverzicht voor belastingaangifte (tax_report.py)
- [ ] Jaarverslag + 3-maands prognose (annual_report.py)
- [ ] Data visualisation — charts embedded in Excel output
- [ ] Forecast scenarios: what if income drops / expenses rise
- [ ] Asset overview: investments, business assets, personal valuables (bike, watch, clothing, jewellery)
- [ ] Net worth calculation: total assets - liabilities, per month
- [ ] Routine: auto-import bank files + receipts on fixed schedule
- [ ] Dashboard: Streamlit or web UI

## Graveyard — decided against
- ~~Supabase + Vercel dashboard~~ — overkill for local use, local Excel is faster to build and easier to debug
