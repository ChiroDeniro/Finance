"""
ABN AMRO Bookkeeping Processor
================================
Drop your ABN AMRO Excel (TXT) export into the /input folder, then run:
    python process.py

Output: /output/boekhouding_YYYYMMDD_HHMM.xlsx  with:
  - Sheet 1: Transacties       — all transactions, sorted by date
  - Sheet 2: Maand Overzicht   — kasboek-style blocks per month
  - Sheet 3: Jaar Overzicht    — same structure per year
  - Sheet 4: Onbekend          — unknown merchants grouped for easy triage
"""

import os
import sys

from config import RULES_FILE, INCOME_CATS, VASTE_LASTEN_CATS, DAGELIJKS_CATS, MAANDEN_NL, dutch_euros
from loader import find_input_files, load_transactions
from categoriser import load_rules, apply_categories, detect_internal_transfers, create_starter_rules, migrate_rules
from excel_output import save_output


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if "--migrate-rules" in sys.argv:
        migrate_rules()
        return

    if "--create-rules" in sys.argv or not os.path.exists(RULES_FILE):
        print("Starter rules.xlsx aanmaken ...")
        create_starter_rules()
        if "--create-rules" in sys.argv:
            return

    # ── Parse --year flag ─────────────────────────────────────────────────────
    year_filter = None
    if "--year" in sys.argv:
        idx = sys.argv.index("--year")
        if idx + 1 < len(sys.argv):
            year_filter = sys.argv[idx + 1]

    print("\nABN AMRO Bookkeeping Processor")
    print("=" * 38)
    if year_filter:
        print(f"Jaar filter: {year_filter}")

    files = find_input_files()
    df    = load_transactions(files)

    if year_filter:
        df = df[df["year"] == year_filter].reset_index(drop=True)
        if df.empty:
            sys.exit(f"Geen transacties gevonden voor jaar {year_filter}")
        print(f"Gefilterd op {year_filter}: {len(df)} transacties")

    rules = load_rules()
    df    = apply_categories(df, rules)
    df    = detect_internal_transfers(df)
    out   = save_output(df, year_label=year_filter or "")

    # ── Terminal summary ──────────────────────────────────────────────────────
    df_real = df[df["category"] != "Interne Overboeking"]
    print()
    for m in sorted(df_real["month"].unique()):
        mdf  = df_real[df_real["month"] == m]
        inc  = mdf[mdf["category"].isin(INCOME_CATS)]["amount"].sum()
        vl   = mdf[mdf["category"].isin(VASTE_LASTEN_CATS)]["amount"].sum()
        dag  = mdf[mdf["category"].isin(DAGELIJKS_CATS)]["amount"].sum()
        net  = mdf["amount"].sum()
        yr, mo = m.split("-")
        label  = f"{MAANDEN_NL[int(mo)]} {yr}"
        print(f"  {label:<10}  |  in: {dutch_euros(inc):<10}"
              f"  |  vaste lasten: {dutch_euros(vl):<10}"
              f"  |  dagelijks: {dutch_euros(dag):<10}"
              f"  |  netto: {dutch_euros(net)}")

    t_inc = df_real[df_real["category"].isin(INCOME_CATS)]["amount"].sum()
    t_vl  = df_real[df_real["category"].isin(VASTE_LASTEN_CATS)]["amount"].sum()
    t_dag = df_real[df_real["category"].isin(DAGELIJKS_CATS)]["amount"].sum()
    t_net = df_real["amount"].sum()
    print("  " + "-" * 80)
    print(f"  {'Totaal':<10}  |  in: {dutch_euros(t_inc):<10}"
          f"  |  vaste lasten: {dutch_euros(t_vl):<10}"
          f"  |  dagelijks: {dutch_euros(t_dag):<10}"
          f"  |  netto: {dutch_euros(t_net)}")
    print(f"\n  Output: {out}\n")


if __name__ == "__main__":
    main()
