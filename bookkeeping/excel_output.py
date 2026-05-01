import os
from datetime import datetime
from openpyxl import Workbook

KLEUR_POSITIEF = "C6EFCE"
KLEUR_NEGATIEF = "FCE4D6"

from config import OUTPUT_DIR, ACCOUNT_LABELS
from sheet_transactions import write_transactions_sheet
from sheet_overview import write_overview_sheet
from sheet_spaar import write_spaar_overview_sheet
from sheet_jaar import write_jaar_samenvatting_sheet
from sheet_controle import write_controle_sheet


def save_output(df, year_label=""):
    import time
    suffix = f" {year_label}" if year_label else ""
    wb = Workbook()

    betaal_acc = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Betaal" in lbl), None)
    spaar_acc  = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Spaar"  in lbl), None)
    df_betaal  = df[df["account"] == betaal_acc] if betaal_acc else df
    df_spaar   = df[df["account"] == spaar_acc]  if spaar_acc  else df.iloc[0:0]

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df_betaal)
    ws_tx.title = f"Transacties Betaal{suffix}"
    print(f"  Sheet Transacties Betaal:    {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ts = wb.create_sheet()
    write_transactions_sheet(ws_ts, df_spaar)
    ws_ts.title = f"Transacties Spaar{suffix}"
    print(f"  Sheet Transacties Spaar:     {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_mo = wb.create_sheet()
    write_overview_sheet(ws_mo, df_betaal, group_by="month",
                         tx_sheet_name=ws_tx.title, year_label=year_label)
    ws_mo.title = f"Maand Overzicht{suffix}"
    print(f"  Sheet Maand Overzicht:       {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ms = wb.create_sheet()
    write_spaar_overview_sheet(ws_ms, df_spaar, year_label=year_label,
                               tx_sheet_name=ws_ts.title)
    print(f"  Sheet Maand Overzicht Spaar: {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_js = wb.create_sheet()
    write_jaar_samenvatting_sheet(ws_js, df_betaal, year_label=year_label,
                                  tx_sheet_name=ws_tx.title)
    print(f"  Sheet Jaar Samenvatting:     {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ctrl = wb.create_sheet()
    write_controle_sheet(ws_ctrl, df_betaal, year_label=year_label)
    print(f"  Sheet Controle:              {time.time() - t0:.2f}s")

    if year_label:
        filename = f"boekhouding_{year_label}.xlsx"
    else:
        now = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"boekhouding_{now}.xlsx"

    t0 = time.time()
    out_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(out_path)
    print(f"  wb.save():              {time.time() - t0:.2f}s")
    print(f"Opgeslagen: {out_path}")
    return out_path
