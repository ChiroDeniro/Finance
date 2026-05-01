import os
import time
from openpyxl import Workbook

KLEUR_POSITIEF = "C6EFCE"
KLEUR_NEGATIEF = "FCE4D6"

from config import OUTPUT_DIR, ACCOUNT_LABELS
from sheet_transactions import write_transactions_sheet
from sheet_overview import write_overview_sheet
from sheet_controle import write_controle_sheet


def _betaal_df(df):
    betaal_acc = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Betaal" in lbl), None)
    return df[df["account"] == betaal_acc] if betaal_acc else df


def save_year_output(df, year):
    wb        = Workbook()
    df_betaal = _betaal_df(df)
    df_onb    = df_betaal[df_betaal["category"] == "Dagelijks Overig"]

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df_betaal)
    ws_tx.title = "Transacties"
    print(f"  Sheet Transacties:           {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_mo = wb.create_sheet()
    write_overview_sheet(ws_mo, df_betaal, group_by="month",
                         tx_sheet_name="Transacties", year_label=year)
    ws_mo.title = "Maand Overzicht"
    print(f"  Sheet Maand Overzicht:       {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_onb = wb.create_sheet()
    write_transactions_sheet(ws_onb, df_onb)
    ws_onb.title = "Onbekende Transacties"
    print(f"  Sheet Onbekende Transacties: {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ctrl = wb.create_sheet()
    write_controle_sheet(ws_ctrl, df_betaal, year_label=year)
    ws_ctrl.title = "Controle"
    print(f"  Sheet Controle:              {time.time() - t0:.2f}s")

    out_path = os.path.join(OUTPUT_DIR, f"boekhouding_{year}.xlsx")
    t0 = time.time()
    wb.save(out_path)
    print(f"  wb.save():              {time.time() - t0:.2f}s")
    print(f"Opgeslagen: {out_path}")
    return out_path


def save_master_output(df):
    wb        = Workbook()
    df_betaal = _betaal_df(df)

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df)
    ws_tx.title = "Transacties"
    print(f"  Sheet Transacties:           {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ctrl = wb.create_sheet()
    write_controle_sheet(ws_ctrl, df_betaal, year_label="")
    ws_ctrl.title = "Controle"
    print(f"  Sheet Controle:              {time.time() - t0:.2f}s")

    out_path = os.path.join(OUTPUT_DIR, "boekhouding_alles.xlsx")
    t0 = time.time()
    wb.save(out_path)
    print(f"  wb.save():              {time.time() - t0:.2f}s")
    print(f"Opgeslagen: {out_path}")
    return out_path
