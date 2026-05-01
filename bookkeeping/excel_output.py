import os
import time
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

KLEUR_POSITIEF = "C6EFCE"
KLEUR_NEGATIEF = "FCE4D6"
KLEUR_GEEL     = "FFFFE0"

from config import OUTPUT_DIR, ACCOUNT_LABELS, MAANDEN_NL, C_HEADER, WHITE, _fill
from sheet_transactions import write_transactions_sheet
from sheet_overview import write_overview_sheet
from sheet_controle import write_controle_sheet
from sheet_knab import write_knab_jaaroverzicht, write_knab_uitschieters


def _betaal_df(df):
    betaal_acc = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Betaal" in lbl), None)
    return df[df["account"] == betaal_acc] if betaal_acc else df


def write_saldo_sheet(ws, df, year):
    euro_fmt   = '€ #,##0.00'
    year_int   = int(year)
    today      = date.today()
    geel_fill  = PatternFill("solid", fgColor=KLEUR_GEEL)

    betaal_acc = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Betaal" in lbl), None)
    spaar_acc  = next((acc for acc, lbl in ACCOUNT_LABELS.items() if "Spaar"  in lbl), None)
    df_betaal  = df[df["account"] == betaal_acc] if betaal_acc else df.iloc[0:0]
    df_spaar   = df[df["account"] == spaar_acc]  if spaar_acc  else df.iloc[0:0]

    headers = ["Maand", "Saldo Betaal per 1e", "Saldo Spaar per 1e", "Totaal"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = _fill(C_HEADER)
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
        cell.alignment = Alignment(horizontal="left" if c == 1 else "center")
    ws.row_dimensions[1].height = 18

    last_spaar = None

    for m in range(1, 13):
        r      = m + 1
        period = f"{year_int}-{m:02d}"
        is_future = (year_int > today.year) or (year_int == today.year and m > today.month)

        ws.cell(row=r, column=1, value=MAANDEN_NL[m]).font = Font(name="Arial", size=9)

        if is_future:
            continue

        def _saldo(val):
            return float(str(val).replace(",", "."))

        df_b = df_betaal[df_betaal["month"] == period].sort_values("date")
        betaal_saldo = _saldo(df_b.iloc[0]["balance_before"]) if not df_b.empty else None

        df_s = df_spaar[df_spaar["month"] == period].sort_values("date")
        if not df_s.empty:
            spaar_saldo = _saldo(df_s.iloc[0]["balance_before"])
            last_spaar  = spaar_saldo
            spaar_estimated = False
        elif last_spaar is not None:
            spaar_saldo     = last_spaar
            spaar_estimated = True
        else:
            spaar_saldo     = None
            spaar_estimated = False

        if betaal_saldo is not None:
            cell = ws.cell(row=r, column=2, value=betaal_saldo)
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="right")

        if spaar_saldo is not None:
            cell = ws.cell(row=r, column=3, value=spaar_saldo)
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="right")
            if spaar_estimated:
                cell.fill = geel_fill

        if betaal_saldo is not None or spaar_saldo is not None:
            cell = ws.cell(row=r, column=4, value=f"=B{r}+C{r}")
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="right")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 16


def save_year_output(df, year):
    wb        = Workbook()
    df_betaal = _betaal_df(df)
    df_onb    = df_betaal[df_betaal["category"] == "Dagelijks Overig"]

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df_betaal.sort_values("date", ascending=False))
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

    t0 = time.time()
    ws_saldo = wb.create_sheet()
    write_saldo_sheet(ws_saldo, df, year)
    ws_saldo.title = "Saldo Overzicht"
    print(f"  Sheet Saldo Overzicht:       {time.time() - t0:.2f}s")

    out_path = os.path.join(OUTPUT_DIR, f"boekhouding_{year}.xlsx")
    t0 = time.time()
    wb.save(out_path)
    print(f"  wb.save():              {time.time() - t0:.2f}s")
    print(f"Opgeslagen: {out_path}")
    return out_path


def save_knab_output(df):
    """Generate boekhouding_knab.xlsx for the Knab zakelijke rekening."""
    wb = Workbook()

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df.sort_values("date", ascending=False))
    ws_tx.title = "Transacties"
    print(f"  Sheet Transacties:           {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_mo = wb.create_sheet()
    write_overview_sheet(ws_mo, df, group_by="month",
                         tx_sheet_name="Transacties", year_label="")
    ws_mo.title = "Maand Overzicht"
    print(f"  Sheet Maand Overzicht:       {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_jr = wb.create_sheet()
    write_knab_jaaroverzicht(ws_jr, df)
    ws_jr.title = "Jaar Vergelijking"
    print(f"  Sheet Jaar Vergelijking:     {time.time() - t0:.2f}s")

    t0 = time.time()
    ws_ui = wb.create_sheet()
    write_knab_uitschieters(ws_ui, df)
    ws_ui.title = "Uitschieters"
    print(f"  Sheet Uitschieters:          {time.time() - t0:.2f}s")

    out_path = os.path.join(OUTPUT_DIR, "boekhouding_knab.xlsx")
    t0 = time.time()
    wb.save(out_path)
    print(f"  wb.save():                   {time.time() - t0:.2f}s")
    print(f"Opgeslagen: {out_path}")
    return out_path


def save_master_output(df):
    wb        = Workbook()
    df_betaal = _betaal_df(df)

    t0 = time.time()
    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df.sort_values("date", ascending=False))
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
