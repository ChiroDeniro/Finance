import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from config import (
    OUTPUT_DIR,
    INCOME_CATS, VASTE_LASTEN_CATS, DAGELIJKS_CATS, OVERIG_CATS,
    WHITE,
    C_INC_HDR, C_INC_SUB, C_INC_ROW,
    C_VL_HDR,  C_VL_SUB,  C_VL_ROW,
    C_DAG_HDR, C_DAG_SUB, C_DAG_ROW,
    C_OVR_HDR, C_OVR_SUB, C_OVR_ROW,
    C_SAM_HDR, C_NETTO_POS, C_NETTO_NEG, C_KOSTEN_ROW,
    C_HEADER, C_ALT_ROW,
    format_period, _fill,
)


def write_transactions_sheet(ws, df):
    ws.title = "Transacties"
    ws.freeze_panes = "A2"

    headers = ["Datum", "Rekening", "Omschrijving", "Merchant",
               "Bedrag (EUR)", "Categorie", "Maand"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = _fill(C_HEADER)
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 20

    inc_fill = _fill(C_INC_ROW)
    alt_fill = _fill(C_ALT_ROW)
    euro_fmt = '€ #,##0.00;-€ #,##0.00'

    for r, (_, row) in enumerate(df.iterrows(), 2):
        is_income = row["amount"] > 0
        fill = (inc_fill if is_income
                else (alt_fill if r % 2 == 0 else None))
        data = [
            row["date"].date(), row["account"], row["description"][:80],
            row["merchant"], row["amount"], row["category"], row["month"],
        ]
        for c, val in enumerate(data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center")
            if fill:
                cell.fill = fill
            if c == 5:
                cell.number_format = euro_fmt

    for c, w in enumerate([12, 12, 50, 28, 14, 22, 10], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def write_overview_sheet(ws, df, group_by="month"):
    if group_by == "month":
        ws.title    = "Maand Overzicht"
        periods     = sorted(df["month"].unique())
        pivot_col   = "month"
    else:
        ws.title    = "Jaar Overzicht"
        periods     = sorted(df["year"].unique())
        pivot_col   = "year"

    ws.freeze_panes = "B2"

    n          = len(periods)
    tot_col    = n + 2
    gem_col    = n + 3 if group_by == "month" else None
    last_col   = gem_col or tot_col
    euro_fmt   = '€ #,##0;-€ #,##0'

    EXCLUDE_FROM_OVERVIEW = {"Interne Overboeking"}
    pivot = df[~df["category"].isin(EXCLUDE_FROM_OVERVIEW)].pivot_table(
        index="category", columns=pivot_col,
        values="amount", aggfunc="sum", fill_value=0,
    )

    def _val(cat, period):
        try:
            return float(pivot.loc[cat, period])
        except KeyError:
            return 0.0

    row_num = [2]

    def _header_row():
        r = row_num[0]
        ws.cell(row=r, column=1, value="Categorie")
        ws.cell(row=r, column=1).fill = _fill(C_HEADER)
        ws.cell(row=r, column=1).font = Font(bold=True, color=WHITE, name="Arial", size=10)
        for c, p in enumerate(periods, 2):
            cell = ws.cell(row=r, column=c, value=format_period(p, group_by))
            cell.fill = _fill(C_HEADER)
            cell.font = Font(bold=True, color=WHITE, name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=tot_col, value="Totaal")
        ws.cell(row=r, column=tot_col).fill = _fill(C_HEADER)
        ws.cell(row=r, column=tot_col).font = Font(bold=True, color=WHITE, name="Arial", size=10)
        ws.cell(row=r, column=tot_col).alignment = Alignment(horizontal="center")
        if gem_col:
            ws.cell(row=r, column=gem_col, value="Gemiddeld")
            ws.cell(row=r, column=gem_col).fill = _fill(C_HEADER)
            ws.cell(row=r, column=gem_col).font = Font(bold=True, color=WHITE, name="Arial", size=10)
            ws.cell(row=r, column=gem_col).alignment = Alignment(horizontal="center")
        ws.row_dimensions[r].height = 20
        row_num[0] += 1

    def _section_hdr(label, color):
        r = row_num[0]
        for c in range(1, last_col + 1):
            ws.cell(row=r, column=c).fill = _fill(color)
            ws.cell(row=r, column=c).font = Font(bold=True, color=WHITE, name="Arial", size=9)
        ws.cell(row=r, column=1).value = label
        ws.row_dimensions[r].height = 16
        row_num[0] += 1

    def _data_row(cat, _unused_color=None):
        r = row_num[0]
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=1).font = Font(name="Arial", size=9)
        total = 0.0
        for c, p in enumerate(periods, 2):
            v    = _val(cat, p)
            cell = ws.cell(row=r, column=c, value=round(v))
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="right")
            total += v
        tc = ws.cell(row=r, column=tot_col, value=round(total))
        tc.number_format = euro_fmt
        tc.font = Font(bold=True, name="Arial", size=9)
        tc.alignment = Alignment(horizontal="right")
        if gem_col:
            avg = total / n if n else 0
            gc = ws.cell(row=r, column=gem_col, value=round(avg))
            gc.number_format = euro_fmt
            gc.font = Font(italic=True, name="Arial", size=9)
            gc.alignment = Alignment(horizontal="right")
        row_num[0] += 1

    def _subtotal_row(label, cats, color):
        r     = row_num[0]
        fill  = _fill(color)
        font  = Font(bold=True, color=WHITE, name="Arial", size=9)
        ws.cell(row=r, column=1, value=label).fill = fill
        ws.cell(row=r, column=1).font = font
        grand = 0.0
        for c, p in enumerate(periods, 2):
            v    = sum(_val(cat, p) for cat in cats)
            cell = ws.cell(row=r, column=c, value=round(v))
            cell.number_format = euro_fmt
            cell.fill = fill
            cell.font = font
            grand += v
        tc = ws.cell(row=r, column=tot_col, value=round(grand))
        tc.number_format = euro_fmt
        tc.fill = fill
        tc.font = font
        if gem_col:
            avg = grand / n if n else 0
            gc = ws.cell(row=r, column=gem_col, value=round(avg))
            gc.number_format = euro_fmt
            gc.fill = fill
            gc.font = font
        ws.row_dimensions[r].height = 15
        row_num[0] += 1

    def _blank():
        ws.row_dimensions[row_num[0]].height = 5
        row_num[0] += 1

    def _samenvatting_row(label, cats, row_color, sub_color):
        r    = row_num[0]
        fill = _fill(sub_color)
        font = Font(bold=True, color=WHITE, name="Arial", size=9)
        ws.cell(row=r, column=1, value=label).fill = fill
        ws.cell(row=r, column=1).font = font
        grand = 0.0
        for c, p in enumerate(periods, 2):
            v    = sum(_val(cat, p) for cat in cats)
            cell = ws.cell(row=r, column=c, value=round(v))
            cell.number_format = euro_fmt
            cell.fill = fill
            cell.font = font
            grand += v
        tc = ws.cell(row=r, column=tot_col, value=round(grand))
        tc.number_format = euro_fmt
        tc.fill = fill
        tc.font = font
        if gem_col:
            avg = grand / n if n else 0
            gc = ws.cell(row=r, column=gem_col, value=round(avg))
            gc.number_format = euro_fmt
            gc.fill = fill
            gc.font = font
        row_num[0] += 1

    _header_row()

    _section_hdr("INKOMEN", C_INC_HDR)
    for cat in INCOME_CATS:
        _data_row(cat, C_INC_ROW)
    _subtotal_row("Totaal Inkomen", INCOME_CATS, C_INC_SUB)
    _blank()

    _section_hdr("VASTE LASTEN", C_VL_HDR)
    for cat in VASTE_LASTEN_CATS:
        _data_row(cat, C_VL_ROW)
    _subtotal_row("Totaal Vaste Lasten", VASTE_LASTEN_CATS, C_VL_SUB)
    _blank()

    _section_hdr("DAGELIJKSE UITGAVEN", C_DAG_HDR)
    for cat in DAGELIJKS_CATS:
        _data_row(cat, C_DAG_ROW)
    _subtotal_row("Totaal Dagelijks", DAGELIJKS_CATS, C_DAG_SUB)
    _blank()

    _section_hdr("OVERIG", C_OVR_HDR)
    for cat in OVERIG_CATS:
        _data_row(cat, C_OVR_ROW)
    _subtotal_row("Totaal Overig", OVERIG_CATS, C_OVR_SUB)
    _blank()

    _blank()

    _section_hdr("SAMENVATTING", C_SAM_HDR)
    _samenvatting_row("Totaal Inkomen",      INCOME_CATS,       C_INC_ROW, C_INC_SUB)
    _samenvatting_row("Totaal Vaste Lasten", VASTE_LASTEN_CATS, C_VL_ROW,  C_VL_SUB)
    _samenvatting_row("Totaal Dagelijks",    DAGELIJKS_CATS,    C_DAG_ROW, C_DAG_SUB)
    _samenvatting_row("Totaal Overig",       OVERIG_CATS,       C_OVR_ROW, C_OVR_SUB)

    EXPENSE_CATS = VASTE_LASTEN_CATS + DAGELIJKS_CATS + OVERIG_CATS
    r    = row_num[0]
    fill = _fill(C_KOSTEN_ROW)
    font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=r, column=1, value="Totaal Kosten").fill = fill
    ws.cell(row=r, column=1).font = font
    grand = 0.0
    for c, p in enumerate(periods, 2):
        v    = abs(sum(_val(cat, p) for cat in EXPENSE_CATS))
        cell = ws.cell(row=r, column=c, value=round(v))
        cell.number_format = euro_fmt
        cell.fill = fill
        cell.font = font
        grand += v
    tc = ws.cell(row=r, column=tot_col, value=round(grand))
    tc.number_format = euro_fmt
    tc.fill = fill
    tc.font = font
    if gem_col:
        gc = ws.cell(row=r, column=gem_col, value=round(grand / n if n else 0))
        gc.number_format = euro_fmt
        gc.fill = fill
        gc.font = font
    ws.row_dimensions[r].height = 15
    row_num[0] += 1

    df_real = df[df["category"] != "Interne Overboeking"]
    r    = row_num[0]
    font = Font(bold=True, color=WHITE, name="Arial", size=11)
    ws.cell(row=r, column=1, value="NETTO").font = font
    ws.row_dimensions[r].height = 18
    netto_grand = 0.0
    for c, p in enumerate(periods, 2):
        v     = df_real[df_real[pivot_col] == p]["amount"].sum()
        color = C_NETTO_POS if v >= 0 else C_NETTO_NEG
        cell  = ws.cell(row=r, column=c, value=round(v))
        cell.number_format = euro_fmt
        cell.fill = _fill(color)
        cell.font = font
        netto_grand += v
    netto_color = C_NETTO_POS if netto_grand >= 0 else C_NETTO_NEG
    ws.cell(row=r, column=1).fill = _fill(netto_color)
    tc = ws.cell(row=r, column=tot_col, value=round(netto_grand))
    tc.number_format = euro_fmt
    tc.fill = _fill(netto_color)
    tc.font = font
    if gem_col:
        avg   = netto_grand / n if n else 0
        color = C_NETTO_POS if avg >= 0 else C_NETTO_NEG
        gc    = ws.cell(row=r, column=gem_col, value=round(avg))
        gc.number_format = euro_fmt
        gc.fill = _fill(color)
        gc.font = font

    ws.column_dimensions["A"].width = 26
    for c in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12


def write_jaar_samenvatting_sheet(ws, df, year_label=""):
    ws.title = f"Jaar Samenvatting{' ' + year_label if year_label else ''}"
    ws.freeze_panes = "B2"

    euro_fmt = '€ #,##0;-€ #,##0'
    pct_fmt  = '0.0"%"'

    EXCLUDE = {"Interne Overboeking"}
    df_real = df[~df["category"].isin(EXCLUDE)]

    cat_totals = df_real.groupby("category")["amount"].sum()

    def _total(cats):
        return sum(cat_totals.get(c, 0.0) for c in cats)

    totaal_inkomen = _total(INCOME_CATS)

    def _pct(val):
        return round(val / totaal_inkomen * 100, 1) if totaal_inkomen else 0.0

    row_num = [2]

    def _hdr():
        r = row_num[0]
        for c, h in enumerate(["Categorie", "Totaal", "% van Inkomen"], 1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.fill = _fill(C_HEADER)
            cell.font = Font(bold=True, color=WHITE, name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left")
        ws.row_dimensions[r].height = 20
        row_num[0] += 1

    def _sec_hdr(label, color):
        r = row_num[0]
        for c in range(1, 4):
            ws.cell(row=r, column=c).fill = _fill(color)
            ws.cell(row=r, column=c).font = Font(bold=True, color=WHITE, name="Arial", size=9)
        ws.cell(row=r, column=1).value = label
        ws.row_dimensions[r].height = 16
        row_num[0] += 1

    def _data_row(cat, _unused_color=None):
        r   = row_num[0]
        val = cat_totals.get(cat, 0.0)
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=1).font = Font(name="Arial", size=9)
        tc = ws.cell(row=r, column=2, value=round(val))
        tc.number_format = euro_fmt
        tc.font = Font(name="Arial", size=9)
        tc.alignment = Alignment(horizontal="right")
        pc = ws.cell(row=r, column=3, value=_pct(val))
        pc.number_format = pct_fmt
        pc.font = Font(italic=True, name="Arial", size=9)
        pc.alignment = Alignment(horizontal="right")
        row_num[0] += 1

    def _subtotal(label, cats, color):
        r    = row_num[0]
        fill = _fill(color)
        font = Font(bold=True, color=WHITE, name="Arial", size=9)
        val  = _total(cats)
        ws.cell(row=r, column=1, value=label).fill = fill
        ws.cell(row=r, column=1).font = font
        tc = ws.cell(row=r, column=2, value=round(val))
        tc.number_format = euro_fmt
        tc.fill = fill
        tc.font = font
        pc = ws.cell(row=r, column=3, value=_pct(val))
        pc.number_format = pct_fmt
        pc.fill = fill
        pc.font = font
        ws.row_dimensions[r].height = 15
        row_num[0] += 1

    def _blank():
        ws.row_dimensions[row_num[0]].height = 5
        row_num[0] += 1

    def _sam_row(label, cats, sub_color):
        r    = row_num[0]
        fill = _fill(sub_color)
        font = Font(bold=True, color=WHITE, name="Arial", size=9)
        val  = _total(cats)
        ws.cell(row=r, column=1, value=label).fill = fill
        ws.cell(row=r, column=1).font = font
        tc = ws.cell(row=r, column=2, value=round(val))
        tc.number_format = euro_fmt
        tc.fill = fill
        tc.font = font
        pc = ws.cell(row=r, column=3, value=_pct(val))
        pc.number_format = pct_fmt
        pc.fill = fill
        pc.font = font
        row_num[0] += 1

    _hdr()

    _sec_hdr("INKOMEN", C_INC_HDR)
    for cat in INCOME_CATS:
        _data_row(cat, C_INC_ROW)
    _subtotal("Totaal Inkomen", INCOME_CATS, C_INC_SUB)
    _blank()

    _sec_hdr("VASTE LASTEN", C_VL_HDR)
    for cat in VASTE_LASTEN_CATS:
        _data_row(cat, C_VL_ROW)
    _subtotal("Totaal Vaste Lasten", VASTE_LASTEN_CATS, C_VL_SUB)
    _blank()

    _sec_hdr("DAGELIJKSE UITGAVEN", C_DAG_HDR)
    for cat in DAGELIJKS_CATS:
        _data_row(cat, C_DAG_ROW)
    _subtotal("Totaal Dagelijks", DAGELIJKS_CATS, C_DAG_SUB)
    _blank()

    _sec_hdr("OVERIG", C_OVR_HDR)
    for cat in OVERIG_CATS:
        _data_row(cat, C_OVR_ROW)
    _subtotal("Totaal Overig", OVERIG_CATS, C_OVR_SUB)
    _blank()
    _blank()

    _sec_hdr("SAMENVATTING", C_SAM_HDR)
    _sam_row("Totaal Inkomen",      INCOME_CATS,       C_INC_SUB)
    _sam_row("Totaal Vaste Lasten", VASTE_LASTEN_CATS, C_VL_SUB)
    _sam_row("Totaal Dagelijks",    DAGELIJKS_CATS,    C_DAG_SUB)
    _sam_row("Totaal Overig",       OVERIG_CATS,       C_OVR_SUB)

    r     = row_num[0]
    netto = df_real["amount"].sum()
    color = C_NETTO_POS if netto >= 0 else C_NETTO_NEG
    fill  = _fill(color)
    font  = Font(bold=True, color=WHITE, name="Arial", size=11)
    ws.cell(row=r, column=1, value="NETTO").fill = fill
    ws.cell(row=r, column=1).font = font
    tc = ws.cell(row=r, column=2, value=round(netto))
    tc.number_format = euro_fmt
    tc.fill = fill
    tc.font = font
    pc = ws.cell(row=r, column=3, value=_pct(netto))
    pc.number_format = pct_fmt
    pc.fill = fill
    pc.font = font
    ws.row_dimensions[r].height = 18

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16


def write_controle_sheet(ws, df, year_label=""):
    ws.title = f"Controle{' ' + year_label if year_label else ''}"
    ws.freeze_panes = "B2"
    euro_fmt = '€ #,##0;-€ #,##0'

    periods = sorted(df["month"].unique())
    df_real = df[df["category"] != "Interne Overboeking"]
    df_int  = df[df["category"] == "Interne Overboeking"]

    def _s(frame, period):
        return float(frame[frame["month"] == period]["amount"].sum())

    headers = [
        "Maand", "Inkomen (cat)", "Uitgaven (cat)", "Netto Alle", "Interne OB",
    ]

    r = 1
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.fill = _fill(C_HEADER)
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
        cell.alignment = Alignment(horizontal="center" if c > 1 else "left")
    ws.row_dimensions[r].height = 18

    EXPENSE_CATS = VASTE_LASTEN_CATS + DAGELIJKS_CATS + OVERIG_CATS
    running = [0.0] * (len(headers) - 1)

    for period in periods:
        r += 1
        inc_cat = _s(df_real[df_real["category"].isin(INCOME_CATS)], period)
        uit_cat = _s(df_real[df_real["category"].isin(EXPENSE_CATS)], period)
        netto   = _s(df_real, period)
        interne = _s(df_int, period)

        vals = [inc_cat, uit_cat, netto, interne]
        ws.cell(row=r, column=1, value=format_period(period, "month")).font = Font(name="Arial", size=9)
        for c, v in enumerate(vals, 2):
            cell = ws.cell(row=r, column=c, value=round(v))
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="right")
        for i, v in enumerate(vals):
            running[i] += v

    r += 1
    ws.cell(row=r, column=1, value="Totaal").font = Font(bold=True, name="Arial", size=9)
    for c, v in enumerate(running, 2):
        cell = ws.cell(row=r, column=c, value=round(v))
        cell.number_format = euro_fmt
        cell.font = Font(bold=True, name="Arial", size=9)
        cell.alignment = Alignment(horizontal="right")

    ws.column_dimensions["A"].width = 12
    for c in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15


def save_output(df, year_label=""):
    suffix = f" {year_label}" if year_label else ""
    wb = Workbook()

    ws_tx = wb.active
    write_transactions_sheet(ws_tx, df)
    ws_tx.title = f"Transacties{suffix}"

    ws_mo = wb.create_sheet()
    write_overview_sheet(ws_mo, df, group_by="month")
    ws_mo.title = f"Maand Overzicht{suffix}"

    ws_js = wb.create_sheet()
    write_jaar_samenvatting_sheet(ws_js, df, year_label=year_label)

    ws_ctrl = wb.create_sheet()
    write_controle_sheet(ws_ctrl, df, year_label=year_label)

    if year_label:
        filename = f"boekhouding_{year_label}.xlsx"
    else:
        now = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"boekhouding_{now}.xlsx"

    out_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(out_path)
    print(f"Opgeslagen: {out_path}")
    return out_path
