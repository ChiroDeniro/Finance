import sys
sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from networth_config import (
    PEILDATUM, ABN_LABELS, ODS_SALDO,
    BELEGGINGEN, ACTIVA, SCHULDEN, VORDERINGEN_OP_ODS,
)
from config import OUTPUT_DIR

EURO = '€ #,##0.00;-€ #,##0.00'
C_DARK  = "1F4E79"
C_MED   = "2E75B6"
C_SUB   = "BDD7EE"
C_GREEN = "375623"
C_ALT   = "F2F2F2"

INPUT_DIR = Path(__file__).parent.parent / "bookkeeping" / "input"


def _fill(hex_c):
    return PatternFill("solid", fgColor=hex_c)


def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, name="Arial", size=size)


def read_abn_balances():
    balances = {}
    for f in INPUT_DIR.glob("*.TAB"):
        try:
            with open(f, encoding="latin-1") as fh:
                for line in fh:
                    parts = line.strip().split("\t")
                    if len(parts) < 5:
                        continue
                    account, _, date_str, _, bal_str = parts[:5]
                    account = account.strip()
                    date_str = date_str.strip()
                    bal = float(bal_str.strip().replace(",", "."))
                    if account not in balances or date_str > balances[account][0]:
                        balances[account] = (date_str, bal)
        except Exception:
            continue
    return {acc: bal for acc, (_, bal) in balances.items()}


def write_vermogen(ws, abn_balances):
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 16
    ws.freeze_panes = "A2"
    r = [1]

    def row():
        n = r[0]; r[0] += 1; return n

    def title(text):
        rn = row()
        c = ws.cell(rn, 1, text)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=12)
        c.fill = _fill(C_DARK)
        c.alignment = Alignment(vertical="center")
        ws.merge_cells(f"A{rn}:B{rn}")
        ws.row_dimensions[rn].height = 24

    def section(text):
        rn = row()
        for col in (1, 2):
            ws.cell(rn, col).fill = _fill(C_MED)
        c = ws.cell(rn, 1, text)
        c.font = _font(bold=True, color="FFFFFF")

    def item(label, amount, alt=False):
        rn = row()
        lc = ws.cell(rn, 1, f"  {label}")
        ac = ws.cell(rn, 2, amount)
        for cell in (lc, ac):
            cell.font = _font()
            if alt:
                cell.fill = _fill(C_ALT)
        ac.number_format = EURO
        ac.alignment = Alignment(horizontal="right")
        return amount or 0

    def subtotal(label, total, color=C_SUB, text_color="000000", size=10):
        rn = row()
        lc = ws.cell(rn, 1, label)
        ac = ws.cell(rn, 2, total)
        for cell in (lc, ac):
            cell.font = Font(bold=True, color=text_color, name="Arial", size=size)
            cell.fill = _fill(color)
        ac.number_format = EURO
        ac.alignment = Alignment(horizontal="right")
        if size > 10:
            ws.row_dimensions[rn].height = size + 10
        return total

    def gap():
        row()

    title(f"Netto Vermogen — {PEILDATUM}")
    gap()

    total_bezittingen = 0

    section("Privé Bankrekeningen  (automatisch uit TAB export)")
    bank_total = sum(item(ABN_LABELS.get(acc, f"ABN AMRO {acc}"), bal, i % 2 == 0)
                     for i, (acc, bal) in enumerate(abn_balances.items()))
    if not abn_balances:
        item("(geen TAB bestanden gevonden — saldo is €0)", 0)
    subtotal("Subtotaal privé bank", bank_total)
    total_bezittingen += bank_total
    gap()

    section("Zakelijke Rekening (ODS / Knab)")
    ods = item("Knab Zakelijke Rekening", ODS_SALDO)
    subtotal("Subtotaal zakelijk", ods)
    total_bezittingen += ods
    gap()

    section("Beleggingen")
    beleg = sum(item(k, v, i % 2 == 0) for i, (k, v) in enumerate(BELEGGINGEN.items()))
    subtotal("Subtotaal beleggingen", beleg)
    total_bezittingen += beleg
    gap()

    section("Activa (geschatte marktwaarde)")
    activa = sum(item(k, v, i % 2 == 0) for i, (k, v) in enumerate(ACTIVA.items()))
    subtotal("Subtotaal activa", activa)
    total_bezittingen += activa
    gap()

    subtotal("TOTAAL BEZITTINGEN", total_bezittingen, color=C_GREEN, text_color="FFFFFF", size=11)
    gap()

    section("Schulden")
    schulden = sum(item(k, v, i % 2 == 0) for i, (k, v) in enumerate(SCHULDEN.items()))
    subtotal("Totaal schulden", schulden)
    gap()

    netto = total_bezittingen - schulden
    subtotal("NETTO VERMOGEN", netto, color=C_DARK, text_color="FFFFFF", size=13)


def write_vorderingen(ws):
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.freeze_panes = "A2"

    for c, h in enumerate(["Datum", "Omschrijving", "Bedrag", "Status"], 1):
        cell = ws.cell(1, c, h)
        cell.fill = _fill(C_DARK)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 20

    open_total = 0
    for i, v in enumerate(VORDERINGEN_OP_ODS, 2):
        bg = _fill(C_ALT) if i % 2 == 0 else None
        for c, val in enumerate([v["datum"], v["omschrijving"], v["bedrag"], v.get("status", "Open")], 1):
            cell = ws.cell(i, c, val)
            cell.font = _font()
            if bg:
                cell.fill = bg
            if c == 3:
                cell.number_format = EURO
                cell.alignment = Alignment(horizontal="right")
        if v.get("status", "Open") == "Open":
            open_total += v["bedrag"]

    r_tot = len(VORDERINGEN_OP_ODS) + 2
    ws.cell(r_tot, 2, "Totaal openstaand (ODS schuldig aan Chris)")
    ac = ws.cell(r_tot, 3, open_total)
    for c in range(1, 5):
        cell = ws.cell(r_tot, c)
        cell.font = _font(bold=True)
        cell.fill = _fill(C_SUB)
    ac.number_format = EURO
    ac.alignment = Alignment(horizontal="right")

    ws.auto_filter.ref = "A1:D1"


def main():
    abn_balances = read_abn_balances()
    if abn_balances:
        print("✅ ABN AMRO saldi (uit TAB export):")
        for acc, bal in abn_balances.items():
            print(f"   {ABN_LABELS.get(acc, acc)}: €{bal:,.2f}")
    else:
        print("⚠️  Geen TAB bestanden — ABN saldo handmatig invullen in networth_config.py")

    wb = Workbook()
    write_vermogen(wb.active, abn_balances)
    wb.active.title = "Vermogen"

    write_vorderingen(wb.create_sheet("Vorderingen op ODS"))

    out = OUTPUT_DIR / f"networth_{PEILDATUM}.xlsx"
    wb.save(out)

    open_total = sum(v["bedrag"] for v in VORDERINGEN_OP_ODS if v.get("status", "Open") == "Open")
    print(f"💰 Openstaande vordering op ODS: €{open_total:.2f}")
    print(f"📁 Opgeslagen: {out}")


if __name__ == "__main__":
    main()
