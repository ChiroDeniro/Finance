"""
ABN AMRO Bookkeeping Processor
================================
Drop your ABN AMRO Excel (TXT) export into the /input folder, then run:
    python process.py

Output: /output/boekhouding_YYYYMMDD_HHMM.xlsx  with:
  - Sheet 1: All transactions with categories
  - Sheet 2: Monthly overview per category (income / expenses split)
  - Sheet 3: Unknowns to review (transactions that need a category)
"""

import os
import re
import sys
import glob
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
RULES_FILE = os.path.join(BASE_DIR, "rules.xlsx")

os.makedirs(INPUT_DIR,  exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Colors ───────────────────────────────────────────────────────────────────
C_HEADER   = "1F4E79"   # dark blue
C_ALT_ROW  = "EEF3FA"   # light blue
C_UNKNOWN  = "FFF2CC"   # yellow – needs attention
C_INCOME   = "E2EFDA"   # green – income
C_TOTAL    = "D6E4BC"   # darker green – totals
C_EXPENSE  = "FDECEA"   # light red – expense total
WHITE      = "FFFFFF"

# ── Step 1: Find & load the TAB file(s) ──────────────────────────────────────
def find_input_files():
    patterns = ["*.TAB", "*.tab", "*.txt", "*.TXT"]
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(INPUT_DIR, p)))
    if not files:
        print(f"❌  Geen invoerbestanden gevonden in {INPUT_DIR}/")
        print("    Download je ABN AMRO afschrift als 'Excel (TXT)' en zet het daar neer.")
        sys.exit(1)
    print(f"✅  {len(files)} invoerbestand(en) gevonden:")
    for f in files:
        print(f"    {os.path.basename(f)}")
    return files


def validate_tab_file(df, filepath):
    """Validate that df matches the expected ABN AMRO TAB export format."""
    name = os.path.basename(filepath)
    errors = []

    if len(df.columns) != 8:
        errors.append(f"verwacht 8 kolommen, maar gevonden {len(df.columns)}")

    if len(df) == 0:
        errors.append("bestand bevat geen transacties")
    else:
        # Date column must look like YYYYMMDD (8 digits)
        for v in df["date"].dropna().head(3):
            if not re.match(r'^\d{8}$', str(v).strip()):
                errors.append(
                    f"datum-kolom niet in JJJJMMDD-formaat (voorbeeld: '{v}')"
                )
                break

        # Amount column must use Dutch decimal notation (comma as decimal)
        if not any("," in str(v) for v in df["amount"].dropna().head(5)):
            errors.append(
                "bedrag-kolom mist komma als decimaalteken "
                "(verwacht Nederlandstalige notatie, bijv. '-19,42')"
            )

    if errors:
        print(f"\n❌  Ongeldig bestandsformaat: {name}")
        for e in errors:
            print(f"    • {e}")
        print("    Verwacht: ABN AMRO 'Excel (TXT)' export – 8 tab-gescheiden kolommen:")
        print("    rekening | valuta | datum (JJJJMMDD) | saldo_voor | saldo_na "
              "| valutadatum | bedrag | omschrijving")
        sys.exit(1)


def load_transactions(files):
    dfs = []
    for f in files:
        df = pd.read_csv(
            f, sep="\t", header=None, encoding="utf-8",
            names=["account", "currency", "date", "balance_before",
                   "balance_after", "value_date", "amount", "description"],
            dtype=str
        )
        validate_tab_file(df, f)
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)

    # Deduplicate (in case both files have same account)
    df = df.drop_duplicates()

    # Clean up types
    df["date"]        = pd.to_datetime(df["date"], format="%Y%m%d")
    df["month"]       = df["date"].dt.to_period("M").astype(str)   # e.g. "2026-03"
    df["amount"]      = df["amount"].str.replace(",", ".").astype(float)
    df["description"] = df["description"].str.strip()

    # Extract a readable merchant name from the description
    df["merchant"] = df["description"].apply(extract_merchant)

    print(f"✅  {len(df)} transacties geladen  "
          f"({df['date'].min().date()} → {df['date'].max().date()})")
    return df


def extract_merchant(desc):
    """Pull a clean merchant/counterparty name out of the raw description."""
    desc = str(desc).strip()
    # BEA (pin/contactless): "BEA, Apple Pay   Merchant Name,PAS..."
    m = re.search(r"BEA,.*?\s{2,}(.+?),PAS", desc)
    if m:
        return m.group(1).strip()
    # SEPA transfers: "/NAME/Counterparty Name/"
    m = re.search(r"/NAME/([^/]+)", desc)
    if m:
        return m.group(1).strip()
    # ABN own costs
    if "ABN AMRO" in desc.upper():
        return "ABN AMRO Bank"
    # Fallback: first 40 chars
    return desc[:40]


# ── Step 2: Load categorisation rules ────────────────────────────────────────
def load_rules():
    """
    rules.xlsx has columns: keyword | category
    keyword is matched (case-insensitive, partial) against the merchant name.
    Order matters – first match wins.
    """
    if not os.path.exists(RULES_FILE):
        print(f"⚠️   Geen rules.xlsx gevonden op {RULES_FILE}")
        print("    Voer uit met --create-rules om een startbestand te maken.")
        return []

    df = pd.read_excel(RULES_FILE, sheet_name="Rules")
    rules = []
    for _, row in df.iterrows():
        kw  = str(row.get("keyword",  "")).strip()
        cat = str(row.get("category", "")).strip()
        if kw and cat:
            rules.append((kw.lower(), cat))
    print(f"✅  {len(rules)} categorisatieregels geladen")
    return rules


def categorise(merchant, rules):
    merchant_lower = merchant.lower()
    for keyword, category in rules:
        if keyword in merchant_lower:
            return category
    return "⚠️ Onbekend"


def apply_categories(df, rules):
    df["category"] = df["merchant"].apply(lambda m: categorise(m, rules))
    n_unknown = (df["category"] == "⚠️ Onbekend").sum()
    n_total   = len(df)
    pct = 100 * (n_total - n_unknown) / n_total if n_total else 0
    print(f"✅  {n_total - n_unknown}/{n_total} transacties gecategoriseerd  ({pct:.0f}%)")
    if n_unknown:
        print(f"⚠️   {n_unknown} transacties hebben nog een categorie nodig → zie het 'Onbekend' tabblad")
    return df


# ── Step 3: Build output Excel ────────────────────────────────────────────────
def style_header(ws, row, cols, bg=C_HEADER, fg=WHITE, bold=True, height=20):
    fill = PatternFill("solid", fgColor=bg)
    font = Font(bold=bold, color=fg, name="Arial", size=10)
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height


def thin_border():
    side = Side(style="thin", color="AAAAAA")
    return Border(left=side, right=side, top=side, bottom=side)


def write_transactions_sheet(ws, df):
    ws.title = "Transacties"
    ws.freeze_panes = "A2"

    headers = ["Datum", "Rekening", "Omschrijving", "Merchant",
               "Bedrag (€)", "Categorie", "Maand"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, len(headers))

    alt_fill = PatternFill("solid", fgColor=C_ALT_ROW)
    unk_fill = PatternFill("solid", fgColor=C_UNKNOWN)
    inc_fill = PatternFill("solid", fgColor=C_INCOME)
    euro_fmt = '#,##0.00 "€";[Red]-#,##0.00 "€"'

    for r, (_, row) in enumerate(df.iterrows(), 2):
        is_unknown = row["category"] == "⚠️ Onbekend"
        is_income  = row["amount"] > 0
        fill = (unk_fill if is_unknown
                else (inc_fill if is_income
                      else (alt_fill if r % 2 == 0 else None)))

        data = [
            row["date"].date(),
            row["account"],
            row["description"][:80],
            row["merchant"],
            row["amount"],
            row["category"],
            row["month"],
        ]
        for c, val in enumerate(data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center")
            if fill:
                cell.fill = fill
            if c == 5:
                cell.number_format = euro_fmt

    widths = [12, 12, 50, 28, 14, 22, 10]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def write_overview_sheet(ws, df):
    ws.title = "Overzicht per Maand"
    ws.freeze_panes = "B2"

    months     = sorted(df["month"].unique())
    categories = sorted(df[df["category"] != "⚠️ Onbekend"]["category"].unique())
    income_cats  = [c for c in categories if df[df["category"] == c]["amount"].sum() > 0]
    expense_cats = [c for c in categories if c not in income_cats]

    euro_fmt = '#,##0.00 "€";[Red]-#,##0.00 "€"'
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    hdr_font = Font(bold=True, color=WHITE, name="Arial", size=10)
    tot_fill = PatternFill("solid", fgColor=C_TOTAL)
    exp_fill = PatternFill("solid", fgColor=C_EXPENSE)
    alt_fill = PatternFill("solid", fgColor=C_ALT_ROW)
    inc_sec  = PatternFill("solid", fgColor="375623")  # dark forest green
    exp_sec  = PatternFill("solid", fgColor="843C0C")  # dark red-brown
    inc_sub  = PatternFill("solid", fgColor="70AD47")  # medium green

    total_col = len(months) + 2

    # ── header row ──────────────────────────────────────────────────────────
    ws.cell(row=1, column=1, value="Categorie")
    ws.cell(row=1, column=1).fill = hdr_fill
    ws.cell(row=1, column=1).font = hdr_font
    for c, m in enumerate(months, 2):
        cell = ws.cell(row=1, column=c, value=m)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")
    ws.cell(row=1, column=total_col, value="Totaal")
    ws.cell(row=1, column=total_col).fill = hdr_fill
    ws.cell(row=1, column=total_col).font = hdr_font

    pivot = df.pivot_table(index="category", columns="month",
                           values="amount", aggfunc="sum", fill_value=0)

    def _section_header(row_num, label, fill):
        for c in range(1, total_col + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.fill = fill
            cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
        ws.cell(row=row_num, column=1).value = label
        ws.row_dimensions[row_num].height = 16

    def _cat_row(row_num, cat, is_income):
        row_fill = alt_fill if row_num % 2 == 0 else None
        ws.cell(row=row_num, column=1, value=cat).font = Font(
            name="Arial", size=9, bold=is_income)
        if row_fill:
            ws.cell(row=row_num, column=1).fill = row_fill
        row_total = 0
        for c, m in enumerate(months, 2):
            val = pivot.loc[cat, m] if (cat in pivot.index and m in pivot.columns) else 0
            cell = ws.cell(row=row_num, column=c, value=round(val, 2))
            cell.number_format = euro_fmt
            cell.font = Font(name="Arial", size=9)
            if row_fill:
                cell.fill = row_fill
            row_total += val
        tot_cell = ws.cell(row=row_num, column=total_col, value=round(row_total, 2))
        tot_cell.number_format = euro_fmt
        tot_cell.font = Font(name="Arial", size=9, bold=True)
        tot_cell.fill = tot_fill if is_income else exp_fill

    row_num = 2

    # ── INKOMSTEN section ────────────────────────────────────────────────────
    _section_header(row_num, "INKOMSTEN", inc_sec)
    row_num += 1
    for cat in income_cats:
        _cat_row(row_num, cat, is_income=True)
        row_num += 1

    # income subtotal
    inc_grand = 0
    for c in range(1, total_col + 1):
        cell = ws.cell(row=row_num, column=c)
        cell.fill = inc_sub
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
    ws.cell(row=row_num, column=1).value = "Totaal inkomsten"
    for c, m in enumerate(months, 2):
        val = df[(df["month"] == m) & (df["category"].isin(income_cats))]["amount"].sum()
        cell = ws.cell(row=row_num, column=c, value=round(val, 2))
        cell.number_format = euro_fmt
        cell.fill = inc_sub
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
        inc_grand += val
    tot_cell = ws.cell(row=row_num, column=total_col, value=round(inc_grand, 2))
    tot_cell.number_format = euro_fmt
    tot_cell.fill = inc_sub
    tot_cell.font = Font(bold=True, color=WHITE, name="Arial", size=9)
    row_num += 2  # subtotal row + blank separator

    # ── UITGAVEN section ─────────────────────────────────────────────────────
    _section_header(row_num, "UITGAVEN", exp_sec)
    row_num += 1
    for cat in expense_cats:
        _cat_row(row_num, cat, is_income=False)
        row_num += 1

    # ── TOTAAL NETTO ─────────────────────────────────────────────────────────
    ws.cell(row=row_num, column=1, value="TOTAAL NETTO")
    ws.cell(row=row_num, column=1).fill = hdr_fill
    ws.cell(row=row_num, column=1).font = Font(bold=True, color=WHITE, name="Arial")
    for c, m in enumerate(months, 2):
        val = df[df["month"] == m]["amount"].sum()
        cell = ws.cell(row=row_num, column=c, value=round(val, 2))
        cell.number_format = euro_fmt
        cell.font = Font(bold=True, name="Arial", color=WHITE)
        cell.fill = hdr_fill
    net_total = df["amount"].sum()
    tot_cell = ws.cell(row=row_num, column=total_col, value=round(net_total, 2))
    tot_cell.number_format = euro_fmt
    tot_cell.font = Font(bold=True, name="Arial", color=WHITE)
    tot_cell.fill = hdr_fill

    # ── column widths ────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 28
    for c in range(2, total_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14


def write_unknowns_sheet(ws, df):
    ws.title = "⚠️ Onbekend"
    unknowns = df[df["category"] == "⚠️ Onbekend"].copy()

    if unknowns.empty:
        ws.cell(row=1, column=1, value="✅ Geen onbekende transacties!")
        return

    headers = ["Datum", "Merchant", "Bedrag (€)", "Omschrijving",
               "Jouw Categorie (vul in)"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, len(headers), bg="C55A11")

    unk_fill = PatternFill("solid", fgColor=C_UNKNOWN)
    euro_fmt = '#,##0.00 "€";[Red]-#,##0.00 "€"'

    for r, (_, row) in enumerate(unknowns.iterrows(), 2):
        vals = [row["date"].date(), row["merchant"],
                row["amount"], row["description"][:80], ""]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = unk_fill
            cell.font = Font(name="Arial", size=9)
            if c == 3:
                cell.number_format = euro_fmt

    widths = [12, 28, 14, 55, 28]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    ws.cell(row=1, column=5).font = Font(bold=True, color="FFFFFF",
                                          name="Arial", size=10)


def save_output(df):
    wb = Workbook()
    ws1 = wb.active
    write_transactions_sheet(ws1, df)
    ws2 = wb.create_sheet()
    write_overview_sheet(ws2, df)
    ws3 = wb.create_sheet()
    write_unknowns_sheet(ws3, df)

    now = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = os.path.join(OUTPUT_DIR, f"boekhouding_{now}.xlsx")
    wb.save(out_path)
    print(f"✅  Opgeslagen: {out_path}")
    return out_path


# ── Step 4: Create starter rules.xlsx ────────────────────────────────────────
def create_starter_rules():
    """Generate a starter rules.xlsx with sensible Dutch categories."""
    starter = [
        # keyword              category
        ("albert heijn",       "Boodschappen"),
        ("dirk",               "Boodschappen"),
        ("jumbo",              "Boodschappen"),
        ("lidl",               "Boodschappen"),
        ("ah ",                "Boodschappen"),
        ("tuda fruta",         "Boodschappen"),
        ("takeaway",           "Eten & Drinken"),
        ("deliveroo",          "Eten & Drinken"),
        ("ming kee",           "Eten & Drinken"),
        ("cn bagijn",          "Eten & Drinken"),
        ("koffiehuis",         "Eten & Drinken"),
        ("crow-bar",           "Eten & Drinken"),
        ("ls nine bar",        "Eten & Drinken"),
        ("cafe",               "Eten & Drinken"),
        ("tabakshop",          "Eten & Drinken"),
        ("gogo tabak",         "Eten & Drinken"),
        ("heilzaam",           "Gezondheid"),
        ("apotheek",           "Gezondheid"),
        ("unive",              "Verzekeringen"),
        ("sportcity",          "Sport"),
        ("david lloyd",        "Sport"),
        ("ns groep",           "OV & Reizen"),
        ("ns reizigers",       "OV & Reizen"),
        ("bck*ns",             "OV & Reizen"),
        ("den haag cs",        "OV & Reizen"),
        ("youfone",            "Telefoon"),
        ("belastingdienst",    "Vaste Lasten"),
        ("abn amro",           "Bankkosten"),
        ("hogeschool",         "Inkomen"),
        ("salaris",            "Inkomen"),
        ("huiver",             "Kamerhuur"),
        ("sparen",             "Sparen"),
        ("bol.com",            "Diversen"),
        ("media markt",        "Diversen"),
        ("laurenskerk",        "Diversen"),
        ("primera",            "Diversen"),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "Rules"

    ws.cell(row=1, column=1, value="keyword")
    ws.cell(row=1, column=2, value="category")
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    hdr_font = Font(bold=True, color=WHITE, name="Arial")
    for c in [1, 2]:
        ws.cell(row=1, column=c).fill = hdr_fill
        ws.cell(row=1, column=c).font = hdr_font

    for r, (kw, cat) in enumerate(starter, 2):
        ws.cell(row=r, column=1, value=kw)
        ws.cell(row=r, column=2, value=cat)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30

    ws2 = wb.create_sheet("Instructies")
    instructions = [
        ("Hoe werkt rules.xlsx?", ""),
        ("", ""),
        ("keyword",  "De tekst die gezocht wordt in de merchant naam (niet hoofdlettergevoelig)"),
        ("category", "De post/categorie die je aan die transactie wilt geven"),
        ("", ""),
        ("Tips:", ""),
        ("• Volgorde telt",     "De EERSTE match wint. Zet specifieke regels bovenaan."),
        ("• Deelwoorden",       "\"albert heijn\" matcht ook \"Albert Heijn 1870\""),
        ("• Inkomen",           "Gebruik categorie \"Inkomen\" voor positieve bedragen"),
        ("• Nieuw toevoegen",   "Voeg gewoon een rij toe aan het Rules tabblad"),
    ]
    for r, (a, b) in enumerate(instructions, 1):
        ws2.cell(row=r, column=1, value=a).font = Font(bold=(r == 1), name="Arial")
        ws2.cell(row=r, column=2, value=b).font = Font(name="Arial")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 60

    wb.save(RULES_FILE)
    print(f"✅  Starter rules.xlsx aangemaakt: {RULES_FILE}")
    print("    Open het bestand, pas de categorieën aan en voer process.py opnieuw uit.")


# ── Step 5: Migrate existing rules.xlsx to new category names ────────────────
CATEGORY_MIGRATION = {
    "inkomsten":                "Inkomen",
    "eten & drinken buiten":    "Eten & Drinken",
    "koffie & café":            "Eten & Drinken",
    "koffie & cafe":            "Eten & Drinken",
    "ov & reizen":              "OV & Reizen",
    "sport & fitness":          "Sport",
    "telefoon & internet":      "Telefoon",
    "huur & wonen":             "Kamerhuur",
    "belastingen":              "Vaste Lasten",
    "online winkelen":          "Diversen",
    "gezondheid & apotheek":    "Gezondheid",
    "tabak":                    "Eten & Drinken",
    "cultuur & entertainment":  "Diversen",
}


def migrate_rules():
    """Update existing rules.xlsx category names to the new kasboek style."""
    if not os.path.exists(RULES_FILE):
        print("⚠️   Geen rules.xlsx gevonden – niets te migreren.")
        return

    wb = load_workbook(RULES_FILE)
    if "Rules" not in wb.sheetnames:
        print("⚠️   Geen 'Rules' tabblad gevonden in rules.xlsx.")
        return

    ws = wb["Rules"]
    changed = 0
    for row in ws.iter_rows(min_row=2):
        cat_cell = row[1]  # column B = category
        old_val  = str(cat_cell.value or "").strip()
        new_val  = CATEGORY_MIGRATION.get(old_val.lower())
        if new_val and new_val != old_val:
            cat_cell.value = new_val
            changed += 1

    wb.save(RULES_FILE)
    print(f"✅  rules.xlsx bijgewerkt: {changed} categorie(ën) hernoemd")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if "--migrate-rules" in sys.argv:
        migrate_rules()
        return

    if "--create-rules" in sys.argv or not os.path.exists(RULES_FILE):
        print("📋  Starter rules.xlsx aanmaken …")
        create_starter_rules()
        if "--create-rules" in sys.argv:
            return

    print("\n🏦  ABN AMRO Bookkeeping Processor")
    print("=" * 40)

    files = find_input_files()
    df    = load_transactions(files)
    rules = load_rules()
    df    = apply_categories(df, rules)

    out = save_output(df)

    months_sorted = sorted(df["month"].unique())
    print("\n📊  Maandoverzicht:")
    print(f"    {'Maand':<10}  {'Inkomsten':>12}  {'Uitgaven':>12}  {'Netto':>12}")
    print("    " + "─" * 54)
    for m in months_sorted:
        mdf = df[df["month"] == m]
        inc = mdf[mdf["amount"] > 0]["amount"].sum()
        exp = mdf[mdf["amount"] < 0]["amount"].sum()
        net = mdf["amount"].sum()
        print(f"    {m:<10}  €{inc:>10,.2f}  €{exp:>10,.2f}  €{net:>10,.2f}")
    print("    " + "─" * 54)
    t_inc = df[df["amount"] > 0]["amount"].sum()
    t_exp = df[df["amount"] < 0]["amount"].sum()
    t_net = df["amount"].sum()
    print(f"    {'Totaal':<10}  €{t_inc:>10,.2f}  €{t_exp:>10,.2f}  €{t_net:>10,.2f}")
    print(f"\n    Output → {out}\n")


if __name__ == "__main__":
    main()
