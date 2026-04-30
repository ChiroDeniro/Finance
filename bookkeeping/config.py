import os
import json
from openpyxl.styles import PatternFill

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR   = os.path.join(BASE_DIR, "input")
RULES_FILE  = os.path.join(BASE_DIR, "rules.xlsx")
CONFIG_JSON = os.path.join(BASE_DIR, "config.json")

_DRIVE_DIR = r"C:\Users\chris\Documents\Finance\Kasboek"
OUTPUT_DIR = _DRIVE_DIR if os.path.isdir(_DRIVE_DIR) else os.path.join(BASE_DIR, "output")

os.makedirs(INPUT_DIR,  exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Category definitions ──────────────────────────────────────────────────────
INCOME_CATS = [
    "Salaris",
    "DUO / Studiefinanciering",
    "Zorgtoeslag",
    "Familie & Giften",
    "ZZP Opname",
    "Inkomsten Overig",
]
VASTE_LASTEN_CATS = [
    "Huur",
    "Inclusief Huur",
    "Zorgverzekering",
    "Telefoon & Internet",
    "Bankkosten",
    "Abonnementen",
    "Sport & Fitness",
    "Onderhoud",
]
DAGELIJKS_CATS = [
    "Boodschappen",
    "Eten & Drinken",
    "Uitgaan",
    "OV & Reizen",
    "Kleding",
    "Kapper",
    "Gezondheid",
    "WbW",
    "Cultuur & Entertainment",
    "Studie",
    "Dagelijks Overig",
]
OVERIG_CATS = ["Sparen", "Beleggen"]

# Categories excluded from Totaal Kosten in SAMENVATTING
KOSTEN_EXCLUDE = frozenset({"Beleggen"})

ALL_KNOWN_CATS = INCOME_CATS + VASTE_LASTEN_CATS + DAGELIJKS_CATS + OVERIG_CATS

OWN_ACCOUNTS = {"536542171", "844835730"}

# ── Colors ────────────────────────────────────────────────────────────────────
WHITE = "FFFFFF"

C_INC_HDR  = "375623"
C_INC_SUB  = "548235"
C_INC_ROW  = "E2EFDA"

C_VL_HDR   = "1F4E79"
C_VL_SUB   = "2E75B6"
C_VL_ROW   = "DEEAF1"

C_DAG_HDR  = "843C0C"
C_DAG_SUB  = "C55A11"
C_DAG_ROW  = "FCE4D6"

C_OVR_HDR  = "595959"
C_OVR_SUB  = "808080"
C_OVR_ROW  = "F2F2F2"

C_IO_HDR   = "4A4A4A"
C_IO_ROW   = "DEDEDE"

C_SAM_HDR   = "1F4E79"
C_NETTO_POS = "375623"
C_NETTO_NEG = "C00000"
C_KOSTEN_ROW = "FFD6E0"

C_HEADER   = "1F4E79"
C_ALT_ROW  = "EEF3FA"

# ── Dutch month labels ────────────────────────────────────────────────────────
MAANDEN_NL = {
    1: "Jan", 2: "Feb", 3: "Mrt", 4: "Apr", 5: "Mei",  6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Dec",
}


def load_config():
    if os.path.exists(CONFIG_JSON):
        with open(CONFIG_JSON, encoding="utf-8") as f:
            return json.load(f)
    return {"accounts": {}, "iban_rules": {}}


def save_config(cfg):
    with open(CONFIG_JSON, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def get_iban_rules():
    return list(load_config().get("iban_rules", {}).items())


def format_period(period_str, group_by):
    if group_by == "year":
        return period_str
    yr, mo = period_str.split("-")
    return f"{MAANDEN_NL[int(mo)]} '{yr[2:]}"


def dutch_euros(value):
    return "€" + f"{int(round(value)):,}".replace(",", ".")


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)
