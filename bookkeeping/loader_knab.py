"""
Knab zakelijke rekening CSV loader.

Knab export format (semicolon-separated, UTF-8 BOM):
  Row 0: "KNAB EXPORT..." (skip)
  Row 1: column headers
  Row 2+: data

Amount sign: CreditDebet column ("C" = positive, "D" = negative).
Date format: DD-MM-YYYY.
"""

import os
import glob
import pandas as pd

from config import INPUT_DIR, INCOME_CATS, load_config, save_config

KNAB_DIR = os.path.join(INPUT_DIR, "ODS")

# Counterparty IBAN → category for this ZZP business account.
# IBAN-first rules override all merchant keyword matching.
KNAB_COUNTERPARTY_RULES = {
    "NL98ABNA0536542171": "Interne Overboeking",   # eigen ABN betaalrekening
    "NL20ABNA0894212354": "ZZP Inkomen",            # Hogeschool der Kunsten (facturen)
    "NL36INGB0003445588": "Inkomsten Overig",        # Belastingdienst teruggave
    "NL86INGB0002445588": "Belasting",               # Belastingdienst afdracht BTW/IB
}


def find_knab_files():
    if not os.path.isdir(KNAB_DIR):
        return []
    seen, files = set(), []
    for p in ["*.csv", "*.CSV"]:
        for f in glob.glob(os.path.join(KNAB_DIR, p)):
            key = os.path.normcase(os.path.abspath(f))
            if key not in seen:
                seen.add(key)
                files.append(f)
    return files


def load_knab_transactions(files):
    if not files:
        return None

    dfs = []
    for f in files:
        df = _load_one(f)
        if df is not None:
            dfs.append(df)

    if not dfs:
        return None

    df = pd.concat(dfs, ignore_index=True)
    rijen_voor = len(df)
    df = df.drop_duplicates(
        subset=["account", "date", "amount", "description"],
        keep="first",
    )
    n_dubbel = rijen_voor - len(df)
    if n_dubbel > 0:
        print(f"  Deduplicatie Knab: {n_dubbel} dubbele rijen verwijderd")

    df = df.sort_values("date").reset_index(drop=True)
    print(f"Knab geladen: {len(df)} transacties  "
          f"({df['date'].min().date()} — {df['date'].max().date()})")

    _update_config_knab(df["account"].unique().tolist())
    return df


def apply_knab_categories(df, rules):
    """IBAN-first categorization, then merchant keyword match, fallback Zakelijke Kosten."""

    def _cat(row):
        iban = str(row.get("counterparty_iban", "")).strip()
        if iban in KNAB_COUNTERPARTY_RULES:
            return KNAB_COUNTERPARTY_RULES[iban]
        merchant = str(row.get("merchant", "")).strip().lower()
        for kw, cat in rules:
            if kw in merchant:
                return cat
        return "Zakelijke Kosten"

    df = df.copy()
    df["category"] = df.apply(_cat, axis=1)

    # Sign guard: income categories only valid for positive amounts
    for cat in INCOME_CATS:
        df.loc[(df["category"] == cat) & (df["amount"] < 0), "category"] = "Zakelijke Kosten"

    n_onbekend = (df["category"] == "Zakelijke Kosten").sum()
    print(f"Knab gecategoriseerd: {len(df)} transacties  "
          f"({n_onbekend} als Zakelijke Kosten)")
    return df


def _load_one(filepath):
    try:
        raw = pd.read_csv(
            filepath,
            sep=";",
            skiprows=1,   # skip "KNAB EXPORT" metadata row
            header=0,
            encoding="utf-8-sig",
            dtype=str,
        )
    except Exception as e:
        print(f"  Fout bij lezen {os.path.basename(filepath)}: {e}")
        return None

    for col in raw.columns:
        if raw[col].dtype == object:
            raw[col] = raw[col].str.strip('"').str.strip()

    raw = raw.dropna(subset=["Rekeningnummer"])
    raw = raw[raw["Rekeningnummer"].str.startswith("NL")]
    if raw.empty:
        return None

    date = pd.to_datetime(raw["Transactiedatum"], format="%d-%m-%Y", errors="coerce")
    bedrag = raw["Bedrag"].str.replace(",", ".").astype(float)
    credit_debet = raw["CreditDebet"].str.upper()
    amount = bedrag.where(credit_debet == "C", -bedrag)
    description = raw.apply(_build_description, axis=1)
    merchant = raw["Tegenrekeninghouder"].fillna("").str.strip()
    counterparty_iban = raw["Tegenrekeningnummer"].fillna("").str.strip()

    result = pd.DataFrame({
        "account":           raw["Rekeningnummer"],
        "currency":          raw.get("Valutacode", "EUR"),
        "date":              date,
        "balance_before":    None,
        "balance_after":     None,
        "value_date":        raw.get("Valutadatum", None),
        "amount":            amount,
        "description":       description,
        "merchant":          merchant,
        "counterparty_iban": counterparty_iban,
        "month":             date.dt.to_period("M").astype(str),
        "year":              date.dt.year.astype(str),
    })

    print(f"  {os.path.basename(filepath)}: {len(result)} transacties")
    return result


def _build_description(row):
    parts = []
    omschr = str(row.get("Omschrijving", "")).strip()
    naam   = str(row.get("Tegenrekeninghouder", "")).strip()
    iban   = str(row.get("Tegenrekeningnummer", "")).strip()
    wijze  = str(row.get("Betaalwijze", "")).strip()
    if omschr and omschr.lower() not in ("nan", ""):
        parts.append(omschr)
    if naam and naam.lower() not in ("nan", ""):
        parts.append(f"Naam: {naam}")
    if iban and iban.lower() not in ("nan", ""):
        parts.append(f"IBAN: {iban}")
    if wijze and wijze.lower() not in ("nan", ""):
        parts.append(wijze)
    return "  ".join(parts) if parts else ""


def _update_config_knab(account_numbers):
    cfg      = load_config()
    accounts = cfg.get("accounts", {})
    new_accs = [a for a in account_numbers if a not in accounts]
    if new_accs:
        for acct in new_accs:
            accounts[acct] = {"label": "Zakelijke Rekening (Knab)", "iban": acct}
        cfg["accounts"] = accounts
        save_config(cfg)
        for acct in new_accs:
            print(f"  Knab account opgeslagen: {acct}")
