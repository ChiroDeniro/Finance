"""
Revolut CSV loader.

Revolut export format (Dutch, comma-separated, UTF-8):
  Type, Product, Startdatum, Datum voltooid, Beschrijving,
  Bedrag, Kosten, Valuta, Status, Saldo

Drop exports in: input/Revolut/
Only rows where Status == "VOLTOOID" are processed.
Bedrag is already signed (negative = expense).
Type == "Overschrijving" → flagged as internal transfer.
"""

import os
import glob
import pandas as pd

from config import INPUT_DIR, load_config

REVOLUT_DIR = os.path.join(INPUT_DIR, "Revolut")


def _own_ibans():
    cfg = load_config()
    return {v["iban"].upper() for v in cfg.get("accounts", {}).values() if v.get("iban")}


def find_revolut_files():
    if not os.path.isdir(REVOLUT_DIR):
        return []
    seen, files = set(), []
    for p in ["*.csv", "*.CSV"]:
        for f in glob.glob(os.path.join(REVOLUT_DIR, p)):
            key = os.path.normcase(os.path.abspath(f))
            if key not in seen:
                seen.add(key)
                files.append(f)
    return files


def load_revolut_transactions(files):
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
        subset=["date", "amount", "description"],
        keep="first",
    )
    n_dubbel = rijen_voor - len(df)
    if n_dubbel > 0:
        print(f"  Deduplicatie Revolut: {n_dubbel} dubbele rijen verwijderd")

    df = df.sort_values("date").reset_index(drop=True)
    print(f"Revolut geladen: {len(df)} transacties  "
          f"({df['date'].min().date()} — {df['date'].max().date()})")
    return df


def _load_one(filepath):
    try:
        raw = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str)
    except Exception as e:
        print(f"  Fout bij lezen {os.path.basename(filepath)}: {e}")
        return None

    raw.columns = raw.columns.str.strip()
    for col in raw.columns:
        if raw[col].dtype == object:
            raw[col] = raw[col].str.strip()

    required = {"Type", "Datum voltooid", "Beschrijving", "Bedrag", "Valuta", "Status"}
    missing = required - set(raw.columns)
    if missing:
        print(f"  {os.path.basename(filepath)}: ontbrekende kolommen {missing} — overgeslagen")
        return None

    raw = raw[raw["Status"].str.upper() == "VOLTOOID"].copy()
    if raw.empty:
        print(f"  {os.path.basename(filepath)}: geen VOLTOOID transacties")
        return None

    date = pd.to_datetime(raw["Datum voltooid"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
    amount = raw["Bedrag"].astype(float)
    description = raw["Beschrijving"].fillna("")

    own = _own_ibans()

    def _is_internal(row):
        typ  = row["Type"].strip().lower()
        desc = row["Beschrijving"].upper()
        if typ == "geld toevoegen":           # top-up from own bank to Revolut
            return True
        if typ == "overschrijving":           # only internal if destination is own IBAN
            return any(iban in desc for iban in own)
        return False

    is_transfer = raw.apply(_is_internal, axis=1)

    result = pd.DataFrame({
        "account":             "Revolut",
        "currency":            raw["Valuta"],
        "date":                date,
        "balance_before":      None,
        "balance_after":       None,
        "value_date":          None,
        "amount":              amount,
        "description":         description,
        "merchant":            description,
        "month":               date.dt.to_period("M").astype(str),
        "year":                date.dt.year.astype(str),
        "source":              "Revolut",
        "_revolut_transfer":   is_transfer,
    })

    print(f"  {os.path.basename(filepath)}: {len(result)} transacties")
    return result
