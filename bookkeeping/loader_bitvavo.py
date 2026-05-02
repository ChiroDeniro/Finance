"""
Bitvavo crypto CSV loader.

Bitvavo export format (komma-gescheiden, UTF-8):
  Timezone,Date,Time,Type,Currency,Amount,Quote Currency,Quote Price,
  Received / Paid Currency,Received / Paid Amount,Fee currency,Fee amount,
  Status,Transaction ID,Address

Types: buy, sell, deposit, withdrawal, ...
Alleen Status == Completed rijen worden verwerkt.

EUR flow: Received/Paid Amount (negatief = betaald EUR, positief = ontvangen EUR)
EUR stortingen: Amount (positief, Received/Paid Amount leeg)
De fee zit al inbegrepen in Received/Paid Amount voor buy/sell.

Drop exports in: input/Bitvavo - crypto/
"""

import os
import glob
import pandas as pd
from config import INPUT_DIR

BITVAVO_DIR = os.path.join(INPUT_DIR, "Bitvavo - crypto")


def find_bitvavo_files():
    if not os.path.isdir(BITVAVO_DIR):
        return []
    seen, files = set(), []
    for p in ["*.csv", "*.CSV"]:
        for f in glob.glob(os.path.join(BITVAVO_DIR, p)):
            key = os.path.normcase(os.path.abspath(f))
            if key not in seen:
                seen.add(key)
                files.append(f)
    return files


def load_bitvavo_transactions(files):
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
    voor = len(df)
    df = df.drop_duplicates(
        subset=["date", "type", "currency", "amount", "eur_amount"], keep="first"
    )
    if voor - len(df):
        print(f"  Deduplicatie Bitvavo: {voor - len(df)} dubbele rijen verwijderd")
    df = df.sort_values("date").reset_index(drop=True)
    print(f"Bitvavo geladen: {len(df)} transacties  "
          f"({df['date'].min().date()} — {df['date'].max().date()})")
    return df


def _num(series):
    return (series.fillna("0").astype(str).str.strip()
            .str.replace(",", ".", regex=False)
            .apply(lambda x: float(x) if x not in ("", "nan", "None") else 0.0))


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

    required = {"Date", "Time", "Type", "Currency", "Amount", "Status"}
    if missing := required - set(raw.columns):
        print(f"  {os.path.basename(filepath)}: ontbrekende kolommen {missing} — overgeslagen")
        return None

    raw = raw[raw["Status"].str.lower() == "completed"].copy()
    if raw.empty:
        print(f"  {os.path.basename(filepath)}: geen Completed transacties")
        return None

    date = pd.to_datetime(
        raw["Date"] + " " + raw["Time"], format="mixed", errors="coerce"
    )

    amount     = _num(raw["Amount"])
    eur_amount = _num(raw.get("Received / Paid Amount", pd.Series("0", index=raw.index)))
    fee_eur    = _num(raw.get("Fee amount", pd.Series("0", index=raw.index)))
    typ        = raw["Type"].str.lower()
    currency   = raw["Currency"].fillna("")

    # EUR stortingen: Received/Paid Amount is leeg → gebruik Amount
    is_eur_deposit = (typ == "deposit") & (currency == "EUR") & (eur_amount == 0)
    eur_amount = eur_amount.copy()
    eur_amount[is_eur_deposit] = amount[is_eur_deposit]

    result = pd.DataFrame({
        "date":       date,
        "type":       typ,
        "currency":   currency,
        "amount":     amount,
        "eur_amount": eur_amount,
        "fee_eur":    fee_eur,
        "month":      date.dt.to_period("M").astype(str),
        "year":       date.dt.year.astype(str),
    })
    result = result.dropna(subset=["date"])
    print(f"  {os.path.basename(filepath)}: {len(result)} transacties")
    return result
