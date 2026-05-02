"""
DeGiro transacties CSV loader.

DeGiro export format (komma-gescheiden, UTF-8):
  Datum,Tijd,Product,ISIN,Beurs,Uitvoeringsplaats,Aantal,Koers,[valuta],
  Lokale waarde,[valuta],Waarde EUR,Wisselkoers,AutoFX Kosten,
  Transactiekosten en/of kosten van derden EUR,Totaal EUR,Order ID,[leeg]

Positief Totaal EUR = verkoop (ontvangen)
Negatief Totaal EUR = aankoop (betaald)

Drop exports in: input/DeGiro/
"""

import os
import glob
import pandas as pd
from config import INPUT_DIR

DEGIRO_DIR = os.path.join(INPUT_DIR, "DeGiro")


def find_degiro_files():
    if not os.path.isdir(DEGIRO_DIR):
        return []
    seen, files = set(), []
    for p in ["*.csv", "*.CSV"]:
        for f in glob.glob(os.path.join(DEGIRO_DIR, p)):
            key = os.path.normcase(os.path.abspath(f))
            if key not in seen:
                seen.add(key)
                files.append(f)
    return files


def load_degiro_transactions(files):
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
    df = df.drop_duplicates(subset=["date", "product", "aantal", "totaal_eur"], keep="first")
    if voor - len(df):
        print(f"  Deduplicatie DeGiro: {voor - len(df)} dubbele rijen verwijderd")
    df = df.sort_values("date").reset_index(drop=True)
    print(f"DeGiro geladen: {len(df)} transacties  "
          f"({df['date'].min().date()} — {df['date'].max().date()})")
    return df


def _num(series):
    return (series.astype(str).str.strip()
            .str.replace(",", ".", regex=False)
            .apply(lambda x: float(x) if x not in ("", "nan", "None", "-") else 0.0))


def _load_one(filepath):
    try:
        raw = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str, header=0)
    except Exception as e:
        print(f"  Fout bij lezen {os.path.basename(filepath)}: {e}")
        return None

    raw.columns = [str(c).strip() for c in raw.columns]
    if len(raw.columns) < 16:
        print(f"  {os.path.basename(filepath)}: onverwacht formaat — overgeslagen")
        return None

    # Datum (col 0) + Tijd (col 1)
    date = pd.to_datetime(
        raw.iloc[:, 0].str.strip() + " " + raw.iloc[:, 1].str.strip(),
        format="%d-%m-%Y %H:%M",
        errors="coerce",
    )

    # Named columns where available, positional fallback for unnamed ones
    def _col(name, idx):
        return raw[name] if name in raw.columns else raw.iloc[:, idx]

    result = pd.DataFrame({
        "date":       date,
        "product":    _col("Product", 2).fillna("").str.strip(),
        "isin":       _col("ISIN", 3).fillna("").str.strip(),
        "aantal":     _num(_col("Aantal", 6)),
        "koers_eur":  _num(_col("Koers", 7)),
        "waarde_eur": _num(_col("Waarde EUR", 11)),
        "kosten_eur": _num(_col("Transactiekosten en/of kosten van derden EUR", 14)),
        "totaal_eur": _num(_col("Totaal EUR", 15)),
        "month":      date.dt.to_period("M").astype(str),
        "year":       date.dt.year.astype(str),
    })

    result = result.dropna(subset=["date"])
    print(f"  {os.path.basename(filepath)}: {len(result)} transacties")
    return result
