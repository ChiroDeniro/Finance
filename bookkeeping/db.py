"""
SQLite storage layer — bookkeeping/finance.db

Tabellen:
  transactions         — ABN, Knab, Revolut (gecategoriseerde kasboektransacties)
  degiro_transactions  — DeGiro beleggingstransacties
  bitvavo_transactions — Bitvavo crypto transacties

De DB is een reproduceerbaar cache van de bronbestanden + rules.xlsx.
Bij categoriewijzigingen: verwijder finance.db en herrun process.py.
"""

import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).parent / "finance.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS transactions (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    source            TEXT    NOT NULL,
    account           TEXT    NOT NULL,
    currency          TEXT,
    date              TEXT    NOT NULL,
    balance_before    REAL,
    balance_after     REAL,
    amount            REAL    NOT NULL,
    description       TEXT,
    merchant          TEXT,
    month             TEXT,
    year              TEXT,
    category          TEXT,
    counterparty_iban TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_tx
    ON transactions(source, account, date, amount, description);

CREATE TABLE IF NOT EXISTS degiro_transactions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    date        TEXT    NOT NULL,
    product     TEXT,
    isin        TEXT,
    aantal      REAL,
    koers_eur   REAL,
    waarde_eur  REAL,
    kosten_eur  REAL,
    totaal_eur  REAL    NOT NULL,
    month       TEXT,
    year        TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_degiro
    ON degiro_transactions(date, product, aantal, totaal_eur);

CREATE TABLE IF NOT EXISTS bitvavo_transactions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    date        TEXT    NOT NULL,
    type        TEXT,
    currency    TEXT,
    amount      REAL,
    eur_amount  REAL,
    fee_eur     REAL,
    month       TEXT,
    year        TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_bitvavo
    ON bitvavo_transactions(date, type, currency, amount, eur_amount);
"""


def _conn(path=DB_PATH):
    c = sqlite3.connect(path)
    c.row_factory = sqlite3.Row
    return c


def _safe(v):
    if v is None:
        return None
    try:
        f = float(str(v).replace(",", "."))
        return None if f != f else f  # NaN check
    except (TypeError, ValueError):
        return None


def init_db(path=DB_PATH):
    with _conn(path) as c:
        c.executescript(_SCHEMA)


# ── Upsert functions ───────────────────────────────────────────────────────────

def upsert_transactions(df, path=DB_PATH):
    """ABN / Knab / Revolut → transactions. Slaat duplicaten over."""
    cols = ["source", "account", "currency", "date", "balance_before", "balance_after",
            "amount", "description", "merchant", "month", "year", "category",
            "counterparty_iban"]

    records = [{
        "source":            str(r.get("source", "ABN")),
        "account":           str(r["account"]),
        "currency":          str(r.get("currency", "EUR")),
        "date":              r["date"].strftime("%Y-%m-%d"),
        "balance_before":    _safe(r.get("balance_before")),
        "balance_after":     _safe(r.get("balance_after")),
        "amount":            float(r["amount"]),
        "description":       str(r.get("description", "")),
        "merchant":          str(r.get("merchant", "")),
        "month":             str(r.get("month", "")),
        "year":              str(r.get("year", "")),
        "category":          str(r.get("category", "")),
        "counterparty_iban": r.get("counterparty_iban") or None,
    } for _, r in df.iterrows()]

    sql = (f"INSERT OR IGNORE INTO transactions ({', '.join(cols)}) "
           f"VALUES ({', '.join(':' + c for c in cols)})")
    with _conn(path) as c:
        c.executemany(sql, records)
    print(f"DB transactions: {len(records)} rijen aangeboden")


def upsert_degiro(df, path=DB_PATH):
    """DeGiro → degiro_transactions. Slaat duplicaten over."""
    cols = ["date", "product", "isin", "aantal", "koers_eur", "waarde_eur",
            "kosten_eur", "totaal_eur", "month", "year"]

    records = [{
        "date":       r["date"].strftime("%Y-%m-%d"),
        "product":    str(r.get("product", "")),
        "isin":       str(r.get("isin", "")),
        "aantal":     _safe(r.get("aantal")),
        "koers_eur":  _safe(r.get("koers_eur")),
        "waarde_eur": _safe(r.get("waarde_eur")),
        "kosten_eur": _safe(r.get("kosten_eur")),
        "totaal_eur": float(r["totaal_eur"]),
        "month":      str(r.get("month", "")),
        "year":       str(r.get("year", "")),
    } for _, r in df.iterrows()]

    sql = (f"INSERT OR IGNORE INTO degiro_transactions ({', '.join(cols)}) "
           f"VALUES ({', '.join(':' + c for c in cols)})")
    with _conn(path) as c:
        c.executemany(sql, records)
    print(f"DB degiro_transactions: {len(records)} rijen aangeboden")


def upsert_bitvavo(df, path=DB_PATH):
    """Bitvavo → bitvavo_transactions. Slaat duplicaten over."""
    cols = ["date", "type", "currency", "amount", "eur_amount", "fee_eur", "month", "year"]

    records = [{
        "date":       r["date"].strftime("%Y-%m-%d"),
        "type":       str(r.get("type", "")),
        "currency":   str(r.get("currency", "")),
        "amount":     _safe(r.get("amount")),
        "eur_amount": _safe(r.get("eur_amount")),
        "fee_eur":    _safe(r.get("fee_eur")),
        "month":      str(r.get("month", "")),
        "year":       str(r.get("year", "")),
    } for _, r in df.iterrows()]

    sql = (f"INSERT OR IGNORE INTO bitvavo_transactions ({', '.join(cols)}) "
           f"VALUES ({', '.join(':' + c for c in cols)})")
    with _conn(path) as c:
        c.executemany(sql, records)
    print(f"DB bitvavo_transactions: {len(records)} rijen aangeboden")


# ── Query functions ────────────────────────────────────────────────────────────

def query_year(year, path=DB_PATH) -> pd.DataFrame:
    with _conn(path) as c:
        return pd.read_sql_query(
            "SELECT * FROM transactions WHERE year = ?", c, params=(str(year),)
        )


def query_all(path=DB_PATH) -> pd.DataFrame:
    with _conn(path) as c:
        return pd.read_sql_query("SELECT * FROM transactions ORDER BY date", c)


def query_account_balances(path=DB_PATH) -> dict:
    """Laatste balance_after per account (ABN rekeningen met saldo-data)."""
    with _conn(path) as c:
        rows = c.execute("""
            SELECT account, balance_after, MAX(date) AS latest_date
            FROM transactions
            WHERE balance_after IS NOT NULL
            GROUP BY account
        """).fetchall()
    return {r["account"]: (r["balance_after"], r["latest_date"]) for r in rows}


def row_counts(path=DB_PATH) -> dict:
    """Aantal rijen per tabel — voor verificatie."""
    with _conn(path) as c:
        return {
            "transactions":         c.execute("SELECT COUNT(*) FROM transactions").fetchone()[0],
            "degiro_transactions":  c.execute("SELECT COUNT(*) FROM degiro_transactions").fetchone()[0],
            "bitvavo_transactions": c.execute("SELECT COUNT(*) FROM bitvavo_transactions").fetchone()[0],
        }
