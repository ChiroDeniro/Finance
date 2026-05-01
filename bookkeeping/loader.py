import os
import re
import sys
import glob
import pandas as pd

from config import INPUT_DIR, load_config, save_config


def find_input_files():
    patterns = ["*.TAB", "*.tab", "*.txt", "*.TXT"]
    seen, files = set(), []
    for p in patterns:
        for f in glob.glob(os.path.join(INPUT_DIR, "**", p), recursive=True):
            key = os.path.normcase(os.path.abspath(f))
            if key not in seen:
                seen.add(key)
                files.append(f)
    if not files:
        print(f"Geen invoerbestanden gevonden in {INPUT_DIR}/ (of submappen)")
        print("Download je ABN AMRO afschrift als 'Excel (TXT)' en zet het daar neer.")
        sys.exit(1)
    print(f"Gevonden: {len(files)} invoerbestand(en)")
    for f in files:
        print(f"  {os.path.relpath(f, INPUT_DIR)}")
    return files


def validate_tab_file(df, filepath):
    name   = os.path.basename(filepath)
    errors = []

    if len(df.columns) != 8:
        errors.append(f"verwacht 8 kolommen, gevonden {len(df.columns)}")

    if len(df) == 0:
        errors.append("bestand bevat geen transacties")
    else:
        for v in df["date"].dropna().head(3):
            if not re.match(r"^\d{8}$", str(v).strip()):
                errors.append(
                    f"datum-kolom niet in JJJJMMDD-formaat (voorbeeld: '{v}')"
                )
                break
        if not any("," in str(v) for v in df["amount"].dropna().head(5)):
            errors.append(
                "bedrag-kolom mist komma als decimaalteken "
                "(verwacht bijv. '-19,42')"
            )

    if errors:
        print(f"\nOngeldig bestandsformaat: {name}")
        for e in errors:
            print(f"  - {e}")
        print("Verwacht: ABN AMRO 'Excel (TXT)' export, 8 tab-kolommen:")
        print("  rekening | valuta | datum (JJJJMMDD) | saldo_voor | saldo_na"
              " | valutadatum | bedrag | omschrijving")
        sys.exit(1)


def load_transactions(files):
    dfs = []
    for f in files:
        raw = pd.read_csv(
            f, sep="\t", header=None, encoding="utf-8",
            names=["account", "currency", "date", "balance_before",
                   "balance_after", "value_date", "amount", "description"],
            dtype=str,
        )
        validate_tab_file(raw, f)
        dfs.append(raw)

    df = pd.concat(dfs, ignore_index=True)
    rijen_voor = len(df)
    df = df.drop_duplicates(
        subset=["account", "date", "amount", "description"],
        keep="first",
    )
    n_dubbel = rijen_voor - len(df)
    if n_dubbel > 0:
        print(f"  Deduplicatie: {n_dubbel} dubbele rijen verwijderd")
    df["date"]        = pd.to_datetime(df["date"], format="%Y%m%d")
    df["month"]       = df["date"].dt.to_period("M").astype(str)
    df["year"]        = df["date"].dt.year.astype(str)
    df["amount"]      = df["amount"].str.replace(",", ".").astype(float)
    df["description"] = df["description"].str.strip()
    df["merchant"]    = df["description"].apply(extract_merchant)
    df = df.sort_values("date").reset_index(drop=True)

    print(f"Geladen: {len(df)} transacties  "
          f"({df['date'].min().date()} - {df['date'].max().date()})")

    _update_config_accounts(df["account"].unique().tolist())
    return df


def _update_config_accounts(account_numbers):
    cfg      = load_config()
    accounts = cfg.get("accounts", {})
    new_accs = [a for a in account_numbers if a not in accounts]
    if new_accs:
        for acct in new_accs:
            accounts[acct] = {"label": "", "iban": ""}
        cfg["accounts"] = accounts
        save_config(cfg)
        for acct in new_accs:
            print(f"  Nieuw account opgeslagen: {acct} — vul label/IBAN in config.json in")


def extract_merchant(desc):
    desc = str(desc).strip()
    # BEA / Apple Pay pin transaction
    m = re.search(r"BEA,.*?\s{2,}(.+?),PAS", desc)
    if m:
        return m.group(1).strip()
    # Standard SEPA /NAME/ format
    m = re.search(r"/NAME/([^/]+)", desc)
    if m:
        return m.group(1).strip()
    # SEPA Incasso "Naam: ..." format (e.g. NS GROEP, gym subscriptions)
    m = re.search(r"Naam:\s*([^\t\n/,]+)", desc)
    if m:
        return m.group(1).strip()
    if "ABN AMRO" in desc.upper():
        return "ABN AMRO Bank"
    return desc[:40]
