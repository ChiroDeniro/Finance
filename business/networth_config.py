from datetime import date

# Datum van deze snapshot — update bij elke run
PEILDATUM = "2026-05-01"

# Labels voor ABN AMRO rekeningen (accountnummer → weergavenaam)
ABN_LABELS = {
    "536542171": "ABN AMRO Betaalrekening",
    "844835730": "ABN AMRO Spaarrekening",
}

# Knab zakelijke rekening — handmatig invullen (staat niet in CSV export)
ODS_SALDO = 0.00

# Beleggingen — handmatig bijhouden
BELEGGINGEN = {
    "DeGiro portfolio":          0.00,
    "Lijfrente (Brand New Day)": 0.00,
}

# Activa — geschatte marktwaarde (niet aanschafprijs)
ACTIVA = {
    "Laptop":       800.00,
    "Monitoren (2x)": 600.00,
    "Fiets":          0.00,
}

# Schulden — handmatig bijhouden
SCHULDEN = {
    "DUO studieschuld": 0.00,
}

# Zakelijke kosten door Chris privé voorgeschoten — ODS is dit verschuldigd
# status: "Open" of "Terugbetaald"
VORDERINGEN_OP_ODS = [
    {
        "datum":        "2025-11-01",
        "omschrijving": "Claude Pro abonnement",
        "bedrag":       21.78,
        "status":       "Open",
    },
    # {
    #     "datum":        "YYYY-MM-DD",
    #     "omschrijving": "Monitoren (2x)",
    #     "bedrag":       0.00,   # vul in zodra bedrag bekend is
    #     "status":       "Open",
    # },
]
