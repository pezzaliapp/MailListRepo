"""Fixture sintetiche piccole per i test (no PII reali)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def synthetic_clients() -> pd.DataFrame:
    """Anagrafica con 7 clienti che coprono i casi limite di interesse."""
    rows = [
        # Cliente 1 — solo storico, una email valida
        {"CONTATTO CLIENTI": 1, "RAGIONE SOCIALE 1": "Alpha SRL", "NOME IN AGENDA": "Alpha",
         "EMAIL": "alpha@example.com", "NR.TELEFONICO": "0123", "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        # Cliente 2 — solo ordine aperto, due email (una multipla nella stessa cella)
        {"CONTATTO CLIENTI": 2, "RAGIONE SOCIALE 1": "Beta SRL", "NOME IN AGENDA": "Beta",
         "EMAIL": "beta1@example.com - beta2@example.com", "NR.TELEFONICO": "0456",
         "NR.CELLULARE": None, "NR.FAX": None, "PROGRESSIVO": 1},
        # Cliente 3 — entrambi (storico + ordine aperto), due righe rubrica con due email
        {"CONTATTO CLIENTI": 3, "RAGIONE SOCIALE 1": "Gamma SRL", "NOME IN AGENDA": "Gamma",
         "EMAIL": "gamma@example.com", "NR.TELEFONICO": "0789", "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        {"CONTATTO CLIENTI": 3, "RAGIONE SOCIALE 1": "Gamma SRL", "NOME IN AGENDA": "Gamma 2",
         "EMAIL": "gamma@example.com", "NR.TELEFONICO": None, "NR.CELLULARE": "333",
         "NR.FAX": None, "PROGRESSIVO": 2},
        # Cliente 4 — nessun match con vendite/ordini, una email
        {"CONTATTO CLIENTI": 4, "RAGIONE SOCIALE 1": "Delta SRL", "NOME IN AGENDA": "Delta",
         "EMAIL": "delta@example.com", "NR.TELEFONICO": None, "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        # Cliente 5 — storico ma email vuota
        {"CONTATTO CLIENTI": 5, "RAGIONE SOCIALE 1": "Epsilon SRL", "NOME IN AGENDA": "Eps",
         "EMAIL": None, "NR.TELEFONICO": "0999", "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        # Cliente 6 — storico, email NON valida (regex)
        {"CONTATTO CLIENTI": 6, "RAGIONE SOCIALE 1": "Zeta SRL", "NOME IN AGENDA": "Zeta",
         "EMAIL": "non-una-email", "NR.TELEFONICO": None, "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        # Cliente 7 — storico con email duplicata su due righe rubrica
        {"CONTATTO CLIENTI": 7, "RAGIONE SOCIALE 1": "Eta SRL", "NOME IN AGENDA": "Eta",
         "EMAIL": "eta@example.com", "NR.TELEFONICO": None, "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 1},
        {"CONTATTO CLIENTI": 7, "RAGIONE SOCIALE 1": "Eta SRL", "NOME IN AGENDA": "Eta 2",
         "EMAIL": "ETA@example.com", "NR.TELEFONICO": None, "NR.CELLULARE": None,
         "NR.FAX": None, "PROGRESSIVO": 2},
    ]
    return pd.DataFrame(rows)


@pytest.fixture
def synthetic_sales() -> pd.DataFrame:
    """Vendite sintetiche: 1, 3, 5, 6, 7 hanno acquisti; 1 ha anche un reso."""
    base = {
        "ANNO SPEDIZIONE": 2025, "TIPO SPEDIZIONE": "AV", "NUMERO SPEDIZIONE": 1,
        "SPESE DI TRASPORTO": 0.0, "PORTO": 0.0, "DESCRIZIONE ELEMENTO": "",
        "RIGA SPEDIZIONI": 1, "DATA SPEDIZIONE": pd.Timestamp("2025-09-01"),
        "ZONA": 1.0, "DESCRIZIONE ELEMENTO.1": "MILANO", "AGENTE": 10.0,
        "DESCRIZIONE ELEMENTO.2": "Agente A", "SOTTOAREA": 1.0,
        "DESCRIZIONE ELEMENTO.3": "Lombardia", "NAZIONE": "IT",
        "DESC. DESTINAZIONE 1": "", "DESCRIZIONE ELEMENTO.4": "SMONTAGOMME",
        "ARTICOLO": "ART01", "CLASSE 3 ARTICOLO": 1.0,
        "DESCRIZIONE ELEMENTO.5": "BASIC",
        "DESCRIZIONE": "Smontagomme basic", "QTA CONSEGNATA": 1.0,
        "PZ NETTO VENDITA": 1000.0, "IMPORTO CONSEGNATO": 1000.0,
        "CAUSALE MAGAZZINO": "V1", "AGENTE.1": 10.0,
        "DESCRIZIONE ELEMENTO.6": "", "NUMERO ORDINE": 100,
        "CLASSE 1 ARTICOLO": 1,
    }
    rows = []
    for code, dt, importo, tipo in [
        (1, "2025-09-01", 1000.0, "AV"),     # storico Alpha
        (1, "2025-10-01", -200.0, "RC"),     # reso Alpha
        (3, "2026-01-15", 500.0, "AV"),      # storico Gamma (recente)
        (5, "2020-03-20", 700.0, "AV"),      # storico Epsilon (vecchio)
        (6, "2025-12-01", 300.0, "AV"),      # storico Zeta (recente)
        (7, "2024-06-10", 900.0, "AV"),      # storico Eta (oltre 12 mesi)
    ]:
        r = dict(base)
        r["CLIENTE"] = code
        r["RAGIONE SOCIALE 1"] = f"Cliente{code}"
        r["DATA SPEDIZIONE"] = pd.Timestamp(dt)
        r["IMPORTO CONSEGNATO"] = importo
        r["TIPO SPEDIZIONE"] = tipo
        r["ANNO SPEDIZIONE"] = pd.Timestamp(dt).year
        r["NUMERO SPEDIZIONE"] = len(rows) + 1
        rows.append(r)
    df = pd.DataFrame(rows)
    df["IS_RESO"] = df["TIPO SPEDIZIONE"].str.startswith("R") | (df["IMPORTO CONSEGNATO"] < 0)
    df["IMPORTO_NETTO"] = df.apply(
        lambda r: -abs(r["IMPORTO CONSEGNATO"]) if r["IS_RESO"] else r["IMPORTO CONSEGNATO"],
        axis=1,
    )
    return df


@pytest.fixture
def synthetic_orders() -> pd.DataFrame:
    """Ordini sintetici: 2 e 3 hanno ordini aperti."""
    base = {
        "ANNO": 2026, "TIPO": "OC", "PORTO": 0, "DESCRIZIONE ELEMENTO": "",
        "SPESE DI TRASPORTO": 0, "SOTTOAREA": 1, "DESCRIZIONE ELEMENTO.1": "Lombardia",
        "PROVINCIA": "MI", "DESCRIZIONE ELEMENTO.2": "Agente A",
        "ARTICOLO": "ART01", "CLASSE 3 ARTICOLO": 1.0,
        "DESCRIZIONE ELEMENTO.3": "BASIC", "CLASSE 2 ARTICOLO": 1,
        "DESCRIZIONE ELEMENTO.4": "Sottocategoria", "CLASSE 1 ARTICOLO": 1,
        "DESCRIZIONE ELEMENTO.5": "SMONTAGOMME", "DESCRIZIONE": "Smontagomme basic",
        "UM VENDITA": "PZ", "PZ NETTO VENDITA": 1000.0,
        "IMPONIBILE ORD. VAL": 1000.0,
    }
    rows = []
    for code, num, qta, importo, dc in [
        (2, 100, 2, 2000.0, "2026-05-15"),
        (3, 101, 1, 1500.0, "2026-04-20"),  # in ritardo
        (3, 102, 1, 800.0, "2026-06-01"),
    ]:
        r = dict(base)
        r["CLIENTE"] = code
        r["NUM."] = num
        r["CLIENTE.1"] = f"Cliente{code}"
        r["DATA CREAZIONE"] = pd.Timestamp("2026-03-01")
        r["DATA CONSEGNA"] = pd.Timestamp(dc)
        r["QTA ORDINATA"] = qta
        r["QTA INEVASA"] = qta
        r["IMPORTO INEVASO"] = importo
        rows.append(r)
    return pd.DataFrame(rows)


@pytest.fixture
def tmp_output_dir(tmp_path: Path) -> Path:
    out = tmp_path / "output"
    out.mkdir()
    return out
