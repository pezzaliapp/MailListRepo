"""Caricamento e pulizia dei dataset Excel.

Tutti i loader normalizzano `CLIENTE` / `CONTATTO CLIENTI` a `int` (chiave di
join condivisa fra anagrafica, vendite e ordini), bonificano gli spazi nelle
stringhe e rendono date e importi numerici.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Optional

import pandas as pd

from . import CLIENTS_FILE, ORDERS_FILE, SALES_FILE

logger = logging.getLogger(__name__)


def configure_logging(level: Optional[str] = None) -> None:
    """Configura il logging globale prendendo il livello da env LOG_LEVEL."""
    lvl = (level or os.getenv("LOG_LEVEL") or "INFO").upper()
    logging.basicConfig(
        level=lvl,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


def _strip_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rimuove spazi superflui nelle colonne di tipo object."""
    obj_cols = df.select_dtypes(include=["object"]).columns
    for c in obj_cols:
        df[c] = df[c].astype("string").str.strip()
    return df


def load_clients(path: Path = CLIENTS_FILE) -> pd.DataFrame:
    """Carica l'anagrafica clienti.

    Args:
        path: percorso del file xlsx.

    Returns:
        DataFrame con colonna `CONTATTO CLIENTI` come int e stringhe trimmate.
    """
    logger.info("Carico anagrafica clienti da %s", path)
    df = pd.read_excel(path, sheet_name="RUBRICA CONTATTI")
    df = _strip_object_columns(df)
    df["CONTATTO CLIENTI"] = pd.to_numeric(
        df["CONTATTO CLIENTI"], errors="coerce"
    ).astype("Int64")
    df = df.dropna(subset=["CONTATTO CLIENTI"]).copy()
    df["CONTATTO CLIENTI"] = df["CONTATTO CLIENTI"].astype("int64")
    if "PROGRESSIVO" in df.columns:
        df["PROGRESSIVO"] = pd.to_numeric(
            df["PROGRESSIVO"], errors="coerce"
        ).fillna(0).astype("int64")
    logger.info("Anagrafica caricata: %d righe, %d clienti distinti",
                len(df), df["CONTATTO CLIENTI"].nunique())
    return df


def load_orders(path: Path = ORDERS_FILE) -> pd.DataFrame:
    """Carica gli ordini clienti dettaglio.

    Args:
        path: percorso del file xlsx.

    Returns:
        DataFrame con date, importi e codici cliente normalizzati.
    """
    logger.info("Carico ordini da %s", path)
    df = pd.read_excel(path, sheet_name="ORDINI CLIENTI DETTAGLIO")
    df = _strip_object_columns(df)
    df["CLIENTE"] = pd.to_numeric(df["CLIENTE"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["CLIENTE"]).copy()
    df["CLIENTE"] = df["CLIENTE"].astype("int64")

    for col in ("DATA CREAZIONE", "DATA CONSEGNA"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    numeric_cols = [
        "QTA ORDINATA",
        "PZ NETTO VENDITA",
        "QTA INEVASA",
        "IMPORTO INEVASO",
        "IMPONIBILE ORD. VAL",
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    logger.info("Ordini caricati: %d righe, %d clienti, %d ordini distinti",
                len(df),
                df["CLIENTE"].nunique(),
                df[["ANNO", "TIPO", "NUM."]].drop_duplicates().shape[0])
    return df


def load_sales(path: Path = SALES_FILE) -> pd.DataFrame:
    """Carica lo storico vendite/spedizioni e resi.

    Args:
        path: percorso del file xlsx.

    Returns:
        DataFrame con date, importi normalizzati e flag `IS_RESO` per
        distinguere i resi dalle vendite.
    """
    logger.info("Carico vendite da %s", path)
    df = pd.read_excel(path, sheet_name="SPEDIZIONI E RESI CLIENTI DETTA")
    df = _strip_object_columns(df)
    df["CLIENTE"] = pd.to_numeric(df["CLIENTE"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["CLIENTE"]).copy()
    df["CLIENTE"] = df["CLIENTE"].astype("int64")

    if "DATA SPEDIZIONE" in df.columns:
        df["DATA SPEDIZIONE"] = pd.to_datetime(df["DATA SPEDIZIONE"], errors="coerce")

    numeric_cols = ["QTA CONSEGNATA", "PZ NETTO VENDITA", "IMPORTO CONSEGNATO"]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    df["IS_RESO"] = (
        df["TIPO SPEDIZIONE"].fillna("").str.startswith("R")
        | (df["IMPORTO CONSEGNATO"] < 0)
    )
    df["IMPORTO_NETTO"] = df.apply(
        lambda r: -abs(r["IMPORTO CONSEGNATO"]) if r["IS_RESO"] else r["IMPORTO CONSEGNATO"],
        axis=1,
    )

    logger.info("Vendite caricate: %d righe, %d clienti, anni %s-%s",
                len(df),
                df["CLIENTE"].nunique(),
                int(df["ANNO SPEDIZIONE"].min()) if not df.empty else "?",
                int(df["ANNO SPEDIZIONE"].max()) if not df.empty else "?")
    return df


def load_all(
    clients_path: Path = CLIENTS_FILE,
    orders_path: Path = ORDERS_FILE,
    sales_path: Path = SALES_FILE,
) -> dict[str, pd.DataFrame]:
    """Carica i tre dataset principali in un dict comodo per la dashboard."""
    return {
        "clients": load_clients(clients_path),
        "orders": load_orders(orders_path),
        "sales": load_sales(sales_path),
    }
