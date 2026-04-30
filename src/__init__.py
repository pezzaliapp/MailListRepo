"""MailListRepo: analisi vendite/ordini e generazione mailing list."""

from pathlib import Path

BASE_DIR: Path = Path(__file__).resolve().parent.parent
DATA_DIR: Path = BASE_DIR / "data"
OUTPUT_DIR: Path = BASE_DIR / "output"

CLIENTS_FILE: Path = DATA_DIR / "elenco_clienti_al_30_aprile_2026.xlsx"
ORDERS_FILE: Path = DATA_DIR / "ordini_al_30_aprile_2026.xlsx"
SALES_FILE: Path = DATA_DIR / "vendite_al_30_aprile_2026.xlsx"

REFERENCE_DATE: str = "2026-04-30"

__all__ = [
    "BASE_DIR",
    "DATA_DIR",
    "OUTPUT_DIR",
    "CLIENTS_FILE",
    "ORDERS_FILE",
    "SALES_FILE",
    "REFERENCE_DATE",
]
