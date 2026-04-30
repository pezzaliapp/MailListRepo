"""Pipeline CLI: rigenera tutti gli output (analisi vendite, ordini, mailing list)."""

from __future__ import annotations

import logging
import sys

from src import OUTPUT_DIR
from src.loaders import configure_logging, load_all
from src.analytics import orders as orders_mod
from src.analytics import sales as sales_mod
from src.mailing import builder as mailing_mod

logger = logging.getLogger("pipeline")


def main() -> int:
    """Esegue il pipeline end-to-end."""
    configure_logging()
    logger.info("Avvio pipeline. Output dir: %s", OUTPUT_DIR)

    data = load_all()
    sales_results = sales_mod.build_all(data["sales"])
    orders_results = orders_mod.build_all(data["orders"])
    mailing_results = mailing_mod.build_all(
        data["clients"], data["sales"], data["orders"]
    )

    mailing_df = mailing_results["mailing"]
    n_clients = data["clients"]["CONTATTO CLIENTI"].nunique()
    n_buyers = mailing_df.loc[mailing_df["ha_acquistato_storico"], "codice_cliente"].nunique()
    n_open = mailing_df.loc[mailing_df["ha_ordine_aperto"], "codice_cliente"].nunique()

    print()
    print("=" * 60)
    print("PIPELINE COMPLETATA")
    print("=" * 60)
    print(f"Clienti totali in anagrafica:                {n_clients}")
    print(f"Clienti con almeno un acquisto storico:      {n_buyers}")
    print(f"Clienti con ordini aperti:                   {n_open}")
    print()
    print(f"File generati in {OUTPUT_DIR}:")
    print(f"  - analisi_vendite.xlsx          ({len(sales_results)} fogli)")
    print(f"  - analisi_ordini.xlsx           ({len(orders_results)} fogli)")
    print(f"  - mailing/mailing_list_completa.xlsx")
    print(f"  - mailing/mailing_list_acquirenti.csv")
    print(f"  - mailing/mailing_list_ordini_aperti.csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
