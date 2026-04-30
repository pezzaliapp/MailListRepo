"""Analisi delle vendite (storico spedizioni e resi).

Espone funzioni che ritornano DataFrame e una funzione `build_all` che
salva tutti i risultati in un unico file Excel multi-foglio.
"""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

from .. import OUTPUT_DIR, REFERENCE_DATE

logger = logging.getLogger(__name__)


def _ref_date(reference_date: str | None) -> pd.Timestamp:
    return pd.Timestamp(reference_date or REFERENCE_DATE)


def revenue_by_year(sales: pd.DataFrame) -> pd.DataFrame:
    """Fatturato netto (vendite - resi) per anno."""
    g = (
        sales.groupby("ANNO SPEDIZIONE", dropna=False)["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={"ANNO SPEDIZIONE": "ANNO", "IMPORTO_NETTO": "FATTURATO_NETTO"})
        .sort_values("ANNO")
    )
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    return g


def revenue_by_year_macro(sales: pd.DataFrame) -> pd.DataFrame:
    """Fatturato per anno × macro-famiglia (DESCRIZIONE ELEMENTO.4)."""
    macro_col = "DESCRIZIONE ELEMENTO.4"
    g = (
        sales.groupby(["ANNO SPEDIZIONE", macro_col], dropna=False)["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={
            "ANNO SPEDIZIONE": "ANNO",
            macro_col: "MACRO_FAMIGLIA",
            "IMPORTO_NETTO": "FATTURATO_NETTO",
        })
    )
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    return g.sort_values(["ANNO", "FATTURATO_NETTO"], ascending=[True, False])


def top_clients(
    sales: pd.DataFrame,
    top_n: int = 50,
    last_months: int | None = None,
    reference_date: str | None = None,
) -> pd.DataFrame:
    """Top N clienti per fatturato netto.

    Args:
        sales: dataset vendite.
        top_n: numero massimo di clienti.
        last_months: se valorizzato, filtra solo gli ultimi N mesi.
        reference_date: data di riferimento (default REFERENCE_DATE).
    """
    df = sales
    if last_months is not None:
        cutoff = _ref_date(reference_date) - relativedelta(months=last_months)
        df = df[df["DATA SPEDIZIONE"] >= cutoff]
    g = (
        df.groupby(["CLIENTE", "RAGIONE SOCIALE 1"], dropna=False)["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={"IMPORTO_NETTO": "FATTURATO_NETTO"})
        .sort_values("FATTURATO_NETTO", ascending=False)
        .head(top_n)
    )
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    return g.reset_index(drop=True)


def revenue_by_geography(sales: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Fatturato per nazione, regione (SOTTOAREA), provincia (ZONA)."""
    out: dict[str, pd.DataFrame] = {}
    for label, cols in {
        "nazione": ["NAZIONE"],
        "regione": ["SOTTOAREA", "DESCRIZIONE ELEMENTO.3"],
        "provincia": ["ZONA", "DESCRIZIONE ELEMENTO.1"],
    }.items():
        g = (
            sales.groupby(cols, dropna=False)["IMPORTO_NETTO"]
            .sum()
            .reset_index()
            .rename(columns={"IMPORTO_NETTO": "FATTURATO_NETTO"})
            .sort_values("FATTURATO_NETTO", ascending=False)
        )
        g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
        out[label] = g.reset_index(drop=True)
    return out


def revenue_by_agent(sales: pd.DataFrame) -> pd.DataFrame:
    """Fatturato per agente (DESCRIZIONE ELEMENTO.2)."""
    agent_col = "DESCRIZIONE ELEMENTO.2"
    g = (
        sales.groupby(["AGENTE", agent_col], dropna=False)["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={agent_col: "AGENTE_DESCR", "IMPORTO_NETTO": "FATTURATO_NETTO"})
        .sort_values("FATTURATO_NETTO", ascending=False)
    )
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    return g.reset_index(drop=True)


def top_articles(sales: pd.DataFrame, top_n: int = 30) -> dict[str, pd.DataFrame]:
    """Top articoli per fatturato e per quantità (escludendo i resi)."""
    only_sales = sales[~sales["IS_RESO"]]
    by_revenue = (
        only_sales.groupby(["ARTICOLO", "DESCRIZIONE"], dropna=False)["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={"IMPORTO_NETTO": "FATTURATO"})
        .sort_values("FATTURATO", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    by_revenue["FATTURATO"] = by_revenue["FATTURATO"].round(2)

    by_qty = (
        only_sales.groupby(["ARTICOLO", "DESCRIZIONE"], dropna=False)["QTA CONSEGNATA"]
        .sum()
        .reset_index()
        .rename(columns={"QTA CONSEGNATA": "QUANTITA"})
        .sort_values("QUANTITA", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    return {"by_revenue": by_revenue, "by_quantity": by_qty}


def monthly_trend(
    sales: pd.DataFrame,
    months: int = 24,
    reference_date: str | None = None,
) -> pd.DataFrame:
    """Trend mensile ultimi N mesi con confronto YoY."""
    ref = _ref_date(reference_date)
    cutoff = ref - relativedelta(months=months)
    df = sales[sales["DATA SPEDIZIONE"] >= cutoff].copy()
    df["MESE"] = df["DATA SPEDIZIONE"].dt.to_period("M").astype(str)
    g = (
        df.groupby("MESE")["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={"IMPORTO_NETTO": "FATTURATO_NETTO"})
        .sort_values("MESE")
    )

    # YoY: confronto con stesso mese anno precedente sull'intero storico.
    full = sales.copy()
    full["MESE"] = full["DATA SPEDIZIONE"].dt.to_period("M").astype(str)
    full_g = (
        full.groupby("MESE")["IMPORTO_NETTO"].sum().reset_index()
        .rename(columns={"IMPORTO_NETTO": "FATTURATO_NETTO"})
    )
    full_g["FATTURATO_NETTO_YOY"] = full_g["FATTURATO_NETTO"].shift(12)
    g = g.merge(full_g[["MESE", "FATTURATO_NETTO_YOY"]], on="MESE", how="left")
    g["VAR_YOY_PCT"] = (
        (g["FATTURATO_NETTO"] - g["FATTURATO_NETTO_YOY"]) / g["FATTURATO_NETTO_YOY"] * 100
    ).round(2)
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    g["FATTURATO_NETTO_YOY"] = g["FATTURATO_NETTO_YOY"].round(2)
    return g


def returns_rate(sales: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Tasso resi per anno e per macro-famiglia."""
    macro_col = "DESCRIZIONE ELEMENTO.4"

    def _rate(grouped: pd.DataFrame) -> pd.DataFrame:
        grouped = grouped.copy()
        grouped["TASSO_RESI_PCT"] = (
            grouped["RESI"].abs() / grouped["VENDITE"].abs() * 100
        ).round(2)
        grouped["VENDITE"] = grouped["VENDITE"].round(2)
        grouped["RESI"] = grouped["RESI"].round(2)
        return grouped

    by_year = (
        sales.groupby(["ANNO SPEDIZIONE", "IS_RESO"])["IMPORTO_NETTO"]
        .sum()
        .unstack(fill_value=0.0)
        .rename(columns={False: "VENDITE", True: "RESI"})
        .reset_index()
        .rename(columns={"ANNO SPEDIZIONE": "ANNO"})
    )
    by_year = _rate(by_year).sort_values("ANNO").reset_index(drop=True)

    by_macro = (
        sales.groupby([macro_col, "IS_RESO"])["IMPORTO_NETTO"]
        .sum()
        .unstack(fill_value=0.0)
        .rename(columns={False: "VENDITE", True: "RESI"})
        .reset_index()
        .rename(columns={macro_col: "MACRO_FAMIGLIA"})
    )
    by_macro = _rate(by_macro).sort_values("VENDITE", ascending=False).reset_index(drop=True)

    return {"by_year": by_year, "by_macro": by_macro}


def dormant_clients(
    sales: pd.DataFrame,
    months: int = 12,
    reference_date: str | None = None,
) -> pd.DataFrame:
    """Clienti che hanno acquistato in passato ma non negli ultimi N mesi."""
    ref = _ref_date(reference_date)
    cutoff = ref - relativedelta(months=months)
    last_purchase = (
        sales.groupby(["CLIENTE", "RAGIONE SOCIALE 1"])["DATA SPEDIZIONE"]
        .max()
        .reset_index()
        .rename(columns={"DATA SPEDIZIONE": "ULTIMO_ACQUISTO"})
    )
    revenue = (
        sales.groupby(["CLIENTE", "RAGIONE SOCIALE 1"])["IMPORTO_NETTO"]
        .sum()
        .reset_index()
        .rename(columns={"IMPORTO_NETTO": "FATTURATO_TOTALE"})
    )
    merged = last_purchase.merge(revenue, on=["CLIENTE", "RAGIONE SOCIALE 1"])
    dormant = merged[merged["ULTIMO_ACQUISTO"] < cutoff].copy()
    dormant["GIORNI_INATTIVITA"] = (ref - dormant["ULTIMO_ACQUISTO"]).dt.days
    dormant["FATTURATO_TOTALE"] = dormant["FATTURATO_TOTALE"].round(2)
    return dormant.sort_values("FATTURATO_TOTALE", ascending=False).reset_index(drop=True)


def purchase_frequency(sales: pd.DataFrame) -> pd.DataFrame:
    """Frequenza acquisti per cliente: numero transazioni distinte e ticket medio."""
    only = sales[~sales["IS_RESO"]]
    g = only.groupby(["CLIENTE", "RAGIONE SOCIALE 1"]).agg(
        NUMERO_TRANSAZIONI=("NUMERO SPEDIZIONE", "nunique"),
        FATTURATO_NETTO=("IMPORTO_NETTO", "sum"),
    ).reset_index()
    g["TICKET_MEDIO"] = (g["FATTURATO_NETTO"] / g["NUMERO_TRANSAZIONI"]).round(2)
    g["FATTURATO_NETTO"] = g["FATTURATO_NETTO"].round(2)
    return g.sort_values("FATTURATO_NETTO", ascending=False).reset_index(drop=True)


def build_all(
    sales: pd.DataFrame,
    output_path: Path | None = None,
    reference_date: str | None = None,
) -> dict[str, pd.DataFrame]:
    """Calcola tutte le analisi vendite e le scrive in un Excel multi-foglio.

    Args:
        sales: DataFrame vendite normalizzato dal loader.
        output_path: percorso Excel di output. Default `output/analisi_vendite.xlsx`.
        reference_date: override data di riferimento (default REFERENCE_DATE).

    Returns:
        Dict con tutti i DataFrame prodotti, comodo per la dashboard.
    """
    output_path = output_path or (OUTPUT_DIR / "analisi_vendite.xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    geo = revenue_by_geography(sales)
    arts = top_articles(sales)
    rates = returns_rate(sales)

    results: dict[str, pd.DataFrame] = {
        "fatturato_per_anno": revenue_by_year(sales),
        "fatturato_anno_macro": revenue_by_year_macro(sales),
        "top50_clienti_storico": top_clients(sales, 50),
        "top50_clienti_12m": top_clients(sales, 50, last_months=12, reference_date=reference_date),
        "geo_nazione": geo["nazione"],
        "geo_regione": geo["regione"],
        "geo_provincia": geo["provincia"],
        "fatturato_per_agente": revenue_by_agent(sales),
        "top30_articoli_fatturato": arts["by_revenue"],
        "top30_articoli_quantita": arts["by_quantity"],
        "trend_mensile_24m": monthly_trend(sales, 24, reference_date),
        "tasso_resi_per_anno": rates["by_year"],
        "tasso_resi_per_macro": rates["by_macro"],
        "clienti_dormienti_12m": dormant_clients(sales, 12, reference_date),
        "clienti_dormienti_24m": dormant_clients(sales, 24, reference_date),
        "frequenza_acquisti": purchase_frequency(sales),
    }

    logger.info("Scrittura Excel analisi vendite -> %s", output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            # Excel limita a 31 char i nomi foglio.
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    # CSV duplicato per ciascun risultato (utile per import esterni).
    csv_dir = output_path.parent / "vendite_csv"
    csv_dir.mkdir(parents=True, exist_ok=True)
    for name, df in results.items():
        df.to_csv(csv_dir / f"{name}.csv", index=False)

    return results
