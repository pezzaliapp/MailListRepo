"""Dashboard Streamlit per l'esplorazione di vendite, ordini e mailing list."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src import OUTPUT_DIR, REFERENCE_DATE
from src.analytics import orders as orders_mod
from src.analytics import sales as sales_mod
from src.loaders import configure_logging, load_all
from src.mailing import builder as mailing_mod

st.set_page_config(page_title="Analisi Vendite & Mailing List", layout="wide")
configure_logging("WARNING")


@st.cache_data(show_spinner="Carico dataset...")
def _load() -> dict[str, pd.DataFrame]:
    return load_all()


@st.cache_data(show_spinner="Calcolo analisi vendite...")
def _sales_results(_sales: pd.DataFrame) -> dict[str, pd.DataFrame]:
    return sales_mod.build_all(_sales, output_path=OUTPUT_DIR / "analisi_vendite.xlsx")


@st.cache_data(show_spinner="Calcolo analisi ordini...")
def _orders_results(_orders: pd.DataFrame) -> dict[str, pd.DataFrame]:
    return orders_mod.build_all(_orders, output_path=OUTPUT_DIR / "analisi_ordini.xlsx")


@st.cache_data(show_spinner="Costruisco mailing list...")
def _mailing(
    _clients: pd.DataFrame, _sales: pd.DataFrame, _orders: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    return mailing_mod.build_all(_clients, _sales, _orders)


def _filter_sales(sales: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtri vendite")
    years = sorted(sales["ANNO SPEDIZIONE"].dropna().unique().tolist())
    sel_years = st.sidebar.multiselect("Anno", years, default=years)
    agents = sorted(
        sales["DESCRIZIONE ELEMENTO.2"].dropna().astype(str).unique().tolist()
    )
    sel_agents = st.sidebar.multiselect("Agente", agents, default=agents)
    macros = sorted(
        sales["DESCRIZIONE ELEMENTO.4"].dropna().astype(str).unique().tolist()
    )
    sel_macros = st.sidebar.multiselect("Macro-famiglia", macros, default=macros)
    regions = sorted(
        sales["DESCRIZIONE ELEMENTO.3"].dropna().astype(str).unique().tolist()
    )
    sel_regions = st.sidebar.multiselect("Regione", regions, default=regions)

    df = sales[sales["ANNO SPEDIZIONE"].isin(sel_years)]
    if sel_agents:
        df = df[df["DESCRIZIONE ELEMENTO.2"].isin(sel_agents)]
    if sel_macros:
        df = df[df["DESCRIZIONE ELEMENTO.4"].isin(sel_macros)]
    if sel_regions:
        df = df[df["DESCRIZIONE ELEMENTO.3"].isin(sel_regions)]
    return df


def render_sales_tab(sales: pd.DataFrame) -> None:
    st.subheader("Analisi vendite")
    filt = _filter_sales(sales)
    st.caption(
        f"Righe filtrate: {len(filt):,} | "
        f"Fatturato netto filtrato: € {filt['IMPORTO_NETTO'].sum():,.2f}"
    )

    res = _sales_results(filt) if len(filt) != len(sales) else _sales_results(sales)

    col1, col2 = st.columns(2)
    with col1:
        df_year = res["fatturato_per_anno"]
        fig = px.bar(df_year, x="ANNO", y="FATTURATO_NETTO", title="Fatturato per anno")
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        df_trend = res["trend_mensile_24m"]
        fig = px.line(
            df_trend, x="MESE", y=["FATTURATO_NETTO", "FATTURATO_NETTO_YOY"],
            title="Trend mensile 24 mesi (vs YoY)",
        )
        st.plotly_chart(fig, use_container_width=True)

    df_macro = res["fatturato_anno_macro"]
    pivot = df_macro.pivot_table(
        index="MACRO_FAMIGLIA", columns="ANNO", values="FATTURATO_NETTO", fill_value=0
    )
    fig = px.imshow(
        pivot, aspect="auto", title="Heatmap fatturato per anno × macro-famiglia",
        color_continuous_scale="Blues",
    )
    st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**Top 50 clienti — storico**")
        st.dataframe(res["top50_clienti_storico"], use_container_width=True, height=350)
    with col4:
        st.markdown("**Top 50 clienti — ultimi 12 mesi**")
        st.dataframe(res["top50_clienti_12m"], use_container_width=True, height=350)

    col5, col6 = st.columns(2)
    with col5:
        st.markdown("**Top 30 articoli per fatturato**")
        st.dataframe(res["top30_articoli_fatturato"], use_container_width=True, height=300)
    with col6:
        st.markdown("**Top 30 articoli per quantità**")
        st.dataframe(res["top30_articoli_quantita"], use_container_width=True, height=300)

    col7, col8 = st.columns(2)
    with col7:
        st.markdown("**Fatturato per regione**")
        st.dataframe(res["geo_regione"], use_container_width=True, height=300)
    with col8:
        st.markdown("**Fatturato per agente**")
        st.dataframe(res["fatturato_per_agente"], use_container_width=True, height=300)

    col9, col10 = st.columns(2)
    with col9:
        st.markdown("**Tasso resi per anno**")
        st.dataframe(res["tasso_resi_per_anno"], use_container_width=True, height=300)
    with col10:
        st.markdown("**Clienti dormienti (>12 mesi)**")
        st.dataframe(res["clienti_dormienti_12m"], use_container_width=True, height=300)


def render_orders_tab(orders: pd.DataFrame) -> None:
    st.subheader("Ordini aperti")
    open_o = orders_mod.open_orders(orders)

    st.sidebar.header("Filtri ordini")
    macros = sorted(
        open_o["DESCRIZIONE ELEMENTO.5"].dropna().astype(str).unique().tolist()
    )
    sel_macros = st.sidebar.multiselect("Macro-famiglia (ordini)", macros, default=macros, key="om_macros")
    provs = sorted(open_o["PROVINCIA"].dropna().astype(str).unique().tolist())
    sel_provs = st.sidebar.multiselect("Provincia", provs, default=provs, key="om_prov")

    filt = open_o
    if sel_macros:
        filt = filt[filt["DESCRIZIONE ELEMENTO.5"].isin(sel_macros)]
    if sel_provs:
        filt = filt[filt["PROVINCIA"].isin(sel_provs)]

    res = orders_mod.build_all(filt) if len(filt) != len(open_o) else _orders_results(orders)

    sumr = res["riepilogo_backlog"].iloc[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Backlog totale (€)", f"{sumr['BACKLOG_TOTALE']:,.2f}")
    c2.metric("Qta inevasa", f"{int(sumr['QTA_INEVASA_TOTALE']):,}")
    c3.metric("Ordini aperti", f"{int(sumr['ORDINI_APERTI_DISTINTI']):,}")
    c4.metric("Righe aperte", f"{int(sumr['RIGHE_APERTE']):,}")

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(
            res["scadenze_consegna"], x="FASCIA_CONSEGNA", y="BACKLOG",
            title="Backlog per fascia di consegna",
        )
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        fig = px.bar(res["eta_ordini"], x="FASCIA_ETA", y="BACKLOG",
                     title="Backlog per età ordine")
        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**Backlog per cliente**")
        st.dataframe(res["backlog_per_cliente"], use_container_width=True, height=350)
    with col4:
        st.markdown("**Backlog per macro-famiglia**")
        st.dataframe(res["backlog_per_macro"], use_container_width=True, height=350)

    col5, col6 = st.columns(2)
    with col5:
        st.markdown("**Backlog per provincia**")
        st.dataframe(res["backlog_per_provincia"], use_container_width=True, height=300)
    with col6:
        st.markdown("**Top 20 articoli aperti**")
        st.dataframe(res["top20_articoli_aperti"], use_container_width=True, height=300)


def render_mailing_tab(mailing: pd.DataFrame, segments: dict[str, pd.DataFrame]) -> None:
    st.subheader("Mailing list")

    seg_choice = st.radio(
        "Segmento",
        options=["Tutti", "Solo acquirenti", "Solo ordini aperti", "Acquirenti E ordini", "Senza email"],
        horizontal=True,
    )
    mapping = {
        "Tutti": "tutti",
        "Solo acquirenti": "solo_acquirenti",
        "Solo ordini aperti": "solo_ordini_aperti",
        "Acquirenti E ordini": "acquirenti_e_ordini",
        "Senza email": "senza_email",
    }
    df = segments[mapping[seg_choice]].copy()

    q = st.text_input("Cerca (ragione sociale o email)").strip().lower()
    if q:
        df = df[
            df["ragione_sociale"].astype(str).str.lower().str.contains(q)
            | df["email"].astype(str).str.lower().str.contains(q)
        ]
    only_valid = st.checkbox("Solo email valide", value=False)
    if only_valid:
        df = df[df["email_valida"]]

    st.caption(f"{len(df):,} righe / {df['codice_cliente'].nunique():,} clienti distinti")
    st.dataframe(df, use_container_width=True, height=500)

    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("Scarica CSV", csv, file_name=f"mailing_{mapping[seg_choice]}.csv", mime="text/csv")

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=seg_choice[:31])
    st.download_button(
        "Scarica Excel", out.getvalue(),
        file_name=f"mailing_{mapping[seg_choice]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_anagrafica_tab(
    clients: pd.DataFrame,
    sales: pd.DataFrame,
    orders: pd.DataFrame,
) -> None:
    st.subheader("Ricerca cliente")
    q = st.text_input("Cerca per ragione sociale o codice cliente").strip()

    if not q:
        st.info("Inserisci un termine di ricerca per visualizzare il dettaglio cliente.")
        return

    if q.isdigit():
        sel = clients[clients["CONTATTO CLIENTI"] == int(q)]
    else:
        sel = clients[clients["RAGIONE SOCIALE 1"].astype(str).str.contains(q, case=False, na=False)]

    if sel.empty:
        st.warning("Nessun cliente trovato.")
        return

    matches = sel.drop_duplicates("CONTATTO CLIENTI")[["CONTATTO CLIENTI", "RAGIONE SOCIALE 1"]]
    st.dataframe(matches, use_container_width=True)

    if matches.shape[0] > 0:
        codice = st.selectbox(
            "Seleziona cliente",
            matches["CONTATTO CLIENTI"].tolist(),
            format_func=lambda c: f"{c} — {matches.loc[matches['CONTATTO CLIENTI']==c, 'RAGIONE SOCIALE 1'].iloc[0]}",
        )
        contacts = clients[clients["CONTATTO CLIENTI"] == codice]
        st.markdown("**Contatti**")
        st.dataframe(contacts, use_container_width=True)

        cl_sales = sales[sales["CLIENTE"] == codice].sort_values("DATA SPEDIZIONE", ascending=False)
        st.markdown(f"**Acquisti ({len(cl_sales):,} righe, fatturato netto € {cl_sales['IMPORTO_NETTO'].sum():,.2f})**")
        st.dataframe(
            cl_sales[
                ["DATA SPEDIZIONE", "NUMERO SPEDIZIONE", "TIPO SPEDIZIONE", "ARTICOLO",
                 "DESCRIZIONE", "QTA CONSEGNATA", "IMPORTO_NETTO"]
            ].head(500),
            use_container_width=True,
            height=300,
        )

        cl_orders = orders[orders["CLIENTE"] == codice]
        cl_orders_open = cl_orders[(cl_orders["QTA INEVASA"] > 0) | (cl_orders["IMPORTO INEVASO"] > 0)]
        st.markdown(f"**Ordini aperti ({len(cl_orders_open):,} righe, backlog € {cl_orders_open['IMPORTO INEVASO'].sum():,.2f})**")
        if cl_orders_open.empty:
            st.info("Nessun ordine aperto.")
        else:
            st.dataframe(
                cl_orders_open[
                    ["ANNO", "TIPO", "NUM.", "DATA CREAZIONE", "DATA CONSEGNA",
                     "ARTICOLO", "DESCRIZIONE", "QTA INEVASA", "IMPORTO INEVASO"]
                ],
                use_container_width=True,
                height=300,
            )


def main() -> None:
    st.title("📦 Analisi Vendite, Ordini e Mailing List")
    st.caption(f"Data di riferimento: **{REFERENCE_DATE}**")

    data = _load()
    mailing_res = _mailing(data["clients"], data["sales"], data["orders"])

    tab_v, tab_o, tab_m, tab_a = st.tabs(["Vendite", "Ordini aperti", "Mailing list", "Anagrafica"])
    with tab_v:
        render_sales_tab(data["sales"])
    with tab_o:
        render_orders_tab(data["orders"])
    with tab_m:
        render_mailing_tab(mailing_res["mailing"], mailing_res["segments"])
    with tab_a:
        render_anagrafica_tab(data["clients"], data["sales"], data["orders"])


if __name__ == "__main__":
    main()
