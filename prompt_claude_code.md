# Prompt per Claude Code — App Analisi Vendite/Ordini + Mail List

## Contesto

Stai lavorando dentro `~/Desktop/MailListRepo/`. La cartella contiene già:
- `data/elenco_clienti_al_30_aprile_2026.xlsx`
- `data/ordini_al_30_aprile_2026.xlsx`
- `data/vendite_al_30_aprile_2026.xlsx`
- questo prompt (`prompt_claude_code.md`)

Il repository GitHub remoto esiste già ed è vuoto: **`https://github.com/pezzaliapp/MailListRepo.git`**. Devi collegarti a quello, non crearne uno nuovo.

## Obiettivo

Costruisci in completa autonomia un'applicazione Python che:

1. **Analizza in dettaglio le vendite** dello storico aziendale.
2. **Analizza in dettaglio gli ordini aperti** (ordini con merce ancora da evadere).
3. **Genera una mailing list** dei clienti incrociando l'anagrafica con vendite e ordini, identificando chi ha acquistato e/o ha merce in ordine.
4. **Pubblica tutto sul repo GitHub esistente** `pezzaliapp/MailListRepo`.

Lavora end-to-end senza chiedermi conferme intermedie. Quando hai finito, mostrami:
- Il commit hash e l'URL del push.
- Un riepilogo di cosa hai prodotto (tabelle, grafici, file CSV/Excel di output).
- Le istruzioni per lanciare la dashboard in locale.

---

## Input — file disponibili in `./data/`

I file sono in italiano e usano la tipica struttura di un gestionale (probabilmente AS/400 o simile). **Non rinominare le colonne**.

### 1. `data/elenco_clienti_al_30_aprile_2026.xlsx`
- Foglio: `RUBRICA CONTATTI`
- Righe: ~2.048 (un cliente può avere più righe = più contatti, distinti dal `PROGRESSIVO`)
- Colonne: `CONTATTO CLIENTI` (codice numerico cliente — **chiave primaria**), `RAGIONE SOCIALE 1`, `NOME IN AGENDA`, `EMAIL`, `NR.TELEFONICO`, `NR.CELLULARE`, `NR.FAX`, `PROGRESSIVO`
- Note: la colonna `EMAIL` può essere vuota (`NaN`) o contenere indirizzi normali e PEC. Un cliente può avere più email (una per riga di contatto).

### 2. `data/ordini_al_30_aprile_2026.xlsx`
- Foglio: `ORDINI CLIENTI DETTAGLIO`
- Righe: ~186 (una riga per articolo per ordine)
- Colonne chiave: `ANNO`, `TIPO`, `NUM.` (numero ordine), `DATA CREAZIONE`, `CLIENTE` (codice numerico — **stesso dominio di `CONTATTO CLIENTI`**), `CLIENTE.1` (ragione sociale), `PROVINCIA`, `ARTICOLO`, `DESCRIZIONE`, `QTA ORDINATA`, `PZ NETTO VENDITA` (prezzo unitario), `QTA INEVASA`, `IMPORTO INEVASO`, `IMPONIBILE ORD. VAL`, `DATA CONSEGNA`
- Tassonomia articolo (3 livelli): `CLASSE 1 ARTICOLO` + `DESCRIZIONE ELEMENTO.5` (macro-famiglia, es. SMONTAGOMME), `CLASSE 2 ARTICOLO` + `DESCRIZIONE ELEMENTO.4` (sottocategoria), `CLASSE 3 ARTICOLO` + `DESCRIZIONE ELEMENTO.3` (modello)
- Geografia: `SOTTOAREA` + `DESCRIZIONE ELEMENTO.1` (zona/regione), `PROVINCIA`, `DESCRIZIONE ELEMENTO.2` (agente/zona)
- **Definizione "ordine aperto"**: righe con `QTA INEVASA > 0` **e/o** `IMPORTO INEVASO > 0`.

### 3. `data/vendite_al_30_aprile_2026.xlsx`
- Foglio: `SPEDIZIONI E RESI CLIENTI DETTA`
- Righe: ~15.862 (storico 2019 → 2026)
- Colonne chiave: `ANNO SPEDIZIONE`, `TIPO SPEDIZIONE` (`AV` = vendita/spedizione, altri valori = resi), `NUMERO SPEDIZIONE`, `DATA SPEDIZIONE`, `CLIENTE` (codice numerico), `RAGIONE SOCIALE 1`, `AGENTE` + `DESCRIZIONE ELEMENTO.2` (nome agente), `NAZIONE`, `SOTTOAREA` + `DESCRIZIONE ELEMENTO.3` (regione), `ZONA` + `DESCRIZIONE ELEMENTO.1` (provincia/città), `ARTICOLO`, `DESCRIZIONE`, `CLASSE 1 ARTICOLO` + `DESCRIZIONE ELEMENTO.4` (macro-famiglia), `CLASSE 3 ARTICOLO` + `DESCRIZIONE ELEMENTO.5` (modello), `QTA CONSEGNATA`, `PZ NETTO VENDITA`, `IMPORTO CONSEGNATO`, `CAUSALE MAGAZZINO`, `NUMERO ORDINE`
- Importante: distingui chiaramente tra **spedizioni** e **resi** usando `TIPO SPEDIZIONE` e/o segno di `IMPORTO CONSEGNATO`. Per il fatturato netto, somma vendite e sottrai resi.

---

## Requisiti funzionali

### A) Modulo `src/analytics/sales.py` — Analisi vendite

Calcola e salva (sia come CSV sia come fogli di un unico Excel `output/analisi_vendite.xlsx`):

1. **Fatturato per anno** (somma `IMPORTO CONSEGNATO` netto resi), con grafico a barre.
2. **Fatturato per anno × macro-famiglia articolo** (`DESCRIZIONE ELEMENTO.4`), heatmap.
3. **Top 50 clienti per fatturato totale** (storico completo) e **Top 50 ultimi 12 mesi**.
4. **Fatturato per area geografica**: per `NAZIONE`, per `SOTTOAREA` (regione), per `ZONA` (provincia/città).
5. **Fatturato per agente** (`DESCRIZIONE ELEMENTO.2`).
6. **Top 30 articoli per fatturato** e **Top 30 articoli per quantità venduta**.
7. **Trend mensile** ultimi 24 mesi (linea), con confronto YoY.
8. **Tasso resi**: % di resi su vendite per anno e per macro-famiglia.
9. **Clienti dormienti**: clienti che hanno comprato in passato ma non negli ultimi 12 mesi (e non negli ultimi 24).
10. **Frequenza acquisti per cliente**: numero di transazioni distinte (per `NUMERO SPEDIZIONE`) e ticket medio.

### B) Modulo `src/analytics/orders.py` — Analisi ordini aperti

Filtra solo righe con `QTA INEVASA > 0` o `IMPORTO INEVASO > 0` e produci `output/analisi_ordini.xlsx`:

1. **Backlog totale** (somma `IMPORTO INEVASO`) e numero ordini aperti distinti (`ANNO` + `TIPO` + `NUM.`).
2. **Backlog per cliente** ordinato decrescente.
3. **Backlog per macro-famiglia articolo** (`DESCRIZIONE ELEMENTO.5`).
4. **Backlog per provincia/regione**.
5. **Scadenze**: distribuzione per `DATA CONSEGNA` (in ritardo / questa settimana / prossimi 30 gg / oltre).
6. **Età ordini**: giorni trascorsi da `DATA CREAZIONE` a oggi (30/04/2026), con classi `<30`, `30-60`, `60-90`, `>90`.
7. **Top 20 articoli più ordinati** ancora da evadere.

### C) Modulo `src/mailing/builder.py` — Generazione mailing list

Logica:

1. Carica anagrafica clienti e raggruppa contatti per `CONTATTO CLIENTI`, aggregando tutte le email valide (ignora vuote, deduplica, trim).
2. Per ogni cliente, calcola:
   - `ha_acquistato_storico` (bool) — esiste almeno una vendita.
   - `ha_acquistato_12m` (bool) — vendita negli ultimi 365 giorni.
   - `ha_ordine_aperto` (bool) — almeno una riga ordine con inevaso > 0.
   - `fatturato_totale`, `fatturato_12m`, `data_ultimo_acquisto`.
   - `backlog_attuale`, `n_ordini_aperti`.
3. Produci tre file in `output/mailing/`:
   - `mailing_list_acquirenti.csv` — clienti con `ha_acquistato_storico=True`, una riga per email.
   - `mailing_list_ordini_aperti.csv` — clienti con `ha_ordine_aperto=True`, una riga per email.
   - `mailing_list_completa.xlsx` — un foglio "Tutti", un foglio "Solo acquirenti", un foglio "Solo ordini aperti", un foglio "Acquirenti E ordini aperti", un foglio "Senza email" (clienti che matchano ma non hanno email — vanno segnalati).
4. Ogni riga deve contenere almeno: `codice_cliente`, `ragione_sociale`, `email`, `provincia` (se ricavabile), `fatturato_totale`, `fatturato_12m`, `data_ultimo_acquisto`, `backlog_attuale`, `n_ordini_aperti`, `ha_acquistato_12m`, `ha_ordine_aperto`.
5. **Validazione email**: applica una regex semplice e marca le email non valide con flag `email_valida=False` (non scartarle, segnalale).

### D) Dashboard `app.py` — Streamlit

Una single-page app Streamlit con tab:
- **Vendite**: tutti i grafici del modulo A, con filtri per anno, agente, regione, macro-famiglia.
- **Ordini aperti**: dashboard del modulo B con filtri analoghi.
- **Mailing list**: tabella interattiva filtrabile (acquirenti / ordini aperti / entrambi), con bottone "Esporta CSV" che scarica i file generati.
- **Anagrafica**: ricerca cliente per ragione sociale o codice, mostra tutti i suoi acquisti, ordini aperti, contatti.

Usa `pandas`, `streamlit`, `plotly` (grafici interattivi), `openpyxl`. Niente database: tutto in-memory caricato all'avvio con `@st.cache_data`.

---

## Struttura del progetto

```
MailListRepo/
├── data/                              # input (gitignored, contiene PII)
│   ├── elenco_clienti_al_30_aprile_2026.xlsx
│   ├── ordini_al_30_aprile_2026.xlsx
│   └── vendite_al_30_aprile_2026.xlsx
├── output/                            # output generati (gitignored)
│   ├── analisi_vendite.xlsx
│   ├── analisi_ordini.xlsx
│   └── mailing/
│       ├── mailing_list_acquirenti.csv
│       ├── mailing_list_ordini_aperti.csv
│       └── mailing_list_completa.xlsx
├── src/
│   ├── __init__.py
│   ├── loaders.py                     # caricamento e pulizia dei 3 file
│   ├── analytics/
│   │   ├── __init__.py
│   │   ├── sales.py
│   │   └── orders.py
│   └── mailing/
│       ├── __init__.py
│       └── builder.py
├── tests/
│   ├── test_loaders.py
│   ├── test_sales.py
│   ├── test_orders.py
│   └── test_mailing.py
├── app.py                             # dashboard Streamlit
├── run_pipeline.py                    # script CLI: rigenera tutti gli output
├── requirements.txt
├── .gitignore                         # esclude data/, output/, .venv/, __pycache__/, .pytest_cache/, .DS_Store
├── README.md
└── LICENSE                            # MIT
```

---

## Requisiti tecnici

- **Python 3.11+**.
- Crea un virtualenv `.venv`, installa le dipendenze, verifica che tutto giri.
- `requirements.txt` con versioni pinned: `pandas`, `openpyxl`, `streamlit`, `plotly`, `python-dateutil`, `pytest`.
- **Type hints** ovunque, docstring stile Google, formattazione `black`-compatible.
- **Tutti i path** vanno gestiti via `pathlib.Path` partendo da una `BASE_DIR` definita in un solo posto.
- **Logging** via modulo `logging` (non `print`), livello configurabile da env var `LOG_LEVEL`.
- **Robustezza dati**: gestisci `NaN` su email, importi, date; sanitizza spazi; coerci `CLIENTE` a `int` su tutti i dataset prima del join.
- **Test pytest**: almeno uno per ciascun modulo. Per il mailing builder, crea fixture sintetiche piccole (5-10 clienti) che coprano i casi: cliente solo storico, solo ordine aperto, entrambi, nessun match, email vuota, email duplicata.
- **README.md** con: descrizione, prerequisiti, setup (`pip install -r requirements.txt`), come lanciare `python run_pipeline.py`, come lanciare `streamlit run app.py`, struttura output.

---

## Pubblicazione su GitHub (repo già esistente)

Il repository remoto esiste già: `https://github.com/pezzaliapp/MailListRepo.git`. **NON usare `gh repo create`**.

Esegui esattamente questa sequenza:

1. Verifica che `gh` sia autenticato: `gh auth status`. Se non lo è, fermati e chiedimi di autenticarmi.
2. **Verifica `.gitignore` PRIMA di qualsiasi `git add`**. Deve contenere almeno:
   ```
   data/
   output/
   .venv/
   __pycache__/
   *.pyc
   .pytest_cache/
   .DS_Store
   .env
   ```
3. Inizializza git locale:
   ```bash
   git init -b main
   git remote add origin https://github.com/pezzaliapp/MailListRepo.git
   ```
4. Prima di committare, esegui `git status` e verifica che NON compaiano file dentro `data/` o `output/`. Se compaiono, sistema il `.gitignore` e ripeti.
5. Esegui `git add -A` poi un secondo controllo con `git ls-files | grep -E "(data/|output/)"` — deve restituire stringa vuota. Se restituisce qualcosa, FERMATI e sistema.
6. Commit: `git commit -m "feat: initial implementation of sales/orders analytics and mailing list builder"`
7. Verifica se il remoto ha già contenuto: `git ls-remote origin`. 
   - Se è vuoto: `git push -u origin main`.
   - Se ha già un branch (es. `main` con un README iniziale): fai `git pull --rebase origin main` e poi `git push -u origin main`. In caso di conflitti banali (solo README), risolvi tenendo il tuo README.
8. Stampa l'URL finale del repo e l'hash del commit.

**Importante**: questo repo è pubblico (è di un account personale `pezzaliapp`). Anche se i dati sono gitignored, fai un controllo finale che nessun file in commit contenga email/codici fiscali/dati sensibili (ad esempio nei test fixture devi usare email finte tipo `cliente1@example.com`, non email reali estratte dai file).

---

## Criteri di accettazione

Considera il lavoro finito quando:
- [ ] `python run_pipeline.py` gira senza errori e produce tutti i file in `output/`.
- [ ] `streamlit run app.py` apre la dashboard e tutte le tab caricano dati reali.
- [ ] `pytest` passa al 100%.
- [ ] Il push su `pezzaliapp/MailListRepo` è andato a buon fine, senza file dentro `data/` o `output/` nel repo remoto.
- [ ] Il README spiega in italiano come usare il progetto.
- [ ] Mi mostri 3 numeri sintetici alla fine: numero clienti totali in anagrafica, numero clienti con almeno un acquisto storico, numero clienti con ordini aperti.

Procedi.
