// MailListRepo — elaborazione client-side dei 3 file Excel.
// Tutto in browser, nessun upload remoto.

const SHEETS = {
  anagrafica: "RUBRICA CONTATTI",
  ordini: "ORDINI CLIENTI DETTAGLIO",
  vendite: "SPEDIZIONI E RESI CLIENTI DETTA",
};

// Email da escludere sempre dalla mail list (case-insensitive).
const EMAIL_ESCLUSE = new Set([
  "a.pezzali@cormachsrl.com",
]);

// Anagrafica: lettura per nome (le colonne che ci servono non sono duplicate).
const COLUMNS = {
  anagrafica: ["CLIENTE", "RAGIONE SOCIALE 1", "EMAIL"],
};

// Ordini e vendite: lettura per indice di colonna. Le intestazioni sono
// duplicate (CLIENTE due volte negli ordini, DESCRIZIONE ELEMENTO più volte
// nelle vendite) quindi sheet_to_json default sovrascrive i duplicati.
// Indici verificati sui file di riferimento del 30/04/2026.
const VENDITE_IDX = {
  anno: 0,        // ANNO SPEDIZIONE
  tipo: 1,        // TIPO SPEDIZIONE   ("R…" = reso)
  data: 7,        // DATA SPEDIZIONE
  cliente: 10,    // CLIENTE (codice numerico)
  ragione: 16,    // RAGIONE SOCIALE 1
  articolo: 19,   // ARTICOLO
  descrizione: 22, // DESCRIZIONE
  qta: 23,        // QTA CONSEGNATA
  importo: 25,    // IMPORTO CONSEGNATO
};
const VENDITE_HEADER_CHECK = {
  0: "ANNO SPEDIZIONE",
  1: "TIPO SPEDIZIONE",
  7: "DATA SPEDIZIONE",
  10: "CLIENTE",
  16: "RAGIONE SOCIALE 1",
  19: "ARTICOLO",
  23: "QTA CONSEGNATA",
  25: "IMPORTO CONSEGNATO",
};
const ORDINI_IDX = {
  anno: 0,        // ANNO
  num: 2,         // NUM.
  data: 6,        // DATA CREAZIONE
  cliente: 7,     // CLIENTE (codice, prima occorrenza)
  ragione: 12,    // CLIENTE (ragione sociale, seconda occorrenza)
  articolo: 13,   // ARTICOLO
  descrizione: 20, // DESCRIZIONE
  qtaInevasa: 24, // QTA INEVASA
  importoInevaso: 25, // IMPORTO INEVASO
};
const ORDINI_HEADER_CHECK = {
  0: "ANNO",
  2: "NUM.",
  6: "DATA CREAZIONE",
  7: "CLIENTE",
  12: "CLIENTE",
  13: "ARTICOLO",
  24: "QTA INEVASA",
  25: "IMPORTO INEVASO",
};

const state = {
  // workbook caricati ma non ancora elaborati
  wbAnagrafica: null,
  wbOrdini: null,
  wbVendite: null,
  // dati parsati (popolati da process())
  anagrafica: null,    // Map<code, { ragione, emails: Set<string> }>
  venditeRows: null,   // Array<{ anno, code, ragione, isReso, importo, date }>
  ordiniRows: null,    // Array<{ anno, code, ragione, num, backlog, articolo, descrizione, data }>
  years: [],           // anni distinti, decrescente
  filtroAnno: null,    // null = tutti, altrimenti number
  // risultato corrente del filtro (popolato da recomputeAndRender)
  result: null,        // { vendite: Map, ordini: Map, maillist: Array }
};

// ---------- Utilità ----------

const fmtEuro = (n) =>
  (Number(n) || 0).toLocaleString("it-IT", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const fmtInt = (n) => (Number(n) || 0).toLocaleString("it-IT");

const fmtDate = (d) => {
  if (!d) return "";
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date.getTime())) return "";
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};

const toNumber = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/\./g, "").replace(",", ".").replace(/[^\d\-.]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

const toDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + v * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    const dt = new Date(Number(y), Number(mo) - 1, Number(d));
    return isNaN(dt.getTime()) ? null : dt;
  }
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? null : dt;
};

const toYear = (annoCell, fallbackDate) => {
  if (annoCell !== null && annoCell !== undefined && annoCell !== "") {
    const n = parseInt(annoCell, 10);
    if (!isNaN(n)) return n;
  }
  if (fallbackDate instanceof Date && !isNaN(fallbackDate.getTime())) {
    return fallbackDate.getFullYear();
  }
  return null;
};

const isValidEmail = (s) => {
  if (!s) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s).trim());
};

const findSheet = (workbook, wanted) => {
  if (workbook.SheetNames.includes(wanted)) return workbook.Sheets[wanted];
  const norm = (s) => s.trim().toUpperCase();
  const target = norm(wanted);
  for (const name of workbook.SheetNames) {
    if (norm(name) === target) return workbook.Sheets[name];
  }
  return null;
};

const checkColumns = (rows, required, label) => {
  if (!rows.length) throw new Error(`${label}: nessuna riga trovata.`);
  const have = new Set(Object.keys(rows[0]));
  const missing = required.filter((c) => !have.has(c));
  if (missing.length) {
    throw new Error(`${label}: colonne mancanti → ${missing.join(", ")}`);
  }
};

const checkHeaderAt = (headers, expected, label) => {
  const norm = (s) => String(s || "").trim().toUpperCase();
  const errs = [];
  for (const idx of Object.keys(expected)) {
    const want = norm(expected[idx]);
    const got = norm(headers[idx]);
    if (got !== want) errs.push(`[${idx}] attesa "${expected[idx]}", trovata "${headers[idx] ?? ""}"`);
  }
  if (errs.length) {
    throw new Error(`${label}: header non corrisponde alla struttura attesa.\n${errs.join("\n")}`);
  }
};

const readWorkbook = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
          cellDates: true,
        });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Errore lettura file " + file.name));
    reader.readAsArrayBuffer(file);
  });

// ---------- Parsing ----------

function parseAnagrafica(workbook) {
  const sheet = findSheet(workbook, SHEETS.anagrafica);
  if (!sheet) throw new Error(`Anagrafica: foglio "${SHEETS.anagrafica}" non trovato.`);
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  console.log("Anagrafica:", rows.length, "righe, prima riga:", rows[0]);
  checkColumns(rows, COLUMNS.anagrafica, "Anagrafica");

  // Più righe possono avere lo stesso CLIENTE, una per email. Aggreghiamo.
  const map = new Map(); // code -> { ragione, emails: Set }
  for (const r of rows) {
    const codeRaw = r["CLIENTE"];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;
    if (!map.has(code)) {
      map.set(code, { ragione: "", emails: new Set() });
    }
    const e = map.get(code);
    const ragione = r["RAGIONE SOCIALE 1"] ? String(r["RAGIONE SOCIALE 1"]).trim() : "";
    if (!e.ragione && ragione) e.ragione = ragione;
    const emailRaw = r["EMAIL"] ? String(r["EMAIL"]).trim().toLowerCase() : "";
    if (emailRaw && isValidEmail(emailRaw) && !EMAIL_ESCLUSE.has(emailRaw)) {
      e.emails.add(emailRaw);
    }
  }
  return map;
}

function parseVendite(workbook) {
  const sheet = findSheet(workbook, SHEETS.vendite);
  if (!sheet) throw new Error(`Vendite: foglio "${SHEETS.vendite}" non trovato.`);
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  if (!rows.length) throw new Error("Vendite: nessuna riga trovata.");
  const headers = rows[0];
  console.log("Vendite headers:", headers);
  checkHeaderAt(headers, VENDITE_HEADER_CHECK, "Vendite");

  const I = VENDITE_IDX;
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const codeRaw = row[I.cliente];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;

    const tipo = row[I.tipo];
    const isReso = !!(tipo && String(tipo).trim().toUpperCase().startsWith("R"));
    const date = toDate(row[I.data]);
    out.push({
      anno: toYear(row[I.anno], date),
      code,
      ragione: row[I.ragione] ? String(row[I.ragione]).trim() : "",
      isReso,
      importo: toNumber(row[I.importo]),
      date,
    });
  }
  return out;
}

function parseOrdini(workbook) {
  const sheet = findSheet(workbook, SHEETS.ordini);
  if (!sheet) throw new Error(`Ordini: foglio "${SHEETS.ordini}" non trovato.`);
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  if (!rows.length) throw new Error("Ordini: nessuna riga trovata.");
  const headers = rows[0];
  console.log("Ordini headers:", headers);
  checkHeaderAt(headers, ORDINI_HEADER_CHECK, "Ordini");

  const I = ORDINI_IDX;
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const qta = toNumber(row[I.qtaInevasa]);
    const imp = toNumber(row[I.importoInevaso]);
    if (qta <= 0 && imp <= 0) continue;

    const codeRaw = row[I.cliente];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;

    const date = toDate(row[I.data]);
    const numRaw = row[I.num];
    out.push({
      anno: toYear(row[I.anno], date),
      code,
      ragione: row[I.ragione] ? String(row[I.ragione]).trim() : "",
      num: numRaw !== null && numRaw !== undefined && numRaw !== "" ? String(numRaw).trim() : "",
      backlog: imp,
      articolo: row[I.articolo] ? String(row[I.articolo]).trim() : "",
      descrizione: row[I.descrizione] ? String(row[I.descrizione]).trim() : "",
      data: date,
    });
  }
  return out;
}

// ---------- Aggregazioni in funzione del filtro anno ----------

function aggregateVendite(records, year) {
  const agg = new Map(); // code -> { ragione, count, revenue, lastDate }
  for (const r of records) {
    if (year !== null && r.anno !== year) continue;
    const signed = r.isReso ? -r.importo : r.importo;
    if (!agg.has(r.code)) {
      agg.set(r.code, { ragione: r.ragione, count: 0, revenue: 0, lastDate: null });
    }
    const e = agg.get(r.code);
    if (!e.ragione && r.ragione) e.ragione = r.ragione;
    e.count += 1;
    e.revenue += signed;
    if (r.date && (!e.lastDate || r.date > e.lastDate)) e.lastDate = r.date;
  }
  return agg;
}

function aggregateOrdini(records, year) {
  const agg = new Map(); // code -> { ragione, orders: Set, backlog, articoli: [] }
  for (const r of records) {
    if (year !== null && r.anno !== year) continue;
    if (!agg.has(r.code)) {
      agg.set(r.code, { ragione: r.ragione, orders: new Set(), backlog: 0, articoli: [] });
    }
    const e = agg.get(r.code);
    if (!e.ragione && r.ragione) e.ragione = r.ragione;
    if (r.num) e.orders.add(r.num);
    e.backlog += r.backlog;
    const label = [r.articolo, r.descrizione].filter(Boolean).join(" — ");
    if (label && !e.articoli.includes(label)) e.articoli.push(label);
  }
  return agg;
}

function buildMailList(anagrafica, vendite, ordini) {
  const out = [];
  for (const [code, info] of anagrafica) {
    const hasSales = vendite.has(code);
    const hasOpenOrders = ordini.has(code);
    if (!hasSales && !hasOpenOrders) continue;
    if (info.emails.size === 0) continue;
    let ragione = info.ragione;
    if (!ragione && hasSales) ragione = vendite.get(code).ragione;
    if (!ragione && hasOpenOrders) ragione = ordini.get(code).ragione;
    for (const email of info.emails) {
      out.push({ email, code, ragione: ragione || "", hasSales, hasOpenOrders });
    }
  }
  out.sort((a, b) => a.email.localeCompare(b.email));
  return out;
}

// ---------- Render ----------

const $ = (id) => document.getElementById(id);

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderTable(target, headers, rows) {
  const t = $(target);
  if (!rows.length) {
    t.innerHTML = `<tr><td class="empty" colspan="${headers.length}">Nessun dato.</td></tr>`;
    return;
  }
  const thead =
    "<thead><tr>" +
    headers.map((h) => `<th class="${h.num ? "num" : ""}">${escapeHtml(h.label)}</th>`).join("") +
    "</tr></thead>";
  const tbody =
    "<tbody>" +
    rows
      .map((r) =>
        "<tr>" +
        headers.map((h) => `<td class="${h.num ? "num" : ""}">${escapeHtml(r[h.key] ?? "")}</td>`).join("") +
        "</tr>"
      ).join("") +
    "</tbody>";
  t.innerHTML = thead + tbody;
}

function annoLabel() {
  return state.filtroAnno === null ? "tutti gli anni" : `anno ${state.filtroAnno}`;
}

function renderVendite(vendite) {
  const arr = [];
  let totale = 0;
  for (const [code, e] of vendite) {
    arr.push({ code, ragione: e.ragione, count: e.count, revenue: e.revenue, lastDate: e.lastDate });
    totale += e.revenue;
  }
  arr.sort((a, b) => b.revenue - a.revenue);

  $("summaryVendite").innerHTML =
    `<span><strong>${fmtInt(arr.length)}</strong> clienti con acquisti (${annoLabel()})</span>` +
    `<span>fatturato totale: <strong>${fmtEuro(totale)}</strong></span>`;

  renderTable(
    "tableVendite",
    [
      { key: "code", label: "Codice" },
      { key: "ragione", label: "Ragione sociale" },
      { key: "countF", label: "Spedizioni", num: true },
      { key: "revenueF", label: "Fatturato", num: true },
      { key: "lastDateF", label: "Ultimo acquisto" },
    ],
    arr.map((r) => ({
      ...r,
      countF: fmtInt(r.count),
      revenueF: fmtEuro(r.revenue),
      lastDateF: fmtDate(r.lastDate),
    }))
  );
  $("cardVendite").style.display = "";
}

function renderOrdini(ordini) {
  const arr = [];
  let backlogTot = 0;
  for (const [code, e] of ordini) {
    arr.push({
      code,
      ragione: e.ragione,
      orders: e.orders.size,
      backlog: e.backlog,
      articoli: e.articoli.slice(0, 3).join(" · ") + (e.articoli.length > 3 ? " …" : ""),
    });
    backlogTot += e.backlog;
  }
  arr.sort((a, b) => b.backlog - a.backlog);

  $("summaryOrdini").innerHTML =
    `<span><strong>${fmtInt(arr.length)}</strong> clienti con ordini aperti (${annoLabel()})</span>` +
    `<span>backlog totale: <strong>${fmtEuro(backlogTot)}</strong></span>`;

  renderTable(
    "tableOrdini",
    [
      { key: "code", label: "Codice" },
      { key: "ragione", label: "Ragione sociale" },
      { key: "ordersF", label: "Ordini aperti", num: true },
      { key: "backlogF", label: "Backlog", num: true },
      { key: "articoli", label: "Articoli (max 3)" },
    ],
    arr.map((r) => ({
      ...r,
      ordersF: fmtInt(r.orders),
      backlogF: fmtEuro(r.backlog),
    }))
  );
  $("cardOrdini").style.display = "";
}

function renderMaillist(list) {
  $("summaryMaillist").innerHTML =
    `<span><strong>${fmtInt(list.length)}</strong> indirizzi email (${annoLabel()})</span>`;
  renderTable(
    "tableMaillist",
    [
      { key: "email", label: "Email" },
      { key: "code", label: "Codice" },
      { key: "ragione", label: "Ragione sociale" },
      { key: "salesF", label: "Ha acquistato?" },
      { key: "ordersF", label: "Ordini aperti?" },
    ],
    list.map((r) => ({
      ...r,
      salesF: r.hasSales ? "Sì" : "No",
      ordersF: r.hasOpenOrders ? "Sì" : "No",
    }))
  );
  $("cardMaillist").style.display = "";
}

function recomputeAndRender() {
  if (!state.anagrafica || !state.venditeRows || !state.ordiniRows) return;
  const year = state.filtroAnno;
  const vendite = aggregateVendite(state.venditeRows, year);
  const ordini = aggregateOrdini(state.ordiniRows, year);
  const maillist = buildMailList(state.anagrafica, vendite, ordini);
  state.result = { vendite, ordini, maillist };
  console.log("Render:", {
    annoFiltro: year,
    clientiVendite: vendite.size,
    clientiOrdini: ordini.size,
    mailListRighe: maillist.length,
  });
  renderVendite(vendite);
  renderOrdini(ordini);
  renderMaillist(maillist);
}

// ---------- Export ----------

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const MAILLIST_HEADERS = ["Email", "Codice cliente", "Ragione sociale", "Ha acquistato?", "Ha ordini aperti?"];
const maillistRow = (r) => [r.email, r.code, r.ragione, r.hasSales ? "Sì" : "No", r.hasOpenOrders ? "Sì" : "No"];

function downloadCsv(list) {
  const escape = (v) => {
    const s = v === null || v === undefined ? "" : String(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [MAILLIST_HEADERS.map(escape).join(",")];
  for (const r of list) lines.push(maillistRow(r).map(escape).join(","));
  const content = "﻿" + lines.join("\r\n");
  triggerDownload(new Blob([content], { type: "text/csv;charset=utf-8" }), "maillist.csv");
}

function downloadXlsx(list) {
  const aoa = [MAILLIST_HEADERS, ...list.map(maillistRow)];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, "MailList");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  triggerDownload(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "maillist.xlsx"
  );
}

function downloadTxt(list) {
  // Solo email, una per riga, UTF-8 senza BOM, fine riga \n.
  const text = list.map((r) => r.email).join("\n") + (list.length ? "\n" : "");
  triggerDownload(new Blob([text], { type: "text/plain;charset=utf-8" }), "maillist.txt");
}

// ---------- UI wiring ----------

function setStatus(id, msg, ok) {
  const el = $(id);
  el.textContent = msg;
  el.className = "status " + (msg ? (ok ? "ok" : "err") : "");
}

function showError(msg) {
  const box = $("errorBox");
  if (!msg) {
    box.style.display = "none";
    box.textContent = "";
    return;
  }
  // "block" esplicito: la regola CSS imposta display:none, quindi clear inline non basta.
  box.style.display = "block";
  box.textContent = msg;
}

function bindFile(inputId, statusId, key) {
  $(inputId).addEventListener("change", async (ev) => {
    const f = ev.target.files[0];
    if (!f) {
      state[key] = null;
      setStatus(statusId, "", true);
      return;
    }
    setStatus(statusId, "Lettura in corso…", true);
    try {
      const wb = await readWorkbook(f);
      state[key] = wb;
      setStatus(statusId, `Caricato: ${f.name}`, true);
    } catch (err) {
      state[key] = null;
      setStatus(statusId, `Errore: ${err.message || err}`, false);
    }
  });
}

function populateYearSelect(years) {
  const sel = $("filtroAnno");
  sel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "Tutti gli anni";
  sel.appendChild(optAll);
  for (const y of years) {
    const o = document.createElement("option");
    o.value = String(y);
    o.textContent = String(y);
    sel.appendChild(o);
  }
  sel.value = "";
}

function reset() {
  state.wbAnagrafica = null;
  state.wbOrdini = null;
  state.wbVendite = null;
  state.anagrafica = null;
  state.venditeRows = null;
  state.ordiniRows = null;
  state.years = [];
  state.filtroAnno = null;
  state.result = null;
  ["fileAnagrafica", "fileOrdini", "fileVendite"].forEach((id) => ($(id).value = ""));
  ["statusAnagrafica", "statusOrdini", "statusVendite"].forEach((id) => setStatus(id, "", true));
  ["cardFiltro", "cardVendite", "cardOrdini", "cardMaillist"].forEach((id) => ($(id).style.display = "none"));
  showError("");
}

function process() {
  console.log("Elabora cliccato");
  try {
    console.log("File caricati:", {
      hasAnagrafica: !!state.wbAnagrafica,
      hasOrdini: !!state.wbOrdini,
      hasVendite: !!state.wbVendite,
    });

    const missing = [];
    if (!state.wbAnagrafica) missing.push("Anagrafica");
    if (!state.wbOrdini) missing.push("Ordini");
    if (!state.wbVendite) missing.push("Vendite");
    if (missing.length) {
      showError("Carica tutti e 3 i file prima di elaborare. Mancano: " + missing.join(", "));
      return;
    }
    showError("");

    const anagrafica = parseAnagrafica(state.wbAnagrafica);
    console.log("Anagrafica parsata. Clienti unici:", anagrafica.size);

    const venditeRows = parseVendite(state.wbVendite);
    console.log("Vendite parsate. Righe utili:", venditeRows.length);

    const ordiniRows = parseOrdini(state.wbOrdini);
    console.log("Ordini parsati. Righe aperte:", ordiniRows.length);

    const yearSet = new Set();
    for (const r of venditeRows) if (r.anno !== null) yearSet.add(r.anno);
    for (const r of ordiniRows) if (r.anno !== null) yearSet.add(r.anno);
    const years = [...yearSet].sort((a, b) => b - a);
    console.log("Anni distinti:", years);

    state.anagrafica = anagrafica;
    state.venditeRows = venditeRows;
    state.ordiniRows = ordiniRows;
    state.years = years;
    state.filtroAnno = null;

    populateYearSelect(years);
    $("cardFiltro").style.display = "";

    recomputeAndRender();
    $("cardFiltro").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    console.error("Errore in process:", err);
    showError(err && err.message ? err.message : String(err));
    ["cardFiltro", "cardVendite", "cardOrdini", "cardMaillist"].forEach((id) => ($(id).style.display = "none"));
  }
}

document.addEventListener("DOMContentLoaded", () => {
  bindFile("fileAnagrafica", "statusAnagrafica", "wbAnagrafica");
  bindFile("fileOrdini", "statusOrdini", "wbOrdini");
  bindFile("fileVendite", "statusVendite", "wbVendite");
  $("btnProcess").addEventListener("click", process);
  $("btnReset").addEventListener("click", reset);
  $("filtroAnno").addEventListener("change", (ev) => {
    const v = ev.target.value;
    state.filtroAnno = v === "" ? null : Number(v);
    recomputeAndRender();
  });
  $("btnDownloadCsv").addEventListener("click", () => {
    if (state.result && state.result.maillist) downloadCsv(state.result.maillist);
  });
  $("btnDownloadXlsx").addEventListener("click", () => {
    if (state.result && state.result.maillist) downloadXlsx(state.result.maillist);
  });
  $("btnDownloadTxt").addEventListener("click", () => {
    if (state.result && state.result.maillist) downloadTxt(state.result.maillist);
  });
});
