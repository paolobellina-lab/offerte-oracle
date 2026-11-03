import * as db from "./firebase.js";
import {
  TIPI, FONTI, CAMPI_SCONTO, ETICHETTE, COLONNE_ORACLE, num, isAttivita, codiceOracle, prezzoListino,
  listinoCalcolato, prezzoBase, calcolaOfferta, righeOracle, controlloCartellino, freschezza, ordinaRighe,
} from "./calc.js";
import { analizzaFile, leggiAggiornamentoListino, righeDaTestoIncollato, leggiExcel, slug } from "./importa.js";
import { scriviExcel } from "./xlsx-lite.js";

/* ================= Stato ================= */
const COLL = { cat: "oracle_catalogo", cfg: "oracle_config", cli: "oracle_clienti", off: "oracle_offerte" };
const S = {
  capitoli: [], gmv: [], clienti: [], chunks: [], catalogo: [], perK: new Map(), offerte: [],
  off: null, vista: "offerte", sub: "righe", dirty: false, salvando: false, utente: null,
  limiteCat: 300, limitePick: 300, pickCap: "", importato: null, aggOracle: null,
};
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));

/* ================= Utilità ================= */
export const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fEuro = new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" });
const euro = (n) => fEuro.format(Number(n) || 0);
const fNum = new Intl.NumberFormat("it-IT", { maximumFractionDigits: 4 });
const numIt = (n) => fNum.format(Number(n) || 0);
const pctTxt = (f) => (num(f) ? fNum.format(Math.round(num(f) * 1e6) / 1e4) + "%" : "");
const pctVal = (f) => (num(f) ? String(Math.round(num(f) * 1e6) / 1e4).replace(".", ",") : "");
const daPct = (s) => num(s) / 100;
const oggiIso = () => new Date().toISOString().slice(0, 10);
const dataIt = (iso) => (iso ? new Date(iso).toLocaleDateString("it-IT") : "");
const dataOra = (ms) => (ms ? new Date(ms).toLocaleString("it-IT", { dateStyle: "short", timeStyle: "short" }) : "");
const ricarichi = () => Object.fromEntries(S.capitoli.map((c) => [c.nome, num(c.ricarico)]));
const ordineCap = () => S.capitoli.map((c) => c.nome);
const opzCalc = () => ({ ricarichi: ricarichi(), ordineCapitoli: ordineCap() });
const clienteOff = () => S.clienti.find((c) => c.id === S.off?.cliente) || null;
const fornitore = (g) => S.gmv.find((x) => x.codice === g)?.fornitore || "";

let toastT;
function toast(msg, ms = 3500) {
  const t = $("#toast");
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastT);
  toastT = setTimeout(() => (t.hidden = true), ms);
}
function errore(e, cosa = "Operazione non riuscita") {
  console.error(e);
  toast(e?.code === "permission-denied"
    ? `L'account ${S.utente?.email || ""} non è autorizzato: aggiungilo alle regole di Firestore.`
    : `${cosa}: ${e?.message || e}`, 6000);
}
function scarica(blob, nome) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nome;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* ================= Finestra di dialogo ================= */
let dlgHandler = null;
function dialogo({ titolo, html, pulsanti = [{ label: "Chiudi", value: "annulla" }], onApri, onConferma }) {
  $("#dlg-titolo").textContent = titolo;
  $("#dlg-body").innerHTML = html;
  $("#dlg-foot").innerHTML = pulsanti.map((p) =>
    `<button class="btn ${p.cls || ""}" value="${p.value}" ${p.value === "annulla" ? 'formnovalidate' : ""}>${esc(p.label)}</button>`).join("");
  dlgHandler = onConferma;
  onApri?.($("#dlg-body"));
  $("#dlg").showModal();
}
$("#dlg-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const v = e.submitter?.value || "annulla";
  if (v === "annulla" || !dlgHandler) return $("#dlg").close();
  try {
    const r = await dlgHandler(v, $("#dlg-body"));
    if (r !== false) $("#dlg").close();
  } catch (x) { errore(x); }
});

/* ================= Avvio ================= */
let ascoltatori = [];
db.avvia({
  onLogin(u) {
    S.utente = u;
    $("#login").hidden = true;
    $("#app").hidden = false;
    $("#banner-prova").hidden = !u.prova;
    $("#utente").textContent = u.email;
    collega();
  },
  onLogout() {
    ascoltatori.forEach((f) => f());
    ascoltatori = [];
    $("#app").hidden = true;
    $("#login").hidden = false;
  },
  onErrore: (e) => ($("#login .err").textContent = "Accesso non riuscito: " + (e.message || e)),
}).catch((e) => {
  $("#login").hidden = false;
  $("#login .err").textContent = "Firebase non raggiungibile: " + (e.message || e);
});

$("#google-login").addEventListener("click", async () => {
  $("#login .err").textContent = "";
  try { await db.accediGoogle(); } catch (x) { $("#login .err").textContent = "Accesso non riuscito: " + (x.message || x); }
});
$("#esci").addEventListener("click", () => db.esci());

function collega() {
  const onErr = (e) => errore(e, "Errore di lettura");
  ascoltatori.push(db.ascolta(COLL.cfg, (docs) => {
    const cap = docs.find((d) => d.id === "capitoli");
    const gmv = docs.find((d) => d.id === "gmv");
    S.capitoli = (cap?.lista || []).slice().sort((a, b) => num(a.ordine) - num(b.ordine));
    S.gmv = gmv?.lista || [];
    ridisegna();
  }, onErr));
  ascoltatori.push(db.ascolta(COLL.cat, (docs) => {
    S.chunks = docs;
    S.catalogo = [];
    for (const d of docs) for (const a of d.articoli || []) S.catalogo.push({ ...a, _chunk: d.id });
    S.catalogo.sort((a, b) => a.capitolo.localeCompare(b.capitolo) || num(a.ordine) - num(b.ordine));
    S.perK = new Map(S.catalogo.map((a) => [a.k, a]));
    ridisegna();
  }, onErr));
  ascoltatori.push(db.ascolta(COLL.cli, (docs) => {
    S.clienti = docs.sort((a, b) => (a.nome || "").localeCompare(b.nome || ""));
    ridisegna();
  }, onErr));
  ascoltatori.push(db.ascolta(COLL.off, (docs) => {
    S.offerte = docs.sort((a, b) => num(b.aggiornatoIl) - num(a.aggiornatoIl));
    if (S.off && !S.dirty && !S.salvando) {
      const remoto = docs.find((d) => d.id === S.off.id);
      if (!remoto) { S.off = null; if (S.vista === "offerta") vai("offerte"); }
      else if (num(remoto.aggiornatoIl) > num(S.off.aggiornatoIl)) { S.off = remoto; if (S.vista === "offerta") apriOfferta(remoto.id, true); }
    }
    if (!S.off && location.hash.startsWith("#offerta=")) {
      const id = decodeURIComponent(location.hash.slice(9));
      if (docs.some((d) => d.id === id)) apriOfferta(id);
    }
    if (S.vista !== "offerta") ridisegna(); // l'offerta aperta si ridisegna solo se cambiata altrove (sopra)
  }, onErr));
}

/* ================= Navigazione ================= */
function vai(vista) {
  if (S.vista === "offerta" && vista !== "offerta") salvaOraSeServe();
  S.vista = vista;
  for (const v of ["offerte", "offerta", "catalogo", "clienti", "importa"]) $("#v-" + v).hidden = v !== vista;
  $$("#tabs button").forEach((b) => b.classList.toggle("on", b.dataset.vista === (vista === "offerta" ? "offerte" : vista)));
  if (vista !== "offerta" && location.hash) history.replaceState(null, "", location.pathname);
  ridisegna();
  window.scrollTo(0, 0);
}
$("#tabs").addEventListener("click", (e) => { const b = e.target.closest("button"); if (b) vai(b.dataset.vista); });

let rafRid = 0, rinviato = false;
const staScrivendo = () => {
  const a = document.activeElement;
  return a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && a.closest("#v-offerta, #v-catalogo, #v-clienti") && !$("#dlg").open;
};
document.addEventListener("focusout", () => setTimeout(() => {
  if (rinviato && !staScrivendo()) { rinviato = false; ridisegna(); }
}, 50));
function ridisegna() {
  cancelAnimationFrame(rafRid);
  rafRid = requestAnimationFrame(() => {
    // dati arrivati mentre si scrive in un campo: si ridisegna appena il campo perde il focus
    if (staScrivendo() && S.vista !== "offerte" && S.vista !== "importa") { rinviato = true; return; }
    if (S.vista === "offerte") disegnaElencoOfferte();
    else if (S.vista === "offerta") disegnaOfferta();
    else if (S.vista === "catalogo") disegnaCatalogo();
    else if (S.vista === "clienti") disegnaClienti();
  });
}

/* ================= Elenco offerte ================= */
function nuovaOffertaVuota() {
  const ora = Date.now();
  return {
    id: db.nuovoId(), nome: "", cliente: "", tipo: "LISTINO", prezzi: "DA ORACLE", sottocapitoli: false,
    lineaPartenza: 10, magazzino: "BG", globali: {}, gmv: {}, righe: [], note: "",
    creatoIl: ora, aggiornatoIl: ora, creatoDa: S.utente?.email || "",
  };
}
$("#nuova-offerta").addEventListener("click", async () => {
  const o = nuovaOffertaVuota();
  try { await db.salva(COLL.off, o.id, o); } catch (x) { return errore(x); }
  apriOfferta(o.id, false, o);
  S.sub = "aggiungi";
  disegnaOfferta(true);
  $("#o-nome").focus();
});
$("#cerca-offerte").addEventListener("input", disegnaElencoOfferte);

function disegnaElencoOfferte() {
  const q = $("#cerca-offerte").value.trim().toLowerCase();
  const nomeCli = (id) => S.clienti.find((c) => c.id === id)?.nome || "";
  const lista = S.offerte.filter((o) => !q || (o.nome || "").toLowerCase().includes(q) || nomeCli(o.cliente).toLowerCase().includes(q));
  $("#lista-offerte").innerHTML = lista.length ? lista.map((o) => {
    const { righe, totale } = calcolaOfferta(o, opzCalc());
    return `<tr class="cliccabile" data-id="${esc(o.id)}">
      <td><b>${esc(o.nome || "(senza riferimento)")}</b></td><td>${esc(nomeCli(o.cliente))}</td>
      <td>${esc(o.tipo)}</td><td>${esc(o.prezzi)}</td><td class="num">${righe.length}</td>
      <td class="num">${euro(totale)}</td><td class="nowrap muted">${dataOra(o.aggiornatoIl)}</td>
      <td class="nowrap"><button class="btn small" data-dup="${esc(o.id)}">Duplica</button></td></tr>`;
  }).join("") : `<tr><td colspan="8" class="vuoto">${S.offerte.length ? "Nessuna offerta trovata." :
    S.catalogo.length ? "Ancora nessuna offerta: crea la prima con “+ Nuova offerta”." :
    "Il catalogo è vuoto: vai in “Importa e aggiorna” e carica il file Excel."}</td></tr>`;
}
$("#lista-offerte").addEventListener("click", (e) => {
  const dup = e.target.closest("[data-dup]");
  if (dup) { e.stopPropagation(); return duplica(S.offerte.find((o) => o.id === dup.dataset.dup)); }
  const tr = e.target.closest("tr[data-id]");
  if (tr) apriOfferta(tr.dataset.id);
});

async function duplica(o) {
  if (!o) return;
  const n = { ...structuredClone(o), id: db.nuovoId(), nome: (o.nome || "Offerta") + " (copia)",
    creatoIl: Date.now(), aggiornatoIl: Date.now(), creatoDa: S.utente?.email || "" };
  try { await db.salva(COLL.off, n.id, n); toast("Offerta duplicata"); apriOfferta(n.id, false, n); } catch (x) { errore(x); }
}

/* ================= Offerta ================= */
function apriOfferta(id, daRemoto = false, dati = null) {
  const o = dati || S.offerte.find((x) => x.id === id);
  if (!o) return;
  S.off = structuredClone(o);
  S.off.globali ||= {}; S.off.gmv ||= {}; S.off.righe ||= [];
  S.dirty = false;
  if (!daRemoto) S.sub = S.off.righe.length ? "righe" : "aggiungi";
  history.replaceState(null, "", "#offerta=" + encodeURIComponent(id));
  if (S.vista !== "offerta") vai("offerta"); else disegnaOfferta(true);
  disegnaTestata();
  statoSalvataggio(daRemoto ? "Aggiornata da un altro dispositivo" : "");
}

let tSalva;
function modificata() {
  S.dirty = true;
  statoSalvataggio("Modifiche non salvate…");
  clearTimeout(tSalva);
  tSalva = setTimeout(salvaOfferta, 700);
}
async function salvaOfferta() {
  if (!S.off || !S.dirty) return;
  clearTimeout(tSalva);
  S.dirty = false;
  S.salvando = true;
  S.off.aggiornatoIl = Date.now();
  const copia = structuredClone(S.off);
  try {
    await db.salva(COLL.off, copia.id, copia);
    statoSalvataggio("Salvata alle " + new Date().toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" }));
  } catch (x) { S.dirty = true; statoSalvataggio("Non salvata"); errore(x, "Salvataggio non riuscito"); }
  finally { S.salvando = false; }
}
const salvaOraSeServe = () => { if (S.dirty) salvaOfferta(); };
window.addEventListener("beforeunload", (e) => { if (S.dirty) { salvaOfferta(); e.preventDefault(); } });
function statoSalvataggio(t) { $("#stato-salvataggio").textContent = t; }

$("#torna-offerte").addEventListener("click", () => { salvaOraSeServe(); S.off = null; vai("offerte"); });
$("#duplica-offerta").addEventListener("click", async () => { await salvaOfferta(); duplica(S.off); });
$("#elimina-offerta").addEventListener("click", () => {
  dialogo({
    titolo: "Eliminare l'offerta?",
    html: `<p>L'offerta <b>${esc(S.off.nome || "(senza riferimento)")}</b> verrà eliminata definitivamente.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Elimina", value: "ok", cls: "danger" }],
    async onConferma() {
      const id = S.off.id;
      S.dirty = false; S.off = null;
      await db.elimina(COLL.off, id);
      vai("offerte");
      toast("Offerta eliminata");
    },
  });
});

/* --- Testata --- */
function seg(el, valori, attuale, etichetta = (v) => v) {
  el.innerHTML = valori.map((v) => `<button type="button" data-v="${esc(String(v))}" class="${String(v) === String(attuale) ? "on" : ""}">${esc(etichetta(v))}</button>`).join("");
}
function disegnaTestata() {
  const o = S.off;
  if (!o) return;
  $("#o-nome").value = o.nome || "";
  $("#o-linea").value = o.lineaPartenza ?? 10;
  $("#o-magazzino").value = o.magazzino ?? "BG";
  $("#o-note").value = o.note || "";
  $("#o-cliente").innerHTML = `<option value="">— nessun cliente (niente controllo cartellino) —</option>` +
    S.clienti.map((c) => `<option value="${esc(c.id)}" ${c.id === o.cliente ? "selected" : ""}>${esc(c.nome)}</option>`).join("");
  seg($("#o-tipo"), TIPI, o.tipo);
  seg($("#o-prezzi"), FONTI, o.prezzi);
  seg($("#o-sotto"), ["SI", "NO"], o.sottocapitoli ? "SI" : "NO");
  $("#o-globali").innerHTML = [...CAMPI_SCONTO, "ric"].map((k) => `
    <label class="f" ${k === "ric" ? 'title="Solo per VENDITA CLIENTE. Vuoto = ricarico automatico del capitolo."' : ""}>${ETICHETTE[k]}${k === "ric" ? " *" : ""}
      <span class="pct-wrap"><input class="pct" data-glob="${k}" inputmode="decimal" value="${esc(pctVal(o.globali[k]))}"
        ${k === "ric" ? (o.tipo !== "VENDITA CLIENTE" ? "disabled" : 'placeholder="auto"') : ""}></span></label>`).join("") +
    `<span class="muted small" style="align-self:flex-end;padding-bottom:7px">* ricarico solo per vendita cliente; vuoto = automatico del capitolo</span>`;
}
$("#o-nome").addEventListener("input", (e) => { S.off.nome = e.target.value; modificata(); });
$("#o-magazzino").addEventListener("input", (e) => { S.off.magazzino = e.target.value.trim(); modificata(); });
$("#o-note").addEventListener("input", (e) => { S.off.note = e.target.value; modificata(); });
$("#o-linea").addEventListener("change", (e) => { S.off.lineaPartenza = num(e.target.value) || 10; modificata(); disegnaOfferta(); });
$("#o-cliente").addEventListener("change", (e) => { S.off.cliente = e.target.value; modificata(); disegnaOfferta(); });
$("#o-tipo").addEventListener("click", (e) => {
  const b = e.target.closest("button"); if (!b) return;
  S.off.tipo = b.dataset.v; modificata(); disegnaTestata(); disegnaOfferta();
});
$("#o-prezzi").addEventListener("click", (e) => {
  const b = e.target.closest("button"); if (!b) return;
  S.off.prezzi = b.dataset.v; modificata(); disegnaTestata(); disegnaOfferta();
});
$("#o-sotto").addEventListener("click", (e) => {
  const b = e.target.closest("button"); if (!b) return;
  S.off.sottocapitoli = b.dataset.v === "SI"; modificata(); disegnaTestata(); disegnaOfferta();
});
$("#o-globali").addEventListener("change", (e) => {
  const k = e.target.dataset.glob; if (!k) return;
  const v = daPct(e.target.value);
  if (v) S.off.globali[k] = v; else delete S.off.globali[k];
  e.target.value = pctVal(v);
  modificata(); ridisegnaOfferta();
});

$("#subtabs").addEventListener("click", (e) => {
  const b = e.target.closest("button"); if (!b) return;
  S.sub = b.dataset.sub; disegnaOfferta(true);
});

/* Dopo un ridisegno rimette il cursore nel campo dove si trovava (Tab da una quantità alla successiva). */
function ricordaFocus() {
  const a = document.activeElement;
  if (!a || a.tagName !== "INPUT") return () => {};
  const tr = a.closest("tr[data-k]");
  let sel = null;
  if (tr && a.dataset.campo) sel = `tr[data-k="${CSS.escape(tr.dataset.k)}"] input[data-campo="${a.dataset.campo}"]`;
  else if (tr && a.dataset.pick) sel = `tr[data-k="${CSS.escape(tr.dataset.k)}"] input[data-pick="${a.dataset.pick}"]`;
  else if (a.dataset.gmv) sel = `input[data-gmv="${CSS.escape(a.dataset.gmv)}"][data-k="${a.dataset.k}"]`;
  else if (a.dataset.glob) sel = `input[data-glob="${a.dataset.glob}"]`;
  if (!sel) return () => {};
  return () => { const n = document.querySelector(sel); if (n && n !== a) { n.focus(); n.select?.(); } };
}

// dopo un "change" il focus si è già spostato sul campo successivo: si ridisegna un attimo dopo
const ridisegnaOfferta = () => setTimeout(() => disegnaOfferta(), 0);
function disegnaOfferta(cambioSub = false) {
  const ripristina = ricordaFocus();
  disegnaOffertaIn(cambioSub);
  ripristina();
}
function disegnaOffertaIn(cambioSub) {
  const o = S.off;
  if (!o) return;
  $$("#subtabs button").forEach((b) => b.classList.toggle("on", b.dataset.sub === S.sub));
  for (const s of ["righe", "aggiungi", "gmv", "export", "note"]) $("#s-" + s).hidden = s !== S.sub;
  const calc = calcolaOfferta(o, opzCalc());
  const cart = controlloCartellino(o, clienteOff(), opzCalc());
  const fuori = cart.filter((c) => c.stato === "fuori").length;
  const gmvPresenti = new Set(calc.righe.map((x) => x.riga.gmv).filter(Boolean));
  $("#conta-righe").textContent = calc.righe.length ? `(${calc.righe.length})` : "";
  $("#conta-gmv").textContent = gmvPresenti.size ? `(${gmvPresenti.size})` : "";
  $("#o-totali").innerHTML = `
    <div class="kpi"><span>Totale offerta</span><span class="big">${euro(calc.totale)}</span></div>
    <div class="kpi"><span>Righe</span><span class="big">${calc.righe.length}</span></div>
    <div class="kpi"><span>Tipo</span><span>${esc(o.tipo)} · ${esc(o.prezzi)}</span></div>
    ${o.cliente ? `<div class="kpi"><span>Cartellino ${esc(clienteOff()?.nome || "")}</span><span>${
      fuori ? `<span class="badge bad">${fuori} GMV fuori cartellino</span>` : `<span class="badge ok">nessun problema</span>`}</span></div>` : ""}`;
  if (S.sub === "righe") disegnaRighe(calc);
  else if (S.sub === "aggiungi") disegnaPicker(cambioSub);
  else if (S.sub === "gmv") disegnaGmv(calc, cart);
  else if (S.sub === "export") disegnaExport();
}

/* --- Righe dell'offerta --- */
function disegnaRighe(calc) {
  const tb = $("#righe-offerta");
  if (!calc.righe.length) {
    tb.innerHTML = `<tr><td colspan="15" class="vuoto">Nessuna riga. Vai su “Aggiungi articoli” e scrivi le quantità.</td></tr>`;
    return;
  }
  let cap = null, sotto = null, html = "";
  let totCap = 0;
  const chiudiCap = () => cap !== null ? `<tr class="sotto"><td colspan="12" class="num">Totale ${esc(cap)}</td><td class="num">${euro(totCap)}</td><td colspan="2"></td></tr>` : "";
  for (const { riga: r, c } of calc.righe) {
    if (r.capitolo !== cap) {
      html += chiudiCap();
      cap = r.capitolo; sotto = null; totCap = 0;
      html += `<tr class="gruppo"><td colspan="15">${esc(cap)}</td></tr>`;
    }
    if ((r.sottocapitolo || "") !== sotto) {
      sotto = r.sottocapitolo || "";
      if (sotto) html += `<tr class="sotto"><td colspan="15">${esc(sotto)}</td></tr>`;
    }
    totCap += c.totale;
    const att = isAttivita(r);
    html += `<tr data-k="${esc(r.k)}">
      <td class="mono nowrap">${esc(codiceOracle(r)) || '<span class="badge bad">senza codice</span>'}</td>
      <td class="desc"><b>${esc(r.nOrdine)}</b><div class="d">${esc(r.descrizione)}</div></td>
      <td>${esc(r.um)}</td>
      <td class="num"><input class="qta" data-campo="qta" inputmode="decimal" value="${esc(numIt(r.qta))}"></td>
      <td class="num">${att ? `<input class="qta" data-campo="giorni" inputmode="decimal" value="${esc(numIt(r.giorni))}">` : ""}</td>
      <td class="num">${euro(c.base)}</td>
      <td class="num">${pctTxt(c.sc3)}</td><td class="num">${pctTxt(c.sc4)}</td><td class="num">${pctTxt(c.sc5)}</td>
      <td class="num">${pctTxt(c.aum1)}</td><td class="num">${pctTxt(c.ric)}</td>
      <td class="num"><b>${euro(c.unitario)}</b></td><td class="num">${euro(c.totale)}</td>
      <td class="mono">${esc(r.gmv)}</td>
      <td><button class="icon" data-togli="${esc(r.k)}" title="Togli la riga">✕</button></td></tr>`;
  }
  html += chiudiCap();
  html += `<tr class="sotto"><td colspan="12" class="num">Totale offerta</td><td class="num">${euro(calc.totale)}</td><td colspan="2"></td></tr>`;
  tb.innerHTML = html;
}
$("#righe-offerta").addEventListener("change", (e) => {
  const inp = e.target.closest("input[data-campo]"); if (!inp) return;
  const k = inp.closest("tr").dataset.k;
  const r = S.off.righe.find((x) => x.k === k); if (!r) return;
  r[inp.dataset.campo] = num(inp.value);
  if (inp.dataset.campo === "qta" && !r.qta) S.off.righe = S.off.righe.filter((x) => x.k !== k);
  modificata(); ridisegnaOfferta();
});
$("#righe-offerta").addEventListener("click", (e) => {
  const b = e.target.closest("[data-togli]"); if (!b) return;
  S.off.righe = S.off.righe.filter((x) => x.k !== b.dataset.togli);
  modificata(); disegnaOfferta();
});
$("#aggiorna-da-catalogo").addEventListener("click", () => {
  let n = 0, mancanti = 0;
  S.off.righe = S.off.righe.map((r) => {
    const a = S.perK.get(r.k) || S.catalogo.find((x) => x.capitolo === r.capitolo && codiceOracle(x) === codiceOracle(r) && x.nOrdine === r.nOrdine);
    if (!a) { mancanti++; return r; }
    n++;
    return { ...rigaDaArticolo(a), qta: r.qta, giorni: isAttivita(r) ? r.giorni : a.giorni };
  });
  modificata(); disegnaOfferta();
  toast(`Prezzi aggiornati su ${n} righe${mancanti ? `; ${mancanti} righe non sono più nel catalogo e restano com'erano` : ""}.`, 5000);
});

/* Copia dei dati dell'articolo dentro la riga: l'offerta resta uguale anche se il catalogo cambia. */
function rigaDaArticolo(a) {
  const { _chunk, _qta, ...r } = a;
  r.descrizione = (r.descrizione || "").slice(0, 400);
  r.listino = prezzoListino(a);
  return r;
}

/* --- Aggiungi articoli --- */
function disegnaPicker(reset = false) {
  const capSel = $("#a-capitolo");
  const caps = ordineCap();
  if (!S.pickCap || !caps.includes(S.pickCap)) S.pickCap = caps[0] || "";
  capSel.innerHTML = caps.map((c) => `<option ${c === S.pickCap ? "selected" : ""}>${esc(c)}</option>`).join("");
  const nelCap = S.catalogo.filter((a) => a.capitolo === S.pickCap);
  const sotti = [...new Set(nelCap.map((a) => a.sottocapitolo).filter(Boolean))];
  const sottoSel = $("#a-sotto");
  const prev = reset ? "" : sottoSel.value;
  sottoSel.innerHTML = `<option value="">Tutti i sottocapitoli</option>` + sotti.map((s) => `<option ${s === prev ? "selected" : ""}>${esc(s)}</option>`).join("");
  const q = $("#a-cerca").value.trim().toLowerCase();
  const solo = $("#a-solo").checked;
  const inOff = new Map(S.off.righe.map((r) => [r.k, r]));
  const lista = nelCap.filter((a) => (!sottoSel.value || a.sottocapitolo === sottoSel.value) &&
    (!solo || inOff.has(a.k)) &&
    (!q || `${codiceOracle(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)));
  const vis = lista.slice(0, S.limitePick);
  let sotto = null, html = "";
  for (const a of vis) {
    if ((a.sottocapitolo || "") !== sotto && !sottoSel.value) {
      sotto = a.sottocapitolo || "";
      html += `<tr class="sotto"><td colspan="9">${esc(sotto || "—")}</td></tr>`;
    }
    const r = inOff.get(a.k);
    const att = isAttivita(a);
    html += `<tr data-k="${esc(a.k)}">
      <td class="mono nowrap">${esc(codiceOracle(a))}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b><div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td><td class="num">${euro(prezzoListino(a))}</td>
      <td class="mono">${esc(a.gmv)}</td>
      <td><span class="dot ${freschezza(a.dataAgg)}" title="${esc(dataIt(a.dataAgg) || "data non indicata")}"></span></td>
      <td class="num"><input class="qta" data-pick="qta" inputmode="decimal" value="${r ? esc(numIt(r.qta)) : ""}" placeholder="0"></td>
      <td class="num">${att ? `<input class="qta" data-pick="giorni" inputmode="decimal" value="${esc(numIt(r ? r.giorni : a.giorni))}">` : ""}</td></tr>`;
  }
  $("#righe-picker").innerHTML = html || `<tr><td colspan="9" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Il catalogo è vuoto: importalo dal file Excel."}</td></tr>`;
  $("#picker-info").textContent = lista.length > vis.length ? `Mostrati ${vis.length} di ${lista.length}: affina la ricerca.` :
    `${lista.length} articoli. Scrivi la quantità e premi Invio o Tab: la riga entra subito nell'offerta.`;
}
$("#a-capitolo").addEventListener("change", (e) => { S.pickCap = e.target.value; disegnaPicker(true); });
$("#a-sotto").addEventListener("change", () => disegnaPicker());
$("#a-cerca").addEventListener("input", () => disegnaPicker());
$("#a-solo").addEventListener("change", () => disegnaPicker());
$("#righe-picker").addEventListener("change", (e) => {
  const inp = e.target.closest("input[data-pick]"); if (!inp) return;
  const k = inp.closest("tr").dataset.k;
  const a = S.perK.get(k); if (!a) return;
  let r = S.off.righe.find((x) => x.k === k);
  const v = num(inp.value);
  if (inp.dataset.pick === "qta") {
    if (!v) S.off.righe = S.off.righe.filter((x) => x.k !== k);
    else if (r) r.qta = v;
    else S.off.righe.push({ ...rigaDaArticolo(a), qta: v });
  } else {
    if (r) r.giorni = v;
    else { // giorni scritti prima della quantità
      S.off.righe.push({ ...rigaDaArticolo(a), giorni: v, qta: 1 });
      inp.closest("tr").querySelector('[data-pick="qta"]').value = "1";
    }
  }
  modificata();
  ridisegnaOfferta(); // aggiorna totali e conteggi; il picker si ridisegna mantenendo filtro e cursore
});

/* --- Sconti per GMV e cartellino --- */
function disegnaGmv(calc, cart) {
  const o = S.off;
  const perGmv = new Map();
  for (const x of calc.righe) {
    const g = x.riga.gmv || "";
    if (!perGmv.has(g)) perGmv.set(g, { n: 0, prima: x });
    perGmv.get(g).n++;
  }
  const esiti = new Map(cart.map((c) => [c.gmv, c]));
  const html = [...perGmv.entries()].map(([g, info]) => {
    if (!g) return `<tr><td class="muted">(senza GMV)</td><td></td><td class="num">${info.n}</td><td colspan="8" class="muted small">Le righe senza GMV usano gli sconti dell'offerta o del catalogo.</td></tr>`;
    const v = o.gmv[g] || {};
    const e = esiti.get(g);
    const inp = (k) => `<span class="pct-wrap"><input class="pct" data-gmv="${esc(g)}" data-k="${k}" inputmode="decimal" value="${esc(pctVal(v[k]))}"
      ${k === "ric" && o.tipo !== "VENDITA CLIENTE" ? "disabled" : ""}></span>`;
    const att = e?.tipo === "attivita";
    const badge = !e ? '<span class="muted small">—</span>' :
      `<span class="badge ${e.stato === "ok" ? "ok" : e.stato === "pari" ? "warn" : "bad"}">${esc(e.testo)}</span>`;
    return `<tr><td class="mono"><b>${esc(g)}</b></td><td>${esc(fornitore(g))}</td><td class="num">${info.n}</td>
      <td>${inp("sc3")}</td><td>${inp("sc4")}</td><td>${inp("sc5")}</td><td>${inp("aum1")}</td><td>${inp("ric")}</td>
      <td class="num">${att ? euro(e.offerta) : pctTxt(1 - info.prima.c.fattore) || "0%"}</td>
      <td class="num">${e ? (att ? euro(e.cartellino) : pctTxt(e.cartellino) || "0%") : ""}</td><td>${badge}</td></tr>`;
  }).join("");
  $("#righe-gmv").innerHTML = html || `<tr><td colspan="11" class="vuoto">Aggiungi articoli all'offerta per vedere i loro GMV.</td></tr>`;
}
$("#righe-gmv").addEventListener("change", (e) => {
  const inp = e.target.closest("input[data-gmv]"); if (!inp) return;
  const g = inp.dataset.gmv, k = inp.dataset.k;
  const v = daPct(inp.value);
  S.off.gmv[g] ||= {};
  if (v) S.off.gmv[g][k] = v; else delete S.off.gmv[g][k];
  if (!Object.keys(S.off.gmv[g]).length) delete S.off.gmv[g];
  modificata(); ridisegnaOfferta();
});

/* --- Esportazione per Oracle --- */
function datiOracle() { return righeOracle(S.off, opzCalc()); }
function disegnaExport() {
  const { righe, avvisi } = datiOracle();
  const av = [];
  if (!righe.length) av.push(`<div class="avviso">Nessuna riga da esportare.</div>`);
  if (avvisi.length) av.push(`<div class="avviso bad"><b>${avvisi.length} righe escluse:</b><ul class="riep">${avvisi.map((a) => `<li>${esc(a)}</li>`).join("")}</ul></div>`);
  const fuori = controlloCartellino(S.off, clienteOff(), opzCalc()).filter((c) => c.stato === "fuori");
  if (fuori.length) av.push(`<div class="avviso">Attenzione: ${fuori.map((f) => esc(f.gmv)).join(", ")} fuori cartellino del cliente (vedi “Sconti per GMV e cartellino”).</div>`);
  if (!S.off.nome) av.push(`<div class="avviso info">Suggerimento: scrivi il riferimento dell'offerta, viene usato nel nome del file.</div>`);
  $("#export-avvisi").innerHTML = av.join("");
  $("#tab-export").innerHTML = `<thead><tr>${COLONNE_ORACLE.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead>
    <tbody>${righe.map((r) => `<tr>${r.map((v, i) => `<td class="${typeof v === "number" ? "num" : i === 1 ? "mono nowrap" : ""}">${esc(typeof v === "number" ? numIt(v) : v)}</td>`).join("")}</tr>`).join("")}</tbody>`;
}
const nomeFile = (est) => `Import_Righe_${(S.off.nome || "offerta").replace(/[^\w.-]+/g, "_")}.${est}`;
$("#scarica-xlsx").addEventListener("click", () => {
  const { righe } = datiOracle();
  scarica(scriviExcel([COLONNE_ORACLE, ...righe], "Sheet1"), nomeFile("xlsx"));
});
const celleTesto = (sep) => (v) => {
  const s = typeof v === "number" ? String(v).replace(".", ",") : String(v ?? "");
  return sep === ";" && /[;"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s.replace(/[\t\n\r]+/g, " ");
};
$("#scarica-csv").addEventListener("click", () => {
  const { righe } = datiOracle();
  const f = celleTesto(";");
  const csv = [COLONNE_ORACLE, ...righe].map((r) => r.map(f).join(";")).join("\r\n");
  scarica(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }), nomeFile("csv"));
});
$("#copia-righe").addEventListener("click", async () => {
  const { righe } = datiOracle();
  const f = celleTesto("\t");
  const tsv = [COLONNE_ORACLE, ...righe].map((r) => r.map(f).join("\t")).join("\n");
  try { await navigator.clipboard.writeText(tsv); toast(`Copiate ${righe.length} righe con le intestazioni`); }
  catch { toast("Il browser non permette di copiare: usa “Scarica Excel”."); }
});

/* ================= Catalogo ================= */
$("#c-capitolo").addEventListener("change", () => { S.limiteCat = 300; disegnaCatalogo(); });
$("#c-cerca").addEventListener("input", () => { S.limiteCat = 300; disegnaCatalogo(); });
$("#c-altri").addEventListener("click", () => { S.limiteCat += 300; disegnaCatalogo(); });

function disegnaCatalogo() {
  const sel = $("#c-capitolo");
  const prev = sel.value;
  sel.innerHTML = `<option value="">Tutti i capitoli</option>` +
    S.capitoli.map((c) => `<option ${c.nome === prev ? "selected" : ""}>${esc(c.nome)}</option>`).join("");
  const cap = sel.value;
  const q = $("#c-cerca").value.trim().toLowerCase();
  const lista = S.catalogo.filter((a) => (!cap || a.capitolo === cap) &&
    (!q || `${codiceOracle(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)));
  const ord = ordinaRighe(lista, ordineCap());
  const vis = ord.slice(0, S.limiteCat);
  let gruppo = null, html = "";
  for (const a of vis) {
    const g = (cap ? "" : a.capitolo + " · ") + (a.sottocapitolo || "—");
    if (g !== gruppo) { gruppo = g; html += `<tr class="sotto"><td colspan="12">${esc(g)}</td></tr>`; }
    html += `<tr class="cliccabile" data-k="${esc(a.k)}">
      <td class="mono nowrap">${esc(codiceOracle(a))}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b><div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td><td class="num">${euro(prezzoListino(a))}</td>
      <td class="num">${pctTxt(a.sc3)}</td><td class="num">${pctTxt(a.sc4)}</td><td class="num">${pctTxt(a.sc5)}</td>
      <td class="num">${pctTxt(a.aum1)}</td><td class="num">${pctTxt(a.riparam)}</td><td class="mono">${esc(a.gmv)}</td>
      <td class="nowrap"><span class="dot ${freschezza(a.dataAgg)}"></span>${esc(dataIt(a.dataAgg))}</td></tr>`;
  }
  $("#righe-catalogo").innerHTML = html || `<tr><td colspan="12" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Catalogo vuoto: importalo da “Importa e aggiorna”."}</td></tr>`;
  $("#c-info").textContent = `${vis.length} di ${lista.length} articoli`;
  $("#c-altri").hidden = vis.length >= lista.length;
}
$("#righe-catalogo").addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-k]");
  if (tr) modificaArticolo(S.perK.get(tr.dataset.k));
});
$("#c-nuovo").addEventListener("click", () => {
  if (!S.capitoli.length) return toast("Prima crea un capitolo (Impostazioni capitolo) o importa il file Excel.");
  modificaArticolo(null);
});

/* Salva tutti gli articoli di un capitolo, divisi in blocchi sotto il limite di 1 MB di Firestore. */
function opsCapitolo(nome, articoli) {
  const puliti = articoli.slice().sort((a, b) => num(a.ordine) - num(b.ordine))
    .map(({ _chunk, _qta, ...a }) => a);
  const blocchi = [];
  let cur = [], dim = 0;
  for (const a of puliti) {
    const d = JSON.stringify(a).length * 1.2 + 50; // margine per i caratteri accentati
    if (cur.length && dim + d > 700000) { blocchi.push(cur); cur = []; dim = 0; }
    cur.push(a); dim += d;
  }
  if (cur.length) blocchi.push(cur);
  const base = slug(nome);
  const ops = blocchi.map((b, i) => ({ tipo: "set", coll: COLL.cat, id: `${base}__${i}`, data: { capitolo: nome, parte: i, articoli: b, aggiornatoIl: Date.now() } }));
  const nuovi = new Set(ops.map((o) => o.id));
  for (const c of S.chunks) if (c.capitolo === nome && !nuovi.has(c.id)) ops.push({ tipo: "del", coll: COLL.cat, id: c.id });
  return ops;
}
const salvaCapitolo = (nome) => db.scrivi(opsCapitolo(nome, S.catalogo.filter((a) => a.capitolo === nome)));

function campoNum(id, label, valore, pct = false) {
  return `<label class="f">${label}${pct ? `<span class="pct-wrap"><input class="pct" style="width:100%" id="${id}" inputmode="decimal" value="${esc(pctVal(valore))}"></span>`
    : `<input id="${id}" inputmode="decimal" value="${esc(valore ? String(valore).replace(".", ",") : "")}">`}</label>`;
}
function modificaArticolo(a) {
  const nuovo = !a;
  const capDef = $("#c-capitolo").value || S.capitoli[0]?.nome || "";
  a = a || { capitolo: capDef, precodice: "", codice: "", nOrdine: "", descrizione: "", um: "n.", listinoAcquisto: 0, sc3: 0, sc4: 0, sc5: 0, aum1: 0, riparam: 0, giorni: 0, gmv: "", dataAgg: oggiIso() };
  const auto = !num(a.listino) || Math.abs(num(a.listino) - listinoCalcolato(a)) < 0.005;
  const sotti = [...new Set(S.catalogo.filter((x) => x.capitolo === a.capitolo).map((x) => x.sottocapitolo).filter(Boolean))];
  dialogo({
    titolo: nuovo ? "Nuovo articolo" : `${codiceOracle(a)} · ${a.nOrdine}`,
    html: `<div class="grid-set">
      <label class="f">Capitolo<select id="m-cap">${S.capitoli.map((c) => `<option ${c.nome === a.capitolo ? "selected" : ""}>${esc(c.nome)}</option>`).join("")}</select></label>
      <label class="f wide">Sottocapitolo<input id="m-sotto" list="m-sotti" value="${esc(a.sottocapitolo || "")}"><datalist id="m-sotti">${sotti.map((s) => `<option value="${esc(s)}">`).join("")}</datalist></label>
      <label class="f">Precodice<input id="m-pre" value="${esc(a.precodice)}"></label>
      <label class="f">Codice<input id="m-cod" value="${esc(a.codice)}"></label>
      <label class="f wide">N° di ordinazione<input id="m-ord" value="${esc(a.nOrdine)}"></label>
      <label class="f">U.M.<input id="m-um" value="${esc(a.um)}"></label>
      <label class="f">GMV<input id="m-gmv" list="m-gmvs" value="${esc(a.gmv)}"><datalist id="m-gmvs">${S.gmv.map((g) => `<option value="${esc(g.codice)}">${esc(g.fornitore)}</option>`).join("")}</datalist></label>
      ${campoNum("m-la", "Listino acquisto €", a.listinoAcquisto)}
      ${campoNum("m-gg", "Giorni (attività)", a.giorni)}
      ${campoNum("m-sc3", "Sconto 3", a.sc3, true)}${campoNum("m-sc4", "Sconto 4", a.sc4, true)}
      ${campoNum("m-sc5", "Sconto 5", a.sc5, true)}${campoNum("m-aum1", "Aumento 1", a.aum1, true)}
      ${campoNum("m-rp", "Riparametrizzazione", a.riparam, true)}
      ${campoNum("m-li", "Prezzo di listino €", a.listino)}
      <label class="f">Data aggiornamento<input id="m-data" type="date" value="${esc(a.dataAgg || "")}"></label>
    </div>
    <label class="row small" style="margin-top:8px"><input type="checkbox" id="m-auto" ${auto ? "checked" : ""}>
      Prezzo di listino calcolato automaticamente (come la colonna K dell'Excel)</label>
    <label class="f" style="margin-top:10px">Descrizione<textarea id="m-des" rows="4">${esc(a.descrizione)}</textarea></label>
    <div class="grid-set" style="margin-top:10px">
      <label class="f wide">Note pre-linea<textarea id="m-npre" rows="2">${esc(a.notePre)}</textarea></label>
      <label class="f wide">Note post-linea<textarea id="m-npost" rows="2">${esc(a.notePost)}</textarea></label>
    </div>`,
    pulsanti: [...(nuovo ? [] : [{ label: "Elimina articolo", value: "del", cls: "danger" }]),
      { label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(b) {
      const sync = () => {
        const li = $("#m-li", b);
        li.disabled = $("#m-auto", b).checked;
        if (li.disabled) li.value = String(Math.round(listinoCalcolato(leggi(b)) * 10000) / 10000).replace(".", ",");
      };
      b.addEventListener("input", sync);
      sync();
    },
    async onConferma(v, b) {
      const vecchioCap = a.capitolo;
      if (v === "del") {
        S.catalogo = S.catalogo.filter((x) => x.k !== a.k);
        await salvaCapitolo(vecchioCap);
        toast("Articolo eliminato. Le offerte che lo contengono non cambiano.");
        return;
      }
      const n = leggi(b);
      if (!n.codice && !n.nOrdine) { toast("Scrivi almeno il codice o il n° di ordinazione."); return false; }
      if (!n.listino || $("#m-auto", b).checked) n.listino = listinoCalcolato(n);
      if (nuovo) {
        const stessi = S.catalogo.filter((x) => x.capitolo === n.capitolo && (!n.sottocapitolo || x.sottocapitolo === n.sottocapitolo));
        n.ordine = (stessi.length ? Math.max(...stessi.map((x) => num(x.ordine))) : 0) + 0.01;
        n.rifCap = stessi[0]?.rifCap || "";
        n.k = `${n.capitolo}|${n.rifCap}|${n.precodice}${n.codice}|${n.nOrdine}|${db.nuovoId()}`;
        S.catalogo.push(n);
      } else {
        Object.assign(S.perK.get(a.k), n);
      }
      await salvaCapitolo(n.capitolo);
      if (vecchioCap !== n.capitolo) await salvaCapitolo(vecchioCap);
      toast(nuovo ? "Articolo aggiunto" : "Articolo salvato");
    },
  });
  function leggi(b) {
    const g = (id) => $(id, b).value;
    return {
      ...a,
      capitolo: g("#m-cap"), sottocapitolo: g("#m-sotto").trim(), precodice: g("#m-pre").trim(), codice: g("#m-cod").trim(),
      nOrdine: g("#m-ord").trim(), um: g("#m-um").trim(), gmv: g("#m-gmv").trim(),
      listinoAcquisto: num(g("#m-la")), giorni: num(g("#m-gg")),
      sc3: daPct(g("#m-sc3")), sc4: daPct(g("#m-sc4")), sc5: daPct(g("#m-sc5")), aum1: daPct(g("#m-aum1")), riparam: daPct(g("#m-rp")),
      listino: num(g("#m-li")), dataAgg: g("#m-data"), descrizione: g("#m-des"), notePre: g("#m-npre"), notePost: g("#m-npost"),
    };
  }
}

$("#c-impost").addEventListener("click", () => {
  dialogo({
    titolo: "Capitoli e ricarico automatico",
    html: `<p class="muted small" style="margin-top:0">Il ricarico automatico si applica alle offerte “VENDITA CLIENTE” quando non è indicato un ricarico nell'offerta o nel GMV.</p>
      <table><thead><tr><th>Capitolo</th><th class="num">Articoli</th><th>Ricarico automatico</th></tr></thead><tbody>
      ${S.capitoli.map((c, i) => `<tr><td>${esc(c.nome)}</td><td class="num">${S.catalogo.filter((a) => a.capitolo === c.nome).length}</td>
        <td><span class="pct-wrap"><input class="pct" data-i="${i}" value="${esc(pctVal(c.ricarico))}"></span></td></tr>`).join("")}
      </tbody></table>
      <label class="f" style="margin-top:12px">Nuovo capitolo (nome)<input id="m-nuovocap" placeholder="es. EVAC"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    async onConferma(v, b) {
      const lista = S.capitoli.map((c, i) => ({ ...c, ricarico: daPct($(`[data-i="${i}"]`, b).value) }));
      const nuovo = $("#m-nuovocap", b).value.trim().toUpperCase().replace(/\s+/g, "_");
      if (nuovo && !lista.some((c) => c.nome === nuovo)) lista.push({ nome: nuovo, ricarico: 0.28, ordine: lista.length });
      await db.salva(COLL.cfg, "capitoli", { lista });
      toast("Capitoli salvati");
    },
  });
});

/* ================= Clienti ================= */
$("#cl-cerca").addEventListener("input", disegnaClienti);
function disegnaClienti() {
  const q = $("#cl-cerca").value.trim().toLowerCase();
  const lista = S.clienti.filter((c) => !q || (c.nome || "").toLowerCase().includes(q) || (c.codice || "").includes(q));
  $("#righe-clienti").innerHTML = lista.map((c) => `<tr class="cliccabile" data-id="${esc(c.id)}">
    <td><b>${esc(c.nome)}</b></td><td class="mono">${esc(c.codice)}</td>
    <td class="num">${c.att === "" || c.att === undefined ? '<span class="muted">—</span>' : euro(c.att)}</td>
    <td class="small">${Object.entries(c.sconti || {}).map(([g, v]) => `<span class="badge info">${esc(g)} ${pctTxt(v)}</span>`).join(" ") || '<span class="muted">nessuno</span>'}</td></tr>`).join("")
    || `<tr><td colspan="4" class="vuoto">Nessun cliente. Importa il file Excel o aggiungine uno.</td></tr>`;
}
$("#righe-clienti").addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-id]");
  if (tr) modificaCliente(S.clienti.find((c) => c.id === tr.dataset.id));
});
$("#cl-nuovo").addEventListener("click", () => modificaCliente(null));
function modificaCliente(c) {
  const nuovo = !c;
  c = c || { nome: "", codice: "", att: "", sconti: {} };
  const gmvs = [...new Set([...S.gmv.map((g) => g.codice), ...Object.keys(c.sconti || {})])].filter((g) => g !== "ATT");
  dialogo({
    titolo: nuovo ? "Nuovo cliente" : c.nome,
    html: `<div class="grid-set">
      <label class="f wide">Ragione sociale<input id="k-nome" value="${esc(c.nome)}" required></label>
      <label class="f">Codice cliente<input id="k-cod" value="${esc(c.codice)}"></label>
      <label class="f">Tariffa attività tecnica €<input id="k-att" inputmode="decimal" value="${esc(c.att === "" ? "" : String(c.att).replace(".", ","))}"></label>
    </div>
    <h3 style="margin:14px 0 8px">Sconti di cartellino per GMV</h3>
    <div class="lista-gmv">${gmvs.map((g) => `<label><span><b class="mono">${esc(g)}</b> <span class="muted">${esc(fornitore(g))}</span></span>
      <span class="pct-wrap"><input class="pct" data-g="${esc(g)}" value="${esc(pctVal(c.sconti?.[g]))}"></span></label>`).join("")}</div>`,
    pulsanti: [...(nuovo ? [] : [{ label: "Elimina cliente", value: "del", cls: "danger" }]),
      { label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    async onConferma(v, b) {
      if (v === "del") { await db.elimina(COLL.cli, c.id); toast("Cliente eliminato"); return; }
      const nome = $("#k-nome", b).value.trim();
      if (!nome) { toast("Scrivi la ragione sociale."); return false; }
      const sconti = {};
      $$("[data-g]", b).forEach((i) => { const x = daPct(i.value); if (x) sconti[i.dataset.g] = x; });
      const attTxt = $("#k-att", b).value.trim();
      const id = c.id || slug(nome) + "-" + db.nuovoId().slice(-4);
      await db.salva(COLL.cli, id, { nome, codice: $("#k-cod", b).value.trim(), att: attTxt === "" ? "" : num(attTxt), sconti });
      toast("Cliente salvato");
    },
  });
}

/* ================= Importazione dal file Excel ================= */
const drop = $("#drop-xlsx");
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); const f = e.dataTransfer.files[0]; if (f) analizza(f); });
$("#file-xlsx").addEventListener("change", (e) => { const f = e.target.files[0]; if (f) analizza(f); e.target.value = ""; });

async function analizza(file) {
  const box = $("#riepilogo-import");
  box.innerHTML = `<div class="avviso info">Lettura di ${esc(file.name)}…</div>`;
  try {
    const r = await analizzaFile(await file.arrayBuffer());
    S.importato = { ...r, nomeFile: file.name };
    const giaCat = new Set(S.catalogo.map((a) => a.capitolo));
    const capQ = [...new Set(r.conQuantita.map((a) => a.capitolo))];
    box.innerHTML = `
      <div class="avviso ok">Letto <b>${esc(file.name)}</b>: ${r.articoli.length} articoli in ${r.capitoli.length} capitoli, ${r.gmv.length} GMV, ${r.clienti.length} clienti.</div>
      <div class="tbl-wrap" style="max-height:300px"><table><thead><tr><th>Capitolo</th><th class="num">Articoli</th><th class="num">Ricarico</th><th class="num">Con quantità</th><th></th></tr></thead><tbody>
      ${r.capitoli.map((c) => `<tr><td>${esc(c.nome)}</td><td class="num">${c.articoli}</td><td class="num">${pctTxt(c.ricarico)}${c.ricaricoDaFile ? "" : ' <span class="badge warn" title="Il capitolo non è nella tabella Capitoli dell\'Excel">predefinito</span>'}</td>
        <td class="num">${r.conQuantita.filter((a) => a.capitolo === c.nome).length || ""}</td>
        <td>${giaCat.has(c.nome) ? '<span class="badge warn">sostituisce l\'attuale</span>' : '<span class="badge info">nuovo</span>'}</td></tr>`).join("")}
      </tbody></table></div>
      <div class="stack" style="margin-top:10px">
        <label class="row"><input type="checkbox" id="i-cat" checked> Catalogo articoli (sostituisce i capitoli elencati sopra; le offerte già fatte non cambiano)</label>
        <label class="row"><input type="checkbox" id="i-cap" checked> Capitoli e ricarichi automatici, elenco GMV</label>
        <label class="row"><input type="checkbox" id="i-cli" ${r.clienti.length ? "checked" : "disabled"}> Clienti e cartellini (aggiorna quelli con lo stesso nome)</label>
        <label class="row"><input type="checkbox" id="i-off" ${capQ.length ? "" : "disabled"}> Crea un'offerta per ogni capitolo con le quantità già scritte nel file
          ${capQ.length ? `<span class="muted small">(${capQ.map(esc).join(", ")})</span>` : ""}</label>
        <div class="row"><button class="btn primary" id="i-vai">Importa</button></div>
      </div>`;
    $("#i-vai").addEventListener("click", eseguiImport);
  } catch (x) {
    console.error(x);
    box.innerHTML = `<div class="avviso bad">Non riesco a leggere il file: ${esc(x.message || x)}</div>`;
  }
}

async function eseguiImport() {
  const r = S.importato; if (!r) return;
  const btn = $("#i-vai");
  btn.disabled = true; btn.textContent = "Importazione…";
  try {
    const ops = [];
    if ($("#i-cap").checked) {
      const lista = [...S.capitoli.filter((c) => !r.capitoli.some((x) => x.nome === c.nome))
        .map((c) => ({ nome: c.nome, ricarico: c.ricarico })), ...r.capitoli.map((c) => ({ nome: c.nome, ricarico: c.ricarico }))];
      // ordine: quello del file, poi i capitoli che esistevano solo nell'app
      const ord = [...r.capitoli.map((c) => c.nome), ...lista.map((c) => c.nome)];
      lista.forEach((c) => (c.ordine = ord.indexOf(c.nome)));
      ops.push({ tipo: "set", coll: COLL.cfg, id: "capitoli", data: { lista } });
      const gm = new Map(S.gmv.map((g) => [g.codice, g]));
      for (const g of r.gmv) gm.set(g.codice, { codice: g.codice, fornitore: g.fornitore || gm.get(g.codice)?.fornitore || "" });
      ops.push({ tipo: "set", coll: COLL.cfg, id: "gmv", data: { lista: [...gm.values()] } });
    } else if ($("#i-cat").checked) {
      const mancanti = r.capitoli.filter((c) => !S.capitoli.some((x) => x.nome === c.nome));
      if (mancanti.length) ops.push({ tipo: "set", coll: COLL.cfg, id: "capitoli",
        data: { lista: [...S.capitoli, ...mancanti.map((c, i) => ({ nome: c.nome, ricarico: c.ricarico, ordine: S.capitoli.length + i }))] } });
    }
    if ($("#i-cat").checked) for (const c of r.capitoli) ops.push(...opsCapitolo(c.nome, r.articoli.filter((a) => a.capitolo === c.nome)));
    if ($("#i-cli").checked) {
      for (const c of r.clienti) {
        const esiste = S.clienti.find((x) => (x.nome || "").trim().toLowerCase() === c.nome.toLowerCase());
        const { id, ...dati } = c;
        ops.push({ tipo: "set", coll: COLL.cli, id: esiste?.id || id, data: dati });
      }
    }
    let nOff = 0;
    if ($("#i-off").checked) {
      const caps = [...new Set(r.conQuantita.map((a) => a.capitolo))];
      for (const cap of caps) {
        const o = nuovaOffertaVuota();
        o.nome = `${cap} (da ${r.nomeFile})`;
        o.righe = r.conQuantita.filter((a) => a.capitolo === cap).map((a) => ({ ...rigaDaArticolo(a), qta: a._qta }));
        ops.push({ tipo: "set", coll: COLL.off, id: o.id, data: o });
        nOff++;
      }
    }
    await db.scrivi(ops);
    $("#riepilogo-import").innerHTML = `<div class="avviso ok">Importazione completata${nOff ? `: create ${nOff} offerte con le quantità del file` : ""}.</div>`;
    S.importato = null;
    toast("Importazione completata");
  } catch (x) {
    errore(x, "Importazione non riuscita");
    btn.disabled = false; btn.textContent = "Importa";
  }
}

/* ================= Aggiornamento prezzi da Oracle ================= */
$("#analizza-oracle").addEventListener("click", () => {
  const t = $("#incolla-oracle").value;
  if (!t.trim()) return toast("Incolla prima le righe copiate da Oracle.");
  confrontaOracle(righeDaTestoIncollato(t));
});
$("#file-oracle").addEventListener("change", async (e) => {
  const f = e.target.files[0]; e.target.value = "";
  if (!f) return;
  try {
    const wb = await leggiExcel(await f.arrayBuffer());
    const nome = wb.nomi.find((n) => /AGGIORNAMENTO LISTINO$/i.test(n)) || wb.nomi[0];
    confrontaOracle(await wb.foglio(nome));
  } catch (x) { errore(x, "File non leggibile"); }
});

const CODICI_GENERICI = new Set(["P000ARTICOLO_GEN", "PITATT.TEC"]);
function confrontaOracle(righe) {
  const box = $("#riepilogo-oracle");
  let mappa;
  try { mappa = leggiAggiornamentoListino(righe); } catch (x) { box.innerHTML = `<div class="avviso bad">${esc(x.message)}</div>`; return; }
  const cambi = [];
  let uguali = 0;
  const trovati = new Set();
  for (const a of S.catalogo) {
    const cod = codiceOracle(a);
    if (CODICI_GENERICI.has(cod)) continue;
    const m = mappa.get(cod);
    if (!m) continue;
    trovati.add(cod);
    const cambiaPrezzo = Math.abs(m.prezzo - num(a.listinoAcquisto)) > 0.004;
    const cambiaGmv = m.gmv && m.gmv !== a.gmv;
    if (cambiaPrezzo || cambiaGmv) cambi.push({ a, m }); else uguali++;
  }
  const nonTrovati = [...mappa.keys()].filter((c) => !trovati.has(c));
  S.aggOracle = cambi;
  box.innerHTML = `
    <div class="avviso ${cambi.length ? "info" : "ok"}">${mappa.size} codici letti: <b>${cambi.length}</b> articoli da aggiornare, ${uguali} già allineati${nonTrovati.length ? `, ${nonTrovati.length} codici non presenti nel catalogo` : ""}.
      Gli articoli generici (P000ARTICOLO_GEN, PITATT.TEC) non si aggiornano: hanno tutti lo stesso codice in Oracle.</div>
    ${cambi.length ? `<div class="tbl-wrap" style="max-height:340px"><table><thead><tr><th>Capitolo</th><th>Codice</th><th>Articolo</th><th class="num">Listino acq. attuale</th><th class="num">Oracle</th><th class="num">Diff.</th><th>GMV</th></tr></thead><tbody>
      ${cambi.map(({ a, m }) => { const d = num(a.listinoAcquisto) ? (m.prezzo / num(a.listinoAcquisto) - 1) : 0;
        return `<tr><td>${esc(a.capitolo)}</td><td class="mono">${esc(codiceOracle(a))}</td><td>${esc(a.nOrdine)}</td>
        <td class="num">${euro(a.listinoAcquisto)}</td><td class="num"><b>${euro(m.prezzo)}</b></td>
        <td class="num"><span class="badge ${d > 0 ? "bad" : "ok"}">${d > 0 ? "+" : ""}${fNum.format(Math.round(d * 1000) / 10)}%</span></td>
        <td class="mono">${m.gmv && m.gmv !== a.gmv ? `${esc(a.gmv || "—")} → <b>${esc(m.gmv)}</b>` : esc(a.gmv)}</td></tr>`; }).join("")}
      </tbody></table></div>
      <div class="row" style="margin-top:10px"><button class="btn primary" id="applica-oracle">Aggiorna ${cambi.length} articoli</button>
        <span class="muted small">La data di aggiornamento diventa quella di oggi.</span></div>` : ""}
    ${nonTrovati.length ? `<details style="margin-top:8px"><summary class="muted small">Codici non presenti nel catalogo (${nonTrovati.length})</summary><p class="mono small">${nonTrovati.map(esc).join(", ")}</p></details>` : ""}`;
  $("#applica-oracle")?.addEventListener("click", applicaOracle);
}
async function applicaOracle() {
  const cambi = S.aggOracle || [];
  const caps = new Set();
  const oggi = oggiIso();
  for (const { a, m } of cambi) {
    const art = S.perK.get(a.k); if (!art) continue;
    const auto = !num(art.listino) || Math.abs(num(art.listino) - listinoCalcolato(art)) < 0.005;
    art.listinoAcquisto = m.prezzo;
    if (m.gmv) art.gmv = m.gmv;
    if (auto) art.listino = listinoCalcolato(art);
    art.dataAgg = oggi;
    caps.add(art.capitolo);
  }
  try {
    await db.scrivi([...caps].flatMap((c) => opsCapitolo(c, S.catalogo.filter((a) => a.capitolo === c))));
    $("#riepilogo-oracle").innerHTML = `<div class="avviso ok">Aggiornati ${cambi.length} articoli in ${caps.size} capitoli.</div>`;
    $("#incolla-oracle").value = "";
    S.aggOracle = null;
  } catch (x) { errore(x, "Aggiornamento non riuscito"); }
}

/* ================= Copia di sicurezza ================= */
$("#scarica-backup").addEventListener("click", () => {
  const dati = {
    app: "offerte-oracle", versione: 1, creato: new Date().toISOString(),
    capitoli: S.capitoli, gmv: S.gmv, clienti: S.clienti,
    catalogo: S.catalogo.map(({ _chunk, ...a }) => a), offerte: S.offerte,
  };
  scarica(new Blob([JSON.stringify(dati)], { type: "application/json" }), `offerte-oracle-backup-${oggiIso()}.json`);
});
$("#file-backup").addEventListener("change", async (e) => {
  const f = e.target.files[0]; e.target.value = "";
  if (!f) return;
  let d;
  try { d = JSON.parse(await f.text()); if (d.app !== "offerte-oracle") throw new Error("non è una copia di questa app"); }
  catch (x) { return toast("File non valido: " + x.message); }
  dialogo({
    titolo: "Ripristinare la copia?",
    html: `<p>Copia del ${esc(new Date(d.creato).toLocaleString("it-IT"))}: ${d.catalogo.length} articoli, ${d.clienti.length} clienti, ${d.offerte.length} offerte.</p>
      <p><b>Catalogo, clienti e offerte attuali verranno sostituiti.</b></p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Ripristina", value: "ok", cls: "danger" }],
    async onConferma() {
      const ops = [
        { tipo: "set", coll: COLL.cfg, id: "capitoli", data: { lista: d.capitoli } },
        { tipo: "set", coll: COLL.cfg, id: "gmv", data: { lista: d.gmv } },
      ];
      for (const c of S.chunks) ops.push({ tipo: "del", coll: COLL.cat, id: c.id });
      const vecchi = S.chunks; S.chunks = [];
      for (const cap of new Set(d.catalogo.map((a) => a.capitolo))) ops.push(...opsCapitolo(cap, d.catalogo.filter((a) => a.capitolo === cap)));
      S.chunks = vecchi;
      const nuoviCat = new Set(ops.filter((o) => o.coll === COLL.cat && o.tipo === "set").map((o) => o.id));
      const finali = ops.filter((o) => !(o.coll === COLL.cat && o.tipo === "del" && nuoviCat.has(o.id)));
      for (const c of S.clienti) if (!d.clienti.some((x) => x.id === c.id)) finali.push({ tipo: "del", coll: COLL.cli, id: c.id });
      for (const c of d.clienti) { const { id, ...x } = c; finali.push({ tipo: "set", coll: COLL.cli, id, data: x }); }
      for (const o of S.offerte) if (!d.offerte.some((x) => x.id === o.id)) finali.push({ tipo: "del", coll: COLL.off, id: o.id });
      for (const o of d.offerte) { const { id, ...x } = o; finali.push({ tipo: "set", coll: COLL.off, id, data: { ...x, id } }); }
      await db.scrivi(finali);
      toast("Copia ripristinata");
    },
  });
});

// per i test automatici
window.__app = { S, calcolaOfferta, prezzoBase };
