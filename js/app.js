import * as db from "./firebase.js";
import {
  CAMPI_SCONTO, MODI_ACQ, MODI_VEN, IMPOSTAZIONI, OUT_PREZZO, OUT_FORMATO, COLONNE_ORACLE, num, isAttivita, codiceOracle, prezzoListino,
  listinoCalcolato, calcolaRiga, calcolaOfferta, righeOracle, controlloCartellino, freschezza, ordinaRighe,
  riepilogoCapitolo, ricaricoPct, marginePct, vendutaPerTarget, isGenerico, notePreRiga, notePostRiga, esclusoDaExport,
  codiceArticolo, statoPrezzo, CODICE_GENERICO, condizioni, migraCondizioni, leggiSconti, leggiRicarico, testoValido,
  fattoreStandard, pctTesto, impostaStandardGma, scontiEffettivi, standardGma, haScontiPropri, voceGma, gmaGenerico, gmaOracle,
} from "./calc.js";
import { analizzaFile, leggiAggiornamentoListino, righeDaTestoIncollato, leggiExcel, slug } from "./importa.js";
import { scriviExcel, scriviExcelStile } from "./xlsx-lite.js";
import * as BIB from "./biblioteca.js";
import * as OFF from "./offforn.js";
import { attivaTendine } from "./tendina.js";
attivaTendine();

/* ================= Stato ================= */
const COLL = { cat: "oracle_catalogo", cfg: "oracle_config", cli: "oracle_clienti", off: "oracle_offerte", bib: "oracle_biblioteca", img: "oracle_immagini" };
const S = {
  capitoli: [], gmv: [], clienti: [], chunks: [], catalogo: [], perK: new Map(), offerte: [],
  off: null, vista: "offerte", sub: "righe", dirty: false, salvando: false, utente: null,
  limiteCat: 300, limitePick: 300, pickCap: "", pickDb: "", importato: null, aggOracle: null,
  capAtt: "", sel: new Set(), precodici: [], pickSel: new Set(), pickUltimo: "", pickVis: [],
  fornitori: [], bib: [], bibChunks: [], bibPerK: new Map(), bibPronta: false, bibAscolto: null,
};
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));

/* ================= Utilità ================= */
export const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fEuro = (() => { // separatore delle migliaia anche sotto i 10.000 (2.700,00 €) come in Excel
  try { return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR", useGrouping: "always" }); }
  catch { return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }); }
})();
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
const clienteOff = () => S.clienti.find((c) => c.id === S.off?.cliente) || null;
const fornitore = (g) => S.gmv.find((x) => x.codice === g)?.fornitore || "";
// tutti i fornitori che hanno quel codice GMA (per le viste che conoscono solo il codice)
const fornitoriTxt = (g) => [...new Set(S.gmv.filter((x) => x.codice === g).map((x) => x.fornitore || "").filter(Boolean))].join(" / ");
/* Fornitore di un articolo: il suo (lo stesso GMA può esistere con più fornitori), altrimenti quello del GMA. */
const fornitoreDi = (a) => (a?.fornitore || "").trim() || fornitore(a?.gmv);
/* Regole di scadenza dei prezzi per precodice */
const regole = () => Object.fromEntries(S.precodici.filter((p) => p.tipo).map((p) => [String(p.precodice).toUpperCase(), p]));
const stato = (a) => statoPrezzo(a, regole());
const pallino = (a) => { const st = stato(a); return `<span class="dot ${st.livello}" title="${esc(st.testo)}"></span>`; };
/* Codice da mostrare: quello del database; per i generici anche quello che va in Oracle */
const codiceVis = (a) => isGenerico(a)
  ? `<span title="In Oracle: ${CODICE_GENERICO}">${esc(codiceArticolo(a))}</span> <span class="badge warn gen" title="Trattato come articolo generico: in Oracle ${CODICE_GENERICO}, prezzo netto, n° ordinazione e descrizione nelle note">GEN</span>`
  : esc(codiceArticolo(a));

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
function dialogo({ titolo, html, pulsanti = [{ label: "Chiudi", value: "annulla" }], onApri, onConferma, largo = false }) {
  $("#dlg").classList.toggle("largo", largo);
  $("#dlg-titolo").textContent = titolo;
  // corpo nuovo a ogni apertura: gli ascoltatori del dialogo precedente non restano attaccati
  const vecchio = $("#dlg-body"), corpo = vecchio.cloneNode(false);
  vecchio.replaceWith(corpo);
  corpo.innerHTML = html;
  $("#dlg-foot").innerHTML = pulsanti.map((p) =>
    `<button class="btn ${p.cls || ""}" value="${p.value}" ${p.value === "annulla" ? 'formnovalidate' : ""}>${esc(p.label)}</button>`).join("");
  dlgHandler = onConferma;
  onApri?.($("#dlg-body"));
  $("#dlg").showModal();
}
// Invio in un campo del dialogo = pulsante principale (non "Annulla")
$("#dlg-form").addEventListener("keydown", (e) => {
  if (e.key !== "Enter" || !e.target.matches("input:not([type=checkbox])")) return;
  const ok = $("#dlg-foot .primary") || $("#dlg-foot .danger"); if (!ok) return;
  e.preventDefault();
  e.target.dispatchEvent(new Event("change", { bubbles: true })); // il valore appena scritto conta
  ok.click();
});
$("#dlg-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const v = e.submitter?.value || "annulla";
  if (v === "annulla" || !dlgHandler) return $("#dlg").close();
  try {
    const r = await dlgHandler(v, $("#dlg-body"));
    if (r !== false) $("#dlg").close();
  } catch (x) { errore(x); }
});

/* ================= Gruppi apri/chiudi nelle tabelle ================= */
/* Le righe di intestazione (data-gruppo) hanno un − / + che nasconde o mostra le righe sottostanti,
   fino all'intestazione successiva o alla riga dei totali. Lo stato resta finché la pagina è aperta. */
const chiusi = new Set();
const ambitoDi = (tb) => tb.id + (tb.id === "righe-offerta" || tb.id === "riepilogo-righe" ? "|" + (S.off?.id || "") + "|" + S.capAtt : "|" + (S.pickDb || $("#c-capitolo")?.value || ""));
function applicaGruppi(tb) {
  if (!tb) return;
  const amb = ambitoDi(tb);
  let chiuso = false, testa = null, n = 0;
  const chiudiConta = () => { if (testa) testa.querySelector(".tg-n").textContent = n === 1 ? "1 riga" : `${n} righe`; };
  for (const tr of tb.rows) {
    if (tr.dataset.gruppo !== undefined) {
      chiudiConta();
      testa = tr; n = 0;
      chiuso = chiusi.has(amb + "|" + tr.dataset.gruppo);
      tr.classList.add("testa-gruppo");
      const td = tr.cells[0];
      if (!td.querySelector(".tg")) td.insertAdjacentHTML("afterbegin", `<button type="button" class="tg"></button>`);
      if (!td.querySelector(".tg-n")) td.insertAdjacentHTML("beforeend", `<span class="tg-n"></span>`);
      const b = td.querySelector(".tg");
      b.textContent = chiuso ? "+" : "−";
      b.title = chiuso ? "Mostra le righe" : "Nascondi le righe";
      tr.classList.toggle("chiuso", chiuso);
      continue;
    }
    if (tr.classList.contains("sotto") || tr.classList.contains("totale")) { chiudiConta(); testa = null; chiuso = false; tr.hidden = false; continue; }
    if (testa) n++;
    tr.hidden = chiuso;
  }
  chiudiConta();
}
document.addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-gruppo]");
  if (!tr || e.target.closest("input, select, a, .tg-ed")) return;
  const tb = tr.parentElement;
  const k = ambitoDi(tb) + "|" + tr.dataset.gruppo;
  chiusi.has(k) ? chiusi.delete(k) : chiusi.add(k);
  applicaGruppi(tb);
});

document.addEventListener("click", (e) => {
  const b = e.target.closest("[data-tutti]"); if (!b) return;
  const tb = document.getElementById(b.dataset.tutti); if (!tb) return;
  const amb = ambitoDi(tb);
  for (const tr of tb.querySelectorAll("tr[data-gruppo]")) {
    const k = amb + "|" + tr.dataset.gruppo;
    if (b.dataset.az === "chiudi") chiusi.add(k); else chiusi.delete(k);
  }
  applicaGruppi(tb);
});

/* ================= Avvio ================= */
let ascoltatori = [];
function schermataLogin(messaggio = "", verifica = false) {
  $("#app").hidden = true;
  $("#login").hidden = false;
  $("#google-login").hidden = verifica;
  $("#login .err").textContent = messaggio;
  $("#login .err").classList.toggle("info", verifica);
}
db.avvia({
  onVerifica: () => schermataLogin("Verifica dell'account…", true),
  onNegato: (email) => schermataLogin(`L'account ${email || ""} non è abilitato a usare questa app. Entra con un account autorizzato.`),
  onLogin(u) {
    S.utente = u;
    $("#google-login").hidden = false;
    $("#login .err").textContent = "";
    $("#login").hidden = true;
    $("#app").hidden = false;
    $("#banner-prova").hidden = !u.prova;
    setTimeout(() => altezzaBarra(), 0);
    $("#utente").textContent = u.email;
    collega();
  },
  onLogout() {
    ascoltatori.forEach((f) => f());
    ascoltatori = [];
    S.utente = null; S.off = null; S.offerte = []; S.catalogo = []; S.clienti = []; S.bib = []; S.bibAscolto = null; S.bibPronta = false;
    $("#app").hidden = true;
    $("#login").hidden = false;
    $("#google-login").hidden = false;
  },
  onErrore: (e) => schermataLogin("Accesso non riuscito: " + (e.message || e)),
}).catch((e) => {
  $("#login").hidden = false;
  $("#login .err").textContent = "Firebase non raggiungibile: " + (e.message || e);
});

$("#google-login").addEventListener("click", async () => {
  $("#login .err").textContent = "";
  try { await db.accediGoogle(); } catch (x) { $("#login .err").textContent = "Accesso non riuscito: " + (x.message || x); }
});
$("#esci").addEventListener("click", () => db.esci());
$("#tema").addEventListener("click", () => {
  const t = document.documentElement.dataset.theme === "scuro" ? "chiaro" : "scuro";
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem("tema", t); } catch {}
});

function collega() {
  const onErr = (e) => errore(e, "Errore di lettura");
  ascoltatori.push(db.ascolta(COLL.cfg, (docs) => {
    const cap = docs.find((d) => d.id === "capitoli");
    const gmv = docs.find((d) => d.id === "gmv");
    S.precodici = docs.find((d) => d.id === "precodici")?.lista || [];
    S.capitoli = (cap?.lista || []).slice().sort((a, b) => num(a.ordine) - num(b.ordine));
    S.gmv = gmv?.lista || [];
    impostaStandardGma(S.gmv);
    S.fornitori = docs.find((d) => d.id === "fornitori")?.lista || [];
    S.censCfg = docs.find((d) => d.id === "censimento") || {};
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
  if (vista !== S.vista) bloccaModificaCat();
  S.vista = vista;
  $("#sb-offerta").hidden = vista !== "offerta";
  for (const v of ["offerte", "offerta", "catalogo", "biblioteca", "clienti", "impostazioni", "importa"]) $("#v-" + v).hidden = v !== vista;
  if (vista === "biblioteca") caricaBib();
  $$("#tabs button").forEach((b) => b.classList.toggle("on", b.dataset.vista === (vista === "offerta" ? "offerte" : vista)));
  if (vista !== "offerta" && location.hash) history.replaceState(null, "", location.pathname);
  ridisegna(true);
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
function ridisegna(forza = false) {
  cancelAnimationFrame(rafRid);
  rafRid = requestAnimationFrame(() => {
    // dati arrivati mentre si scrive in un campo: si ridisegna appena il campo perde il focus
    // (non quando si cambia pagina: lì si disegna sempre)
    if (!forza && staScrivendo() && S.vista !== "offerte" && S.vista !== "importa") { rinviato = true; return; }
    if (S.vista === "offerte") disegnaElencoOfferte();
    else if (S.vista === "offerta") disegnaOfferta();
    else if (S.vista === "catalogo") disegnaCatalogo();
    else if (S.vista === "clienti") disegnaClienti();
    else if (S.vista === "biblioteca") disegnaBiblioteca();
  });
}

/* ================= Modello dell'offerta =================
   offerta = { id, nome, cliente, note, capitoli: [capitolo], creatoIl, aggiornatoIl, creatoDa }
   capitolo = { nome, acq, ven, out (vedi calc.js: condizioni), sottocapitoli, lineaPartenza, magazzino, notaNuove,
                capitolati: [nomi dei capitolati per Oracle], capNuove, righe: [riga] }
   Il capitolato (colonna "Nome Capitolato" in Oracle) si decide offerta per offerta e non c'entra con
   i sottocapitoli del database, che servono solo a organizzare gli articoli.
   riga = copia dei dati dell'articolo + { id, qta, giorni, notePre, notePost }. Le righe sono nell'ordine scelto. */
function nuovoCapitolo(nome, base = {}) {
  // i selettori si copiano dal capitolo di partenza, i valori no
  return {
    nome, sottocapitoli: !!base.sottocapitoli, lineaPartenza: base.lineaPartenza ?? 10, magazzino: base.magazzino ?? "BG",
    acq: { modo: base.acq?.modo || "standard", gen: "", gma: {} },
    ven: { modo: base.ven?.modo || "ricarico", ricGen: "", ricGma: {}, scGen: "", scGma: {} },
    out: { prezzo: base.out?.prezzo || "vendita", formato: base.out?.formato || "sconti" },
    imp: base.imp || "av", lis: { gen: "", gma: {} },
    notaNuove: "", capitolati: [], capNuove: "", righe: [],
  };
}
/* Capitoli salvati con il vecchio modello (tipo offerta, Sconto 3/4/5…): si convertono una volta. */
function aggiornaCondizioni(c) {
  if (!(c.acq && c.ven && c.out)) {
    Object.assign(c, migraCondizioni(c));
    for (const k of ["tipo", "prezzi", "globali", "gmv"]) delete c[k];
  }
  c.imp ||= "av";
  c.lis ||= { gen: "", gma: {} };
}
function nuovaOffertaVuota() {
  const ora = Date.now();
  return { id: db.nuovoId(), nome: "", cliente: "", note: "", capitoli: [], creatoIl: ora, aggiornatoIl: ora, creatoDa: S.utente?.email || "" };
}
/* Converte le offerte della prima versione (un solo elenco di righe) nel modello a capitoli. */
function normalizza(o) {
  o = structuredClone(o);
  if (!Array.isArray(o.capitoli)) {
    const base = { tipo: o.tipo, prezzi: o.prezzi, sottocapitoli: o.sottocapitoli, lineaPartenza: o.lineaPartenza, magazzino: o.magazzino, globali: o.globali, gmv: o.gmv };
    o.capitoli = [];
    for (const r of o.righe || []) {
      let c = o.capitoli.find((x) => x.nome === r.capitolo);
      if (!c) {
        c = nuovoCapitolo(r.capitolo, base);
        delete c.acq; delete c.ven; delete c.out; // si convertono dal vecchio modello qui sotto
        Object.assign(c, { tipo: base.tipo, prezzi: base.prezzi, globali: base.globali || {}, gmv: base.gmv || {} });
        o.capitoli.push(c);
      }
      c.righe.push(r);
    }
    for (const k of ["righe", "tipo", "prezzi", "sottocapitoli", "lineaPartenza", "magazzino", "globali", "gmv"]) delete o[k];
  }
  for (const c of o.capitoli) {
    aggiornaCondizioni(c);
    c.righe ||= []; c.capitolati ||= [];
    for (const r of c.righe) {
      r.id ||= db.nuovoId();
      // offerte di prima: il capitolato era il sottocapitolo della riga
      if (r.capitolato === undefined) r.capitolato = c.sottocapitoli ? (r.sottocapitolo || "") : "";
      if (r.capitolato && !c.capitolati.includes(r.capitolato)) c.capitolati.push(r.capitolato);
    }
    raggruppa(c);
  }
  return o;
}
const capitoliOff = (o) => (Array.isArray(o.capitoli) ? o.capitoli : normalizza(o).capitoli);
/* Il ricarico automatico è quello del capitolo dell'offerta, qualunque sia il database da cui arriva la riga. */
function opzCap(cap) {
  const ric = num(S.capitoli.find((c) => c.nome === dbNome(cap))?.ricarico);
  return { ricarichi: Object.fromEntries([...S.capitoli.map((c) => c.nome), cap.nome].map((n) => [n, ric])) };
}
function totaliOfferta(o) {
  let venduto = 0, acquistato = 0, righe = 0;
  for (const c of capitoliOff(o)) {
    const r = riepilogoCapitolo(c, opzCap(c));
    venduto += r.venduto; acquistato += r.acquistato; righe += r.righe;
  }
  return { venduto, acquistato, righe };
}
const capAtt = () => S.off?.capitoli.find((c) => c.nome === S.capAtt) || null;
/* Database di un capitolo dell'offerta: il nome con cui è nato, anche se poi è stato rinominato. */
const dbNome = (c) => c.base || c.nome;
function rinominaCapitolo(c) {
  dialogo({
    titolo: `Rinomina il capitolo ${c.nome}`,
    html: `<label class="f">Nuovo nome<input id="rc-nome" value="${esc(c.nome)}"></label>
      ${S.capitoli.some((x) => x.nome === dbNome(c)) ? `<p class="muted small">Il capitolo resta legato al database <b>${esc(dbNome(c))}</b> (articoli proposti e ricarico automatico).</p>` : ""}`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Rinomina", value: "ok", cls: "primary" }],
    onApri: (d) => setTimeout(() => { const i = $("#rc-nome", d); i.focus(); i.select(); }, 50),
    onConferma(v, d) {
      const nome = $("#rc-nome", d).value.trim();
      if (!nome) { toast("Scrivi il nome."); return false; }
      if (nome === c.nome) return;
      if (S.off.capitoli.some((x) => x !== c && x.nome === nome)) { toast("C'è già un capitolo con questo nome."); return false; }
      const vecchio = c.nome;
      if (!c.base) c.base = vecchio;
      if (c.base === nome) delete c.base;
      c.nome = nome;
      if (S.capAtt === vecchio) S.capAtt = nome;
      modificata(); disegnaOfferta(true);
      toast(`Capitolo rinominato in ${nome}.`);
    },
  });
}

/* ================= Elenco offerte ================= */
$("#nuova-offerta").addEventListener("click", async () => {
  const o = nuovaOffertaVuota();
  try { await db.salva(COLL.off, o.id, o); } catch (x) { return errore(x); }
  apriOfferta(o.id, false, o);
  $("#o-rif").focus();
});
$("#cerca-offerte").addEventListener("input", disegnaElencoOfferte);
$("#filtro-stato-off").addEventListener("change", disegnaElencoOfferte);
$("#filtro-cli-off").addEventListener("change", disegnaElencoOfferte);
$("#filtro-fam-off").addEventListener("change", disegnaElencoOfferte);

/* Ricerca nelle offerte: ogni parola deve comparire da qualche parte (intestazione o righe). */
const testoOfferta = (o, cli) => [o.nome, campiRif(o).famiglia, cli?.nome, cli?.codice, o.stato === "consegnata" ? "consegnata" : "lavorazione", o.note,
  ...capitoliOff(o).flatMap((c) => [c.nome, c.titoloComputo, ...(c.capitolati || [])])].filter(Boolean).join(" ").toLowerCase();
const testoRighe = (o) => capitoliOff(o).flatMap((c) => (c.righe || []).map((r) => [codiceArticolo(r), r.precodice, r.codice, r.nOrdine, r.descrizione, r.notePre, r.notePost, r.gmv, r.fornitore].filter(Boolean).join(" "))).join(" ").toLowerCase();
function disegnaElencoOfferte() {
  const parole = $("#cerca-offerte").value.trim().toLowerCase().split(/\s+/).filter(Boolean);
  const stato = $("#filtro-stato-off").value;
  // clienti presenti nelle offerte (più "senza cliente" se serve)
  { const sel = $("#filtro-cli-off"), v = sel.value;
    const ids = [...new Set(S.offerte.map((o) => o.cliente || ""))];
    const voci = ids.filter(Boolean).map((id) => [id, S.clienti.find((c) => c.id === id)?.nome || "(cliente eliminato)"]).sort((a, b) => a[1].localeCompare(b[1]));
    sel.innerHTML = `<option value="">Tutti i clienti</option>` + voci.map(([id, n]) => `<option value="${esc(id)}">${esc(n)}</option>`).join("") + (ids.includes("") ? `<option value="__nessuno">— senza cliente —</option>` : "");
    sel.value = [...sel.options].some((o) => o.value === v) ? v : ""; }
  const cliF = $("#filtro-cli-off").value;
  { const sel = $("#filtro-fam-off"), v = sel.value;
    const fams = [...new Set(S.offerte.map((o) => campiRif(o).famiglia))];
    sel.innerHTML = `<option value="">Tutte le famiglie</option>` + fams.filter(Boolean).sort().map((f) => `<option value="${esc(f)}">${esc(f)}</option>`).join("") + (fams.includes("") ? `<option value="__nessuna">— senza famiglia —</option>` : "");
    sel.value = [...sel.options].some((o) => o.value === v) ? v : ""; }
  const famF = $("#filtro-fam-off").value;
  const cliDi = (id) => S.clienti.find((c) => c.id === id);
  const nomeCli = (id) => cliDi(id)?.nome || "";
  const nelleRighe = new Set();
  const lista = S.offerte.filter((o) => {
    if (stato && (o.stato === "consegnata" ? "consegnata" : "lavorazione") !== stato) return false;
    if (cliF && (cliF === "__nessuno" ? !!o.cliente : o.cliente !== cliF)) return false;
    if (famF && (famF === "__nessuna" ? !!campiRif(o).famiglia : campiRif(o).famiglia !== famF)) return false;
    if (!parole.length) return true;
    const t = testoOfferta(o, cliDi(o.cliente));
    if (parole.every((p) => t.includes(p))) return true;
    const tr = t + " " + testoRighe(o);
    if (parole.every((p) => tr.includes(p))) { nelleRighe.add(o.id); return true; }
    return false;
  });
  $("#lista-offerte").innerHTML = lista.length ? lista.map((o) => {
    const t = totaliOfferta(o);
    const cons = o.stato === "consegnata";
    return `<tr class="cliccabile ${cons ? "off-consegnata" : "off-lavorazione"}" data-id="${esc(o.id)}">
      <td class="fam-cel">${esc(campiRif(o).famiglia)}</td>
      <td>${(() => { const r = campiRif(o); return r.rif ? `<b class="mono">${esc(r.rif)}</b>${r.lavoro ? " - " + esc(r.lavoro) : ""}` : `<b>${esc(o.nome || "(senza riferimento)")}</b>`; })()}${nelleRighe.has(o.id) ? ` <span class="badge info" title="La ricerca è stata trovata negli articoli dell'offerta">trovato negli articoli</span>` : ""}</td><td>${esc(nomeCli(o.cliente))}</td>
      <td class="small">${capitoliOff(o).map((c) => `<span class="badge info">${esc(c.nome)}</span>`).join(" ")}</td>
      <td class="num">${t.righe}</td><td class="num">${euro(t.venduto)}</td>
      <td class="num">${t.venduto ? pctTxt(marginePct(t.venduto, t.acquistato)) || "0%" : ""}</td>
      <td class="nowrap muted">${dataOra(o.aggiornatoIl)}</td>
      <td class="nowrap"><select class="stato-off" data-stato="${esc(o.id)}" title="Stato dell'offerta">
        <option value="lavorazione" ${cons ? "" : "selected"}>In lavorazione</option><option value="consegnata" ${cons ? "selected" : ""}>Consegnata</option></select></td>
      <td class="nowrap"><button class="btn small" data-dup="${esc(o.id)}">Duplica</button></td></tr>`;
  }).join("") : `<tr><td colspan="10" class="vuoto">${S.offerte.length ? "Nessuna offerta trovata." :
    S.catalogo.length ? "Ancora nessuna offerta: crea la prima con “+ Nuova offerta”." :
    "Il database articoli è vuoto: vai in “Importa e aggiorna” e carica il file Excel."}</td></tr>`;
}
$("#lista-offerte").addEventListener("change", async (e) => {
  const sel = e.target.closest("[data-stato]"); if (!sel) return;
  const o = S.offerte.find((x) => x.id === sel.dataset.stato); if (!o) return;
  const stato = sel.value, prima = o.stato;
  o.stato = stato; if (S.off?.id === o.id) S.off.stato = stato;
  disegnaElencoOfferte();
  try { await db.salva(COLL.off, o.id, { ...o }); } catch (x) { o.stato = prima; disegnaElencoOfferte(); errore(x, "Stato non salvato"); }
});
$("#lista-offerte").addEventListener("click", (e) => {
  if (e.target.closest("[data-stato]")) return e.stopPropagation();
  const dup = e.target.closest("[data-dup]");
  if (dup) { e.stopPropagation(); return duplica(S.offerte.find((o) => o.id === dup.dataset.dup)); }
  const tr = e.target.closest("tr[data-id]");
  if (tr) apriOfferta(tr.dataset.id);
});

async function duplica(o) {
  if (!o) return;
  // il riferimento Itecnici è univoco: la copia lo lascia vuoto, la famiglia resta (stesso lavoro)
  const rr = campiRif(o), lav = (rr.lavoro || o.nome || "Offerta") + " (copia)";
  const n = { ...normalizza(o), id: db.nuovoId(), famiglia: rr.famiglia, rif: "", lavoro: lav, nome: lav, stato: "lavorazione",
    creatoIl: Date.now(), aggiornatoIl: Date.now(), creatoDa: S.utente?.email || "" };
  try { await db.salva(COLL.off, n.id, n); toast("Offerta duplicata"); apriOfferta(n.id, false, n); } catch (x) { errore(x); }
}

/* ================= Offerta ================= */
function apriOfferta(id, daRemoto = false, dati = null) {
  const o = dati || S.offerte.find((x) => x.id === id);
  if (!o) return;
  const capPrima = S.capAtt;
  S.off = normalizza(o);
  S.dirty = false;
  if (!daRemoto) {
    S.sel = new Set();
    S.capAtt = S.off.capitoli.length ? "" : null; // "" = riepilogo, null = nessun capitolo ancora
    S.sub = "righe";
  } else if (!S.off.capitoli.some((c) => c.nome === capPrima)) S.capAtt = "";
  history.replaceState(null, "", "#offerta=" + encodeURIComponent(id));
  if (S.vista !== "offerta") vai("offerta"); else disegnaOfferta(true);
  disegnaTestata();
  statoSalvataggio(daRemoto ? "Aggiornata da un altro dispositivo" : "");
}

let tSalva;
/* Le righe di uno stesso capitolato stanno sempre in fila sotto il suo titolo.
   Ordine dei blocchi = ordine dei capitolati (barra "Capitolati per Oracle"); le righe senza capitolato in fondo.
   Dentro ogni blocco resta l'ordine scelto. */
function raggruppa(c) {
  for (const r of c.righe) if (r.capitolato && !c.capitolati.includes(r.capitolato)) c.capitolati.push(r.capitolato);
  const pos = (r) => (r.capitolato ? c.capitolati.indexOf(r.capitolato) : c.capitolati.length);
  c.righe = c.righe.map((r, i) => ({ r, i })).sort((a, b) => pos(a.r) - pos(b.r) || a.i - b.i).map((x) => x.r);
}
function modificata() {
  for (const c of S.off?.capitoli || []) raggruppa(c);
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

/* --- Riferimento dell'offerta: famiglia, riferimento Itecnici (univoco), lavoro (nota) ---
   "nome" resta il testo completo "riferimento - lavoro" usato ovunque (elenco, nomi dei file…). */
const RE_RIF = /^\s*([A-Za-z]{0,4}\d{5,}[A-Za-z0-9]*)\s*(?:[-–—:]\s*(.*))?$/;
function campiRif(o) {
  if (o.rif !== undefined || o.lavoro !== undefined) return { famiglia: o.famiglia || "", rif: o.rif || "", lavoro: o.lavoro || "" };
  const m = String(o.nome || "").match(RE_RIF);
  return m ? { famiglia: o.famiglia || "", rif: m[1].toUpperCase(), lavoro: (m[2] || "").trim() } : { famiglia: o.famiglia || "", rif: "", lavoro: o.nome || "" };
}
const nomeDaRif = (r) => [r.rif, r.lavoro].filter(Boolean).join(" - ");
function infoRif() {
  const o = S.off; if (!o) return;
  const r = campiRif(o), msg = [];
  const doppio = r.rif && S.offerte.filter((x) => x.id !== o.id && campiRif(x).rif === r.rif);
  if (doppio?.length) msg.push(`<span class="bad">⚠ Il riferimento ${esc(r.rif)} è già usato da: ${doppio.map((x) => esc(x.nome || "")).join("; ")}</span>`);
  const fam = r.famiglia && S.offerte.filter((x) => x.id !== o.id && campiRif(x).famiglia === r.famiglia);
  if (fam?.length) msg.push(`<span class="muted">Stessa famiglia: ${fam.map((x) => `${esc(x.nome || "")}${S.clienti.find((c) => c.id === x.cliente) ? " (" + esc(S.clienti.find((c) => c.id === x.cliente).nome) + ")" : ""}`).join("; ")}</span>`);
  $("#o-rif-info").innerHTML = msg.join("<br>");
}
function cambioRif() {
  const o = S.off; if (!o) return;
  o.famiglia = $("#o-fam").value.replace(/\s+/g, "").toUpperCase();
  o.rif = $("#o-rif").value.replace(/\s+/g, "").toUpperCase();
  o.lavoro = $("#o-lav").value.trim();
  o.nome = nomeDaRif(o);
  infoRif(); modificata();
}
for (const id of ["#o-fam", "#o-rif", "#o-lav"]) $(id).addEventListener("input", cambioRif);
for (const id of ["#o-fam", "#o-rif"]) $(id).addEventListener("change", (e) => { e.target.value = e.target.value.replace(/\s+/g, "").toUpperCase(); });

/* --- Testata dell'offerta --- */
function seg(el, valori, attuale, etichetta = (v) => v) {
  el.innerHTML = valori.map((v) => `<button type="button" data-v="${esc(String(v))}" class="${String(v) === String(attuale) ? "on" : ""}">${esc(etichetta(v))}</button>`).join("");
}
function disegnaTestata() {
  const o = S.off;
  if (!o) return;
  { const r = campiRif(o); $("#o-fam").value = r.famiglia; $("#o-rif").value = r.rif; $("#o-lav").value = r.lavoro; }
  $("#o-fam-lista").innerHTML = [...new Set(S.offerte.map((x) => campiRif(x).famiglia).filter(Boolean))].sort().map((f) => `<option value="${esc(f)}">`).join("");
  infoRif();
  $("#o-note").value = o.note || "";
  $("#o-cliente").innerHTML = `<option value="">— nessun cliente (niente controllo cartellino) —</option>` +
    S.clienti.map((c) => `<option value="${esc(c.id)}" class="${c.censito ? "cli-ok" : "cli-no"}" ${c.id === o.cliente ? "selected" : ""}>${c.censito ? "● " : "○ "}${esc(c.nome)}</option>`).join("");
  statoCensimento();
  seg($("#o-imp"), Object.keys(IMPOSTAZIONI), impOfferta(o), (v) => IMPOSTAZIONI[v]);
  $("#o-imp-nota").textContent = impOfferta(o) === "listino"
    ? "Listino: per ogni capitolo si definisce uno sconto di listino e si ricalcola il prezzo di listino di ogni articolo (riparametrizzato) in modo che, scontato, dia il prezzo di acquisto."
    : "Acquisto / vendita: ogni capitolo ha condizioni di acquisto e condizioni di vendita (ricarico o sconto sul listino).";
}
/* Impostazione dei prezzi: vale per tutta l'offerta (tutti i capitoli). */
const impOfferta = (o) => o.imp || (o.capitoli.length && o.capitoli.every((c) => c.imp === "listino") ? "listino" : "av");
$("#o-imp").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-v]"); if (!b || !S.off) return;
  S.off.imp = b.dataset.v;
  for (const c of S.off.capitoli) { aggiornaCondizioni(c); c.imp = b.dataset.v; }
  modificata(); disegnaTestata(); disegnaOfferta();
});
$("#o-note").addEventListener("input", (e) => { S.off.note = e.target.value; modificata(); });
$("#o-cliente").addEventListener("change", (e) => { S.off.cliente = e.target.value; statoCensimento(); modificata(); disegnaOfferta(); });
/* Cliente dell'offerta: verde se il cartellino è stato censito da Oracle, rosso se no. */
function statoCensimento() {
  const sel = $("#o-cliente"), cli = clienteOff(), nota = $("#o-cli-stato");
  sel.classList.toggle("cli-ok", !!cli?.censito);
  sel.classList.toggle("cli-no", !!cli && !cli.censito);
  nota.className = "small " + (!cli ? "muted" : cli.censito ? "cli-ok-txt" : "cli-no-txt");
  $("#o-cli-cod").value = cli?.codice || "";
  $("#o-cli-copia").disabled = !cli?.codice;
  nota.textContent = !cli ? "" : cli.censito ? `✓ Cartellino censito da Oracle il ${dataIt(cli.censito)}` : "✗ Cartellino non censito da Oracle";
}

const copiaCodCliente = async () => {
  const v = $("#o-cli-cod").value; if (!v) return;
  $("#o-cli-cod").select();
  try { await navigator.clipboard.writeText(v); } catch { document.execCommand?.("copy"); }
  toast(`Codice cliente ${v} copiato.`, 2000);
};
$("#o-cli-cod").addEventListener("click", copiaCodCliente);
$("#o-cli-copia").addEventListener("click", copiaCodCliente);

/* --- Barra dei capitoli --- */
function disegnaBarraCapitoli() {
  const o = S.off;
  const tot = (c) => riepilogoCapitolo(c, opzCap(c)).venduto;
  // sigla: una lettera per parola (KNX_MDT → KM, CONTROLLO_ACCESSI → CA), oppure le prime tre lettere (FIRE → FIR)
  const sigla = (n) => { const p = n.split(/[_\s.-]+/).filter(Boolean); return (p.length > 1 ? p.map((x) => x[0]).join("") : n).slice(0, 3).toUpperCase(); };
  $("#cap-tabs").innerHTML =
    `<button type="button" data-cap="" class="sb-voce ${S.capAtt === "" ? "on" : ""}" title="Riepilogo dell'offerta"><span class="sb-ico">${ICONA_RIEP}</span><span class="sb-testo">Riepilogo</span></button>` +
    o.capitoli.map((c) => `<button type="button" data-cap="${esc(c.nome)}" class="sb-voce ${S.capAtt === c.nome ? "on" : ""}" title="${esc(c.nome)} · ${euro(tot(c))} (doppio clic per rinominare)">
      <span class="sb-sigla">${esc(sigla(c.nome))}</span><span class="sb-testo sb-cap"><span>${esc(c.nome)}</span><span class="conta">${euro(tot(c))}</span></span></button>`).join("");
  const liberi = S.capitoli.filter((c) => !o.capitoli.some((x) => dbNome(x) === c.nome || x.nome === c.nome));
  $("#add-cap").innerHTML = `<option value="">Scegli il capitolo da aggiungere…</option>` + liberi.map((c) => `<option>${esc(c.nome)}</option>`).join("");
  $("#sb-add").hidden = !liberi.length;
  $("#add-cap").hidden = !liberi.length;
}
const ICONA_RIEP = '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/></svg>';

/* ----- Menu laterale: chiuso = solo icone; aperto = si allarga sopra la pagina ----- */
const menuAperto = (v) => $("#sidebar").classList.toggle("aperta", v);
$("#sb-toggle").addEventListener("click", () => menuAperto(!$("#sidebar").classList.contains("aperta")));
$("#sb-add").addEventListener("click", () => { menuAperto(true); setTimeout(() => { $("#add-cap").focus(); $("#add-cap").showPicker?.(); }, 180); });
$("#sidebar").addEventListener("click", (e) => {
  // scelta una voce di navigazione, il menu si richiude
  if (e.target.closest("[data-cap], [data-sub], [data-vista], #torna-offerte")) menuAperto(false);
});
document.addEventListener("click", (e) => {
  if ($("#sidebar").classList.contains("aperta") && !e.target.closest("#sidebar") && !e.target.closest("dialog")) menuAperto(false);
});
document.addEventListener("keydown", (e) => { if (e.key === "Escape") menuAperto(false); });
// spazio occupato dal banner della modalità prova, così il menu laterale entra tutto nello schermo
const altezzaBarra = () => document.documentElement.style.setProperty("--banner-h", ($("#banner-prova")?.hidden ? 0 : $("#banner-prova").offsetHeight) + "px");
window.addEventListener("resize", altezzaBarra);
setTimeout(altezzaBarra, 0);

$("#cap-tabs").addEventListener("dblclick", (e) => {
  const b = e.target.closest("button[data-cap]"); if (!b || !b.dataset.cap) return;
  rinominaCapitolo(S.off.capitoli.find((c) => c.nome === b.dataset.cap));
});
$("#cap-tabs").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-cap]"); if (!b) return;
  S.capAtt = b.dataset.cap; S.sel = new Set();
  disegnaOfferta(true);
});
$("#add-cap").addEventListener("change", (e) => {
  const nome = e.target.value; if (!nome) return;
  const modello = capAtt() || S.off.capitoli[S.off.capitoli.length - 1];
  // il nuovo capitolo parte con tipo offerta e prezzi del capitolo attuale, ma senza sconti
  S.off.capitoli.push(nuovoCapitolo(nome, { ...(modello ? { acq: modello.acq, ven: modello.ven, out: modello.out, sottocapitoli: modello.sottocapitoli } : {}), imp: impOfferta(S.off) }));
  S.capAtt = nome; S.sub = "righe"; S.sel = new Set();
  menuAperto(false);
  modificata(); disegnaOfferta(true);
});

/* --- Condizioni del capitolo (pagina Prezzi) e impostazioni dell'output --- */
const ricAutoCap = (c) => num(S.capitoli.find((x) => x.nome === dbNome(c))?.ricarico);
function disegnaImpostCapitolo() {
  const c = capAtt(); if (!c) return;
  aggiornaCondizioni(c);
  seg($("#c-ven-modo"), Object.keys(MODI_VEN), c.ven.modo, (v) => MODI_VEN[v]);
  const lis = c.imp === "listino";
  $("#c-ven-modo").hidden = lis;
  $("#cond-ven-titolo").textContent = lis ? "Condizioni di listino" : "Condizioni di vendita";
  $("#c-ven-card").classList.toggle("blocco-lis", lis);

  seg($("#c-out-prezzo"), Object.keys(OUT_PREZZO), c.out.prezzo, (v) => (v === "vendita" && c.imp === "listino" ? "Listino" : OUT_PREZZO[v]));
  seg($("#c-out-formato"), Object.keys(OUT_FORMATO), c.out.formato, (v) => OUT_FORMATO[v]);
  seg($("#c-sottocap"), ["SI", "NO"], c.sottocapitoli ? "SI" : "NO");
  $("#c-linea").value = c.lineaPartenza ?? 10;
  $("#c-magazzino").value = c.magazzino ?? "BG";
  const ricarico = c.ven.modo === "ricarico" && !lis;
  $("#c-target-calcola").textContent = ricarico ? "Calcola il ricarico" : "Calcola lo sconto";
  $("#c-target-nota").textContent = lis ? "Imposta lo sconto di listino generale per arrivare a quel totale di listino."
    : ricarico ? "Imposta il ricarico generale per arrivare a quel totale (condizione a pacchetto)."
    : "Imposta lo sconto generale sul listino per arrivare a quel totale (condizione a pacchetto).";
  $("#c-out-nota").textContent = c.out.formato === "netti"
    ? "Si compila solo la colonna Prezzo Netto."
    : "Si compila solo Sconto 3 [%]: un unico sconto sul listino acquisto, che somma tutti gli sconti e i ricarichi (negativo se il prezzo supera il listino). Attività tecnica e generici vanno sempre a prezzo netto.";
}
const suCap = (fn) => (e) => { const c = capAtt(); if (!c) return; aggiornaCondizioni(c); fn(c, e); };
const suSeg = (fn) => suCap((c, e) => { const b = e.target.closest("button[data-v]"); if (b) { fn(c, b.dataset.v); modificata(); disegnaOfferta(); } });
$("#c-ven-modo").addEventListener("click", suSeg((c, v) => (c.ven.modo = v)));
$("#c-out-prezzo").addEventListener("click", suSeg((c, v) => (c.out.prezzo = v)));
$("#c-out-formato").addEventListener("click", suSeg((c, v) => (c.out.formato = v)));
$("#c-sottocap").addEventListener("click", suSeg((c, v) => (c.sottocapitoli = v === "SI")));
$("#c-linea").addEventListener("change", suCap((c, e) => { c.lineaPartenza = num(e.target.value) || 10; modificata(); disegnaOfferta(); }));
$("#c-magazzino").addEventListener("input", suCap((c, e) => { c.magazzino = e.target.value.trim(); modificata(); }));
/* Testo di sconti o ricarico: si accetta solo se è scritto bene (es. 75+15, 55, +15, 0). */
function scriviCondizione(inp, ric, salva) {
  const t = inp.value.trim().replace(/\s+/g, "").replace(/%/g, "");
  if (!testoValido(t, ric)) {
    inp.classList.add("errato");
    toast(ric ? "Ricarico non valido: scrivi per esempio 28 oppure 20+5." : "Sconto non valido: scrivi per esempio 15, 55 oppure 75+15.", 6000);
    return false;
  }
  inp.classList.remove("errato");
  inp.value = t;
  salva(t);
  modificata(); ridisegnaOfferta();
  return true;
}
/* Condizione di una riga (Generale o GMA): cella Standard modificabile + celle Extra.
   Standard = sconto del database (o ricarico automatico). Se lo cambi vale al posto dello standard
   (0 = niente standard, prezzo di listino). Le celle Extra si aggiungono in cascata: compilata una ne compare un'altra.
   Sconti: 15 = sconto, -25 = sconto negativo (ricarico del 25%). Ricarichi: 5 = ricarico, -5 = riduzione.
   Si salva come testo: "+15+5" = standard più extra; "55+10" = 55 al posto dello standard, poi 10. */
const normTesto = (v) => String(v ?? "").trim().replace(/\s+/g, "").replace(/%/g, "").replace(/\./g, ",");
function scomponi(val) {
  const t = normTesto(val);
  if (!t) return { over: null, extra: [] };
  if (t.startsWith("+")) return { over: null, extra: t.slice(1).split("+").filter(Boolean) };
  const p = t.split("+").filter(Boolean);
  return { over: p[0] ?? null, extra: p.slice(1) };
}
// il numero nella cella è lo sconto (o il ricarico) così com'è: -25 in uno sconto = ricarico del 25%
const cellaDaTermine = (x) => x;
const termineDaCella = (v) => normTesto(v).replace(/^\+/, "");
const RE_CASCATA = /^-?\d+(?:,\d+)?(?:\+-?\d+(?:,\d+)?)*$/;
const htmlCella = (v, ph = "") => `<input class="cond cella" value="${esc(v)}" placeholder="${esc(ph)}" inputmode="decimal">`;
function htmlCondizione({ dove, gma = "", val, tipo, std = "", eredita = false, phStd = "", phExtra = "extra" }) {
  const { over, extra } = scomponi(val);
  const mod = over !== null;
  const valStd = mod ? over : eredita ? "" : std;
  const celle = [...extra.map((x) => cellaDaTermine(x, tipo)), ""];
  return `<span class="cond-box" data-dove="${dove}" data-gma="${esc(gma)}" data-tipo="${tipo}" data-std="${esc(std)}" data-eredita="${eredita ? 1 : ""}">
    <span class="std-wrap"><input class="cond std${mod ? " mod" : ""}" value="${esc(valStd)}" placeholder="${esc(eredita ? "= Generale" : phStd)}" inputmode="decimal"
      title="${mod ? `Modificato: lo standard era ${std || "—"}%` : "Standard: cambialo per sostituirlo (0 = niente standard)"}">${mod && std ? `<s class="std-barrato">${esc(std)}%</s>` : ""}</span>
    <span class="celle" title="Extra in cascata: compilata una cella ne compare un'altra">${celle.map((v, i) => htmlCella(v, i === 0 && !extra.length ? (mod || extra.length ? "extra" : phExtra) : "")).join("")}</span></span>`;
}
function salvaCondizione(box) {
  const c = capAtt(); if (!c) return;
  aggiornaCondizioni(c);
  const g = box.dataset.gma, tipo = box.dataset.tipo;
  const salva = {
    acqGen: (t) => (c.acq.gen = t),
    venScGen: (t) => (c.ven.scGen = t),
    ricGen: (t) => (c.ven.ricGen = t),
    acqGma: (t) => { if (t) c.acq.gma[g] = t; else delete c.acq.gma[g]; },
    venScGma: (t) => { if (t) c.ven.scGma[g] = t; else delete c.ven.scGma[g]; },
    lisGen: (t) => (c.lis.gen = t),
    lisGma: (t) => { if (t) c.lis.gma[g] = t; else delete c.lis.gma[g]; },
    ricGma: (t) => { if (t) c.ven.ricGma[g] = t; else delete c.ven.ricGma[g]; },
  }[box.dataset.dove];
  const stdInp = box.querySelector("input.std"), celle = [...box.querySelectorAll("input.cella")];
  const t = normTesto(stdInp.value), std = normTesto(box.dataset.std);
  // uguale allo standard (o vuoto) = non modificato; dove vale il Generale, qualunque valore scritto è proprio del GMA
  const over = !t || (t === std && !box.dataset.eredita) ? null : t;
  const extra = celle.map((x) => termineDaCella(x.value, tipo)).filter(Boolean);
  const errStd = over !== null && !RE_CASCATA.test(over);
  const errate = celle.filter((x) => x.value.trim() && !RE_CASCATA.test(termineDaCella(x.value, tipo)));
  stdInp.classList.toggle("errato", errStd);
  celle.forEach((x) => x.classList.toggle("errato", errate.includes(x)));
  const finale = over === null ? (extra.length ? "+" + extra.join("+") : "") : over + (extra.length ? "+" + extra.join("+") : "");
  if (errStd || errate.length || !testoValido(finale, tipo === "ric")) return toast("Valore non valido: in ogni cella un numero (es. 15, oppure -25 per uno sconto negativo).", 6000);
  salva(finale);
  modificata(); ridisegnaOfferta();
}
// compilando l'ultima cella extra ne compare una nuova accanto
document.addEventListener("input", (e) => {
  const inp = e.target.closest(".cond-box input.cella"); if (!inp) return;
  const celle = inp.parentElement;
  if (inp === celle.lastElementChild && inp.value.trim()) celle.insertAdjacentHTML("beforeend", htmlCella(""));
});
document.addEventListener("change", (e) => { const box = e.target.closest(".cond-box"); if (box && e.target.matches("input")) salvaCondizione(box); });
$("#c-target-calcola").addEventListener("click", suCap((c) => {
  const target = num($("#c-target").value);
  if (!target) return toast("Scrivi il prezzo target del capitolo.");
  const ris = vendutaPerTarget(c, target, opzCap(c));
  if (ris === null) return toast("Non riesco ad arrivarci: servono righe che prendono il valore Generale (senza un valore proprio sul GMA), oppure il target è fuori portata.", 6000);
  const ricarico = ris.campo === "ricGen";
  const applica = () => {
    if (ris.campo === "lisGen") c.lis.gen = ris.testo; else c.ven[ris.campo] = ris.testo;
    modificata(); disegnaOfferta();
    toast(`Aggiunto nella riga Generale un ${ricarico ? "ricarico" : ris.campo === "lisGen" ? "sconto di listino" : "sconto"} extra di ${pctTxt(ris.extra) || "0%"} per arrivare a ${euro(target)} (lo standard resta com'è).`, 6000);
  };
  const rc = riepilogoCapitolo(c, opzCap(c));
  if (target >= rc.acquistato) return applica();
  // prezzo target sotto il costo: si vende in perdita (di solito c'è una quantità sbagliata)
  const grandi = c.righe.filter((r) => num(r.qta) >= 1000).map((r) => `${esc(r.nOrdine || codiceArticolo(r))}: q.tà ${esc(numIt(r.qta))}`);
  dialogo({
    titolo: "Il prezzo target è sotto il costo",
    html: `<p>Per arrivare a <b>${euro(target)}</b> bisognerebbe vendere sotto il prezzo di acquisto
      (acquistato del capitolo: <b>${euro(rc.acquistato)}</b>).</p>
      ${grandi.length ? `<p>Controlla le quantità: ${grandi.length === 1 ? "questa riga ha" : "queste righe hanno"} una quantità molto alta:</p><ul class="riep">${grandi.slice(0, 8).map((x) => `<li>${x}</li>`).join("")}</ul>` : ""}`,
    pulsanti: [{ label: "Non applicare", value: "annulla" }, { label: "Applica comunque", value: "ok", cls: "danger" }],
    onConferma: applica,
  });
}));
function togliCapitolo(c) {
  dialogo({
    titolo: `Togliere ${c.nome} dall'offerta?`,
    html: `<p>Il capitolo ha ${c.righe.length} righe: verranno tolte dall'offerta insieme ai suoi sconti. Il database articoli non cambia.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Togli capitolo", value: "ok", cls: "danger" }],
    onConferma() {
      S.off.capitoli = S.off.capitoli.filter((x) => x !== c);
      S.capAtt = S.off.capitoli.length ? "" : null;
      modificata(); disegnaOfferta(true);
    },
  });
}

$("#subtabs").addEventListener("click", (e) => {
  const b = e.target.closest("button"); if (!b) return;
  S.sub = b.dataset.sub; disegnaOfferta(true);
});

/* Dopo un ridisegno rimette il cursore nel campo dove si trovava (Tab da una quantità alla successiva). */
function ricordaFocus() {
  const a = document.activeElement;
  if (!a || a.tagName !== "INPUT") return () => {};
  const tr = a.closest("tr[data-id], tr[data-k]");
  let sel = null;
  if (tr && a.dataset.campo) sel = `tr[data-id="${CSS.escape(tr.dataset.id)}"] [data-campo="${a.dataset.campo}"]`;
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
  // totali dell'offerta
  const cli = clienteOff();
  let venduto = 0, acquistato = 0, righe = 0, fuori = 0;
  for (const c of o.capitoli) {
    const r = riepilogoCapitolo(c, opzCap(c));
    venduto += r.venduto; acquistato += r.acquistato; righe += r.righe;
    fuori += controlloCartellino(c, cli, opzCap(c)).filter((x) => x.stato === "fuori").length;
  }
  $("#o-totali").innerHTML = `
    <div class="kpi kpi-main"><span>${impOfferta(o) === "listino" ? "Totale listino" : "Venduto"}</span><span class="big">${euro(venduto)}</span></div>
    ${impOfferta(o) === "listino" ? "" : `<div class="kpi k-acq"><span>Acquistato</span><span class="big">${euro(acquistato)}</span></div>
    <div class="kpi k-mar ${venduto && marginePct(venduto, acquistato) < 0 ? "k-neg" : ""}"><span>Margine</span><span class="big">${venduto ? pctTxt(marginePct(venduto, acquistato)) || "0%" : "—"}</span></div>
    <div class="kpi k-ric"><span>Ricarico</span><span class="big">${acquistato ? pctTxt(ricaricoPct(venduto, acquistato)) || "0%" : "—"}</span></div>`}
    ${o.cliente ? `<div class="kpi"><span>Cartellino ${esc(cli?.nome || "")}</span><span>${
      fuori ? `<span class="badge bad">${fuori} GMA fuori cartellino</span>` : `<span class="badge ok">nessun problema</span>`}</span></div>` : ""}`;
  disegnaBarraCapitoli();
  const c = capAtt();
  // riferimento, cliente e totali dell'offerta: solo nel riepilogo generale
  $("#o-testata").hidden = !!c;
  $("#o-titolo").textContent = `${o.nome || "(senza riferimento)"}${cli ? " · " + cli.nome : ""}`;
  $("#sb-capitolo").hidden = !c;
  if (c) $("#sb-cap-nome").textContent = c.nome;
  $("#p-vuota").hidden = S.capAtt !== null;
  $("#p-riepilogo").hidden = S.capAtt !== "";
  $("#p-capitolo").hidden = !c;
  if (S.capAtt === "") return disegnaRiepilogo();
  if (!c) return;
  disegnaImpostCapitolo();
  const rc = riepilogoCapitolo(c, opzCap(c));
  // venduto / acquistato / margine del capitolo: solo nella pagina Sconti e ricarichi
  $(".cap-testa").hidden = S.sub !== "prezzi";
  $("#c-totali").innerHTML = (S.sub !== "prezzi" ? "" : `
    <span class="mini-kpi k-ven"><span>${c.imp === "listino" ? "Totale listino" : "Venduto"}</span><b>${euro(rc.venduto)}</b></span>
    ${c.imp === "listino" ? "" : `<span class="mini-kpi k-acq"><span>Acquistato</span><b>${euro(rc.acquistato)}</b></span>
    <span class="mini-kpi k-mar ${rc.venduto && marginePct(rc.venduto, rc.acquistato) < 0 ? "k-neg" : ""}"><span>Margine</span><b>${rc.venduto ? pctTxt(marginePct(rc.venduto, rc.acquistato)) || "0%" : "—"}</b></span>`}
    <span class="mini-kpi k-info"><span>Condizioni</span><b>${esc(descrCondizioni(c))}</b></span>`);
  $$("#subtabs button").forEach((b) => b.classList.toggle("on", b.dataset.sub === S.sub));
  for (const s of ["righe", "prezzi", "export"]) $("#s-" + s).hidden = s !== S.sub;
  $("#conta-righe").textContent = c.righe.length ? `(${c.righe.length})` : "";
  $("#conta-gmv").textContent = rc.perGmv.size ? `(${rc.perGmv.size})` : "";
  if (S.sub === "righe") { disegnaRighe(c); if (!S.noPicker) disegnaPicker(c, cambioSub); }
  else if (S.sub === "prezzi") { disegnaGmv(c, rc); disegnaRiepilogoRighe(c); }
  else if (S.sub === "export") disegnaExport(c);
}

/* Descrizione breve delle condizioni del capitolo */
function descrCondizioni(c) {
  const k = condizioni(c);
  const conExtra = k.acq.gen || Object.keys(k.acq.gma).length;
  if (k.imp === "listino") return `Listino riparametrizzato · Acq. ${conExtra ? "con condizioni proprie" : "standard"} · Oracle: ${k.out.prezzo === "acquisto" ? "acquisto" : "listino"}, ${OUT_FORMATO[k.out.formato].toLowerCase()}`;
  return `Acq. ${conExtra ? "con condizioni proprie" : "standard"} · Vend. ${k.ven.modo === "ricarico" ? "ricarico" : "listino/sconto"} · Oracle: ${OUT_PREZZO[k.out.prezzo].toLowerCase()}, ${OUT_FORMATO[k.out.formato].toLowerCase()}`;
}
/* --- Riepilogo dell'offerta (ex ANALISI PACCHETTO) --- */
function disegnaRiepilogo() {
  const o = S.off, cli = clienteOff();
  const tot = { venduto: 0, acquistato: 0, attV: 0, matV: 0, righe: 0 };
  const lisO = impOfferta(o) === "listino";
  $("#riepilogo-capitoli").closest("table").classList.toggle("modo-listino", lisO);
  $("#rc-vend").textContent = lisO ? "Listino" : "Venduto";
  $("#rc-nota").hidden = lisO;
  const righe = o.capitoli.map((c) => {
    const r = riepilogoCapitolo(c, opzCap(c));
    const fuori = controlloCartellino(c, cli, opzCap(c)).filter((x) => x.stato === "fuori").length;
    tot.venduto += r.venduto; tot.acquistato += r.acquistato; tot.attV += r.attivita.venduto; tot.matV += r.materiale.venduto; tot.righe += r.righe;
    return `<tr class="${c.esporta === false ? "qta-zero" : ""}"><td><input type="checkbox" data-esporta="${esc(c.nome)}" ${c.esporta === false ? "" : "checked"} title="Esporta questo capitolo"></td>
      <td><span class="nowrap"><button class="btn small" data-vai-cap="${esc(c.nome)}">${esc(c.nome)}</button><button type="button" class="tg-ed" data-rinomina-cap="${esc(c.nome)}" title="Rinomina il capitolo">✎</button></span>${c.base ? `<div class="muted small">database ${esc(c.base)}</div>` : ""}${c.fonte ? `<div class="muted small" title="Importato dall'offerta del fornitore">da ${esc(c.fonte)}</div>` : ""}</td>
      <td class="small nowrap">${esc(descrCondizioni(c))}</td><td class="num">${r.righe}</td>
      <td class="num"><b>${euro(r.venduto)}</b></td><td class="num solo-av">${euro(r.acquistato)}</td>
      <td class="num">${euro(r.materiale.venduto)}</td><td class="num">${euro(r.attivita.venduto)}</td>
      <td class="num solo-av">${r.acquistato ? pctTxt(ricaricoPct(r.venduto, r.acquistato)) || "0%" : "—"}</td>
      <td class="num solo-av">${r.venduto ? pctTxt(marginePct(r.venduto, r.acquistato)) || "0%" : "—"}</td>
      <td>${!cli ? "" : fuori ? `<span class="badge bad">${fuori} fuori</span>` : '<span class="badge ok">OK</span>'}</td>
      <td class="nowrap"><button class="btn small" data-xlsx-cap="${esc(c.nome)}" ${r.righe ? "" : "disabled"} title="Scarica il foglio per Oracle di questo capitolo">⤓ Excel</button>
        <button class="icon" data-togli-cap="${esc(c.nome)}" title="Togli il capitolo dall'offerta">✕</button></td></tr>`;
  }).join("");
  $("#riepilogo-capitoli").innerHTML = righe ? righe + `<tr class="totale"><td></td><td>Totale offerta</td><td></td><td class="num">${tot.righe}</td>
      <td class="num">${euro(tot.venduto)}</td><td class="num solo-av">${euro(tot.acquistato)}</td><td class="num">${euro(tot.matV)}</td><td class="num">${euro(tot.attV)}</td>
      <td class="num solo-av">${tot.acquistato ? pctTxt(ricaricoPct(tot.venduto, tot.acquistato)) || "0%" : ""}</td>
      <td class="num solo-av">${tot.venduto ? pctTxt(marginePct(tot.venduto, tot.acquistato)) || "0%" : ""}</td><td></td><td></td></tr>`
    : `<tr><td colspan="12" class="vuoto">Nessun capitolo: aggiungine uno con “+ Aggiungi capitolo” o importa un'offerta fornitore.</td></tr>`;
  const daEsp = o.capitoli.filter((c) => c.esporta !== false);
  $("#scarica-tutti").disabled = !daEsp.some((c) => calcolaOfferta(c).righe.length);
  const et = $("#esp-tutti"); et.checked = daEsp.length === o.capitoli.length && o.capitoli.length > 0; et.indeterminate = daEsp.length > 0 && !et.checked;
}
$("#riepilogo-capitoli").addEventListener("change", (e) => {
  const t = e.target.closest("[data-esporta]"); if (!t) return;
  const c = S.off.capitoli.find((x) => x.nome === t.dataset.esporta); if (!c) return;
  c.esporta = t.checked; modificata(); disegnaRiepilogo();
});
$("#esp-tutti").addEventListener("change", (e) => { for (const c of S.off.capitoli) c.esporta = e.target.checked; modificata(); disegnaRiepilogo(); });
$("#riepilogo-capitoli").addEventListener("click", (e) => {
  const rn = e.target.closest("[data-rinomina-cap]");
  if (rn) return rinominaCapitolo(S.off.capitoli.find((c) => c.nome === rn.dataset.rinominaCap));
  const v = e.target.closest("[data-vai-cap]");
  if (v) { S.capAtt = v.dataset.vaiCap; S.sub = "righe"; S.sel = new Set(); return disegnaOfferta(true); }
  const t = e.target.closest("[data-togli-cap]");
  if (t) return togliCapitolo(S.off.capitoli.find((c) => c.nome === t.dataset.togliCap));
  const x = e.target.closest("[data-xlsx-cap]");
  if (x) impostaEScarica(S.off.capitoli.find((c) => c.nome === x.dataset.xlsxCap));
});
$("#scarica-tutti").addEventListener("click", () => {
  const caps = S.off.capitoli.filter((c) => c.esporta !== false && calcolaOfferta(c).righe.length);
  const vai = async () => { for (const c of caps) { scaricaXlsx(c); await new Promise((r) => setTimeout(r, 600)); } };
  const senza = caps.filter((c) => righeSenzaCapitolato(c));
  if (!senza.length) return vai();
  dialogo({
    titolo: "Non tutte le righe hanno un capitolato",
    html: `<p>Righe senza capitolato in: ${senza.map((c) => `<b>${esc(c.nome)}</b> (${righeSenzaCapitolato(c)})`).join(", ")}.</p><p>Vuoi scaricare comunque?</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Scarica comunque", value: "ok", cls: "primary" }],
    onConferma: () => { vai(); },
  });
});

/* --- Righe del capitolo --- */
function lineeOracle(c) {
  // numero di linea Oracle di ogni riga (le righe senza codice o a quantità 0 non hanno linea)
  const m = new Map();
  let l = num(c.lineaPartenza) || 10;
  for (const r of c.righe) if (num(r.qta) !== 0 && codiceOracle(r) && !esclusoDaExport(r, c)) { m.set(r.id, l); l += 10; }
  return m;
}
function aggiungiCapitolato(c, nome) {
  nome = (nome || "").trim();
  if (!nome) return false;
  if (!c.capitolati.includes(nome)) c.capitolati.push(nome);
  return true;
}
/* Capitolati: rinomina (✎) e sposta il blocco (▲ ▼) dalla barra blu del capitolato nelle righe dell'offerta */
function rinominaCapitolato(c, vecchio) {
  dialogo({
    titolo: "Rinomina capitolato",
    html: `<label class="f">Nome che va in Oracle (Nome Capitolato)<input id="r-cap" value="${esc(vecchio)}"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Rinomina", value: "ok", cls: "primary" }],
    onApri: (d) => setTimeout(() => { $("#r-cap", d).focus(); $("#r-cap", d).select(); }, 50),
    onConferma(v, body) {
      const nuovo = $("#r-cap", body).value.trim();
      if (!nuovo) return false;
      if (nuovo !== vecchio && c.capitolati.includes(nuovo)) { // unisce i due capitolati
        c.capitolati = c.capitolati.filter((x) => x !== vecchio);
      } else {
        c.capitolati = c.capitolati.map((x) => (x === vecchio ? nuovo : x));
      }
      c.righe.forEach((r) => { if (r.capitolato === vecchio) r.capitolato = nuovo; });
      if (c.capNuove === vecchio) c.capNuove = nuovo;
      c.esclusiExport = (c.esclusiExport || []).map((x) => (x === vecchio ? nuovo : x));
      modificata(); disegnaOfferta();
    },
  });
}
function spostaCapitolato(c, nome, verso) {
  // solo tra i capitolati che hanno righe, così ogni clic sposta davvero il blocco
  const usati = c.capitolati.filter((k) => c.righe.some((r) => r.capitolato === k));
  const i = usati.indexOf(nome), j = i + verso;
  if (i < 0 || j < 0 || j >= usati.length) return;
  [usati[i], usati[j]] = [usati[j], usati[i]];
  c.capitolati = [...usati, ...c.capitolati.filter((k) => !usati.includes(k))];
  modificata(); disegnaOfferta();
}


/* Gestione dei sottocapitoli (capitolati) del capitolo: rinomina, ordine, raggruppa, nuovi, elimina.
   Le modifiche si fanno su una copia e si applicano con "Salva". */
function gestisciSottocapitoli(c) {
  raggruppa(c);
  const opz = opzCap(c);
  const infoDi = (nomi) => { let n = 0, t = 0; for (const r of c.righe) if (nomi.includes(r.capitolato || "")) { n++; if (num(r.qta)) t += calcolaRiga(r, c, opz.ricarichi).totale; } return { n, t }; };
  let voci = c.capitolati.map((k) => ({ orig: [k], nome: k, del: "" }));
  const senza = infoDi([""]);
  const disegna = (d) => {
    $("#gs-righe", d).innerHTML = voci.map((v, i) => { const x = infoDi(v.orig); return `<tr class="${v.del ? "qta-zero" : ""}">
      <td><input type="checkbox" data-gs-sel="${i}" ${v.sel ? "checked" : ""} ${v.del ? "disabled" : ""}></td>
      <td><input data-gs-nome="${i}" value="${esc(v.nome)}" style="width:100%;min-width:240px" ${v.del ? "disabled" : ""}>${v.orig.length > 1 ? `<div class="muted small">uniti: ${v.orig.map(esc).join(", ")}</div>` : ""}</td>
      <td class="num">${x.n}</td><td class="num">${euro(x.t)}</td>
      <td class="nowrap"><button type="button" class="tg-ed" data-gs-su="${i}" title="Sposta su" ${i ? "" : "disabled"}>▲</button><button type="button" class="tg-ed" data-gs-giu="${i}" title="Sposta giù" ${i < voci.length - 1 ? "" : "disabled"}>▼</button></td>
      <td><select data-gs-del="${i}"><option value="">tieni</option><option value="cap" ${v.del === "cap" ? "selected" : ""}>elimina (le righe restano senza sottocapitolo)</option><option value="tutto" ${v.del === "tutto" ? "selected" : ""}>elimina con le sue righe</option></select></td></tr>`; }).join("")
      || `<tr><td colspan="6" class="vuoto">Nessun sottocapitolo.</td></tr>`;
    const sel = voci.filter((v) => v.sel), nSel = sel.length;
    const tSel = infoDi(sel.flatMap((v) => v.orig));
    $("#gs-unisci", d).disabled = nSel < 2;
    $("#gs-nsel", d).innerHTML = !nSel ? "spunta almeno due sottocapitoli"
      : `${nSel === 1 ? "1 sottocapitolo selezionato" : `${nSel} sottocapitoli selezionati`} · ${tSel.n} righe · totale <b>${euro(tSel.t)}</b>`;
  };
  dialogo({
    titolo: `Sottocapitoli di ${c.nome}`, largo: true,
    html: `<p class="muted small" style="margin-top:0">Il nome del sottocapitolo va in Oracle come “Nome Capitolato”. Cambia il nome nella casella, sposta con ▲ ▼,
        spunta più sottocapitoli per raggrupparli in uno solo. ${senza.n ? `Righe senza sottocapitolo: <b>${senza.n}</b>.` : ""} Le modifiche si applicano con <b>Salva</b>.</p>
      <div class="tbl-wrap" style="max-height:48vh"><table><thead><tr><th></th><th>Sottocapitolo</th><th class="num">Righe</th><th class="num">Totale</th><th>Ordine</th><th>Elimina</th></tr></thead>
        <tbody id="gs-righe"></tbody></table></div>
      <div class="card row" style="margin-top:10px;gap:8px;align-items:flex-end;flex-wrap:wrap">
        <b style="align-self:center">Raggruppa</b><span class="muted small" id="gs-nsel" style="align-self:center"></span>
        <label class="f grow">Nome del sottocapitolo unico<input id="gs-unisci-nome" placeholder="es. PIANO TERRA"></label>
        <button type="button" class="btn" id="gs-unisci">Raggruppa i selezionati</button></div>
      <div class="card row" style="margin-top:8px;gap:8px;align-items:flex-end;flex-wrap:wrap">
        <b style="align-self:center">+ Nuovo sottocapitolo</b>
        <label class="f grow">Nome<input id="gs-nuovo" placeholder="es. PIANO PRIMO"></label>
        <button type="button" class="btn" id="gs-aggiungi">Aggiungi</button></div>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(d) {
      d.addEventListener("change", (e) => {
        const t = e.target;
        if (t.dataset.gsSel !== undefined) { voci[+t.dataset.gsSel].sel = t.checked; disegna(d); }
        else if (t.dataset.gsDel !== undefined) { const v = voci[+t.dataset.gsDel]; v.del = t.value; if (v.del) v.sel = false; disegna(d); }
      });
      d.addEventListener("input", (e) => { const t = e.target; if (t.dataset.gsNome !== undefined) voci[+t.dataset.gsNome].nome = t.value; });
      d.addEventListener("click", (e) => {
        const su = e.target.closest("[data-gs-su]"), giu = e.target.closest("[data-gs-giu]");
        if (su || giu) { const i = +(su || giu).dataset[su ? "gsSu" : "gsGiu"], j = i + (su ? -1 : 1); [voci[i], voci[j]] = [voci[j], voci[i]]; disegna(d); return; }
        if (e.target.id === "gs-unisci") {
          const sel = voci.filter((v) => v.sel);
          const nome = $("#gs-unisci-nome", d).value.trim() || sel[0].nome;
          const primo = sel[0];
          primo.orig = sel.flatMap((v) => v.orig); primo.nome = nome; primo.sel = false;
          voci = voci.filter((v) => v === primo || !sel.includes(v));
          $("#gs-unisci-nome", d).value = "";
          disegna(d); return;
        }
        if (e.target.id === "gs-aggiungi") {
          const n = $("#gs-nuovo", d).value.trim();
          if (!n) return toast("Scrivi il nome.");
          if (voci.some((v) => v.nome === n)) return toast("C'è già un sottocapitolo con questo nome.");
          voci.push({ orig: [], nome: n, del: "" }); $("#gs-nuovo", d).value = ""; disegna(d);
        }
      });
      $("#gs-nuovo", d).addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); $("#gs-aggiungi", d).click(); } });
      $("#gs-unisci-nome", d).addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); if (!$("#gs-unisci", d).disabled) $("#gs-unisci", d).click(); } });
      disegna(d);
    },
    onConferma() {
      if (voci.some((v) => !v.del && !v.nome.trim())) { toast("Ogni sottocapitolo deve avere un nome."); return false; }
      const nuovoNome = new Map(); // nome vecchio → nuovo ("" = senza; null = righe eliminate)
      for (const v of voci) for (const k of v.orig) nuovoNome.set(k, v.del === "tutto" ? null : v.del === "cap" ? "" : v.nome.trim());
      const prima = c.righe.length;
      c.righe = c.righe.filter((r) => nuovoNome.get(r.capitolato || "") !== null || !r.capitolato);
      for (const r of c.righe) if (r.capitolato && nuovoNome.has(r.capitolato)) r.capitolato = nuovoNome.get(r.capitolato);
      c.capitolati = [...new Set(voci.filter((v) => !v.del).map((v) => v.nome.trim()))];
      const mappa = (k) => (nuovoNome.has(k) ? nuovoNome.get(k) : k);
      c.esclusiExport = [...new Set((c.esclusiExport || []).map(mappa).filter(Boolean))];
      if (c.capNuove) c.capNuove = mappa(c.capNuove) || "";
      for (const id of [...S.sel]) if (!c.righe.some((r) => r.id === id)) S.sel.delete(id);
      raggruppa(c);
      modificata(); disegnaOfferta();
      const tolte = prima - c.righe.length;
      toast(`Sottocapitoli aggiornati${tolte ? `; ${tolte} righe eliminate` : ""}.`);
    },
  });
}
$("#gestisci-sotto").addEventListener("click", () => { const c = capAtt(); if (c) gestisciSottocapitoli(c); });

/* Pagina Righe: database a tutta pagina, righe dell'offerta a tutta pagina, oppure tutte e due */
function applicaVistaRighe(c) {
  if (!S.vistaRighe) S.vistaRighe = c && c.righe.length ? "righe" : "db";
  const v = S.vistaRighe;
  $("#vr-db").hidden = v === "righe" || v === "bib";
  $("#vr-off").hidden = v === "db" || v === "bib";
  $("#vr-bib").hidden = v !== "bib";
  if (v === "bib") { caricaBib(); disegnaBibRighe(); }
  $("#s-righe").classList.toggle("vr-piena", v !== "tutte");
  $$("#vista-righe [data-vr]").forEach((b) => b.classList.toggle("on", b.dataset.vr === v));
  if (c) $("#vr-conta").textContent = c.righe.length ? `(${c.righe.length})` : "";
}
function cambiaVistaRighe(v) {
  S.vistaRighe = v;
  applicaVistaRighe(capAtt());
  window.scrollTo(0, 0);
  if (v === "db") setTimeout(() => $("#a-cerca").focus(), 0);
  if (v === "bib") setTimeout(() => $("#rb-cerca").focus(), 0);
}
$("#vista-righe").addEventListener("click", (e) => { const b = e.target.closest("[data-vr]"); if (b) cambiaVistaRighe(b.dataset.vr); });
document.addEventListener("keydown", (e) => {
  if (!e.altKey || S.vista !== "offerta" || S.sub !== "righe" || !capAtt()) return;
  const v = { Digit1: "db", Digit2: "righe", Digit3: "tutte", Digit4: "bib" }[e.code];
  if (v) { e.preventDefault(); cambiaVistaRighe(v); }
});

/* Ctrl/Cmd + X, C, V sulle righe selezionate (non mentre si scrive in un campo) */
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || S.vista !== "offerta" || S.sub !== "righe" || !capAtt()) return;
  const t = e.target;
  if (t.closest?.("input:not([type=checkbox]), textarea, select, [contenteditable], dialog")) return;
  if (String(window.getSelection?.() || "").trim()) return; // testo selezionato col mouse: copia normale
  const az = { KeyX: "taglia", KeyC: "copia", KeyV: "incolla" }[e.code];
  if (!az || (az !== "incolla" && !S.sel.size)) return;
  e.preventDefault(); chiudiMenuDx(); azioneRighe(az);
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && S.appunti?.taglio && !document.querySelector("dialog[open]")) { S.appunti = null; if (capAtt()) disegnaOfferta(); }
});

/* Esc: chiude il menu se è aperto, altrimenti deseleziona le righe dell'offerta e gli articoli del database */
window.addEventListener("keydown", (e) => {
  if (e.key !== "Escape" || document.querySelector("dialog[open]")) return;
  if (!$("#menu-dx").hidden) return; // il primo Esc chiude solo il menu
  if (S.vista !== "offerta" || !capAtt()) return;
  const cerano = S.sel.size || S.pickSel.size;
  S.sel = new Set(); S.pickSel.clear();
  if (cerano) disegnaOfferta();
}, true);

function disegnaRighe(c) {
  applicaVistaRighe(c);
  const tb = $("#righe-offerta");
  const sel = S.sel;
  for (const id of [...sel]) if (!c.righe.some((r) => r.id === id)) sel.delete(id);
  $("#bulk-n").textContent = !sel.size ? "Nessuna riga selezionata" : sel.size === 1 ? "1 riga selezionata" : `${sel.size} righe selezionate`;
  $("#bulk").classList.toggle("attiva", sel.size > 0);
  $("#sel-tutte").checked = c.righe.length > 0 && sel.size === c.righe.length;
  if (!c.righe.length) {
    tb.innerHTML = `<tr><td colspan="13" class="vuoto">Nessuna riga: cerca gli articoli qui sopra e scrivi le quantità.</td></tr>`;
    return;
  }
  const linee = lineeOracle(c);
  const opz = opzCap(c);
  let tot = 0, capPrec = null;
  const opzioniCap = (v) => `<option value="">—</option>` + c.capitolati.map((n) => `<option ${n === v ? "selected" : ""}>${esc(n)}</option>`).join("");
  tb.innerHTML = c.righe.map((r, i) => {
    const x = calcolaRiga(r, c, opz.ricarichi);
    if (num(r.qta) !== 0) tot += x.totale;
    const att = isAttivita(r);
    // "a giornate": la quantità mostrata è il numero di giorni e la colonna Giorni mostra la quantità (i campi si invertono)
    const scambio = att && r.aGiornate, cQ = scambio ? "giorni" : "qta", cG = scambio ? "qta" : "giorni";
    let testa = "";
    if ((r.capitolato || "") !== capPrec) {
      capPrec = r.capitolato || "";
      testa = `<tr class="gruppo" data-gruppo="${esc(capPrec)}"><td colspan="13"><input type="checkbox" class="sel-gruppo" data-selgruppo title="Seleziona tutte le righe di questo capitolato"> ${capPrec
        ? `${esc(capPrec)}<button type="button" class="tg-ed" data-rinomina="${esc(capPrec)}" title="Rinomina il capitolato">✎</button><button type="button" class="tg-ed" data-capsposta="-1" title="Sposta il capitolato prima">▲</button><button type="button" class="tg-ed" data-capsposta="1" title="Sposta il capitolato dopo">▼</button>`
        : '<span style="font-weight:500">senza capitolato</span>'}</td></tr>`;
    }
    return testa + `<tr data-id="${esc(r.id)}" class="${sel.has(r.id) ? "sel" : ""}${att ? " riga-att" : ""}${num(r.qta) ? "" : " qta-zero"}${S.appunti?.taglio && S.appunti.ids.has(r.id) ? " tagliata" : ""}">
      <td><input type="checkbox" data-sel ${sel.has(r.id) ? "checked" : ""}></td>
      <td class="num muted small">${linee.get(r.id) ?? ""}</td>
      <td class="mono nowrap">${codiceArticolo(r) ? codiceVis(r) : '<span class="badge bad">senza codice</span>'} ${statoRiga(r)}</td>
      <td class="img-cel">${(() => { const im = r.img || S.perK.get(r.k)?.img; return im ? `<img class="img-mini" data-img="${esc(im)}" alt="">` : ""; })()}</td>
      <td class="desc"><b>${esc(r.nOrdine)}</b>${r.capitolo !== c.nome && r.capitolo !== dbNome(c) && !r.daOfferta ? ` <span class="badge info" title="Articolo preso da un altro database">${esc(r.capitolo)}</span>` : ""}<div class="d">${esc(r.descrizione)}</div></td>
      <td>${esc(r.um)}</td>
      <td class="num"><input class="qta${num(r[cQ]) >= 1000 ? " qta-alta" : ""}" data-campo="${cQ}" inputmode="decimal" value="${esc(numIt(r[cQ]))}"${num(r[cQ]) >= 1000 ? ' title="Quantità molto alta: controlla che sia giusta"' : ""}>${num(x.qta) !== num(r[cQ]) && num(r.qta) && !scambio ? `<div class="muted small nowrap" title="Quantità che va in Oracle">Oracle: ${esc(numIt(x.qta))}</div>` : ""}</td>
      <td class="num nowrap">${att ? `<input class="qta" data-campo="${cG}" inputmode="decimal" value="${esc(numIt(r[cG]))}" title="${scambio ? "A giornate: qui il numero di volte (la quantità è il numero di giorni)" : "Numero di giorni"}">
        <label class="giornate" title="In vendita e a listino: in Oracle quantità = giorni e prezzo = valore giornaliero (invece di un valore complessivo con quantità 1)">
          <input type="checkbox" data-campo="aGiornate" ${r.aGiornate ? "checked" : ""} > a giornate</label>` : ""}</td>
      <td class="num"><b>${euro(x.unitario)}</b></td><td class="num">${euro(num(r.qta) ? x.totale : 0)}</td>
      <td><select class="capsel" data-campo="capitolato">${opzioniCap(r.capitolato)}</select></td>
      ${isGenerico(r)
        ? `<td><span class="nota-auto" title="Articolo generico: nota pre-linea automatica = n° di ordinazione">auto: ${esc(r.nOrdine)}</span></td>
           <td><span class="nota-auto" title="Articolo generico: nota post-linea automatica = descrizione">auto: descrizione</span></td>`
        : `<td><input class="nota" data-campo="notePre" value="${esc(r.notePre)}" placeholder="nota pre-linea"></td>
           <td><input class="nota" data-campo="notePost" value="${esc(r.notePost)}" placeholder="nota post-linea"></td>`}</tr>`;
  }).join("") + `<tr class="sotto"><td colspan="9" class="num">Totale ${esc(c.nome)}</td><td class="num">${euro(tot)}</td><td colspan="3"></td></tr>`;
  applicaGruppi(tb);
  caricaMiniature(tb);
  // casella del capitolato spuntata se tutte le sue righe sono selezionate
  for (const h of tb.querySelectorAll("tr[data-gruppo]")) {
    const ids = righeDelBlocco(h);
    const cb = h.querySelector("[data-selgruppo]");
    if (cb) { cb.checked = ids.length > 0 && ids.every((id) => sel.has(id)); cb.indeterminate = !cb.checked && ids.some((id) => sel.has(id)); }
  }
}

/* Pagina 2: riepilogo delle righe con sconti e ricarico applicati (come la vista di controllo dell'Excel). */
function disegnaRiepilogoRighe(c) {
  const tb = $("#riepilogo-righe");
  const righe = c.righe.filter((r) => num(r.qta));
  if (!righe.length) { tb.innerHTML = `<tr><td colspan="12" class="vuoto">Nessuna riga: aggiungile nella pagina “Righe”.</td></tr>`; return; }
  const linee = lineeOracle(c), opz = opzCap(c), k = condizioni(c);
  const lisM = k.imp === "listino";
  const ricarico = k.ven.modo === "ricarico" && !lisM;
  $("#rr-ven").textContent = lisM ? "Sc. listino" : ricarico ? "Ricarico" : "Sc. vendita";
  $("#rr-vend").textContent = lisM ? "Listino ripar." : "Vendita";
  let tot = 0, totA = 0, capPrec = null;
  tb.innerHTML = righe.map((r) => {
    const x = calcolaRiga(r, c, opz.ricarichi, k);
    tot += x.totale; totA += x.acq.totale;
    const att = isAttivita(r);
    let testa = "";
    if ((r.capitolato || "") !== capPrec) {
      capPrec = r.capitolato || "";
      testa = `<tr class="gruppo" data-gruppo="${esc(capPrec)}"><td colspan="12">${capPrec ? esc(capPrec) : '<span style="font-weight:500">senza capitolato</span>'}</td></tr>`;
    }
    const scAcq = 1 - x.acq.fattore, fonte = { gma: "dal GMA", generale: "generale", standard: "standard del database", automatico: "automatico del capitolo" };
    return testa + `<tr data-id="${esc(r.id)}"${isAttivita(r) ? ' class="riga-att"' : ""}>
      <td class="num muted small">${linee.get(r.id) ?? ""}</td>
      <td class="mono nowrap">${codiceArticolo(r) ? codiceVis(r) : '<span class="badge bad">senza codice</span>'} ${statoRiga(r)}</td>
      <td class="desc"><b>${esc(r.nOrdine)}</b><div class="d">${esc(r.descrizione)}</div></td>
      <td class="num">${esc(numIt(x.qta))}</td>
      <td class="num">${isGenerico(r) ? `<input class="qta prezzo" data-campo="prezzoGen" inputmode="decimal" value="${esc(numIt(r.listinoAcquisto))}" title="Prezzo del generico: se lo cambi si aggiorna anche nel database con la data di oggi">` : euro(r.listinoAcquisto)}${att ? '<span class="muted small">/g</span>' : ""}</td>
      <td class="num" title="${esc(fonte[x.acq.fonte])}">${att && x.acq.fonte === "standard" ? "" : pctTxt(scAcq) || "0%"}${x.acq.fonte !== "standard" ? ' <span class="fonte">●</span>' : ""}</td>
      <td class="num">${euro(x.acq.unitario)}${att ? '<span class="muted small">/g</span>' : ""}</td>
      <td class="num" title="${esc(fonte[x.fonteVen])}">${lisM ? pctTxt(x.scLis) || "0%" : ricarico ? pctTxt(x.ric) || "0%" : att ? "" : pctTxt(1 - x.fVen) || "0%"}${x.fonteVen === "gma" || x.fonteVen === "generale" ? ' <span class="fonte">●</span>' : ""}</td>
      <td class="num"><b>${euro(x.unitario)}</b></td><td class="num">${euro(x.totale)}</td>
      <td class="num">${x.scontoListino === null || (isGenerico(r) && (r.daOfferta || gmaGenerico(r.gmv, r.fornitore))) ? "" : pctTxt(x.scontoListino) || "0%"}</td>
      <td class="mono">${esc(r.gmv)}</td></tr>`;
  }).join("") + `<tr class="sotto"><td colspan="9" class="num">Totale ${esc(c.nome)} · acquistato ${euro(totA)}</td><td class="num">${euro(tot)}</td><td colspan="2"></td></tr>`;
  applicaGruppi(tb);
}
/* Stato del prezzo di una riga: si guarda l'articolo nel database (data aggiornata), altrimenti la copia nella riga. */
function statoRiga(r) {
  const st = stato(S.perK.get(r.k) || r);
  if (st.livello === "vecchio" || st.livello === "medio") return `<span class="dot ${st.livello}" title="${esc(st.testo)}"></span>`;
  return "";
}
/* Nuovo prezzo di un articolo: aggiorna il database con la data di oggi (e il listino se è calcolato). */
function nuovoPrezzoArticolo(art, prezzo) {
  const auto = !num(art.listino) || Math.abs(num(art.listino) - listinoCalcolato(art)) < 0.005;
  art.listinoAcquisto = prezzo;
  if (auto) art.listino = listinoCalcolato(art);
  art.dataAgg = oggiIso();
}
const rigaDi = (el) => { const c = capAtt(); const id = el.closest("tr[data-id]")?.dataset.id; return { c, id, r: c?.righe.find((x) => x.id === id) }; };
function cambioRiga(e) {
  const inp = e.target;
  const { c, id, r } = rigaDi(inp); if (!r) return;
  if (inp.matches("[data-sel], [data-selgruppo]")) return;
  const campo = inp.dataset.campo; if (!campo) return;
  if (campo === "notePost") { r.notePost = inp.value.trim(); modificata(); return; }
  if (campo === "notePre") { r.notePre = inp.value.trim(); modificata(); return; }
  if (campo === "capitolato") { r.capitolato = inp.value; modificata(); return ridisegnaOfferta(); }
  if (campo === "aGiornate") { r.aGiornate = inp.checked; modificata(); return ridisegnaOfferta(); }
  if (campo === "prezzoGen") {
    const p = num(inp.value);
    if (!p || Math.abs(p - num(r.listinoAcquisto)) < 0.0001) return ridisegnaOfferta();
    nuovoPrezzoArticolo(r, p);
    const art = S.perK.get(r.k);
    if (art) {
      nuovoPrezzoArticolo(art, p);
      salvaCapitoli([art.capitolo, ...propagaClone(art)]).then(() => toast(`Prezzo aggiornato anche nel database (${art.capitolo}) con la data di oggi.`))
        .catch((x) => errore(x, "Aggiornamento del database non riuscito"));
    }
    modificata(); return ridisegnaOfferta();
  }
  r[campo] = num(inp.value);
  // q.tà 0: la riga resta nell'offerta (non va in Oracle e non conta nei totali)
  modificata(); ridisegnaOfferta();
}
$("#righe-offerta").addEventListener("change", cambioRiga);
$("#riepilogo-righe").addEventListener("change", cambioRiga);
/* ----- Menu del tasto destro sulle righe dell'offerta ----- */
function chiudiMenuDx() { $("#menu-dx").hidden = true; }
const TASTO = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent) ? "⌘" : "Ctrl+";
function apriMenuDx(x, y) {
  const c = capAtt(); if (!c) return;
  const n = S.sel.size, m = $("#menu-dx");
  const caps = c.capitolati.map((k) => `<button type="button" data-dx="cap" data-cap-dx="${esc(k)}">${esc(k)}</button>`).join("");
  const selR = c.righe.filter((r) => S.sel.has(r.id));
  const nelDb = selR.length === 1 && !!(S.perK.get(selR[0]._kCatalogo || selR[0].k) || S.catalogo.find((x) => x.k === (selR[0]._kCatalogo || selR[0].k)));
  const convertibili = { gen: selR.filter((r) => isGenerico(r)).length, att: selR.filter((r) => isAttivita(r) && r.daOfferta).length };
  m.innerHTML = `
    <div class="dx-titolo">${n === 1 ? "1 riga selezionata" : `${n} righe selezionate`}</div>
    <button type="button" data-dx="tutte">Seleziona tutte</button>
    <button type="button" data-dx="nessuna">Deseleziona</button>
    <hr>
    <button type="button" data-dx="taglia">✂ Taglia <kbd>${TASTO}X</kbd></button>
    <button type="button" data-dx="copia">⎘ Copia <kbd>${TASTO}C</kbd></button>
    <button type="button" data-dx="incolla" ${S.appunti?.righe.length ? "" : "disabled"}>⇩ Incolla sotto${S.appunti?.righe.length ? ` (${S.appunti.righe.length})` : ""} <kbd>${TASTO}V</kbd></button>
    <hr>
    <button type="button" data-dx="su1">↑ Sposta su</button>
    <button type="button" data-dx="giu1">↓ Sposta giù</button>
    <button type="button" data-dx="su">⤒ Porta in cima</button>
    <button type="button" data-dx="giu">⤓ Porta in fondo</button>
    <hr>
    <button type="button" data-dx="dupqui">⧉ Duplica</button>
    <button type="button" data-dx="dupcon">⧉ Duplica in fondo con capitolato e nota…</button>
    <div class="dx-sub"><button type="button" class="dx-apri">▸ Assegna il capitolato</button>
      <div class="dx-lista">${caps || '<span class="dx-vuoto">nessun capitolato</span>'}<button type="button" data-dx="capnuovo">+ Nuovo capitolato…</button><button type="button" data-dx="cap" data-cap-dx="">— nessun capitolato</button></div></div>
    <button type="button" data-dx="note">✎ Note…</button>
    ${n > 1 ? `<button type="button" data-dx="modmulti">✎ Modifica le ${n} righe selezionate…</button>` : ""}
    ${n === 1 ? (nelDb ? `<button type="button" data-dx="moddb">✎ Modifica nel database</button>` : `<button type="button" data-dx="modriga">✎ Modifica articolo${selR[0]?.daOfferta ? " (da offerta fornitore)" : ""}…</button>`) : ""}
    <button type="button" data-dx="nota">Applica la nota post-linea…</button>
    ${convertibili.gen ? `<button type="button" data-dx="inatt">⇄ Trasforma in attività (PITATT)${convertibili.gen > 1 ? ` · ${convertibili.gen}` : ""}</button>` : ""}
    ${convertibili.att ? `<button type="button" data-dx="ingen">⇄ Trasforma in articolo generico (P000)${convertibili.att > 1 ? ` · ${convertibili.att}` : ""}</button>` : ""}
    <hr>
    <button type="button" data-dx="togli" class="dx-danger">✕ Elimina</button>`;
  m.hidden = false;
  menuDxAperto = Date.now();
  const w = m.offsetWidth, h = m.offsetHeight;
  m.style.left = Math.min(x, innerWidth - w - 8) + "px";
  m.classList.toggle("sx", Math.min(x, innerWidth - w - 8) + w + 230 > innerWidth);
  m.style.top = Math.min(y, innerHeight - h - 8) + "px";
}
$("#righe-offerta").addEventListener("contextmenu", (e) => {
  const tr = e.target.closest("tr[data-id]"); if (!tr) return;
  e.preventDefault();
  if (!S.sel.has(tr.dataset.id)) { S.sel = new Set([tr.dataset.id]); S.ultimoSel = tr.dataset.id; disegnaOfferta(); }
  apriMenuDx(e.clientX, e.clientY);
});
document.addEventListener("click", (e) => { if (!e.target.closest("#menu-dx")) chiudiMenuDx(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape") chiudiMenuDx(); });
let menuDxAperto = 0;
window.addEventListener("scroll", () => { if (Date.now() - menuDxAperto > 400) chiudiMenuDx(); }, true);
$("#menu-dx").addEventListener("click", (e) => {
  const b = e.target.closest("[data-dx]"); if (!b) return;
  const az = b.dataset.dx, c = capAtt();
  chiudiMenuDx();
  if (!c) return;
  const sceltaCap = (sel = "") => `<label class="f">Capitolato<select id="dx-cap"><option value="">— lascia com'è —</option>${c.capitolati.map((k) => `<option ${k === sel ? "selected" : ""}>${esc(k)}</option>`).join("")}</select></label>`;
  if (az === "cap") return azioneRighe("cap", { cap: b.dataset.capDx, vuoto: true });
  if (az === "inatt" || az === "ingen") {
    let n = 0;
    c.righe = c.righe.map((r) => {
      if (!S.sel.has(r.id)) return r;
      if (az === "inatt" && isGenerico(r)) { n++; return convertiRiga(r, true); }
      if (az === "ingen" && isAttivita(r) && r.daOfferta) { n++; return convertiRiga(r, false); }
      return r;
    });
    modificata(); disegnaOfferta();
    return toast(az === "inatt" ? `${n} righe trasformate in attività tecnica (PITATT).` : `${n} righe trasformate in articoli generici (P000 ARTICOLO_GEN).`);
  }
  if (az === "modmulti") return modificaMultipla(c);
  if (az === "modriga") { const r = c.righe.find((x) => S.sel.has(x.id)); return r && modificaRigaOfferta(c, r); }
  if (az === "moddb") { const r = c.righe.find((x) => S.sel.has(x.id)); return r && modificaNelDatabase(r._kCatalogo || r.k); }
  if (az === "capnuovo") return dialogo({
    titolo: "Nuovo capitolato", html: `<label class="f">Nome che va in Oracle (Nome Capitolato)<input id="dx-nuovo" placeholder="es. PIANO TERRA"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Crea e assegna", value: "ok", cls: "primary" }],
    onApri: (d) => setTimeout(() => $("#dx-nuovo", d).focus(), 50),
    onConferma(v, d) { const n = $("#dx-nuovo", d).value.trim(); if (!aggiungiCapitolato(c, n)) return false; azioneRighe("cap", { cap: n }); },
  });
  if (az === "dupcon") return dialogo({
    titolo: "Duplica in fondo", html: `${sceltaCap()}<label class="f" style="margin-top:10px">Nota post-linea delle copie<input id="dx-nota" placeholder="es. posizione A_2"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Duplica", value: "ok", cls: "primary" }],
    onConferma(v, d) { azioneRighe("dup", { cap: $("#dx-cap", d).value, nota: $("#dx-nota", d).value }); },
  });
  if (az === "nota") return dialogo({
    titolo: "Nota post-linea", html: `<label class="f">Nota per le righe selezionate (vuoto = cancella)<input id="dx-nota" placeholder="es. posizione A_2"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Applica", value: "ok", cls: "primary" }],
    onApri: (d) => setTimeout(() => $("#dx-nota", d).focus(), 50),
    onConferma(v, d) { azioneRighe("nota", { nota: $("#dx-nota", d).value }); },
  });
  azioneRighe(az);
});

/* righe del blocco sotto un'intestazione di capitolato (fino alla successiva) */
function righeDelBlocco(h) {
  const ids = [];
  for (let tr = h.nextElementSibling; tr && !tr.hasAttribute("data-gruppo") && !tr.classList.contains("sotto"); tr = tr.nextElementSibling) if (tr.dataset.id) ids.push(tr.dataset.id);
  return ids;
}
/* Selezione: casella, clic su codice/descrizione, Maiusc+clic per un intervallo, casella del capitolato. */
$("#righe-offerta").addEventListener("click", (e) => {
  const c = capAtt(); if (!c) return;
  const ed = e.target.closest("[data-rinomina]");
  if (ed) return rinominaCapitolato(c, ed.dataset.rinomina);
  const sp = e.target.closest("[data-capsposta]");
  if (sp) return spostaCapitolato(c, sp.closest("tr").dataset.gruppo, +sp.dataset.capsposta);
  const g = e.target.closest("[data-selgruppo]");
  if (g) {
    const ids = righeDelBlocco(g.closest("tr"));
    ids.forEach((id) => (g.checked ? S.sel.add(id) : S.sel.delete(id)));
    return disegnaOfferta();
  }
  const tr = e.target.closest("tr[data-id]"); if (!tr) return;
  if (appenaTrascinato()) return;
  const suCasella = e.target.matches("[data-sel]");
  const suTesto = !e.target.closest("input, select, button, label") && e.target.closest("td.desc, td.mono, td.num:nth-child(2)");
  if (!suCasella && !suTesto) return;
  const id = tr.dataset.id, ctrl = e.ctrlKey || e.metaKey;
  // clic semplice: solo questa riga; Ctrl/⌘: aggiunge o toglie; Maiusc: intervallo dall'ultima riga cliccata
  if (e.shiftKey && S.ultimoSel && S.ultimoSel !== id) {
    const a = c.righe.findIndex((r) => r.id === S.ultimoSel), b = c.righe.findIndex((r) => r.id === id);
    if (a >= 0 && b >= 0) { if (!ctrl) S.sel = new Set(); const [x, y] = a < b ? [a, b] : [b, a]; for (let k = x; k <= y; k++) S.sel.add(c.righe[k].id); }
    window.getSelection?.().removeAllRanges();
    return disegnaOfferta();
  } else if (ctrl || suCasella) {
    // la casella di selezione aggiunge o toglie la riga, come Ctrl+clic
    S.sel.has(id) ? S.sel.delete(id) : S.sel.add(id);
  } else {
    S.sel = S.sel.size === 1 && S.sel.has(id) ? new Set() : new Set([id]);
  }
  S.ultimoSel = id;
  disegnaOfferta();
});

/* Note di più righe insieme: un campo vuoto lascia le note com'erano. */
function noteRigheSelezionate(scelte) {
  const modificabili = scelte.filter((r) => !isGenerico(r));
  if (!modificabili.length) return toast("Le note degli articoli generici sono automatiche (n° di ordinazione e descrizione).");
  if (modificabili.length === 1) return modificaNoteRiga(modificabili[0]);
  dialogo({
    titolo: `Note di ${modificabili.length} righe`,
    html: `<p class="muted small" style="margin-top:0">Quello che scrivi va su tutte le righe selezionate (gli articoli generici restano automatici).
      Lascia vuoto un campo per non cambiarlo; spunta "svuota" per cancellarlo.</p>
      <label class="f">Note pre-linea<textarea id="n-pre" rows="3"></textarea></label>
      <label class="row small"><input type="checkbox" id="n-pre-x"> svuota le note pre-linea</label>
      <label class="f" style="margin-top:10px">Note post-linea<textarea id="n-post" rows="3"></textarea></label>
      <label class="row small"><input type="checkbox" id="n-post-x"> svuota le note post-linea</label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Applica", value: "ok", cls: "primary" }],
    onConferma(v, b) {
      const pre = $("#n-pre", b).value.trim(), post = $("#n-post", b).value.trim();
      for (const r of modificabili) {
        if ($("#n-pre-x", b).checked) r.notePre = ""; else if (pre) r.notePre = pre;
        if ($("#n-post-x", b).checked) r.notePost = ""; else if (post) r.notePost = post;
      }
      modificata(); disegnaOfferta();
    },
  });
}
function modificaNoteRiga(r) {
  dialogo({
    titolo: `Note · ${codiceArticolo(r)} ${r.nOrdine}`,
    html: `<label class="f">Note pre-linea<textarea id="n-pre" rows="4">${esc(r.notePre)}</textarea></label>
      <label class="f" style="margin-top:10px">Note post-linea<textarea id="n-post" rows="4">${esc(r.notePost)}</textarea></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onConferma(v, b) {
      r.notePre = $("#n-pre", b).value.trim(); r.notePost = $("#n-post", b).value.trim();
      modificata(); disegnaOfferta();
    },
  });
}
$("#sel-tutte").addEventListener("change", (e) => {
  const c = capAtt(); if (!c) return;
  S.sel = e.target.checked ? new Set(c.righe.map((r) => r.id)) : new Set();
  disegnaOfferta();
});
/* Azioni sulle righe selezionate (menu del tasto destro). opz = { cap, nota } */
function azioneRighe(az, opz = {}) {
  const c = capAtt(); if (!c) return;
  const scelte = c.righe.filter((r) => S.sel.has(r.id));
  const nota = (opz.nota || "").trim();
  const capSc = opz.cap || "";
  if (az === "copia" || az === "taglia") {
    if (!scelte.length) return toast("Seleziona prima le righe.");
    S.appunti = { righe: scelte.map((r) => structuredClone(r)), taglio: az === "taglia", offId: S.off.id, ids: new Set(scelte.map((r) => r.id)) };
    try { navigator.clipboard?.writeText(scelte.map((r) => [codiceArticolo(r), r.descrizione || "", r.qta].join("\t")).join("\n")).catch(() => {}); } catch {}
    toast(az === "taglia"
      ? `Tagliate ${scelte.length} righe: seleziona la riga sotto cui spostarle e premi Incolla (${TASTO}V). Esc annulla.`
      : `Copiate ${scelte.length} righe: seleziona la riga sotto cui incollarle e premi Incolla (${TASTO}V).`, 4500);
    return disegnaOfferta();
  } else if (az === "incolla") {
    const ap = S.appunti;
    if (!ap?.righe.length) return toast("Non hai copiato o tagliato nessuna riga.");
    const sposta = ap.taglio && ap.offId === S.off.id;
    // la destinazione è l'ultima riga selezionata (le righe tagliate non contano); senza selezione si incolla in fondo
    const dest = scelte.filter((r) => !(sposta && ap.ids.has(r.id))).at(-1);
    if (sposta) for (const k of S.off.capitoli) k.righe = k.righe.filter((r) => !ap.ids.has(r.id));
    const nuove = ap.righe.map((r) => ({ ...structuredClone(r), id: sposta ? r.id : db.nuovoId(),
      ...(dest ? { capitolato: dest.capitolato || "" } : {}) }));
    const i = dest ? c.righe.indexOf(dest) + 1 : c.righe.length;
    c.righe.splice(i, 0, ...nuove);
    S.sel = new Set(nuove.map((r) => r.id));
    if (sposta) S.appunti = null; // dopo uno spostamento gli appunti si svuotano, come in Excel
    toast(`${sposta ? "Spostate" : "Incollate"} ${nuove.length} righe${dest ? (dest.capitolato ? ` nel capitolato “${dest.capitolato}”` : " sotto la riga selezionata") : " in fondo"}.`
      + (ap.taglio && !sposta ? " Le righe originali restano nell'altra offerta." : ""), 4000);
  } else if (az === "dup") {
    // le copie vanno in fondo, nello stesso ordine, con il capitolato e la nota scelti (se indicati)
    const copie = scelte.map((r) => ({ ...structuredClone(r), id: db.nuovoId(),
      ...(capSc ? { capitolato: capSc } : {}), ...(nota && !isGenerico(r) ? { notePost: nota } : {}) }));
    c.righe.push(...copie);
    S.sel = new Set(copie.map((r) => r.id));
    toast(`Duplicate ${copie.length} righe${capSc ? ` nel capitolato “${capSc}”` : ""}${nota ? ` con nota “${nota}”` : ""}: le trovi in fondo, già selezionate.`, 4500);
  } else if (az === "cap") {
    if (!capSc && !opz.vuoto) return toast("Scegli il capitolato.");
    scelte.forEach((r) => (r.capitolato = capSc));
  } else if (az === "nota") {
    scelte.filter((r) => !isGenerico(r)).forEach((r) => (r.notePost = nota));
  } else if (az === "su" || az === "giu") {
    const resto = c.righe.filter((r) => !S.sel.has(r.id));
    c.righe = az === "su" ? [...scelte, ...resto] : [...resto, ...scelte];
  } else if (az === "su1") {
    // ogni riga selezionata sale di una posizione (il gruppo selezionato si muove insieme)
    for (let k = 1; k < c.righe.length; k++) if (S.sel.has(c.righe[k].id) && !S.sel.has(c.righe[k - 1].id)) [c.righe[k - 1], c.righe[k]] = [c.righe[k], c.righe[k - 1]];
  } else if (az === "giu1") {
    for (let k = c.righe.length - 2; k >= 0; k--) if (S.sel.has(c.righe[k].id) && !S.sel.has(c.righe[k + 1].id)) [c.righe[k + 1], c.righe[k]] = [c.righe[k], c.righe[k + 1]];
  } else if (az === "dupqui") {
    // le copie vanno subito sotto l'ultima riga selezionata, nello stesso ordine
    const copie = scelte.map((r) => ({ ...structuredClone(r), id: db.nuovoId() }));
    const ultimo = Math.max(...scelte.map((r) => c.righe.indexOf(r)));
    c.righe.splice(ultimo + 1, 0, ...copie);
    S.sel = new Set(copie.map((r) => r.id));
    toast(`Duplicate ${copie.length} righe: le copie sono sotto, già selezionate.`);
  } else if (az === "note") {
    return noteRigheSelezionate(scelte);
  } else if (az === "tutte") {
    S.sel = new Set(c.righe.map((r) => r.id));
  } else if (az === "togli") {
    if (scelte.length > 1 && !opz.confermato) {
      return dialogo({ titolo: `Eliminare ${scelte.length} righe?`, html: `<p>Le ${scelte.length} righe selezionate verranno tolte dall'offerta.</p>`,
        pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Elimina", value: "ok", cls: "danger" }],
        onConferma: () => azioneRighe("togli", { confermato: true }) });
    }
    c.righe = c.righe.filter((r) => !S.sel.has(r.id));
    S.sel = new Set();
  } else if (az === "nessuna") S.sel = new Set();
  modificata(); disegnaOfferta();
}
$("#aggiorna-da-catalogo").addEventListener("click", () => {
  const c = capAtt(); if (!c) return;
  let n = 0, mancanti = 0;
  c.righe = c.righe.map((r) => {
    const a = S.perK.get(r.k) || S.catalogo.find((x) => x.capitolo === r.capitolo && codiceArticolo(x) === codiceArticolo(r) && x.nOrdine === r.nOrdine);
    if (!a) { mancanti++; return r; }
    n++;
    const nuovo = rigaDaArticolo(a);
    // restano quelle della riga: quantità, giorni dell'attività, note, capitolato e "a giornate"
    const out = { ...nuovo, id: r.id, qta: r.qta, giorni: isAttivita(r) ? r.giorni : a.giorni, notePre: r.notePre, notePost: r.notePost, capitolato: r.capitolato || "" };
    if (r.aGiornate) out.aGiornate = r.aGiornate; else delete out.aGiornate;
    return out;
  });
  modificata(); disegnaOfferta();
  toast(`Prezzi e sconti aggiornati su ${n} righe${mancanti ? `; ${mancanti} righe non sono più nel database e restano com'erano` : ""}.`, 5000);
});

/* Copia dei dati dell'articolo dentro la riga: l'offerta resta uguale anche se il database cambia. */
function rigaDaArticolo(a) {
  const { _chunk, _qta, _kCatalogo, _blocco0, ...r } = a;
  if (_kCatalogo) r.k = _kCatalogo;
  // la descrizione completa serve solo ai generici (diventa la nota post-linea); per gli altri basta un estratto
  if (!isGenerico(r)) r.descrizione = (r.descrizione || "").slice(0, 400);
  r.listino = prezzoListino(a);
  return r;
}

/* --- Aggiungi articoli --- */
function disegnaPicker(c, reset = false) {
  const dbs = ordineCap();
  // il database proposto è quello del capitolo; se ne scegli un altro resta finché non cambi capitolo
  if (S.pickCap !== c.nome || !dbs.includes(S.pickDb)) { S.pickCap = c.nome; S.pickDb = dbs.includes(dbNome(c)) ? dbNome(c) : dbs[0] || ""; reset = true; }
  $("#a-db").innerHTML = dbs.map((d) => `<option value="${esc(d)}" ${d === S.pickDb ? "selected" : ""}>${esc(d)}${d === dbNome(c) ? " (questo capitolo)" : ""}</option>`).join("");
  $("#a-cap").innerHTML = `<option value="">— nessuno —</option>` + c.capitolati.map((n) => `<option ${n === c.capNuove ? "selected" : ""}>${esc(n)}</option>`).join("") +
    `<option value="__nuovo">+ nuovo capitolato…</option>`;
  const nelDb = S.catalogo.filter((a) => a.capitolo === S.pickDb);
  const sotti = [...new Set(nelDb.map((a) => a.sottocapitolo).filter(Boolean))];
  const sottoSel = $("#a-sotto");
  const prev = reset ? "" : sottoSel.value;
  sottoSel.innerHTML = `<option value="">Tutti i sottocapitoli</option>` + sotti.map((s) => `<option ${s === prev ? "selected" : ""}>${esc(s)}</option>`).join("");
  const q = $("#a-cerca").value.trim().toLowerCase();
  const solo = false;
  const conta = contaInOfferta(c);
  const lista = nelDb.filter((a) => (!sottoSel.value || a.sottocapitolo === sottoSel.value) &&
    (!solo || conta.has(a.k)) &&
    (!q || `${codiceArticolo(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)));
  const vis = lista.slice(0, S.limitePick);
  let sotto = null, html = "";
  for (const a of vis) {
    if ((a.sottocapitolo || "") !== sotto && !sottoSel.value) {
      sotto = a.sottocapitolo || "";
      html += `<tr class="sotto" data-gruppo="${esc(sotto)}"><td colspan="9"><input type="checkbox" class="sel-gruppo" data-pickgruppo title="Seleziona tutti gli articoli di questo sottocapitolo"> ${esc(sotto || "—")}</td></tr>`;
    }
    const att = isAttivita(a);
    const pSel = S.pickSel.has(a.k);
    html += `<tr data-k="${esc(a.k)}" class="${pSel ? "sel" : ""}">
      <td class="sel-col"><input type="checkbox" data-pick-sel ${pSel ? "checked" : ""}></td>
      <td class="mono nowrap">${codiceVis(a)}</td>
      <td class="img-cel">${a.img ? `<img class="img-mini" data-img="${esc(a.img)}" alt="">` : ""}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b><div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td>
      <td class="mono">${esc(a.gmv)}</td>
      <td>${pallino(a)}</td>
      <td class="num nowrap">${att ? `<input class="qta" data-pick="giorni" inputmode="decimal" title="Giorni" value="${esc(numIt(a.giorni))}"> gg ` : ""}<input class="qta" data-pick="qta" inputmode="decimal" placeholder="q.tà"></td></tr>`;
  }
  $("#righe-picker").innerHTML = html || `<tr><td colspan="9" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Il database è vuoto: importalo dal file Excel."}</td></tr>`;
  applicaGruppi($("#righe-picker"));
  caricaMiniature($("#righe-picker"));
  S.pickVis = vis.map((a) => a.k);
  aggiornaPickSel();
  // articoli della biblioteca (listini fornitore) che non sono nel database
  if (q.length >= 2) {
    caricaBib();
    const tb = $("#righe-picker");
    if (!S.bibPronta) tb.insertAdjacentHTML("beforeend", `<tr class="sotto"><td colspan="8" class="muted">Cerco anche nella biblioteca…</td></tr>`);
    else {
      const trovati = cercaBib(q, { escludiDb: true }).slice(0, 40);
      if (trovati.length) tb.insertAdjacentHTML("beforeend", `<tr class="sotto bib-testa" data-gruppo="__biblioteca"><td colspan="8">Dalla biblioteca (listini fornitore, non nel database)</td></tr>` +
        trovati.map((a) => rigaBibHtml(a, "picker")).join(""));
      applicaGruppi(tb);
    }
  }
  $("#picker-info").textContent = lista.length > vis.length ? `Mostrati ${vis.length} di ${lista.length}: affina la ricerca.` : "";
}
/* --- Selezione di più articoli del database e aggiunta in un colpo solo --- */
function aggiornaPickSel() {
  const n = S.pickSel.size, tb = $("#righe-picker");
  $("#pick-bar").hidden = !n;
  $("#pick-n").textContent = n === 1 ? "1 articolo selezionato" : `${n} articoli selezionati`;
  for (const tr of tb.querySelectorAll("tr[data-k]")) {
    const on = S.pickSel.has(tr.dataset.k);
    tr.classList.toggle("sel", on);
    const cb = tr.querySelector("[data-pick-sel]"); if (cb) cb.checked = on;
  }
  for (const h of tb.querySelectorAll("tr[data-gruppo]")) {
    const ks = kDelBloccoPick(h), cb = h.querySelector("[data-pickgruppo]"); if (!cb) continue;
    const quanti = ks.filter((k) => S.pickSel.has(k)).length;
    cb.checked = ks.length > 0 && quanti === ks.length; cb.indeterminate = quanti > 0 && quanti < ks.length;
  }
  const vis = S.pickVis || [], qv = vis.filter((k) => S.pickSel.has(k)).length;
  const t = $("#pick-tutti"); t.checked = vis.length > 0 && qv === vis.length; t.indeterminate = qv > 0 && qv < vis.length;
}
function kDelBloccoPick(h) {
  const ks = [];
  for (let tr = h.nextElementSibling; tr && !tr.hasAttribute("data-gruppo"); tr = tr.nextElementSibling) if (tr.dataset.k) ks.push(tr.dataset.k);
  return ks;
}
/* Clic e trascina sulle righe: seleziona tutte le righe tra quella di partenza e quella sotto il mouse (come in Excel).
   Con Ctrl/⌘ tenuto all'inizio, l'intervallo si aggiunge a quello che era già selezionato. */
const trascina = { attivo: false, mosso: false };
function avviaTrascina(e, tabella) {
  if (e.button !== 0 || e.shiftKey || e.target.closest("input, select, textarea, button, label, a")) return;
  const tr = e.target.closest(tabella === "off" ? "tr[data-id]" : "tr[data-k]"); if (!tr) return;
  const sel = tabella === "off" ? S.sel : tabella === "cat" ? S.catSel : S.pickSel;
  Object.assign(trascina, { attivo: true, mosso: false, tabella, da: tr.dataset[tabella === "off" ? "id" : "k"], base: e.ctrlKey || e.metaKey ? new Set(sel) : new Set() });
}
function seguiTrascina(e) {
  if (!trascina.attivo) return;
  const off = trascina.tabella === "off", cat = trascina.tabella === "cat";
  const tr = e.target.closest(off ? "#righe-offerta tr[data-id]" : cat ? "#righe-catalogo tr[data-k]" : "#righe-picker tr[data-k]"); if (!tr) return;
  const a = tr.dataset[off ? "id" : "k"];
  if (a === trascina.da && !trascina.mosso) return;
  trascina.mosso = true;
  const ordine = off ? (capAtt()?.righe || []).map((r) => r.id) : cat ? S.catVisK || [] : S.pickVis || [];
  const i = ordine.indexOf(trascina.da), j = ordine.indexOf(a); if (i < 0 || j < 0) return;
  const nuova = new Set(trascina.base);
  for (let k = Math.min(i, j); k <= Math.max(i, j); k++) nuova.add(ordine[k]);
  if (off) {
    S.sel = nuova; S.ultimoSel = a;
    for (const r of $$("#righe-offerta tr[data-id]")) { const on = nuova.has(r.dataset.id); r.classList.toggle("sel", on); const cb = r.querySelector("[data-sel]"); if (cb) cb.checked = on; }
  } else if (cat) { S.catSel = nuova; S.catUltimo = a; aggiornaSelCat(); }
  else { S.pickSel = nuova; S.pickUltimo = a; aggiornaPickSel(); }
}
$("#righe-offerta").addEventListener("mousedown", (e) => avviaTrascina(e, "off"));
$("#righe-picker").addEventListener("mousedown", (e) => avviaTrascina(e, "pick"));
document.addEventListener("mouseover", seguiTrascina);
document.addEventListener("mouseup", () => {
  if (!trascina.attivo) return;
  trascina.attivo = false;
  if (trascina.mosso) { trascina.finito = Date.now(); if (trascina.tabella === "off") disegnaOfferta(); }
});
const appenaTrascinato = () => trascina.finito && Date.now() - trascina.finito < 300;

/* Maiusc+clic o Ctrl+clic sulle righe: niente selezione del testo (il browser la mostrerebbe per un attimo). */
for (const id of ["#righe-offerta", "#righe-picker"]) $(id).addEventListener("mousedown", (e) => {
  if ((e.shiftKey || e.ctrlKey || e.metaKey) && e.target.closest("tr[data-id], tr[data-k]") && !e.target.closest("input:not([type=checkbox]), select, textarea")) e.preventDefault();
});
$("#righe-picker").addEventListener("click", (e) => {
  const g = e.target.closest("[data-pickgruppo]");
  if (g) { kDelBloccoPick(g.closest("tr")).forEach((k) => (g.checked ? S.pickSel.add(k) : S.pickSel.delete(k))); return aggiornaPickSel(); }
  const tr = e.target.closest("tr[data-k]"); if (!tr) return;
  if (appenaTrascinato()) return;
  const suCasella = e.target.matches("[data-pick-sel]");
  const suTesto = !e.target.closest("input, select, button, label") && e.target.closest("td.desc, td.mono");
  if (!suCasella && !suTesto) return;
  const k = tr.dataset.k, vis = S.pickVis || [], ctrl = e.ctrlKey || e.metaKey;
  if (e.shiftKey && S.pickUltimo && S.pickUltimo !== k && vis.includes(S.pickUltimo)) {
    const a = vis.indexOf(S.pickUltimo), b = vis.indexOf(k);
    const [x, y] = a < b ? [a, b] : [b, a];
    if (!ctrl) S.pickSel.clear();
    for (let i = x; i <= y; i++) S.pickSel.add(vis[i]);
    window.getSelection?.().removeAllRanges();
    return aggiornaPickSel();
  } else if (ctrl || suCasella) S.pickSel.has(k) ? S.pickSel.delete(k) : S.pickSel.add(k);
  else { const solo = S.pickSel.size === 1 && S.pickSel.has(k); S.pickSel.clear(); if (!solo) S.pickSel.add(k); }
  S.pickUltimo = k;
  aggiornaPickSel();
});
$("#pick-tutti").addEventListener("change", (e) => { (S.pickVis || []).forEach((k) => (e.target.checked ? S.pickSel.add(k) : S.pickSel.delete(k))); aggiornaPickSel(); });
$("#pick-nessuno").addEventListener("click", () => { S.pickSel.clear(); aggiornaPickSel(); });
$("#pick-aggiungi").addEventListener("click", () => aggiungiSelezionati());
$("#righe-picker").addEventListener("contextmenu", (e) => {
  const tr = e.target.closest("tr[data-k]"); if (!tr || e.target.closest("input:not([type=checkbox])")) return;
  e.preventDefault();
  if (!S.pickSel.has(tr.dataset.k)) { S.pickSel.add(tr.dataset.k); S.pickUltimo = tr.dataset.k; aggiornaPickSel(); }
  S.pickDx = tr.dataset.k;
  const n = S.pickSel.size, m = $("#menu-dx");
  m.innerHTML = `<div class="dx-titolo">${n === 1 ? "1 articolo selezionato" : `${n} articoli selezionati`}</div>
    <button type="button" data-dxp="aggiungi"><b>+ Aggiungi all'offerta…</b></button>
    <hr>
    <button type="button" data-dxp="modifica">✎ Modifica nel database <span class="muted small">(${esc(S.perK.get(tr.dataset.k)?.nOrdine || "")})</span></button>
    <hr>
    <button type="button" data-dxp="tutti">Seleziona tutti quelli visibili</button>
    <button type="button" data-dxp="nessuno">Deseleziona tutti</button>`;
  m.hidden = false; menuDxAperto = Date.now();
  const w = m.offsetWidth, h = m.offsetHeight;
  m.style.left = Math.min(e.clientX, innerWidth - w - 8) + "px";
  m.style.top = Math.min(e.clientY, innerHeight - h - 8) + "px";
});
$("#menu-dx").addEventListener("click", (e) => {
  const b = e.target.closest("[data-dxp]"); if (!b) return;
  chiudiMenuDx();
  const az = b.dataset.dxp;
  if (az === "aggiungi") return aggiungiSelezionati();
  if (az === "modifica") return modificaNelDatabase(S.pickDx);
  if (az === "tutti") (S.pickVis || []).forEach((k) => S.pickSel.add(k));
  if (az === "nessuno") S.pickSel.clear();
  aggiornaPickSel();
});
/* Dal tasto destro: apre il catalogo su quell'articolo, con il modulo di modifica; poi "Torna all'offerta" */
/* Modifica di un campo comune a tutte le righe selezionate (vuoto = resta com'è). Vale solo per le righe
   dell'offerta: il database non cambia. */
function modificaMultipla(c) {
  const righe = c.righe.filter((r) => S.sel.has(r.id));
  if (!righe.length) return;
  const nAtt = righe.filter(isAttivita).length;
  dialogo({
    titolo: `Modifica ${righe.length} righe`, largo: true,
    html: `<p class="muted small" style="margin-top:0">Compila solo i campi da cambiare: quelli vuoti restano come sono su ogni riga.
        La modifica vale per le righe dell'offerta, il database non cambia.</p>
      <div class="grid-set">
        <label class="f">Quantità<input id="mm-qta" inputmode="decimal" placeholder="—"></label>
        <label class="f">U.M.<input id="mm-um" placeholder="—"></label>
        <label class="f">Sottocapitolo<select id="mm-cap"><option value="__no">— non cambiare —</option><option value="">(nessuno)</option>${c.capitolati.map((k) => `<option>${esc(k)}</option>`).join("")}</select></label>
        <label class="f">Fornitore<select id="mm-forn">${opzioniFornitoriGma("")}</select></label>
        <label class="f">GMA<select id="mm-gmv"></select></label>
        <label class="f">Listino acquisto € (stesso prezzo per tutte)<input id="mm-la" inputmode="decimal" placeholder="—"></label>
        <label class="f">Variazione prezzo acquisto %<input id="mm-var" inputmode="decimal" placeholder="es. 5 oppure -3"></label>
        ${nAtt ? `<label class="f">Giorni (attività)<input id="mm-gg" inputmode="decimal" placeholder="—"></label>` : ""}
      </div>
      <div class="grid-set" style="margin-top:10px">
        <label class="f wide">Nota pre-linea <span class="muted small">(non per i generici)</span><input id="mm-npre" placeholder="—"></label>
        <label class="f wide">Nota post-linea <span class="muted small">(non per i generici)</span><input id="mm-npost" placeholder="—"></label>
      </div>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: `Applica a ${righe.length} righe`, value: "ok", cls: "primary" }],
    onApri(d) {
      const gma = (f) => { $("#mm-gmv", d).innerHTML = `<option value="__no">— non cambiare —</option>` + opzioniGma(f, "").replace(/<option value="">.*?<\/option>/, "").replace(/<option value="__nuovo">.*?<\/option>/, ""); };
      $("#mm-forn", d).addEventListener("change", (e) => gma(e.target.value));
      gma("");
    },
    onConferma(v, d) {
      const val = (id) => ($(id, d) ? $(id, d).value.trim() : "");
      const q = val("#mm-qta"), um = val("#mm-um"), cap = $("#mm-cap", d).value, g = $("#mm-gmv", d).value, gf = g !== "__no" ? fornScelto($("#mm-gmv", d), $("#mm-forn", d)) : "";
      const la = val("#mm-la"), vr = val("#mm-var"), gg = val("#mm-gg"), npre = val("#mm-npre"), npost = val("#mm-npost");
      if (la && vr) { toast("Scegli il prezzo oppure la variazione %, non tutti e due."); return false; }
      let n = 0;
      for (const r of righe) {
        if (q) r.qta = num(q);
        if (um) r.um = um;
        if (cap !== "__no") { r.capitolato = cap; }
        if (g !== "__no") { r.gmv = g; r.fornitore = gf; }
        if (la) r.listinoAcquisto = num(la);
        if (vr) r.listinoAcquisto = Math.round(num(r.listinoAcquisto) * (1 + num(vr) / 100) * 1e4) / 1e4;
        if (gg && isAttivita(r)) r.giorni = num(gg);
        if (!isGenerico(r)) { if (npre) r.notePre = npre; if (npost) r.notePost = npost; }
        if (la || vr || gg) r.listino = listinoCalcolato(r);
        n++;
      }
      if (cap !== "__no") raggruppa(c);
      modificata(); disegnaOfferta();
      toast(`${n} righe modificate.`);
    },
  });
}
/* Righe che non sono nel database (es. importate da un'offerta fornitore): si modificano nella riga stessa,
   e volendo si salvano anche nel database come articolo nuovo. */
function modificaRigaOfferta(c, r) {
  const att = isAttivita(r);
  const codF = att ? (r.notePre || r.nOrdine || "") : (r.nOrdine || "");
  const desF = att ? (r.notePost || r.descrizione || "") : (r.descrizione || "");
  const capDb = S.capitoli.some((x) => x.nome === dbNome(c)) ? dbNome(c) : (S.capitoli.find((x) => /GENERIC/i.test(x.nome))?.nome || S.capitoli[0]?.nome || "");
  dialogo({
    titolo: `${att ? "Attività" : "Articolo generico"} · ${codF || "senza codice"}`, largo: true,
    html: `<p class="muted small" style="margin-top:0">${r.daOfferta ? `Riga importata dall'offerta fornitore <b>${esc(r.daOfferta)}</b>: non è nel database. ` : ""}
        Le modifiche valgono per questa riga dell'offerta; spunta “Salva anche nel database” per ritrovarlo nel catalogo.</p>
      <div class="grid-set">
        <label class="f">Codice fornitore (n° di ordinazione)<input id="mr-cod" value="${esc(codF)}"></label>
        <label class="f">U.M.<input id="mr-um" value="${esc(r.um || "n.")}"></label>
        <label class="f">Fornitore<select id="mr-forn">${opzioniFornitoriGma(fornitoreDi(r))}</select></label>
        <label class="f">GMA<select id="mr-gmv">${opzioniGma(fornitoreDi(r), r.gmv)}</select></label>
        ${campoNum("mr-la", att ? "Valore attività €" : "Listino acquisto € (prezzo del fornitore)", r.listinoAcquisto)}
        ${att ? campoNum("mr-gg", "Giorni / quantità del fornitore", r.giorni) : ""}
      </div>
      <label class="f" style="margin-top:10px">Descrizione<textarea id="mr-des" rows="4">${esc(desF)}</textarea></label>
      <div class="card" style="margin-top:10px">
        <label class="row"><input type="checkbox" id="mr-db"> <b>Salva anche nel database</b> come ${att ? "attività tecnica PITATT.TEC" : "articolo generico P000 ARTICOLO_GEN"}</label>
        <div class="grid-set" id="mr-db-box" hidden style="margin-top:6px">
          <label class="f">Capitolo del database<select id="mr-cap">${S.capitoli.map((x) => `<option ${x.nome === capDb ? "selected" : ""}>${esc(x.nome)}</option>`).join("")}</select></label>
          <label class="f wide">Sottocapitolo<input id="mr-sotto" value="${esc(r.capitolato || "")}"></label>
        </div></div>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(d) {
      $("#mr-forn", d).addEventListener("change", (e) => { const g = $("#mr-gmv", d).value, f = e.target.value; $("#mr-gmv", d).innerHTML = opzioniGma(f, !f || voceGma(g, f) ? g : "").replace(/<option value="__nuovo">.*?<\/option>/, ""); });
      $("#mr-gmv", d).innerHTML = $("#mr-gmv", d).innerHTML.replace(/<option value="__nuovo">.*?<\/option>/, "");
      $("#mr-db", d).addEventListener("change", (e) => { $("#mr-db-box", d).hidden = !e.target.checked; });
    },
    async onConferma(v, d) {
      const cod = $("#mr-cod", d).value.trim(), des = $("#mr-des", d).value.trim();
      const x = { codice: cod, descrizione: des, um: $("#mr-um", d).value.trim() || "n.", qta: att ? num($("#mr-gg", d).value) : num(r.qta), unitario: num($("#mr-la", d).value) };
      const n = { ...rigaDaOffertaForn(x, att, r.daOfferta || "", r.capitolo), id: r.id, capitolato: r.capitolato || "", qta: att ? 1 : r.qta };
      n.gmv = $("#mr-gmv", d).value || (att ? "ATT" : "ARTGEN");
      n.fornitore = fornScelto($("#mr-gmv", d), $("#mr-forn", d));
      if (att) { n.notePre = cod; n.notePost = des; n.aGiornate = !!r.aGiornate; }
      if ($("#mr-db", d).checked) {
        const art = { ...n, capitolo: $("#mr-cap", d).value, sottocapitolo: $("#mr-sotto", d).value.trim(), rifCap: "", dataAgg: oggiIso() };
        delete art.id; delete art.qta; delete art.capitolato; delete art.daOfferta;
        if (att) { art.nOrdine = cod; art.descrizione = des; art.notePre = ""; art.notePost = ""; }
        const stessi = S.catalogo.filter((y) => y.capitolo === art.capitolo);
        art.ordine = (stessi.length ? Math.max(...stessi.map((y) => num(y.ordine))) : 0) + 1;
        art.k = `${art.capitolo}||${codiceArticolo(art)}|${art.nOrdine}|${db.nuovoId()}`;
        try { await scriviNelDatabase(art); } catch (e) { errore(e, "Salvataggio nel database non riuscito"); return false; }
        n.k = art.k; delete n.daOfferta;
      }
      Object.assign(r, n);
      modificata(); disegnaOfferta();
      toast($("#mr-db", d).checked ? "Riga modificata e articolo salvato nel database." : "Riga modificata.");
    },
  });
}

function modificaNelDatabase(k) {
  const a = S.perK.get(k) || S.catalogo.find((x) => x.k === k);
  if (!a) return toast("Questo articolo non è più nel database.");
  S.ritorno = S.off ? { offId: S.off.id, dati: S.off, capAtt: S.capAtt, sub: S.sub } : null;
  vai("catalogo");
  $("#c-capitolo").value = a.capitolo; $("#c-sotto").value = ""; $("#c-cerca").value = a.nOrdine || codiceArticolo(a);
  disegnaCatalogo();
  modificaArticolo(a);
}
function aggiornaRitorno() {
  const b = $("#c-torna"); if (!b) return;
  b.hidden = !S.ritorno;
  if (S.ritorno) b.textContent = `← Torna all'offerta`;
}
$("#c-torna").addEventListener("click", () => {
  const r = S.ritorno; S.ritorno = null; aggiornaRitorno();
  if (r) {
    apriOfferta(r.offId, false, S.offerte.find((o) => o.id === r.offId) || r.dati);
    if (S.off?.capitoli.some((c) => c.nome === r.capAtt)) { S.capAtt = r.capAtt; S.sub = r.sub; disegnaOfferta(true); }
    toast("Le righe già nell'offerta tengono i prezzi di prima: per riallinearle usa “Aggiorna prezzi dal database”.", 5000);
  }
});

function aggiungiSelezionati() {
  const c = capAtt(); if (!c) return;
  const arts = S.catalogo.filter((a) => S.pickSel.has(a.k)); // nell'ordine del database
  if (!arts.length) return toast("Seleziona prima gli articoli.");
  const conAtt = arts.some(isAttivita), conGen = arts.some(isGenerico);
  const opzCapit = `<option value="">— nessuno —</option>` + c.capitolati.map((n) => `<option ${n === c.capNuove ? "selected" : ""}>${esc(n)}</option>`).join("") + `<option value="__nuovo">+ nuovo capitolato…</option>`;
  dialogo({
    titolo: arts.length === 1 ? "Aggiungi 1 articolo" : `Aggiungi ${arts.length} articoli`,
    html: `<p class="muted small" style="margin-top:0">Quantità, capitolato e note valgono per tutte le righe che aggiungi. Le righe vanno in fondo, nell'ordine del database.</p>
      <div class="row">
        <label class="f">Quantità (per ogni articolo)<input id="ps-qta" inputmode="decimal" placeholder="es. 2" style="width:140px"></label>
        ${conAtt ? `<label class="f">Giorni delle attività<input id="ps-gg" inputmode="decimal" placeholder="quelli del database" style="width:160px"></label>` : ""}
      </div>
      <label class="f" style="margin-top:10px">Capitolato<select id="ps-cap">${opzCapit}</select></label>
      <label class="f" id="ps-nuovo-box" hidden style="margin-top:6px">Nome del nuovo capitolato<input id="ps-nuovo" placeholder="es. PIANO TERRA"></label>
      <label class="f" style="margin-top:10px">Nota pre-linea<input id="ps-pre" placeholder="vuoto = quella del database"></label>
      <label class="f" style="margin-top:10px">Nota post-linea<input id="ps-post" placeholder="vuoto = quella del database"></label>
      ${conGen ? `<p class="muted small">Gli articoli generici tengono le note automatiche (n° di ordinazione e descrizione).</p>` : ""}`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Aggiungi", value: "ok", cls: "primary" }],
    onApri(d) {
      $("#ps-cap", d).addEventListener("change", (e) => { $("#ps-nuovo-box", d).hidden = e.target.value !== "__nuovo"; if (e.target.value === "__nuovo") $("#ps-nuovo", d).focus(); });
      setTimeout(() => $("#ps-qta", d).focus(), 50);
    },
    onConferma(v, d) {
      const q = num($("#ps-qta", d).value);
      if (!q) { toast("Scrivi la quantità."); $("#ps-qta", d).focus(); return false; }
      let cap = $("#ps-cap", d).value;
      if (cap === "__nuovo") { cap = $("#ps-nuovo", d).value.trim(); if (!aggiungiCapitolato(c, cap)) return false; }
      const gg = conAtt ? $("#ps-gg", d).value.trim() : "";
      const pre = $("#ps-pre", d).value.trim(), post = $("#ps-post", d).value.trim();
      const nuove = arts.map((a) => {
        const r = { ...rigaDaArticolo(a), id: db.nuovoId(), qta: q, capitolato: cap };
        if (isAttivita(a) && gg) r.giorni = num(gg);
        if (!isGenerico(r)) { if (pre) r.notePre = pre; if (post) r.notePost = post; }
        return r;
      });
      c.righe.push(...nuove);
      if (cap) c.capNuove = cap;
      S.pickSel.clear();
      modificata(); disegnaOfferta();
      toast(`Aggiunte ${nuove.length} righe${cap ? ` nel capitolato “${cap}”` : ""}, quantità ${numIt(q)}.`, 4000);
    },
  });
}

function contaInOfferta(c) {
  const m = new Map();
  for (const r of c.righe) { const x = m.get(r.k) || { n: 0, q: 0 }; x.n++; x.q += num(r.qta); m.set(r.k, x); }
  return m;
}
const badgeConta = (x) => (x ? `<span class="badge info" title="Righe già nell'offerta">${x.n === 1 ? "1 riga" : x.n + " righe"} · q.tà ${numIt(x.q)}</span>` : "");
$("#a-db").addEventListener("change", (e) => { S.pickDb = e.target.value; const c = capAtt(); if (c) disegnaPicker(c, true); });
$("#a-sotto").addEventListener("change", () => { const c = capAtt(); if (c) disegnaPicker(c); });
$("#a-cerca").addEventListener("input", () => { const c = capAtt(); if (c) disegnaPicker(c); });
$("#a-cap").addEventListener("change", suCap((c, e) => {
  if (e.target.value !== "__nuovo") { c.capNuove = e.target.value; modificata(); return; }
  dialogo({
    titolo: "Nuovo capitolato",
    html: `<label class="f">Nome che va in Oracle (Nome Capitolato)<input id="r-cap" placeholder="es. PIANO TERRA"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Crea", value: "ok", cls: "primary" }],
    onApri: (b) => setTimeout(() => $("#r-cap", b).focus(), 50),
    onConferma(v, b) {
      const n = $("#r-cap", b).value.trim();
      if (!aggiungiCapitolato(c, n)) return false;
      c.capNuove = n; modificata(); disegnaPicker(c);
    },
  });
  e.target.value = c.capNuove || "";
}));
$("#righe-picker").addEventListener("change", (e) => {
  const inp = e.target.closest('input[data-pick="qta"]'); if (!inp) return;
  const c = capAtt(); if (!c) return;
  const tr = inp.closest("tr");
  const a = S.perK.get(tr.dataset.k); if (!a) return;
  const q = num(inp.value);
  if (!q) { inp.value = ""; return; }
  const r = { ...rigaDaArticolo(a), id: db.nuovoId(), qta: q };
  if (isAttivita(a)) r.giorni = num(tr.querySelector('[data-pick="giorni"]')?.value);
  // nota per le nuove righe: sugli articoli senza una nota già scritta nel database; mai sui generici (le loro note sono automatiche)
  r.capitolato = c.capNuove || "";
  c.righe.push(r);
  inp.value = "";
  tr.classList.add("aggiunta"); setTimeout(() => tr.classList.remove("aggiunta"), 700);
  modificata();
  aggiornaSoloTotali();
});
function aggiornaSoloTotali() {
  // dopo un'aggiunta: aggiorna totali ed elenco righe senza ridisegnare la ricerca in cui si sta scrivendo
  S.noPicker = true;
  try { disegnaOffertaIn(false); } finally { S.noPicker = false; }
}

/* --- GMA: condizioni specifiche, peso e cartellino --- */
/* Sconti standard del database per un gruppo di righe: "75" se uguali per tutte, altrimenti "vari". */
function standardDi(righe) {
  const testi = new Set(righe.filter((r) => !isAttivita(r)).map((r) => {
    const e = scontiEffettivi(r);
    const p = CAMPI_SCONTO.filter((k) => num(e[k])).map((k) => pctTesto(k === "aum1" ? -num(e[k]) : num(e[k])));
    return p.length ? p.join("+") : "0";
  }));
  if (!testi.size) return "—";
  return testi.size === 1 ? [...testi][0] + "%" : "vari";
}
/* Standard del database per un gruppo di righe, come testo ("75", "50+10"); "" se diverso tra le righe; "0" per l'attività. */
function stdTesto(righe) {
  const testi = new Set(righe.map((r) => {
    if (isAttivita(r)) return "0";
    const e = scontiEffettivi(r);
    const p = CAMPI_SCONTO.filter((k) => num(e[k])).map((k) => pctTesto(k === "aum1" ? -num(e[k]) : num(e[k])));
    return p.length ? p.join("+") : "0";
  }));
  return testi.size === 1 ? [...testi][0] : "";
}
function disegnaGmv(c, rc) {
  const k = condizioni(c), ricarico = k.ven.modo === "ricarico";
  const righeCap = c.righe.filter((r) => num(r.qta));
  const gmaOrd = [...rc.perGmv.entries()].sort((a, b) => b[1].venduto - a[1].venduto);
  const conGma = gmaOrd.filter(([g]) => g);
  // condizioni: riga Generale + una riga per GMA, nel blocco acquisto e nel blocco vendita
  const stdDi = (g) => stdTesto(righeCap.filter((r) => (r.gmv || "") === g));
  const soloAtt = (g) => righeCap.filter((r) => (r.gmv || "") === g).every(isAttivita);
  // il Generale vale per il GMA senza valori propri (sull'attività solo se è in aggiunta):
  // se cambia lo standard, la cella standard del GMA dice "come Generale"; se aggiunge solo extra, resta lo standard del GMA
  const vale = (gen, g) => !!normTesto(gen) && (!soloAtt(g) || normTesto(gen).startsWith("+"));
  const opzGen = (gen, g, tipo = "sc", tutti = false) => {
    if (!(tutti ? !!normTesto(gen) : vale(gen, g))) return {};
    const { over, extra } = scomponi(gen);
    return { eredita: over !== null, phExtra: extra.length ? `Gen.: ${extra.map((x) => cellaDaTermine(x, tipo)).join(" · ")}` : "extra" };
  };
  const riga = (g, campo, gen = false) => `<tr class="${gen ? "cond-gen" : ""}"><td class="mono"><b>${gen ? "Generale" : esc(g)}</b></td>
    <td>${gen ? '<span class="muted small">GMA senza valori propri</span>' : esc(fornitoriTxt(g))}</td><td>${campo}</td></tr>`;
  $("#cond-acq").innerHTML = riga("", htmlCondizione({ dove: "acqGen", val: c.acq.gen, tipo: "sc", phStd: "standard" }), true) +
    conGma.map(([g]) => riga(g, htmlCondizione({ dove: "acqGma", gma: g, val: c.acq.gma[g], tipo: "sc", std: stdDi(g), ...opzGen(c.acq.gen, g) }))).join("");
  const ricStd = pctTesto(ricAutoCap(c));
  const lisM = k.imp === "listino";
  // standard dello sconto di listino: lo sconto di acquisto del GMA con più righe, uguale per tutti (anche l'attività)
  const lisStd = pctTesto(k.lis.std || 0);
  const stdRip = () => lisStd;
  $("#cond-ven").innerHTML = lisM
    ? riga("", htmlCondizione({ dove: "lisGen", val: c.lis.gen, tipo: "sc", std: lisStd }), true) +
      conGma.map(([g]) => riga(g, htmlCondizione({ dove: "lisGma", gma: g, val: c.lis.gma[g], tipo: "sc", std: stdRip(g), ...opzGen(c.lis.gen, g, "sc", true) }))).join("")
    : ricarico
    ? riga("", htmlCondizione({ dove: "ricGen", val: c.ven.ricGen, tipo: "ric", std: ricStd }), true) +
      conGma.map(([g]) => riga(g, htmlCondizione({ dove: "ricGma", gma: g, val: c.ven.ricGma[g], tipo: "ric", std: ricStd, ...(normTesto(c.ven.ricGen) ? (() => { const { over, extra } = scomponi(c.ven.ricGen); return { eredita: over !== null, phExtra: extra.length ? `Gen.: ${extra.join(" · ")}` : "extra" }; })() : {}) }))).join("")
    : riga("", htmlCondizione({ dove: "venScGen", val: c.ven.scGen, tipo: "sc", phStd: "standard" }), true) +
      conGma.map(([g]) => riga(g, htmlCondizione({ dove: "venScGma", gma: g, val: c.ven.scGma[g], tipo: "sc", std: stdDi(g), ...opzGen(c.ven.scGen, g) }))).join("");
  // intestazione "Standard" barrata in rosso se in quel riquadro uno standard è stato cambiato
  $("#cond-acq-std").classList.toggle("barrato", !!$("#cond-acq input.std.mod"));
  $("#cond-ven-std").classList.toggle("barrato", !!$("#cond-ven input.std.mod"));
  $("#rg-vend").textContent = lisM ? "Listino" : "Venduto";
  $("#cond-ven-std").textContent = lisM ? `Sconto di listino (standard: sconto di acquisto di ${k.lis.stdGma || "—"}, il GMA con più righe)` : ricarico ? "Ricarico automatico" : "Standard";
  $("#cond-ven-extra").textContent = lisM ? "Extra" : ricarico ? "Ricarichi extra" : "Extra";
  // riepilogo per GMA (solo lettura)
  const esiti = new Map(controlloCartellino(c, clienteOff(), opzCap(c)).map((x) => [x.gmv, x]));
  const html = gmaOrd.map(([g, info]) => {
    const peso = rc.venduto ? info.venduto / rc.venduto : 0;
    const e = esiti.get(g), att = e?.tipo === "attivita";
    const gen = gmaGenerico(g, info.prima.riga.fornitore) || (isGenerico(info.prima.riga) && info.prima.riga.daOfferta);
    const badge = !e ? `<span class="muted small">${gen ? "generici: nessun cartellino" : "—"}</span>` :
      `<span class="badge ${e.stato === "ok" ? "ok" : e.stato === "pari" ? "warn" : e.stato === "nc" ? "info" : "bad"}">${esc(e.testo)}</span>`;
    const scList = gen ? null : info.prima.c.scontoListino;
    return `<tr><td class="mono"><b>${g ? esc(g) : '<span class="muted">(senza GMA)</span>'}</b></td><td>${esc(fornitoriTxt(g))}</td>
      <td class="num">${info.righe}</td><td class="num">${euro(info.acquistato)}</td><td class="num">${euro(info.venduto)}</td>
      <td class="num"><div class="peso"><span style="width:${Math.max(0, Math.min(100, peso * 100)).toFixed(1)}%"></span></div>${pctTxt(peso) || "0%"}</td>
      <td class="num">${info.acquistato ? pctTxt(ricaricoPct(info.venduto, info.acquistato)) || "0%" : "—"}</td>
      <td class="num">${info.venduto ? pctTxt(marginePct(info.venduto, info.acquistato)) || "0%" : "—"}</td>
      <td class="num">${att ? (e ? euro(e.offerta) + "/g" : "") : scList === null ? "" : pctTxt(scList) || "0%"}</td>
      <td class="num">${e ? (e.stato === "nc" ? "NC" : att ? euro(e.cartellino) + "/g" : pctTxt(e.cartellino) || "0%") : ""}</td><td>${badge}</td></tr>`;
  }).join("");
  $("#righe-gmv").innerHTML = html || `<tr><td colspan="11" class="vuoto">Aggiungi articoli al capitolo per vedere i loro GMA.</td></tr>`;
}

/* --- Esportazione per Oracle (un foglio per capitolo) --- */
const datiOracle = (c) => righeOracle(c, opzCap(c));
function disegnaExport(c) {
  // scelta dei capitolati da caricare in questa offerta Oracle
  const gruppi = [];
  for (const r of c.righe) {
    if (!num(r.qta)) continue;
    const n = r.capitolato || "";
    let g = gruppi.find((x) => x.nome === n);
    if (!g) { g = { nome: n, righe: 0, totale: 0 }; gruppi.push(g); }
    const x = calcolaRiga(r, c, opzCap(c).ricarichi);
    g.righe++; g.totale += condizioni(c).out.prezzo === "acquisto" ? x.acq.totale : x.totale;
  }
  const escl = new Set(c.esclusiExport || []);
  $("#export-cap-card").hidden = !(gruppi.length > 1 || escl.size);
  $("#export-capitolati").innerHTML = gruppi.length > 1 || escl.size ? "" +
    gruppi.map((g) => `<label class="chip-check"><input type="checkbox" data-escl="${esc(g.nome)}" ${escl.has(g.nome) ? "" : "checked"}>
      <b>${g.nome ? esc(g.nome) : "senza capitolato"}</b> <span class="muted">${g.righe} righe · ${euro(g.totale)}</span></label>`).join("") +
    `<button class="btn small" id="export-tutti" type="button">Tutti</button>` : "";
  const { righe, avvisi } = datiOracle(c);
  const av = [];
  if (escl.size) av.push(`<div class="avviso info">Esportazione parziale: ${[...escl].map((n) => `“${esc(n || "senza capitolato")}”`).join(", ")} ${escl.size === 1 ? "escluso" : "esclusi"}.</div>`);
  if (!righe.length) av.push(`<div class="avviso">Nessuna riga da esportare.</div>`);
  if (avvisi.length) av.push(`<div class="avviso bad"><b>${avvisi.length} righe escluse:</b><ul class="riep">${avvisi.map((a) => `<li>${esc(a)}</li>`).join("")}</ul></div>`);
  const fuori = controlloCartellino(c, clienteOff(), opzCap(c)).filter((x) => x.stato === "fuori");
  if (fuori.length) av.push(`<div class="avviso">Attenzione: ${fuori.map((f) => esc(f.gmv)).join(", ")} fuori cartellino del cliente (vedi “2 · Sconti e ricarichi”).</div>`);
  // prezzi da aggiornare o cambio listino in arrivo
  const daAgg = [], cambi = [];
  for (const r of c.righe) {
    if (!num(r.qta) || esclusoDaExport(r, c)) continue;
    const st = stato(S.perK.get(r.k) || r);
    if (st.livello === "vecchio") daAgg.push(`${codiceArticolo(r)} ${r.nOrdine}: ${st.testo}`);
    else if (st.livello === "medio") cambi.push(`${codiceArticolo(r)} ${r.nOrdine}: ${st.testo}`);
  }
  const elenco = (l) => `<details><summary>vedi</summary><ul class="riep">${l.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></details>`;
  if (daAgg.length) av.push(`<div class="avviso bad"><b>${daAgg.length} righe con prezzo da aggiornare.</b> ${elenco(daAgg)}</div>`);
  if (cambi.length) av.push(`<div class="avviso"><b>${cambi.length} righe con prezzo in scadenza o cambio listino entro un mese.</b> ${elenco(cambi)}</div>`);
  const senzaCap = righeSenzaCapitolato(c);
  if (senzaCap) av.push(`<div class="avviso bad"><b>Non tutte le righe hanno un capitolato:</b> ${senzaCap} ${senzaCap === 1 ? "riga è" : "righe sono"} senza capitolato${c.sottocapitoli ? " e in Oracle avranno “Nome Capitolato” vuoto" : ""}.</div>`);
  $("#export-avvisi").innerHTML = av.join("");
  $("#export-titolo").textContent = `Foglio per Oracle del capitolo ${c.nome} — ${righe.length} righe`;
  $("#tab-export").innerHTML = `<thead><tr>${COLONNE_ORACLE.map((x) => `<th>${esc(x)}</th>`).join("")}</tr></thead>
    <tbody>${righe.map((r) => `<tr>${r.map((v, i) => `<td class="${typeof v === "number" ? "num" : i === 1 ? "mono" : ""}"${typeof v === "string" && v.length > 30 ? ` title="${esc(v)}"` : ""}>${esc(typeof v === "number" ? numIt(v) : v)}</td>`).join("")}</tr>`).join("")}</tbody>`;
}
const nomeFile = (c, est) => {
  // con un'esportazione parziale nel nome del file vanno i capitolati inclusi
  const escl = new Set(c.esclusiExport || []);
  const inclusi = escl.size ? [...new Set(c.righe.filter((r) => num(r.qta) && !escl.has(r.capitolato || "")).map((r) => r.capitolato || "senza-capitolato"))] : [];
  const parte = inclusi.length ? "_" + inclusi.join("+") : "";
  return `Import_Righe_${`${S.off.nome || "offerta"}_${c.nome}${parte}`.replace(/[^\w.+-]+/g, "_")}.${est}`;
};
$("#export-capitolati").addEventListener("change", suCap((c, e) => {
  const cb = e.target.closest("[data-escl]"); if (!cb) return;
  const n = cb.dataset.escl;
  const escl = new Set(c.esclusiExport || []);
  if (cb.checked) escl.delete(n); else escl.add(n);
  c.esclusiExport = [...escl];
  modificata(); disegnaOfferta();
}));
$("#export-capitolati").addEventListener("click", suCap((c, e) => {
  if (e.target.id !== "export-tutti") return;
  c.esclusiExport = []; modificata(); disegnaOfferta();
}));
/* righe che vanno in Oracle senza capitolato, quando nel capitolo ci sono capitolati */
const righeSenzaCapitolato = (c) => (c.capitolati.length || c.sottocapitoli)
  ? c.righe.filter((r) => num(r.qta) && !r.capitolato && !esclusoDaExport(r, c)).length : 0;
/* prima di scaricare: avviso se non tutte le righe hanno un capitolato */
function controllaCapitolati(c, procedi) {
  const n = righeSenzaCapitolato(c);
  if (!n) return procedi();
  dialogo({
    titolo: "Non tutte le righe hanno un capitolato",
    html: `<p>Nel capitolo <b>${esc(c.nome)}</b> ${n === 1 ? "c'è 1 riga" : `ci sono ${n} righe`} senza capitolato.
      ${c.sottocapitoli ? "In Oracle avranno la colonna “Nome Capitolato” vuota." : ""}</p><p>Vuoi scaricare comunque?</p>`,
    pulsanti: [{ label: "Torna alle righe", value: "righe" }, { label: "Scarica comunque", value: "ok", cls: "primary" }],
    onConferma(v) {
      if (v === "righe") { S.sub = "righe"; disegnaOfferta(true); return; }
      procedi();
    },
  });
}
/* Dal riepilogo: prima le impostazioni dell'output (le stesse della pagina Oracle), poi il download. */
function impostaEScarica(c) {
  if (!c) return;
  aggiornaCondizioni(c);
  const st = { prezzo: c.out.prezzo, formato: c.out.formato, sotto: !!c.sottocapitoli, escl: new Set(c.esclusiExport || []), file: c.out.file || "computo" };
  const gruppi = new Map();
  for (const r of c.righe) { if (!num(r.qta)) continue; const g = r.capitolato || ""; gruppi.set(g, (gruppi.get(g) || 0) + 1); }
  const lis = c.imp === "listino";
  const segHtml = (id, valori, att, et) => `<div class="seg" id="${id}">${valori.map((v) => `<button type="button" data-v="${v}" class="${v === att ? "on" : ""}">${esc(et(v))}</button>`).join("")}</div>`;
  const campo = (titolo, html) => `<div style="display:flex;flex-direction:column;gap:3px"><span class="muted small" style="font-weight:600">${titolo}</span>${html}</div>`;
  dialogo({
    titolo: lis ? `Scarica ${c.nome}` : `Scarica il foglio Oracle di ${c.nome}`,
    html: `<div class="row" style="gap:18px;align-items:flex-end;flex-wrap:wrap">
        ${lis ? campo("File", segHtml("ie-file", ["computo", "oracle"], st.file, (v) => (v === "computo" ? "Computo a listino" : "Foglio Oracle"))) : ""}
        ${lis ? `<label class="f ie-computo">Titolo del computo<input id="ie-titolo" style="width:260px" value="${esc(c.titoloComputo || c.nome)}"></label>` : ""}
        ${campo("Prezzi da esportare", segHtml("ie-prezzo", ["vendita", "acquisto"], st.prezzo, (v) => (v === "vendita" && lis ? "Listino" : OUT_PREZZO[v])))}
        ${campo("Come", segHtml("ie-formato", ["sconti", "netti"], st.formato, (v) => OUT_FORMATO[v]))}
        ${campo("Capitolati in Oracle", segHtml("ie-sotto", ["SI", "NO"], st.sotto ? "SI" : "NO", (v) => v))}
        <label class="f">N° partenza righe<input id="ie-linea" type="number" min="0" step="10" style="width:100px" value="${esc(c.lineaPartenza ?? 10)}"></label>
        <label class="f">Magazzino<input id="ie-mag" maxlength="10" style="width:80px" value="${esc(c.magazzino ?? "BG")}"></label>
      </div>
      ${gruppi.size > 1 || st.escl.size ? `<div style="margin-top:14px"><span class="muted small" style="font-weight:600">${lis ? "Sottocapitoli da includere" : "Capitolati da caricare in questa offerta Oracle"}</span>
        <div class="capitolati" style="margin-top:4px">${[...gruppi.entries()].map(([g, n]) => `<label class="chip-check"><input type="checkbox" data-ie-cap="${esc(g)}" ${st.escl.has(g) ? "" : "checked"}> <b>${g ? esc(g) : "senza capitolato"}</b> <span class="muted">${n} righe</span></label>`).join("")}</div></div>` : ""}
      <p class="muted small" style="margin:12px 0 0">Le impostazioni restano salvate nel capitolo (sono le stesse della pagina Oracle).</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "⤓ Scarica Excel", value: "ok", cls: "primary" }],
    onApri(d) {
      d.addEventListener("click", (e) => {
        const b = e.target.closest(".seg button[data-v]"); if (!b) return;
        const sg = b.parentElement;
        sg.querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
        if (sg.id === "ie-prezzo") st.prezzo = b.dataset.v;
        if (sg.id === "ie-formato") st.formato = b.dataset.v;
        if (sg.id === "ie-sotto") st.sotto = b.dataset.v === "SI";
        if (sg.id === "ie-file") { st.file = b.dataset.v; vedi(); }
      });
      // con il computo contano solo il titolo e i capitolati da includere
      const vedi = () => { if (!lis) return; const comp = st.file === "computo";
        for (const id of ["#ie-prezzo", "#ie-formato", "#ie-sotto"]) $(id, d).parentElement.hidden = comp;
        for (const id of ["#ie-linea", "#ie-mag"]) $(id, d).closest("label").hidden = comp;
        $(".ie-computo", d).hidden = !comp; };
      vedi();
    },
    onConferma(v, d) {
      c.out.prezzo = st.prezzo; c.out.formato = st.formato; c.sottocapitoli = st.sotto;
      if (lis) { c.out.file = st.file; c.titoloComputo = $("#ie-titolo", d).value.trim(); }
      c.lineaPartenza = num($("#ie-linea", d).value) || 10;
      c.magazzino = $("#ie-mag", d).value.trim() || "BG";
      const escl = $$("[data-ie-cap]", d).filter((x) => !x.checked).map((x) => x.dataset.ieCap);
      if (gruppi.size && escl.length === gruppi.size) { toast("Lascia almeno un capitolato da caricare."); return false; }
      c.esclusiExport = escl;
      modificata(); ridisegna(true);
      setTimeout(() => controllaCapitolati(c, () => scaricaXlsx(c)), 50);
    },
  });
}
/* ================= Computo a listino (offerte con prezzi "Listino") =================
   Foglio come PB…_COMPUTO_FIRE: Pos. · RIF CAP · Numero di ordinazione · Descrizione · U.M. · Q.tà · Prezzo di listino
   unitario e totale; titolo del capitolo, una riga per ogni sottocapitolo, totale dell'offerta. */
const STILI_COMPUTO = [
  {},
  { font: { b: true }, fondo: "6969FF", alH: "center", alV: "center", a_capo: true },            // 1 intestazione
  { font: { b: true }, fondo: "6969FF", alH: "left", alV: "center" },                            // 2 intestazione a sinistra
  { font: { b: true }, fondo: "9191FF", alH: "left", alV: "center", a_capo: true },              // 3 sottocapitolo
  { font: { b: true }, fondo: "9191FF", alH: "center", alV: "center" },                          // 4 sottocapitolo centrato
  { fondo: "F8F8F8", alH: "left", alV: "top", a_capo: true },                                     // 5 testo
  { fondo: "F8F8F8", alH: "center", alV: "top" },                                                 // 6 centrato
  { fondo: "F8F8F8", alH: "center", alV: "top", fmt: "#,##0.##" },                                // 7 quantità
  { fondo: "F8F8F8", alH: "right", alV: "top", fmt: '#,##0.00\\ "€"' },                         // 8 prezzo
  { font: { b: true, sz: 12 }, fondo: "9191FF", alH: "right", alV: "center", bordo: "medium" },   // 9 totale (scritta)
  { font: { b: true, sz: 12 }, fondo: "9191FF", alH: "right", alV: "center", bordo: "medium", fmt: '#,##0.00\\ "€"' }, // 10 totale (importo)
  { fondo: "9191FF", bordo: "medium" },                                                           // 11 totale (vuote)
  { font: { b: true }, fondo: "6969FF", alH: "center", alV: "center", fmt: '#,##0.00\\ "€"' },  // 12 intestazione prezzi
];
/* altezza della riga per una descrizione su più righe (colonna D larga ~67 caratteri) */
const altezzaTesto = (t, larg = 72) => { const n = String(t || "").split("\n").reduce((a, x) => a + Math.max(1, Math.ceil(x.length / larg)), 0); return Math.min(409, Math.max(15, Math.round(n * 12.9 + 3))); };
function computoListino(c) {
  aggiornaCondizioni(c);
  const opz = opzCap(c), k = condizioni(c);
  const righe = c.righe.filter((r) => num(r.qta) && !esclusoDaExport(r, c));
  const R = [], unioni = [];
  const vuote = (s, n = 8) => Array.from({ length: n }, () => ({ s }));
  R.push({ h: 15, celle: [{ v: "Pos.", s: 2 }, { v: "RIF CAP", s: 1 }, { v: "Numero di ordinazione", s: 2 }, { v: "Descrizione", s: 2 }, { v: "U.M.", s: 1 }, { v: "Q.tà", s: 1 }, { v: "Prezzo di Listino in Euro", s: 1 }, { s: 1 }] });
  R.push({ h: 15, celle: [{ s: 2 }, { s: 1 }, { s: 2 }, { s: 2 }, { s: 1 }, { s: 1 }, { v: "Unitario", s: 12 }, { v: "Totale", s: 12 }] });
  unioni.push("G1:H1");
  R.push({ h: 15, celle: [{ v: (c.titoloComputo || c.nome).toUpperCase(), s: 1 }, ...vuote(1, 7)] });
  unioni.push("A3:H3");
  const primaRiga = R.length + 1;
  let gruppo = null;
  for (const r of righe) {
    const g = r.capitolato || "";
    if (g !== gruppo) {
      gruppo = g;
      if (g || c.capitolati.length) {
        const rif = [...new Set(righe.filter((x) => (x.capitolato || "") === g).map((x) => x.rifCap || ""))];
        R.push({ h: 15, celle: [{ s: 3 }, { v: rif.length === 1 ? rif[0] : "", s: 4 }, { s: 3 }, { v: g || "Altro", s: 3 }, { s: 4 }, { s: 4 }, { s: 4 }, { s: 4 }] });
      }
    }
    const x = calcolaRiga(r, c, opz.ricarichi, k), att = isAttivita(r);
    const nota = (t) => t && R.push({ h: altezzaTesto(t), celle: [{ s: 6 }, { s: 6 }, { s: 5 }, { v: t, s: 5 }, { s: 6 }, { s: 7 }, { s: 8 }, { s: 8 }] });
    if (!isGenerico(r) && !r.daOfferta) nota(r.notePre);
    const cod = att ? "ATTIVITA'" : r.nOrdine || codiceArticolo(r);
    const des = att ? (r.daOfferta ? r.notePost || r.descrizione : r.descrizione) : r.descrizione;
    const rifRiga = c.capitolati.length ? "" : r.rifCap || "";
    const n = R.length + 1;
    R.push({ h: altezzaTesto(des), celle: [{ s: 6 }, { v: rifRiga, s: 6 }, { v: cod, s: 5 }, { v: des || "", s: 5 }, { v: r.um || "n.", s: 6 }, { v: x.qta, s: 7 }, { v: Math.round(x.unitario * 100) / 100, s: 8 }, { f: `F${n}*G${n}`, s: 8 }] });
    if (!isGenerico(r) && !r.daOfferta) nota(r.notePost);
  }
  const ultima = R.length;
  R.push({ h: 18, celle: [{ s: 11 }, { s: 11 }, { s: 11 }, { v: "Totale Offerta", s: 9 }, { s: 11 }, { s: 11 }, { s: 11 }, { f: `SUM(H${primaRiga}:H${ultima})`, s: 10 }] });
  return { nome: c.nome, colonne: [8.7, 8.7, 22.9, 67.1, 5, 8, 15.7, 15.7], righe: R, unioni, stili: STILI_COMPUTO };
}
function scaricaComputo(c) {
  scarica(scriviExcelStile(computoListino(c)), `COMPUTO_${String(c.nome).replace(/[^A-Za-z0-9_-]+/g, "_")}${S.off?.rif ? "_" + String(S.off.rif).replace(/[^A-Za-z0-9_-]+/g, "_") : ""}.xlsx`);
}

function scaricaXlsx(c) {
  if (!c) return;
  if (c.imp === "listino" && (c.out?.file || "computo") === "computo") return scaricaComputo(c);
  const { righe } = datiOracle(c);
  scarica(scriviExcel([COLONNE_ORACLE, ...righe], "Sheet1"), nomeFile(c, "xlsx"));
}
$("#scarica-xlsx").addEventListener("click", suCap((c) => controllaCapitolati(c, () => scaricaXlsx(c))));
const celleTesto = (sep) => (v) => {
  const s = typeof v === "number" ? String(v).replace(".", ",") : String(v ?? "");
  return sep === ";" && /[;"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s.replace(/[\t\n\r]+/g, " ");
};
$("#scarica-csv").addEventListener("click", suCap((c) => controllaCapitolati(c, () => {
  const { righe } = datiOracle(c);
  const f = celleTesto(";");
  const csv = [COLONNE_ORACLE, ...righe].map((r) => r.map(f).join(";")).join("\r\n");
  scarica(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }), nomeFile(c, "csv"));
})));
$("#copia-righe").addEventListener("click", suCap(async (c) => {
  const { righe } = datiOracle(c);
  const f = celleTesto("\t");
  const tsv = [COLONNE_ORACLE, ...righe].map((r) => r.map(f).join("\t")).join("\n");
  try { await navigator.clipboard.writeText(tsv); toast(`Copiate ${righe.length} righe con le intestazioni`); }
  catch { toast("Il browser non permette di copiare: usa “Scarica Excel”."); }
}));

/* ================= Catalogo ================= */
$("#c-capitolo").addEventListener("change", () => { S.limiteCat = 300; $("#c-sotto").value = ""; disegnaCatalogo(); });
$("#c-sotto").addEventListener("change", () => { S.limiteCat = 300; disegnaCatalogo(); });
$("#c-cerca").addEventListener("input", () => { S.limiteCat = 300; disegnaCatalogo(); });
$("#c-altri").addEventListener("click", () => { S.limiteCat += 300; disegnaCatalogo(); });

/* Filtri del database: precodice, fornitore, GMA, tipo, stato del prezzo. */
function opzFiltro(sel, valori, tutti) {
  const v = sel.value;
  sel.innerHTML = `<option value="">${tutti}</option>` + valori.map((x) => `<option value="${esc(x)}" ${x === v ? "selected" : ""}>${esc(x)}</option>`).join("");
  if (v && !valori.includes(v)) { sel.insertAdjacentHTML("beforeend", `<option value="${esc(v)}" selected>${esc(v)}</option>`); }
}
function aggiornaFiltriCat() {
  const pre = [...new Set(S.catalogo.map((a) => String(a.precodice || "").trim().toUpperCase()).filter(Boolean))].sort();
  opzFiltro($("#cf-pre"), pre, "Tutti");
  opzFiltro($("#cf-forn"), fornitoriGma(), "Tutti");
  const f = $("#cf-forn").value;
  const gmas = [...new Set(S.catalogo.filter((a) => a.gmv && (!f || fornitoreDi(a) === f)).map((a) => a.gmv))].sort();
  opzFiltro($("#cf-gma"), gmas, "Tutti");
  const attivi = ["#cf-pre", "#cf-forn", "#cf-gma", "#cf-tipo", "#cf-prezzo", "#cf-dop"].filter((id) => $(id).value).length;
  $("#cf-azzera").disabled = !attivi;
  $(".filtri-cat").classList.toggle("attivi", !!attivi);
}
/* Doppioni: stesso codice (precodice + codice, esclusi i codici generici) o stesso n° di ordinazione. */
const chiaveCod = (a) => { const c = String(a.codice || "").trim().toUpperCase(); return !c || c === "ARTICOLO_GEN" || c === "ATT.TEC" ? "" : (String(a.precodice || "").trim() + c).toUpperCase(); };
const chiaveOrd = (a) => String(a.nOrdine || "").replace(/\s+/g, "").toUpperCase();
function contaDoppioni() {
  if (S.dopCache?.cat === S.catalogo) return S.dopCache;
  // i cloni sono voluti: un gruppo di cloni conta come un articolo solo
  const cod = new Map(), ord = new Map(), visti = new Set();
  for (const a of S.catalogo) {
    if (a.clone) { if (visti.has(a.clone)) continue; visti.add(a.clone); }
    const c = chiaveCod(a), o = chiaveOrd(a); if (c) cod.set(c, (cod.get(c) || 0) + 1); if (o) ord.set(o, (ord.get(o) || 0) + 1);
  }
  return (S.dopCache = { cat: S.catalogo, cod, ord });
}
const doppioCod = (a) => { const c = chiaveCod(a); return !!c && contaDoppioni().cod.get(c) > 1; };
const doppioOrd = (a) => { const o = chiaveOrd(a); return !!o && contaDoppioni().ord.get(o) > 1; };
function passaFiltriCat(a) {
  const dp = $("#cf-dop").value;
  if (dp === "cod" && !doppioCod(a)) return false;
  if (dp === "nord" && !doppioOrd(a)) return false;
  if (dp === "tutti" && !doppioCod(a) && !doppioOrd(a)) return false;
  const pre = $("#cf-pre").value, f = $("#cf-forn").value, g = $("#cf-gma").value, t = $("#cf-tipo").value, pz = $("#cf-prezzo").value;
  if (pre && String(a.precodice || "").trim().toUpperCase() !== pre) return false;
  if (f && fornitoreDi(a) !== f) return false;
  if (g && a.gmv !== g) return false;
  if (t === "gen" && !isGenerico(a)) return false;
  if (t === "att" && !isAttivita(a)) return false;
  if (t === "art" && (isGenerico(a) || isAttivita(a))) return false;
  if (pz) { const l = stato(a).livello || "nd"; if ((pz === "nd" ? !["fresco", "medio", "vecchio"].includes(l) : l !== pz)) return false; }
  return true;
}
for (const id of ["#cf-pre", "#cf-forn", "#cf-gma", "#cf-tipo", "#cf-prezzo", "#cf-dop"]) $(id).addEventListener("change", () => { S.limiteCat = 300; disegnaCatalogo(); });
$("#cf-azzera").addEventListener("click", () => { for (const id of ["#cf-pre", "#cf-forn", "#cf-gma", "#cf-tipo", "#cf-prezzo", "#cf-dop"]) $(id).value = ""; S.limiteCat = 300; disegnaCatalogo(); });


function disegnaCatalogo() {
  aggiornaRitorno();
  aggiornaFiltriCat();
  const sel = $("#c-capitolo");
  const prev = sel.value;
  sel.innerHTML = `<option value="">Tutti i capitoli</option>` +
    S.capitoli.map((c) => `<option ${c.nome === prev ? "selected" : ""}>${esc(c.nome)}</option>`).join("");
  const cap = sel.value;
  const ss = $("#c-sotto"), prevS = ss.value;
  const sotti = cap ? [...new Set(S.catalogo.filter((a) => a.capitolo === cap).map((a) => a.sottocapitolo).filter(Boolean))] : [];
  ss.innerHTML = `<option value="">Tutti i sottocapitoli</option>` + sotti.map((x) => `<option ${x === prevS ? "selected" : ""}>${esc(x)}</option>`).join("");
  ss.hidden = !cap;
  const sotto = ss.value;
  const q = $("#c-cerca").value.trim().toLowerCase();
  const lista = S.catalogo.filter((a) => (!cap || a.capitolo === cap) && (!sotto || a.sottocapitolo === sotto) &&
    (!q || `${codiceArticolo(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)) && passaFiltriCat(a));
  S.catListaFiltrata = lista;
  const dp = $("#cf-dop").value;
  const chiaveDop = (a) => (dp === "nord" || (dp === "tutti" && !doppioCod(a)) ? "N° ordinazione " + (a.nOrdine || "") : "Codice " + codiceArticolo(a));
  const ord = dp ? lista.slice().sort((x, y) => chiaveDop(x).localeCompare(chiaveDop(y)) || String(x.capitolo).localeCompare(String(y.capitolo))) : ordinaRighe(lista, ordineCap());
  S.catOrdinata = ord;
  const vis = ord.slice(0, S.limiteCat);
  S.catVisK = vis.map((a) => a.k);
  let gruppo = null, html = "";
  for (const a of vis) {
    const g = dp ? "Doppione · " + chiaveDop(a) : (cap ? "" : a.capitolo + " · ") + (a.sottocapitolo || "—");
    if (g !== gruppo) { gruppo = g; html += `<tr class="sotto" data-gruppo="${esc(g)}"><td colspan="11">${esc(g)}</td></tr>`; }
    html += `<tr class="cliccabile${S.catEdit && S.catSel.has(a.k) ? " sel" : ""}${S.catApp?.taglio && S.catApp.ks.includes(a.k) ? " taglio" : ""}" data-k="${esc(a.k)}">
      <td class="mono nowrap">${codiceVis(a)}</td>
      <td class="img-cel">${a.img ? `<img class="img-mini" data-img="${esc(a.img)}" alt="">` : ""}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b>${a.clone && cloniDi(a).length > 1 ? ` <span class="badge info" title="Riga clonata: è collegata ad altre ${cloniDi(a).length - 1} copie nel database">🔗 ${cloniDi(a).length}</span>` : ""}${dp ? ` <span class="badge info">${esc(a.capitolo)} · ${esc(a.sottocapitolo || "—")}</span>` : ""}<div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td>
      ${(() => { const e = scontiEffettivi(a), c = e.daGma ? ' class="num std-gma" title="Acquisto standard del GMA"' : ' class="num"';
        return `<td${c}>${pctTxt(e.sc3)}</td><td${c}>${pctTxt(e.sc4)}</td><td${c}>${pctTxt(e.sc5)}</td><td${c}>${pctTxt(e.aum1)}</td>`; })()}<td class="mono">${esc(a.gmv)}</td>
      <td class="nowrap">${pallino(a)}${esc(dataIt(a.dataAgg))}</td></tr>`;
  }
  $("#righe-catalogo").innerHTML = html || `<tr><td colspan="11" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Catalogo vuoto: importalo da “Importa e aggiorna”."}</td></tr>`;
  applicaGruppi($("#righe-catalogo"));
  caricaMiniature($("#righe-catalogo"));
  $("#c-info").textContent = `${vis.length} di ${lista.length} articoli` + (S.catEdit && S.catSel.size ? ` · ${S.catSel.size} selezionati` : "");
  $("#c-altri").hidden = vis.length >= lista.length;
}
$("#righe-catalogo").addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-k]"); if (!tr) return;
  if (!S.catEdit) return modificaArticolo(S.perK.get(tr.dataset.k));
  if (appenaTrascinato()) return;
  // modalità modifica: clic = seleziona (Ctrl/⌘ aggiunge o toglie, Maiusc intervallo)
  const k = tr.dataset.k, vis = S.catVisK || [], ctrl = e.ctrlKey || e.metaKey;
  if (e.shiftKey && S.catUltimo && S.catUltimo !== k && vis.includes(S.catUltimo)) {
    const a = vis.indexOf(S.catUltimo), b = vis.indexOf(k);
    if (!ctrl) S.catSel.clear();
    for (let i = Math.min(a, b); i <= Math.max(a, b); i++) S.catSel.add(vis[i]);
    window.getSelection?.().removeAllRanges();
  } else if (ctrl) S.catSel.has(k) ? S.catSel.delete(k) : S.catSel.add(k);
  else { const solo = S.catSel.size === 1 && S.catSel.has(k); S.catSel.clear(); if (!solo) S.catSel.add(k); }
  S.catUltimo = k;
  aggiornaSelCat();
});
$("#righe-catalogo").addEventListener("dblclick", (e) => {
  const tr = e.target.closest("tr[data-k]");
  if (tr && S.catEdit) modificaArticolo(S.perK.get(tr.dataset.k));
});

/* ================= Modifica database (tasto destro sul catalogo) =================
   Bloccata finché non si preme "Modifica database"; si riblocca cambiando pagina. */
S.catSel = new Set(); S.catEdit = false; S.catApp = null;
function aggiornaSelCat() {
  for (const tr of $$("#righe-catalogo tr[data-k]")) tr.classList.toggle("sel", S.catEdit && S.catSel.has(tr.dataset.k));
  $("#c-info").textContent = `${(S.catVisK || []).length} di ${(S.catListaFiltrata || []).length} articoli` + (S.catEdit && S.catSel.size ? ` · ${S.catSel.size} selezionati` : "");
}
function impostaModificaCat(on) {
  S.catEdit = on;
  if (!on) { S.catSel.clear(); if (S.catApp?.taglio) S.catApp = null; }
  $("#c-modifica").classList.toggle("on", on);
  $("#c-modifica").textContent = on ? "🔓 Modifica attiva · Blocca" : "🔒 Modifica database";
  $("#c-mod-avviso").hidden = !on;
  $("#v-catalogo").classList.toggle("modifica", on);
}
function bloccaModificaCat() { if (S.catEdit) { impostaModificaCat(false); chiudiMenuDx(); } }
$("#c-modifica").addEventListener("click", () => { impostaModificaCat(!S.catEdit); disegnaCatalogo(); });

$("#righe-catalogo").addEventListener("mousedown", (e) => {
  if (!S.catEdit) return;
  if ((e.shiftKey || e.ctrlKey || e.metaKey) && e.target.closest("tr[data-k]")) e.preventDefault();
  avviaTrascina(e, "cat");
});
$("#righe-catalogo").addEventListener("contextmenu", (e) => {
  if (!S.catEdit) return;
  const tr = e.target.closest("tr[data-k]"); if (!tr) return;
  e.preventDefault();
  if (!S.catSel.has(tr.dataset.k)) { S.catSel = new Set([tr.dataset.k]); S.catUltimo = tr.dataset.k; aggiornaSelCat(); }
  S.catDx = tr.dataset.k;
  const n = S.catSel.size, m = $("#menu-dx"), app = S.catApp, na = app?.ks.length || 0;
  m.innerHTML = `<div class="dx-titolo">${n === 1 ? "1 articolo selezionato" : `${n} articoli selezionati`}</div>
    <button type="button" data-dxc="taglia">✂ Taglia <kbd>${TASTO}X</kbd></button>
    <button type="button" data-dxc="copia">⎘ Copia <kbd>${TASTO}C</kbd></button>
    <button type="button" data-dxc="incolla" ${na ? "" : "disabled"}>⇩ Incolla sotto${na ? ` (${na})` : ""} <kbd>${TASTO}V</kbd></button>
    <button type="button" data-dxc="incollaLink" ${na && !app.taglio ? "" : "disabled"} title="Incolla dei cloni: modificando una copia cambiano tutte">🔗 Incolla collegato sotto${na && !app.taglio ? ` (${na})` : ""}</button>
    <hr>
    ${n === 1 ? `<button type="button" data-dxc="scheda">✎ Apri la scheda articolo</button>` : ""}
    <button type="button" data-dxc="tutti">Seleziona tutti quelli visibili</button>
    <button type="button" data-dxc="nessuno">Deseleziona</button>
    <hr>
    <button type="button" data-dxc="elimina" class="dx-danger">✕ Elimina${n > 1 ? ` ${n} articoli` : ""}…</button>`;
  m.hidden = false; menuDxAperto = Date.now();
  const w = m.offsetWidth, h = m.offsetHeight;
  m.style.left = Math.min(e.clientX, innerWidth - w - 8) + "px";
  m.style.top = Math.min(e.clientY, innerHeight - h - 8) + "px";
});
$("#menu-dx").addEventListener("click", (e) => {
  const b = e.target.closest("[data-dxc]"); if (!b) return;
  chiudiMenuDx();
  azioneCat(b.dataset.dxc, S.catDx);
});
document.addEventListener("keydown", (e) => {
  if (!S.catEdit || S.vista !== "catalogo" || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
  if (e.target.closest?.("input:not([type=checkbox]), textarea, select, [contenteditable], dialog")) return;
  const az = { KeyX: "taglia", KeyC: "copia", KeyV: "incolla" }[e.code]; if (!az) return;
  if (az !== "incolla" && !S.catSel.size) return;
  e.preventDefault(); chiudiMenuDx();
  azioneCat(az, S.catUltimo && S.catSel.has(S.catUltimo) ? S.catUltimo : [...S.catSel].pop());
});
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape" || !S.catEdit || S.vista !== "catalogo" || document.querySelector("dialog[open]") || !$("#menu-dx").hidden) return;
  if (S.catApp?.taglio) S.catApp = null;
  S.catSel.clear(); disegnaCatalogo();
});

/* Articoli selezionati nell'ordine del database */
const selCatOrdinati = () => ordinaRighe(S.catalogo.filter((a) => S.catSel.has(a.k)), ordineCap());
async function azioneCat(az, kDx) {
  if (!S.catEdit) return;
  if (az === "tutti") { (S.catVisK || []).forEach((k) => S.catSel.add(k)); return aggiornaSelCat(); }
  if (az === "nessuno") { S.catSel.clear(); return aggiornaSelCat(); }
  if (az === "scheda") return modificaArticolo(S.perK.get(kDx));
  if (az === "taglia" || az === "copia") {
    const arts = selCatOrdinati(); if (!arts.length) return;
    S.catApp = { taglio: az === "taglia", ks: arts.map((a) => a.k) };
    disegnaCatalogo();
    return toast(`${arts.length} ${arts.length === 1 ? "articolo" : "articoli"} ${az === "taglia" ? "tagliati" : "copiati"}: tasto destro su una riga › Incolla sotto${az === "copia" ? " (o Incolla collegato)" : ""}.`, 4000);
  }
  if (az === "incolla" || az === "incollaLink") return incollaCat(kDx, az === "incollaLink");
  if (az === "elimina") return eliminaCat();
}
async function incollaCat(kDx, link) {
  const app = S.catApp, dove = S.perK.get(kDx);
  if (!app?.ks.length) return toast("Prima taglia o copia degli articoli.");
  if (!dove) return toast("Clicca col destro sulla riga sotto cui incollare.");
  const sorgenti = app.ks.map((k) => S.perK.get(k)).filter(Boolean);
  if (!sorgenti.length) { S.catApp = null; return toast("Gli articoli copiati non ci sono più nel database."); }
  if (app.taglio && sorgenti.includes(dove)) return toast("Non puoi incollare sotto un articolo che stai spostando.");
  const caps = new Set([dove.capitolo]);
  let nuovi;
  if (app.taglio) {
    nuovi = sorgenti;
    for (const a of nuovi) caps.add(a.capitolo);
  } else {
    nuovi = sorgenti.map((a) => {
      if (link && !a.clone) { a.clone = db.nuovoId(); caps.add(a.capitolo); }
      const { _chunk, ...d } = structuredClone(a);
      if (!link) delete d.clone;
      return d;
    });
  }
  // il blocco va subito sotto la riga scelta, nel suo capitolo e sottocapitolo; il capitolo viene rinumerato
  const fuori = new Set(nuovi);
  const lista = ordinaRighe(S.catalogo.filter((a) => a.capitolo === dove.capitolo && !fuori.has(a)), [dove.capitolo]);
  lista.splice(lista.indexOf(dove) + 1, 0, ...nuovi);
  for (const a of nuovi) {
    a.capitolo = dove.capitolo; a.sottocapitolo = dove.sottocapitolo || ""; a.rifCap = dove.rifCap || "";
    if (!app.taglio) a.k = `${a.capitolo}|${a.rifCap}|${codiceArticolo(a)}|${a.nOrdine}|${db.nuovoId()}`;
  }
  lista.forEach((a, i) => (a.ordine = i + 1));
  if (!app.taglio) for (const a of nuovi) { S.catalogo.push(a); S.perK.set(a.k, a); }
  try { await salvaCapitoli(caps); } catch (x) { return errore(x, "Incolla non riuscito"); }
  if (app.taglio) S.catApp = null;
  S.catSel = new Set(nuovi.map((a) => a.k)); S.catUltimo = nuovi[nuovi.length - 1].k;
  disegnaCatalogo();
  toast(`${nuovi.length} ${nuovi.length === 1 ? "articolo" : "articoli"} ${app.taglio ? "spostati" : link ? "incollati collegati (cloni)" : "incollati"} in ${dove.capitolo} · ${dove.sottocapitolo || "—"}.`, 4500);
}
function eliminaCat() {
  const arts = selCatOrdinati(); if (!arts.length) return;
  const conCloni = arts.filter((a) => cloniDi(a).length > 1).length;
  dialogo({
    titolo: arts.length === 1 ? "Eliminare l'articolo?" : `Eliminare ${arts.length} articoli?`,
    html: `<p>${arts.slice(0, 8).map((a) => `<b>${esc(codiceArticolo(a))}</b> · ${esc(a.nOrdine || "")}`).join("<br>")}${arts.length > 8 ? `<br>… e altri ${arts.length - 8}` : ""}</p>
      ${conCloni ? `<p class="muted small">${conCloni} sono cloni: si elimina solo la copia selezionata, le altre copie restano.</p>` : ""}
      <p class="muted small">Le offerte che li contengono non cambiano.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Elimina", value: "ok", cls: "danger" }],
    async onConferma() {
      const via = new Set(arts.map((a) => a.k)), caps = new Set(arts.map((a) => a.capitolo));
      S.catalogo = S.catalogo.filter((a) => !via.has(a.k));
      for (const k of via) S.perK.delete(k);
      // cloni rimasti da soli: tornano articoli normali
      const gruppi = new Map();
      for (const a of S.catalogo) if (a.clone) gruppi.set(a.clone, [...(gruppi.get(a.clone) || []), a]);
      for (const l of gruppi.values()) if (l.length === 1 && arts.some((x) => x.clone === l[0].clone)) { delete l[0].clone; caps.add(l[0].capitolo); }
      try { await salvaCapitoli(caps); } catch (x) { errore(x, "Eliminazione non riuscita"); return false; }
      S.catSel.clear(); if (S.catApp) S.catApp.ks = S.catApp.ks.filter((k) => !via.has(k));
      disegnaCatalogo();
      toast(`${arts.length} ${arts.length === 1 ? "articolo eliminato" : "articoli eliminati"} dal database.`);
    },
  });
}
$("#c-nuovo").addEventListener("click", () => {
  if (!S.capitoli.length) return toast("Prima crea un capitolo (Impostazioni › Ricarico standard) o importa il file Excel.");
  modificaArticolo(null);
});

/* Salva tutti gli articoli di un capitolo, divisi in blocchi sotto il limite di 1 MB di Firestore. */
function opsCapitolo(nome, articoli) {
  const puliti = articoli.slice().sort((a, b) => num(a.ordine) - num(b.ordine))
    .map(({ _chunk, _qta, _blocco0, _kCatalogo, ...a }) => a);
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
/* Menu a tendina dei GMA, filtrati per fornitore (elenco GMA della configurazione). */
const fornitoriGma = () => [...new Set(S.gmv.map((g) => (g.fornitore || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
const opzioniFornitoriGma = (sel = "") => `<option value="">— tutti i fornitori —</option>` + fornitoriGma().map((f) => `<option ${f === sel ? "selected" : ""}>${esc(f)}</option>`).join("");
function opzioniGma(forn = "", sel = "") {
  // una voce per ogni GMA (fornitore + codice): con "tutti i fornitori" lo stesso codice può comparire più volte
  const voci = S.gmv.filter((g) => !forn || (g.fornitore || "").trim() === forn).map((g) => ({ c: g.codice, f: (g.fornitore || "").trim() }))
    .sort((a, b) => a.c.localeCompare(b.c) || a.f.localeCompare(b.f));
  let scelto = false;
  const html = voci.map((v) => { const on = !scelto && v.c === sel; if (on) scelto = true; return `<option value="${esc(v.c)}" data-forn="${esc(v.f)}" ${on ? "selected" : ""}>${esc(v.c)}${forn ? "" : v.f ? " · " + esc(v.f) : ""}</option>`; }).join("");
  return `<option value="">— nessun GMA —</option>` + (sel && !scelto ? `<option value="${esc(sel)}" selected>${esc(sel)}</option>` : "") + html + `<option value="__nuovo">+ Nuovo GMA…</option>`;
}
/* Fornitore della voce scelta in un menu dei GMA (o quello del menu fornitore accanto). */
const fornScelto = (selGma, selForn) => selForn?.value || selGma.selectedOptions[0]?.dataset.forn || fornitore(selGma.value);
/* Aggiunge un GMA all'elenco (oracle_config/gmv). */
async function aggiungiGma(codice, forn, sconto = "") {
  codice = codice.trim().toUpperCase();
  if (!codice || S.gmv.some((g) => g.codice === codice && (g.fornitore || "").trim().toUpperCase() === forn.trim().toUpperCase())) return codice;
  const lista = [...S.gmv.map((g) => ({ ...g })), { codice, fornitore: forn.trim(), ...(sconto ? { sconto } : {}) }];
  await db.salva(COLL.cfg, "gmv", { lista });
  S.gmv = lista; impostaStandardGma(S.gmv);
  return codice;
}
/* ================= Immagini degli articoli =================
   Salvate a parte (collezione oracle_immagini) per non appesantire il database:
   "<id>_p" miniatura (~160 px) e "<id>_g" grande (~800 px). L'articolo tiene solo "img": id. */
const IMG = new Map(); // id_p / id_g → dataURL | Promise
function immagine(id, taglia = "p") {
  const k = `${id}_${taglia}`;
  if (!IMG.has(k)) IMG.set(k, db.leggi(COLL.img, k).then((d) => { const u = d?.u || ""; IMG.set(k, u); return u; }).catch(() => { IMG.delete(k); return ""; }));
  const v = IMG.get(k);
  return v instanceof Promise ? v : Promise.resolve(v);
}
function caricaMiniature(root) {
  for (const im of root.querySelectorAll("img[data-img]:not([src])")) immagine(im.dataset.img).then((u) => { if (u) im.src = u; else im.replaceWith(); });
}
async function riduciImmagine(blob, lato, qualita) {
  const url = URL.createObjectURL(blob);
  try {
    const im = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error("immagine non leggibile")); i.src = url; });
    const r = Math.min(1, lato / Math.max(im.naturalWidth, im.naturalHeight));
    const c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(im.naturalWidth * r)); c.height = Math.max(1, Math.round(im.naturalHeight * r));
    const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height); x.drawImage(im, 0, 0, c.width, c.height);
    return c.toDataURL("image/jpeg", qualita);
  } finally { URL.revokeObjectURL(url); }
}
// anteprima grande passando sopra una miniatura
document.addEventListener("mouseover", (e) => {
  const im = e.target.closest?.("img.img-mini[data-img]"), box = $("#img-anteprima");
  if (!im) { box.hidden = true; return; }
  const r = im.getBoundingClientRect();
  box.innerHTML = `<img src="${im.src}" alt="">`; box.hidden = false;
  box.style.left = Math.min(r.right + 10, innerWidth - 330) + "px"; box.style.top = Math.max(8, Math.min(r.top - 40, innerHeight - 330)) + "px";
  immagine(im.dataset.img, "g").then((u) => { if (u && !box.hidden) { const g = box.querySelector("img"); if (g) g.src = u; } });
});
/* Riquadro immagine nella scheda articolo: incolla (Ctrl+V), trascina o scegli un file. */
function legaRiquadroImg(b, a) {
  const zona = $("#m-img", b), vedi = () => {
    const u = b._img?.p || "";
    zona.querySelector(".img-vedi").innerHTML = u ? `<img src="${u}" alt="">` : b._img === null || !a.img ? `<span class="muted small">Nessuna immagine</span>` : `<img class="img-mini" data-img="${esc(a.img)}" alt="">`;
    caricaMiniature(zona);
    $("#m-img-togli", b).hidden = !(b._img?.p || (a.img && b._img !== null));
  };
  const usa = async (blob) => {
    if (!blob?.type?.startsWith("image/")) return toast("Questo non è un'immagine.");
    try { b._img = { p: await riduciImmagine(blob, 160, 0.8), g: await riduciImmagine(blob, 800, 0.82) }; vedi(); }
    catch (x) { toast("Immagine non leggibile: " + x.message); }
  };
  b.addEventListener("paste", (e) => {
    const it = [...(e.clipboardData?.items || [])].find((x) => x.type.startsWith("image/"));
    if (!it) return; e.preventDefault(); usa(it.getAsFile());
  });
  zona.addEventListener("dragover", (e) => { e.preventDefault(); zona.classList.add("sopra"); });
  zona.addEventListener("dragleave", () => zona.classList.remove("sopra"));
  zona.addEventListener("drop", (e) => { e.preventDefault(); zona.classList.remove("sopra"); const f = [...e.dataTransfer.files].find((x) => x.type.startsWith("image/")); if (f) usa(f); else toast("Trascina un file immagine."); });
  $("#m-img-file", b).addEventListener("change", (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) usa(f); });
  $("#m-img-togli", b).addEventListener("click", () => { b._img = null; vedi(); });
  vedi();
}
/* Salva l'immagine scelta nella scheda e aggiorna n.img */
async function salvaImmagineScheda(b, n) {
  if (b._img === undefined) return;
  const vecchia = n.img;
  if (b._img === null) delete n.img;
  else {
    const id = db.nuovoId();
    await db.scrivi([{ tipo: "set", coll: COLL.img, id: id + "_p", data: { u: b._img.p } }, { tipo: "set", coll: COLL.img, id: id + "_g", data: { u: b._img.g } }]);
    IMG.set(id + "_p", b._img.p); IMG.set(id + "_g", b._img.g);
    n.img = id;
  }
  // la vecchia immagine si cancella se nessun altro articolo (fuori dai cloni di questo) la usa
  if (vecchia && vecchia !== n.img && !S.catalogo.some((x) => x.img === vecchia && x.k !== n.k && !(n.clone && x.clone === n.clone)))
    db.scrivi([{ tipo: "del", coll: COLL.img, id: vecchia + "_p" }, { tipo: "del", coll: COLL.img, id: vecchia + "_g" }]).catch(() => {});
}

/* Menu a tendina dei sottocapitoli (stesso aspetto degli altri menu) con la voce "+ Nuovo sottocapitolo…". */
function opzioniSotto(sotti, sel = "") {
  const l = [...new Set(sotti)].filter(Boolean);
  if (sel && !l.includes(sel)) l.unshift(sel);
  return `<option value="">— nessun sottocapitolo —</option>` + l.map((x) => `<option ${x === sel ? "selected" : ""}>${esc(x)}</option>`).join("") + `<option value="__nuovo">+ Nuovo sottocapitolo…</option>`;
}
const campoSotto = (id, sel, sotti) => `<select id="${id}">${opzioniSotto(sotti, sel)}</select><input id="${id}-nuovo" placeholder="nome del nuovo sottocapitolo" hidden style="margin-top:6px">`;
function legaSotto(id, root) {
  $("#" + id, root).addEventListener("change", (e) => { const n = e.target.value === "__nuovo"; $(`#${id}-nuovo`, root).hidden = !n; if (n) setTimeout(() => $(`#${id}-nuovo`, root).focus(), 30); });
}
const aggiornaSotto = (id, root, sotti, tieni = false) => { const s = $("#" + id, root), v = s.value === "__nuovo" ? "" : s.value; s.innerHTML = opzioniSotto(sotti, tieni || sotti.includes(v) ? v : ""); $(`#${id}-nuovo`, root).hidden = true; };
const valSotto = (id, root) => { const v = $("#" + id, root).value; return (v === "__nuovo" ? $(`#${id}-nuovo`, root).value : v).trim(); };
const sottoDelCap = (cap) => [...new Set(S.catalogo.filter((x) => x.capitolo === cap).map((x) => x.sottocapitolo).filter(Boolean))];

/* ================= Cloni =================
   Un clone è la stessa riga del database messa in più punti (es. una base sotto ogni rivelatore compatibile).
   I cloni hanno lo stesso "clone" (id del gruppo): tutto è in comune tranne la posizione. */
const CAMPI_POSTO = ["capitolo", "sottocapitolo", "rifCap", "ordine", "k", "_chunk", "clone"];
const cloniDi = (a) => (a?.clone ? S.catalogo.filter((x) => x.clone === a.clone) : []);
/* Copia i dati di "src" su tutti i suoi cloni. Restituisce i capitoli da salvare. */
function propagaClone(src) {
  const caps = new Set();
  for (const x of cloniDi(src)) {
    if (x === src) continue;
    for (const key of Object.keys(x)) if (!CAMPI_POSTO.includes(key) && !(key in src)) delete x[key];
    for (const [key, v] of Object.entries(src)) if (!CAMPI_POSTO.includes(key)) x[key] = structuredClone(v);
    caps.add(x.capitolo);
  }
  return caps;
}
const salvaCapitoli = (caps) => db.scrivi([...new Set(caps)].filter(Boolean).flatMap((c) => opsCapitolo(c, S.catalogo.filter((a) => a.capitolo === c))));
function creaClone(a) {
  const opzDopo = (cap) => {
    const arts = ordinaRighe(S.catalogo.filter((x) => x.capitolo === cap), [cap]);
    let html = "", g = null;
    for (const x of arts) {
      const sg = x.sottocapitolo || "—";
      if (sg !== g) { if (g !== null) html += "</optgroup>"; html += `<optgroup label="${esc(sg)}">`; g = sg; }
      html += `<option value="${esc(x.k)}" ${x.k === a.k ? "disabled" : ""}>${esc(codiceArticolo(x))} · ${esc(x.nOrdine || "")} — ${esc(String(x.descrizione || "").slice(0, 60))}</option>`;
    }
    return html + (g !== null ? "</optgroup>" : "") + `<option value="">— in fondo al capitolo —</option>`;
  };
  dialogo({
    titolo: `Clona ${codiceArticolo(a)} · ${a.nOrdine || ""}`,
    html: `<p class="muted small" style="margin-top:0">Il clone è <b>la stessa riga collegata</b>: compare anche nel punto che scegli, e quando modifichi una copia
        (prezzo, sconti, descrizione, GMA…) cambiano tutte. Ogni copia ha solo la sua posizione.</p>
      <label class="f">Capitolo<select id="cl-cap">${S.capitoli.map((c) => `<option ${c.nome === a.capitolo ? "selected" : ""}>${esc(c.nome)}</option>`).join("")}</select></label>
      <label class="f" style="margin-top:8px">Mettilo subito sotto a<select id="cl-dopo">${opzDopo(a.capitolo)}</select></label>
      <p class="muted small">Il clone prende il sottocapitolo dell'articolo che scegli.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Crea clone", value: "ok", cls: "primary" }],
    onApri(d) {
      $("#cl-cap", d).addEventListener("change", (e) => { $("#cl-dopo", d).innerHTML = opzDopo(e.target.value); });
      $("#cl-dopo", d).value = "";
    },
    async onConferma(v, d) {
      const cap = $("#cl-cap", d).value, dopo = S.perK.get($("#cl-dopo", d).value);
      const arts = ordinaRighe(S.catalogo.filter((x) => x.capitolo === cap), [cap]);
      let ordine, sotto;
      if (dopo) {
        const i = arts.indexOf(dopo), succ = arts[i + 1];
        ordine = succ ? (num(dopo.ordine) + num(succ.ordine)) / 2 : num(dopo.ordine) + 0.01;
        if (succ && !(ordine > num(dopo.ordine) && ordine < num(succ.ordine))) { arts.forEach((x, j) => (x.ordine = j + 1)); ordine = num(dopo.ordine) + 0.5; }
        sotto = dopo.sottocapitolo || "";
      } else {
        ordine = (arts.length ? Math.max(...arts.map((x) => num(x.ordine))) : 0) + 1;
        sotto = arts[arts.length - 1]?.sottocapitolo || "";
      }
      if (!a.clone) a.clone = db.nuovoId();
      const { _chunk, ...dati } = structuredClone(a);
      const n = { ...dati, capitolo: cap, sottocapitolo: sotto, rifCap: (dopo || arts[0])?.rifCap || "", ordine };
      n.k = `${cap}|${n.rifCap}|${codiceArticolo(n)}|${n.nOrdine}|${db.nuovoId()}`;
      S.catalogo.push(n); S.perK.set(n.k, n);
      try { await salvaCapitoli([a.capitolo, cap]); } catch (x) { errore(x, "Clone non salvato"); return false; }
      toast(`Clone creato in ${cap} · ${sotto || "—"}. Ora ci sono ${cloniDi(a).length} copie collegate.`, 5000);
    },
  });
}

function modificaArticolo(a) {
  const nuovo = !a;
  const capDef = $("#c-capitolo").value || S.capitoli[0]?.nome || "";
  a = a || { capitolo: capDef, sottocapitolo: $("#c-sotto").value || "", precodice: "", codice: "", nOrdine: "", descrizione: "", um: "n.", listinoAcquisto: 0, sc3: 0, sc4: 0, sc5: 0, aum1: 0, riparam: 0, giorni: 0, gmv: "", dataAgg: oggiIso() };
  const auto = !num(a.listino) || Math.abs(num(a.listino) - listinoCalcolato(a)) < 0.005;
  const sotti = [...new Set(S.catalogo.filter((x) => x.capitolo === a.capitolo).map((x) => x.sottocapitolo).filter(Boolean))];
  dialogo({
    titolo: nuovo ? "Nuovo articolo" : `${codiceArticolo(a)} · ${a.nOrdine}`,
    html: `${cloniDi(a).length > 1 ? `<div class="avviso info small" style="margin-top:0">🔗 <b>Articolo clonato: ${cloniDi(a).length} copie collegate</b>
        (${cloniDi(a).map((x) => `${esc(x.capitolo)} · ${esc(x.sottocapitolo || "—")}${x === a ? " <i>(questa)</i>" : ""}`).join("; ")}).
        Le modifiche valgono per tutte; capitolo e sottocapitolo sono solo di questa copia.</div>` : ""}<div class="grid-set">
      <label class="f">Capitolo<select id="m-cap">${S.capitoli.map((c) => `<option ${c.nome === a.capitolo ? "selected" : ""}>${esc(c.nome)}</option>`).join("")}</select></label>
      <label class="f wide" title="Scegline uno dall'elenco o creane uno nuovo">Sottocapitolo${campoSotto("m-sotto", a.sottocapitolo || "", sotti)}</label>
      <label class="f">Precodice<input id="m-pre" value="${esc(a.precodice)}"></label>
      <label class="f">Codice<input id="m-cod" value="${esc(a.codice)}"></label>
      <label class="f wide">N° di ordinazione<input id="m-ord" value="${esc(a.nOrdine)}"></label>
      <label class="f">U.M.<input id="m-um" value="${esc(a.um)}"></label>
      <label class="f">Fornitore<select id="m-forn">${opzioniFornitoriGma(fornitoreDi(a))}</select></label>
      <label class="f">GMA<select id="m-gmv">${opzioniGma(fornitoreDi(a), a.gmv)}</select></label>
      <div class="f wide nuovo-gma" id="m-nuovo-gma" hidden><span>Nuovo GMA</span><div class="row" style="gap:6px">
        <input id="m-ng-cod" placeholder="codice GMA" style="width:120px"><input id="m-ng-forn" list="m-ng-forns" placeholder="fornitore" class="grow">
        <datalist id="m-ng-forns">${fornitoriGma().map((f) => `<option value="${esc(f)}">`).join("")}</datalist></div></div>
      ${campoNum("m-la", "Listino acquisto €", a.listinoAcquisto)}
      ${campoNum("m-gg", "Giorni (attività)", a.giorni)}
      ${campoNum("m-sc3", "Sconto 1", a.sc3, true)}${campoNum("m-sc4", "Sconto 2", a.sc4, true)}
      ${campoNum("m-sc5", "Sconto 3", a.sc5, true)}${campoNum("m-aum1", "Aumento 1", a.aum1, true)}
      <p class="muted small wide" id="m-std-nota" style="margin:-4px 0 0"></p>
      <div hidden>${campoNum("m-li", "Prezzo di listino €", a.listino)}</div>
      <label class="f">Data aggiornamento<input id="m-data" type="date" value="${esc(a.dataAgg || "")}"></label>
    </div>
    <label class="row small" style="margin-top:8px"><input type="checkbox" id="m-gen" ${isGenerico(a) ? "checked" : ""}>
      <b>Tratta come articolo generico</b> — in Oracle: ${CODICE_GENERICO}, prezzo netto, n° di ordinazione nella nota pre-linea e descrizione nella post-linea</label>
    <p class="muted small" style="margin:4px 0 0">Se cambi un prezzo o uno sconto, la data di aggiornamento diventa quella di oggi (puoi correggerla).</p>
    <label class="row small" style="margin-top:8px" hidden><input type="checkbox" id="m-auto" ${auto ? "checked" : ""}>
      Prezzo di listino calcolato automaticamente (come la colonna K dell'Excel)</label>
    <div class="m-img-riga" style="margin-top:10px">
      <div class="f m-img-box"><span>Immagine</span>
        <div class="img-zona" id="m-img" tabindex="0" title="Incolla un'immagine (Ctrl+V), trascinala qui o scegli un file">
          <div class="img-vedi"></div>
          <div class="img-az"><label class="btn small">Scegli file…<input type="file" id="m-img-file" accept="image/*" hidden></label>
            <button type="button" class="btn small" id="m-img-togli" hidden>Togli</button></div>
          <span class="muted small">oppure incolla (Ctrl+V) o trascina qui</span>
        </div></div>
      <label class="f" style="flex:1">Descrizione<textarea id="m-des" rows="5">${esc(a.descrizione)}</textarea></label>
    </div>
    <div class="grid-set" style="margin-top:10px" id="m-note-box">
      <label class="f wide">Note pre-linea<textarea id="m-npre" rows="2">${esc(a.notePre)}</textarea></label>
      <label class="f wide">Note post-linea<textarea id="m-npost" rows="2">${esc(a.notePost)}</textarea></label>
    </div>
    <p class="muted small" id="m-note-gen">Articolo generico (P000 ARTICOLO_GEN): in Oracle la nota pre-linea è il n° di ordinazione e la nota post-linea è la descrizione.</p>`,
    pulsanti: [...(nuovo ? [] : [{ label: cloniDi(a).length > 1 ? "Elimina questa copia" : "Elimina articolo", value: "del", cls: "danger" },
        ...(cloniDi(a).length > 1 ? [{ label: "Scollega questa copia", value: "scollega" }] : []),
        { label: "🔗 Crea clone…", value: "clone" }]),
      { label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(b) {
      legaRiquadroImg(b, a || {});
      legaSotto("m-sotto", b);
      $("#m-cap", b).addEventListener("change", (e) => aggiornaSotto("m-sotto", b, sottoDelCap(e.target.value)));
      // il fornitore riduce l'elenco dei GMA; scegliendo un GMA il fornitore si allinea
      $("#m-forn", b).addEventListener("change", (e) => {
        const att = $("#m-gmv", b).value, f = e.target.value;
        $("#m-gmv", b).innerHTML = opzioniGma(f, !f || voceGma(att, f) ? att : "");
        $("#m-nuovo-gma", b).hidden = true;
      });
      $("#m-gmv", b).addEventListener("change", (e) => {
        const v = e.target.value;
        $("#m-nuovo-gma", b).hidden = v !== "__nuovo";
        if (v === "__nuovo") { $("#m-ng-forn", b).value = $("#m-forn", b).value; setTimeout(() => $("#m-ng-cod", b).focus(), 30); return; }
        const f = $("#m-gmv", b).selectedOptions[0]?.dataset.forn || "";
        if (f && $("#m-forn", b).value !== f) { $("#m-forn", b).innerHTML = opzioniFornitoriGma(f); $("#m-gmv", b).innerHTML = opzioniGma(f, v); }
      });
      const sync = () => {
        const li = $("#m-li", b);
        li.disabled = $("#m-auto", b).checked;
        if (li.disabled) li.value = String(Math.round(listinoCalcolato(leggi(b)) * 10000) / 10000).replace(".", ",");
      };
      const gen = () => { const g = isGenerico({ codice: $("#m-cod", b).value, generico: $("#m-gen", b).checked }); $("#m-note-box", b).hidden = g; $("#m-note-gen", b).hidden = !g; };
      // sconti vuoti = acquisto standard del GMA: si mostrano come suggerimento grigio
      const stdUpd = () => {
        const g = $("#m-gmv", b).value, st = standardGma(g, fornScelto($("#m-gmv", b), $("#m-forn", b)));
        const vals = st ? [st.sc1, st.sc2, st.sc3, st.aum] : [];
        ["#m-sc3", "#m-sc4", "#m-sc5", "#m-aum1"].forEach((id, i) => { $(id, b).placeholder = st && num(vals[i]) ? pctVal(vals[i]) : ""; });
        const propri = ["#m-sc3", "#m-sc4", "#m-sc5", "#m-aum1"].some((id) => num($(id, b).value));
        $("#m-std-nota", b).innerHTML = !g ? "" : !st ? `Il GMA ${esc(g)} non ha acquisti standard (Impostazioni › Acquisti standard).`
          : propri ? `<b>Sconti propri dell'articolo</b> (eccezione): lo standard di ${esc(g)} è ${esc(testoStd(st))}. Svuota i campi per usare lo standard.`
          : `Vale l'acquisto standard di ${esc(g)}: <b>${esc(testoStd(st))}</b> (in grigio). Scrivi un valore solo se questo articolo fa eccezione.`;
      };
      $("#m-gmv", b).addEventListener("change", stdUpd);
      b.addEventListener("input", (e) => { if (e.target.matches("#m-sc3, #m-sc4, #m-sc5, #m-aum1")) stdUpd(); });
      stdUpd();
      // data automatica: quando cambia un prezzo o uno sconto, la data di aggiornamento diventa oggi
      const prezzi = ["#m-la", "#m-li", "#m-sc3", "#m-sc4", "#m-sc5", "#m-aum1"];
      const iniziali = prezzi.map((id) => $(id, b).value);
      let dataToccata = false;
      $("#m-data", b).addEventListener("input", () => (dataToccata = true));
      b.addEventListener("input", (e) => {
        sync(); gen();
        if (!dataToccata && prezzi.some((id) => e.target.matches(id))) {
          const cambiato = prezzi.some((id, i) => $(id, b).value !== iniziali[i]);
          $("#m-data", b).value = cambiato ? oggiIso() : (a.dataAgg || "");
        }
      });
      b.addEventListener("change", (e) => { if (e.target.id === "m-gen") gen(); });
      sync(); gen();
    },
    async onConferma(v, b) {
      const vecchioCap = a.capitolo;
      if (v === "clone") { setTimeout(() => creaClone(S.perK.get(a.k) || a), 50); return; }
      if (v === "scollega" || v === "del") {
        const caps = [vecchioCap];
        const resto = cloniDi(a).filter((x) => x.k !== a.k);
        if (resto.length === 1) { delete resto[0].clone; caps.push(resto[0].capitolo); }
        if (v === "del") S.catalogo = S.catalogo.filter((x) => x.k !== a.k);
        else delete S.perK.get(a.k).clone;
        await salvaCapitoli(caps);
        toast(v === "del" ? (resto.length ? "Copia eliminata: le altre copie restano." : "Articolo eliminato. Le offerte che lo contengono non cambiano.") : "Copia scollegata: ora è un articolo a sé.");
        return;
      }
      const n = leggi(b);
      if (!n.codice && !n.nOrdine) { toast("Scrivi almeno il codice o il n° di ordinazione."); return false; }
      if (n.gmv === "__nuovo") {
        const cod = $("#m-ng-cod", b).value.trim();
        if (!cod) { toast("Scrivi il codice del nuovo GMA."); $("#m-ng-cod", b).focus(); return false; }
        n.gmv = await aggiungiGma(cod, $("#m-ng-forn", b).value);
      }
      if (!n.listino || $("#m-auto", b).checked) n.listino = listinoCalcolato(n);
      try { await salvaImmagineScheda(b, n); } catch (x) { errore(x, "Immagine non salvata (controlla le regole di Firestore: serve oracle_immagini)"); return false; }
      if (nuovo) {
        const stessi = S.catalogo.filter((x) => x.capitolo === n.capitolo && (!n.sottocapitolo || x.sottocapitolo === n.sottocapitolo));
        n.ordine = (stessi.length ? Math.max(...stessi.map((x) => num(x.ordine))) : 0) + 0.01;
        n.rifCap = stessi[0]?.rifCap || "";
        n.k = `${n.capitolo}|${n.rifCap}|${n.precodice}${n.codice}|${n.nOrdine}|${db.nuovoId()}`;
        S.catalogo.push(n);
      } else {
        Object.assign(S.perK.get(a.k), n);
      }
      const altri = nuovo ? new Set() : propagaClone(S.perK.get(a.k));
      await salvaCapitoli([n.capitolo, vecchioCap, ...altri]);
      toast(nuovo ? "Articolo aggiunto" : altri.size || cloniDi(a).length > 1 ? `Articolo salvato in tutte le ${cloniDi(a).length} copie collegate.` : "Articolo salvato");
    },
  });
  function leggi(b) {
    const g = (id) => $(id, b).value;
    return {
      ...a,
      capitolo: g("#m-cap"), sottocapitolo: valSotto("m-sotto", b), precodice: g("#m-pre").trim(), codice: g("#m-cod").trim(),
      nOrdine: g("#m-ord").trim(), um: g("#m-um").trim(), gmv: g("#m-gmv").trim(), fornitore: fornScelto($("#m-gmv", b), $("#m-forn", b)),
      listinoAcquisto: num(g("#m-la")), giorni: num(g("#m-gg")),
      sc3: daPct(g("#m-sc3")), sc4: daPct(g("#m-sc4")), sc5: daPct(g("#m-sc5")), aum1: daPct(g("#m-aum1")), riparam: num(a.riparam) || 0,
      listino: num(g("#m-li")), dataAgg: g("#m-data"), descrizione: g("#m-des"), notePre: g("#m-npre"), notePost: g("#m-npost"),
      generico: $("#m-gen", b).checked,
    };
  }
}

/* Regole di scadenza dei prezzi per precodice */
$("#c-scadenze").addEventListener("click", () => {
  const conta = new Map();
  for (const a of S.catalogo) {
    const p = String(a.precodice || "").trim().toUpperCase();
    if (!p) continue;
    const x = conta.get(p) || { n: 0, gen: 0 };
    x.n++; if (isGenerico(a)) x.gen++;
    conta.set(p, x);
  }
  for (const r of S.precodici) if (!conta.has(String(r.precodice).toUpperCase())) conta.set(String(r.precodice).toUpperCase(), { n: 0, gen: 0 });
  const elenco = [...conta.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const reg = (p) => S.precodici.find((x) => String(x.precodice).toUpperCase() === p) || {};
  dialogo({
    titolo: "Scadenza dei prezzi per precodice",
    largo: true,
    html: `<p class="muted small" style="margin-top:0"><b>Prezzi a richiesta</b>: il prezzo scade dopo i mesi indicati dalla data di aggiornamento.
      <b>Listino</b>: indica quando è previsto il prossimo cambio listino; un mese prima gli articoli vengono segnalati (possibile aumento)
      e dopo il cambio sono "da aggiornare" finché non aggiorni il prezzo. Senza regola vale la legenda: verde meno di 1 mese, giallo 1–2 mesi, rosso oltre.</p>
      <div class="tbl-wrap" style="max-height:55vh"><table><thead><tr><th>Precodice</th><th class="num">Articoli</th><th>Fornitore / note</th><th>Regola</th><th>Validità (mesi)</th><th>Prossimo cambio listino</th></tr></thead><tbody>
      ${elenco.map(([p, x]) => { const r = reg(p); return `<tr data-p="${esc(p)}"><td class="mono"><b>${esc(p)}</b></td>
        <td class="num">${x.n}${x.gen ? ` <span class="badge warn" title="di cui trattati come generici">${x.gen} GEN</span>` : ""}</td>
        <td><input data-r="descrizione" value="${esc(r.descrizione || "")}" style="width:150px"></td>
        <td><select data-r="tipo"><option value="">— nessuna —</option><option value="richiesta" ${r.tipo === "richiesta" ? "selected" : ""}>prezzi a richiesta</option><option value="listino" ${r.tipo === "listino" ? "selected" : ""}>listino</option></select></td>
        <td><input data-r="validitaMesi" inputmode="decimal" value="${esc(r.validitaMesi ?? "")}" style="width:70px" placeholder="es. 3"></td>
        <td><input data-r="cambioListino" type="date" value="${esc(r.cambioListino || "")}"></td></tr>`; }).join("")}
      </tbody></table></div>
      <label class="f" style="margin-top:10px">Aggiungi un precodice<input id="m-nuovopre" placeholder="es. VAN" style="width:120px"></label>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    async onConferma(v, b) {
      const lista = $$("tr[data-p]", b).map((tr) => {
        const g = (k) => $(`[data-r="${k}"]`, tr).value.trim();
        return { precodice: tr.dataset.p, descrizione: g("descrizione"), tipo: g("tipo"), validitaMesi: num(g("validitaMesi")) || "", cambioListino: g("cambioListino") };
      }).filter((r) => r.tipo || r.descrizione);
      const nuovo = $("#m-nuovopre", b).value.trim().toUpperCase();
      if (nuovo && !lista.some((r) => r.precodice === nuovo)) lista.push({ precodice: nuovo, descrizione: "", tipo: "", validitaMesi: "", cambioListino: "" });
      await db.salva(COLL.cfg, "precodici", { lista });
      toast("Regole salvate");
    },
  });
});

$("#c-impost").addEventListener("click", () => {
  dialogo({
    titolo: "Ricarico standard per capitolo",
    html: `<p class="muted small" style="margin-top:0">Il ricarico automatico si applica alle offerte “VENDITA CLIENTE” quando non è indicato un ricarico nell'offerta o nel GMA.</p>
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
    <td class="small">${Object.entries(c.sconti || {}).map(([g, v]) => `<span class="badge info">${esc(g)} ${c.scontiTesto?.[g] ? esc(c.scontiTesto[g]) + "%" : pctTxt(v)}</span>`).join(" ") || '<span class="muted">nessuno</span>'}</td>
    <td class="small nowrap">${c.censito ? dataIt(c.censito) : '<span class="muted">—</span>'}</td></tr>`).join("")
    || `<tr><td colspan="5" class="vuoto">Nessun cliente. Importa il file Excel, aggiungine uno o carica un censimento da Oracle.</td></tr>`;
}
$("#righe-clienti").addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-id]");
  if (tr) modificaCliente(S.clienti.find((c) => c.id === tr.dataset.id));
});
$("#cl-nuovo").addEventListener("click", () => modificaCliente(null));
/* GMA raggruppati per fornitore (in ordine alfabetico; quelli senza fornitore in fondo). */
function perFornitore(gmvs) {
  const m = new Map();
  for (const g of gmvs) { const f = (fornitoriTxt(g) || "").trim().toUpperCase() || "ALTRI"; if (!m.has(f)) m.set(f, []); m.get(f).push(g); }
  for (const l of m.values()) l.sort();
  return [...m.entries()].sort((a, b) => (a[0] === "ALTRI") - (b[0] === "ALTRI") || a[0].localeCompare(b[0]));
}
function modificaCliente(c) {
  const nuovo = !c;
  c = c || { nome: "", codice: "", att: "", sconti: {} };
  const gmvs = [...new Set([...S.gmv.map((g) => g.codice), ...Object.keys(c.sconti || {})])].filter((g) => String(g).toUpperCase() !== "ATT");
  dialogo({
    largo: true, titolo: nuovo ? "Nuovo cliente" : c.nome,
    html: `<div class="grid-set">
      <label class="f wide">Ragione sociale<input id="k-nome" value="${esc(c.nome)}" required></label>
      <label class="f">Codice cliente<input id="k-cod" value="${esc(c.codice)}"></label>
      <label class="f">Tariffa attività tecnica €<input id="k-att" inputmode="decimal" value="${esc(c.att === "" ? "" : String(c.att).replace(".", ","))}"></label>
    </div>
    <h3 style="margin:14px 0 8px">Sconti di cartellino per GMA</h3>
    <p class="muted small" style="margin:-4px 0 8px">Scrivi <b>NC</b> per un GMA non censito: nel controllo dell'offerta comparirà "Non censito" invece di OK o fuori cartellino.</p>
    <div class="gmv-colonne">${perFornitore(gmvs).map(([forn, gs]) => `<div class="gmv-forn"><div class="gmv-forn-tit">${esc(forn)}</div><div class="lista-gmv">${gs.map((g) => `<label><span><b class="mono">${esc(g)}</b></span>
      <span class="pct-wrap"><input class="pct" data-g="${esc(g)}" value="${esc(scontoCliTxt(c, g))}" ${(() => { const o = gmaOracle(g); return o !== g ? `placeholder="= ${esc(o)} ${esc(scontoCliTxt(c, o) || "")}" title="In Oracle è ${esc(o)}: vale il suo cartellino. Scrivi un valore solo per fare eccezione."` : 'placeholder="es. 45+10 · NC = non censito"'; })()}></span></label>`).join("")}</div></div>`).join("")}</div>`,
    pulsanti: [...(nuovo ? [] : [{ label: "Elimina cliente", value: "del", cls: "danger" }]),
      { label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    async onConferma(v, b) {
      if (v === "del") { await db.elimina(COLL.cli, c.id); toast("Cliente eliminato"); return; }
      const nome = $("#k-nome", b).value.trim();
      if (!nome) { toast("Scrivi la ragione sociale."); return false; }
      const sconti = {}, scontiTesto = {};
      $$("[data-g]", b).forEach((i) => {
        // NC = non censito: il controllo cartellino dice "Non censito" invece di OK / fuori cartellino
        if (/^\s*n\.?\s*c\.?\s*$/i.test(i.value)) { sconti[i.dataset.g] = "NC"; return; }
        const x = cascata(i.value); if (x.f) { sconti[i.dataset.g] = x.f; if (x.testo.includes("+")) scontiTesto[i.dataset.g] = x.testo; } });
      const attTxt = $("#k-att", b).value.trim();
      const id = c.id || slug(nome) + "-" + db.nuovoId().slice(-4);
      await db.salva(COLL.cli, id, { ...(c.censito ? { censito: c.censito } : {}), nome, codice: $("#k-cod", b).value.trim(), att: attTxt === "" ? "" : num(attTxt), sconti, scontiTesto });
      toast("Cliente salvato");
    },
  });
}

/* ================= Censimento dei cartellini da Oracle =================
   1. "Offerta per il censimento": un file da importare in Oracle con un articolo per GMA a quantità 1, senza prezzi
      né sconti, così Oracle applica le condizioni del cliente dell'offerta.
   2. Si scaricano da Oracle le linee di quell'offerta e si caricano qui: per ogni GMA si leggono gli sconti che Oracle
      ha applicato (Sconto GMA, Sconto Cliente, Sconto 3/4/5) e diventano il cartellino del cliente (anche in cascata). */
const cascata = (t) => { const tt = BIB.terminiSconto(t); return { f: tt.length ? 1 - BIB.fattoreDi(tt) : 0, testo: tt.map((x) => String(Math.round(x * 1e4) / 100).replace(".", ",")).join("+") }; };
const scontoCliTxt = (c, g) => (c.sconti?.[g] === "NC" ? "NC" : c.scontiTesto?.[g] || pctVal(c.sconti?.[g]));
/* Righe dell'offerta di censimento: un GMA per riga, con l'articolo proposto (o quello scelto l'ultima volta).
   Si può cambiare articolo, scrivere a mano un codice Oracle, escludere un GMA o aggiungerne. */
function righeCensimento() {
  const salvate = S.censCfg?.righe;
  const gmas = [...new Set([...S.gmv.map((g) => g.codice), ...S.catalogo.map((a) => a.gmv)].filter(Boolean).map((g) => (String(g).toUpperCase() === "ATT" ? "ATT" : g)))].sort();
  const candidati = (g) => g === "ATT" ? S.catalogo.filter((x) => isAttivita(x))
    : S.catalogo.filter((x) => x.gmv === g && !isGenerico(x) && !isAttivita(x) && codiceArticolo(x)).sort((x, y) => String(y.dataAgg || "").localeCompare(String(x.dataAgg || "")));
  const out = gmas.map((g) => {
    const c = candidati(g), prima = salvate?.find((x) => x.g === g);
    const cod = prima ? prima.cod : c[0] ? codiceOracle(c[0]) : "";
    // un GMA che in Oracle ha un altro codice (es. SENSITRON → SIFSFI) non serve nel censimento: vale quello di Oracle
    return { g, cod, on: prima ? prima.on !== false : !!cod && gmaOracle(g) === g, cand: [...new Set(c.map((x) => codiceOracle(x)))].slice(0, 40) };
  });
  for (const x of salvate || []) if (!out.some((y) => y.g === x.g)) out.push({ g: x.g, cod: x.cod, on: x.on !== false, cand: [] });
  return out;
}
$("#cl-cens-scarica").addEventListener("click", () => {
  const righe = righeCensimento();
  if (!righe.length) return toast("Non ci sono GMA: importa prima il file Excel.");
  const descr = (cod) => { if (!cod) return ""; const a = S.catalogo.find((x) => codiceOracle(x) === cod); return a ? String(a.descrizione || a.nOrdine || "").slice(0, 70) : cod ? "codice scritto a mano" : ""; };
  const disegna = (d) => {
    const n = righe.filter((x) => x.on && x.cod).length;
    $("#ce-righe", d).innerHTML = righe.map((x, i) => `<tr class="${x.on && x.cod ? "" : "qta-zero"}">
      <td><input type="checkbox" data-ce-on="${i}" ${x.on ? "checked" : ""} title="Includi questo GMA"></td>
      <td class="mono"><b>${esc(x.g)}</b></td><td class="small">${esc(fornitoriTxt(x.g))}</td>
      <td><input class="mono" data-ce-cod="${i}" list="ce-l${i}" value="${esc(x.cod)}" placeholder="codice articolo Oracle" style="width:220px">
        <datalist id="ce-l${i}">${x.cand.map((c) => `<option value="${esc(c)}">`).join("")}</datalist></td>
      <td class="small" data-ce-des="${i}">${esc(descr(x.cod))}${x.on && !x.cod ? '<span class="badge warn">scrivi un codice</span>' : ""}</td></tr>`).join("");
    if ($("#ce-n")) $("#ce-n").textContent = n;
  };
  dialogo({
    titolo: "Offerta per il censimento dei cartellini", largo: true,
    html: `<p style="margin-top:0">Un articolo per ogni GMA, a quantità 1, <b>senza prezzi né sconti</b>: importato in un'offerta Oracle intestata al cliente,
        Oracle applica le sue condizioni. Poi scarichi le linee dell'offerta da Oracle e le carichi qui con <b>Carica censimento</b>.</p>
      <p class="muted small" style="margin:-4px 0 8px">Puoi cambiare l'articolo di ogni GMA (scegli dall'elenco o scrivi qualunque codice Oracle), togliere la spunta ai GMA da non censire
        e aggiungere GMA. Le scelte restano salvate per i prossimi clienti.</p>
      <div class="tbl-wrap" style="max-height:50vh"><table><thead><tr><th></th><th>GMA</th><th>Fornitore</th><th>Articolo (codice Oracle)</th><th>Descrizione</th></tr></thead>
        <tbody id="ce-righe"></tbody></table></div>
      <div class="card row" style="margin-top:10px;gap:8px;align-items:flex-end;flex-wrap:wrap"><b style="align-self:center">+ GMA</b>
        <label class="f">GMA<input id="ce-ng" placeholder="es. HESA" style="width:120px"></label>
        <label class="f grow">Codice articolo Oracle<input id="ce-nc" class="mono" placeholder="es. HESXXXX"></label>
        <button type="button" class="btn" id="ce-add">Aggiungi</button></div>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva le scelte", value: "salva" }, { label: "Scarica il file per Oracle", value: "ok", cls: "primary" }],
    onApri(d) {
      d.addEventListener("change", (e) => {
        const t = e.target;
        if (t.dataset.ceOn !== undefined) { righe[+t.dataset.ceOn].on = t.checked; disegna(d); }
        else if (t.dataset.ceCod !== undefined) { const x = righe[+t.dataset.ceCod]; x.cod = t.value.trim().toUpperCase(); if (x.cod) x.on = true; disegna(d); }
      });
      d.addEventListener("click", (e) => {
        if (e.target.id !== "ce-add") return;
        const g = $("#ce-ng", d).value.trim().toUpperCase(), cod = $("#ce-nc", d).value.trim().toUpperCase();
        if (!g || !cod) return toast("Scrivi GMA e codice.");
        const x = righe.find((y) => y.g === g);
        if (x) { x.cod = cod; x.on = true; } else righe.push({ g, cod, on: true, cand: [] });
        $("#ce-ng", d).value = ""; $("#ce-nc", d).value = ""; disegna(d);
      });
      for (const id of ["#ce-ng", "#ce-nc"]) $(id, d).addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); $("#ce-add", d).click(); } });
      $("#dlg-foot").insertAdjacentHTML("afterbegin", `<span class="muted small grow" style="align-self:center"><b id="ce-n"></b> righe nel file</span>`);
      disegna(d);
    },
    async onConferma(v) {
      const salva = righe.map((x) => ({ g: x.g, cod: x.cod, on: x.on }));
      try { await db.salva(COLL.cfg, "censimento", { righe: salva }); S.censCfg = { righe: salva }; } catch (x) { errore(x, "Scelte non salvate"); }
      if (v === "salva") { toast("Scelte salvate."); return; }
      const scelte = righe.filter((x) => x.on && x.cod);
      if (!scelte.length) { toast("Nessuna riga da scaricare."); return false; }
      const out = scelte.map((x, i) => {
        const r = new Array(COLONNE_ORACLE.length).fill("");
        r[0] = (i + 1) * 10; r[1] = x.cod; r[2] = "BG"; r[3] = 1; r[16] = "GMA " + x.g;
        return r;
      });
      scarica(scriviExcel([COLONNE_ORACLE, ...out], "Sheet1"), `CENSIMENTO_CARTELLINI_${oggiIso()}.xlsx`);
    },
  });
});

/* Linee di un'offerta Oracle → sconti per GMA. */
function scontiDaLineeOracle(righe) {
  const norm = (s) => String(s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  const hi = righe.findIndex((r) => r && r.some((c) => norm(c).includes("codice articolo")));
  if (hi < 0) throw new Error('Non trovo la colonna "Codice Articolo Fornitore": scarica da Oracle le linee dell\'offerta con le intestazioni.');
  const h = righe[hi].map(norm);
  const col = (re) => h.findIndex((x) => re.test(x));
  const iCod = col(/codice articolo/), iGma = col(/^(gma|gmv)$/), iUnit = col(/^prezzo unitario/), iNetto = col(/^prezzo netto/);
  const iSc = [["Sconto GMA", col(/^sconto gma/)], ["Sconto Cliente", col(/^sconto cliente/)], ["Sconto 3", col(/^sconto 3( \[%\])?$/)], ["Sconto 4", col(/^sconto 4( \[%\])?$/)], ["Sconto 5", col(/^sconto 5( \[%\])?$/)]].filter(([, i]) => i >= 0);
  const iAum = [col(/^aumento 1( \[%\])?$/), col(/^aumento 2( \[%\])?$/)].filter((i) => i >= 0);
  if (!iSc.length) throw new Error("Nel file non ci sono colonne di sconto (Sconto GMA, Sconto Cliente, Sconto 3…).");
  const pct = (v) => {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v > 1 ? v / 100 : v;
    const m = String(v).match(/-?\d+(?:[.,]\d+)?/); if (!m) return 0;
    const n = num(m[0]); return String(v).includes("%") || Math.abs(n) > 1 ? n / 100 : n;
  };
  // GMA dell'articolo: dalla colonna del file o dal database
  const gmaDi = new Map(S.catalogo.map((a) => [codiceOracle(a), a.gmv]));
  for (const x of S.censCfg?.righe || []) if (x.cod) gmaDi.set(x.cod, x.g);
  const out = new Map();
  for (const r of righe.slice(hi + 1)) {
    if (!r) continue;
    const cod = String(r[iCod] ?? "").trim(); if (!cod) continue;
    let g = iGma >= 0 ? String(r[iGma] ?? "").trim().toUpperCase() : "";
    if (!g) g = gmaDi.get(cod) || "";
    if (cod.toUpperCase() === "PITATT.TEC") g = "ATT";
    if (!g || out.has(g)) continue;
    const termini = iSc.map(([n, i]) => ({ n, v: pct(r[i]) })).filter((x) => x.v);
    const aum = iAum.map((i) => pct(r[i])).filter(Boolean);
    const testo = termini.map((x) => String(Math.round(x.v * 1e4) / 100).replace(".", ",")).join("+");
    const f = 1 - termini.reduce((a, x) => a * (1 - x.v), 1) * aum.reduce((a, x) => a * (1 + x), 1);
    const unit = num(iNetto >= 0 && num(r[iNetto]) ? r[iNetto] : iUnit >= 0 ? r[iUnit] : 0);
    out.set(g, { g, cod, testo: aum.length ? "" : testo, f, termini, aum, unit, forzata: iNetto >= 0 && !!num(r[iNetto]) });
  }
  return out;
}
async function righeFileOracle(f) {
  if (/\.pdf$/i.test(f.name)) { const fogli = BIB.righeDaPdf(await testoPdf(f)); if (!fogli.length) throw new Error("Nel PDF non trovo righe."); return fogli[0].righe; }
  if (/\.csv$|\.txt$/i.test(f.name)) return BIB.righeDaTesto(await f.text());
  const wb = await leggiExcel(await f.arrayBuffer());
  return wb.foglio(wb.nomi[0]);
}
$("#cl-cens-carica").addEventListener("click", () => $("#cl-cens-file").click());
$("#cl-cens-file").addEventListener("change", async (e) => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  let trovati;
  try { trovati = scontiDaLineeOracle(await righeFileOracle(f)); } catch (x) { return errore(x, "File non leggibile"); }
  if (!trovati.size) return toast("Nel file non trovo righe con un GMA.", 6000);
  const etich = (c) => `${c.nome}${c.codice ? " · " + c.codice : ""}`;
  // il cliente si può riconoscere dal nome del file (es. "Linee_ROSSI SPA.xlsx")
  const nomeF = f.name.toLowerCase();
  const ipotesi = S.clienti.find((c) => (c.codice && nomeF.includes(String(c.codice).toLowerCase())) || (c.nome && c.nome.length > 3 && nomeF.includes(c.nome.toLowerCase())));
  const lista = [...trovati.values()].sort((a, b) => a.g.localeCompare(b.g));
  const tabella = (cli) => lista.map((x) => {
    const att = x.g === "ATT";
    const prima = cli ? (att ? (cli.att === "" || cli.att === undefined ? "" : euro(cli.att)) : scontoCliTxt(cli, x.g) ? scontoCliTxt(cli, x.g) + "%" : "") : "";
    const dopo = att ? (x.unit ? euro(x.unit * (1 - x.f)) : "") : x.f ? (x.testo || pctVal(x.f)) + "%" : "0%";
    const cambia = cli && prima !== dopo;
    return `<tr><td class="mono"><b>${esc(x.g)}</b></td><td class="small">${esc(fornitoriTxt(x.g))}</td><td class="mono small">${esc(x.cod)}</td>
      <td class="small">${x.termini.map((t) => `${esc(t.n)} ${pctTxt(t.v)}`).join(" · ") || '<span class="muted">nessuno sconto</span>'}${x.aum.length ? ` · aumento ${x.aum.map(pctTxt).join("+")}` : ""}${x.forzata ? ' <span class="badge warn" title="Nella riga c\'è un prezzo netto forzato">netto forzato</span>' : ""}</td>
      <td class="num">${prima ? esc(prima) : '<span class="muted">—</span>'}</td><td class="num"><b>${esc(dopo)}</b>${cambia ? ' <span class="badge warn">cambia</span>' : ""}</td></tr>`;
  }).join("");
  dialogo({
    titolo: `Censimento cartellino · ${f.name}`, largo: true,
    html: `<div class="row" style="flex-wrap:wrap;gap:8px 14px">
        <label class="f grow" style="min-width:280px">Cliente (scegli dall'elenco o scrivi un nuovo cliente)<input id="cs-cli" list="cs-clis" value="${esc(ipotesi ? etich(ipotesi) : "")}" placeholder="Ragione sociale"></label>
        <label class="f">Codice cliente (nuovo cliente)<input id="cs-cod" style="width:140px"></label>
      </div><datalist id="cs-clis">${S.clienti.map((c) => `<option value="${esc(etich(c))}">`).join("")}</datalist>
      <p class="muted small" id="cs-stato" style="margin:6px 0 10px"></p>
      <div class="tbl-wrap" style="max-height:50vh"><table><thead><tr><th>GMA</th><th>Fornitore</th><th>Articolo</th><th>Sconti applicati da Oracle</th><th class="num">Cartellino attuale</th><th class="num">Nuovo</th></tr></thead>
        <tbody id="cs-righe"></tbody></table></div>
      <p class="muted small" style="margin:8px 0 0">Si aggiornano solo i GMA presenti nel file; gli altri sconti del cliente restano com'erano.
        Per l'attività tecnica (ATT) si salva la tariffa = prezzo che Oracle ha messo sulla riga PITATT.TEC.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva il cartellino", value: "ok", cls: "primary" }],
    onApri(d) {
      const agg = () => {
        const v = $("#cs-cli", d).value.trim();
        const cli = S.clienti.find((c) => etich(c) === v || c.nome === v) || null;
        $("#cs-cod", d).closest("label").hidden = !!cli;
        $("#cs-stato", d).innerHTML = !v ? "Scegli il cliente." : cli ? `Cliente esistente: si aggiorna il suo cartellino.${cli.censito ? ` Ultimo censimento: ${dataIt(cli.censito)}.` : ""}` : "<b>Nuovo cliente</b>: verrà creato.";
        $("#cs-righe", d).innerHTML = tabella(cli);
      };
      $("#cs-cli", d).addEventListener("input", agg);
      agg();
      setTimeout(() => $("#cs-cli", d).focus(), 50);
    },
    async onConferma(v, d) {
      const val = $("#cs-cli", d).value.trim();
      if (!val) { toast("Scegli o scrivi il cliente."); return false; }
      const cli = S.clienti.find((c) => etich(c) === val || c.nome === val);
      const base = cli ? structuredClone(cli) : { nome: val, codice: $("#cs-cod", d).value.trim(), att: "", sconti: {} };
      const { id: _id, ...dati } = base;
      dati.sconti ||= {}; dati.scontiTesto ||= {};
      for (const x of lista) {
        if (x.g === "ATT") { if (x.unit) dati.att = Math.round(x.unit * (1 - x.f) * 100) / 100; continue; }
        if (x.f) { dati.sconti[x.g] = x.f; if (x.testo && x.testo.includes("+")) dati.scontiTesto[x.g] = x.testo; else delete dati.scontiTesto[x.g]; }
        else { delete dati.sconti[x.g]; delete dati.scontiTesto[x.g]; }
      }
      dati.censito = oggiIso();
      const id = cli?.id || slug(val) + "-" + db.nuovoId().slice(-4);
      try { await db.salva(COLL.cli, id, dati); } catch (x) { errore(x, "Salvataggio non riuscito"); return false; }
      toast(`Cartellino di ${dati.nome} salvato (${lista.length} GMA).`, 5000);
    },
  });
});

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
      <div class="avviso ok">Letto <b>${esc(file.name)}</b>: ${r.articoli.length} articoli in ${r.capitoli.length} capitoli, ${r.gmv.length} GMA, ${r.clienti.length} clienti.</div>
      <div class="tbl-wrap" style="max-height:300px"><table><thead><tr><th>Capitolo</th><th class="num">Articoli</th><th class="num">Ricarico</th><th class="num">Con quantità</th><th></th></tr></thead><tbody>
      ${r.capitoli.map((c) => `<tr><td>${esc(c.nome)}${c.blocco0 ? ` <span class="badge info" title="Righe copiate in testa al foglio per l'offerta in corso: non entrano nel database perché ci sono già">${c.blocco0} righe del blocco 0 escluse</span>` : ""}</td><td class="num">${c.articoli}</td><td class="num">${pctTxt(c.ricarico)}${c.ricaricoDaFile ? "" : ' <span class="badge warn" title="Il capitolo non è nella tabella Capitoli dell\'Excel">predefinito</span>'}</td>
        <td class="num">${r.conQuantita.filter((a) => a.capitolo === c.nome).length || ""}</td>
        <td>${giaCat.has(c.nome) ? '<span class="badge warn">sostituisce l\'attuale</span>' : '<span class="badge info">nuovo</span>'}</td></tr>`).join("")}
      </tbody></table></div>
      <div class="stack" style="margin-top:10px">
        <label class="row"><input type="checkbox" id="i-cat" checked> Catalogo articoli (sostituisce i capitoli elencati sopra; le offerte già fatte non cambiano)</label>
        <label class="row"><input type="checkbox" id="i-cap" checked> Capitoli e ricarichi automatici, elenco GMA</label>
        <label class="row"><input type="checkbox" id="i-cli" ${r.clienti.length ? "checked" : "disabled"}> Clienti e cartellini (aggiorna quelli con lo stesso nome)</label>
        <label class="row"><input type="checkbox" id="i-off" ${capQ.length ? "" : "disabled"}> Crea un'offerta con le quantità già scritte nel file (un capitolo per ogni foglio)
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
      // un'unica offerta con un capitolo per ogni foglio che ha quantità
      const o = nuovaOffertaVuota();
      o.nome = `Offerta da ${r.nomeFile}`;
      for (const cap of [...new Set(r.conQuantita.map((a) => a.capitolo))]) {
        const c = nuovoCapitolo(cap);
        // come nell'Excel: il capitolato è la descrizione della riga di testata sopra le righe (es. "OFFERTA ANTINTRUSIONE")
        c.righe = r.conQuantita.filter((a) => a.capitolo === cap).map((a) => ({ ...rigaDaArticolo(a), id: db.nuovoId(), qta: a._qta, capitolato: a.sottocapitolo || "" }));
        c.capitolati = [...new Set(c.righe.map((x) => x.capitolato).filter(Boolean))];
        o.capitoli.push(c);
      }
      ops.push({ tipo: "set", coll: COLL.off, id: o.id, data: o });
      nOff = o.capitoli.length;
    }
    await db.scrivi(ops);
    $("#riepilogo-import").innerHTML = `<div class="avviso ok">Importazione completata${nOff ? `: creata un'offerta con ${nOff} capitoli e le quantità del file (la trovi in Offerte)` : ""}.</div>`;
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
    if (/\.pdf$/i.test(f.name)) {
      const fogli = BIB.righeDaPdf(await testoPdf(f));
      if (!fogli.length) return toast("Nel PDF non trovo righe con un prezzo.", 6000);
      return confrontaOracle(fogli[0].righe);
    }
    if (/\.csv$|\.txt$/i.test(f.name)) return confrontaOracle(BIB.righeDaTesto(await f.text()));
    const wb = await leggiExcel(await f.arrayBuffer());
    const nome = wb.nomi.find((n) => /AGGIORNAMENTO LISTINO$/i.test(n)) || wb.nomi[0];
    confrontaOracle(await wb.foglio(nome));
  } catch (x) { errore(x, "File non leggibile"); }
});

const CODICI_GENERICI = new Set([CODICE_GENERICO, "PITATT.TEC"]);
/* Si confrontano i prezzi con il database e con la biblioteca (stesso codice Oracle = precodice + codice). */
const bibCaricata = () => new Promise((ok) => {
  caricaBib();
  if (S.bibPronta) return ok();
  const t = setInterval(() => { if (S.bibPronta) { clearInterval(t); ok(); } }, 100);
  setTimeout(() => { clearInterval(t); ok(); }, 15000);
});
/* Prezzo diverso? Con le righe scontate il confronto si fa sul prezzo unitario arrotondato al centesimo,
   così un listino ricavato da un prezzo arrotondato (es. 15,88 / 0,25) non sembra cambiato. */
const diversoDaOracle = (m, listino) => (m.sconto > 0.0001 ? Math.abs(m.unitario - listino * (1 - m.sconto)) > 0.0051 : Math.abs(m.prezzo - listino) > 0.004);
async function confrontaOracle(righe) {
  const box = $("#riepilogo-oracle");
  let mappa;
  try { mappa = leggiAggiornamentoListino(righe); } catch (x) { box.innerHTML = `<div class="avviso bad">${esc(x.message)}</div>`; return; }
  box.innerHTML = `<div class="avviso info">Confronto ${mappa.size} codici con il database e la biblioteca…</div>`;
  await bibCaricata();
  const cambi = [];
  let uguali = 0;
  const trovati = new Set();
  for (const a of S.catalogo) {
    const cod = codiceArticolo(a);
    if (isGenerico(a) || CODICI_GENERICI.has(cod)) continue;
    const m = mappa.get(cod);
    if (!m) continue;
    trovati.add(cod);
    const cambiaPrezzo = diversoDaOracle(m, num(a.listinoAcquisto));
    const cambiaGmv = m.gmv && m.gmv !== a.gmv;
    if (cambiaPrezzo || cambiaGmv) cambi.push({ a, m, dove: a.capitolo, prima: num(a.listinoAcquisto), gma: a.gmv }); else uguali++;
  }
  for (const b of S.bib) {
    const cod = codiceBib(b);
    const m = mappa.get(cod);
    if (!m) continue;
    trovati.add(cod);
    if (diversoDaOracle(m, num(b.listino))) cambi.push({ b, m, dove: `Biblioteca · ${b.fornitore}`, prima: num(b.listino), gma: condBib(b).gma });
    else uguali++;
  }
  const nonTrovati = [...mappa.keys()].filter((c) => !trovati.has(c));
  S.aggOracle = cambi;
  const nDb = cambi.filter((x) => x.a).length, nBib = cambi.length - nDb;
  box.innerHTML = `
    <div class="avviso ${cambi.length ? "info" : "ok"}">${mappa.size} codici letti: <b>${cambi.length}</b> da aggiornare (${nDb} nel database, ${nBib} in biblioteca), ${uguali} già allineati${nonTrovati.length ? `, ${nonTrovati.length} codici che non sono né nel database né in biblioteca` : ""}${mappa.saltate ? `, ${mappa.saltate} righe saltate (prezzo netto forzato o sconto in euro: il listino non si ricava)` : ""}.
      ${mappa.dubbio ? `<div class="avviso bad" style="margin-top:6px">Sembrano righe di un'offerta ma <b>non ci sono le colonne degli sconti</b> (Sconto 3/4/5): il prezzo unitario viene preso come listino. Se in Oracle le righe erano scontate, i listini risulteranno sbagliati: esporta anche le colonne degli sconti.</div>` : ""}
      ${[...mappa.values()].some((m) => m.sconto > 0.0001) ? "<br>Sono righe di un'offerta: il prezzo unitario di Oracle è già scontato, quindi il listino si ricava togliendo gli sconti della riga (Sconto 3/4/5, Aumenti)." : ""}
      Gli articoli generici (P000ARTICOLO_GEN, PITATT.TEC) non si aggiornano: hanno tutti lo stesso codice in Oracle.</div>
    ${cambi.length ? `<div class="tbl-wrap" style="max-height:340px"><table><thead><tr><th>Dove</th><th>Codice</th><th>Articolo</th><th class="num">Listino attuale</th><th class="num">Listino Oracle</th><th class="num">Diff.</th><th>GMA</th></tr></thead><tbody>
      ${cambi.map(({ a, b, m, dove, prima, gma }) => { const d = prima ? (m.prezzo / prima - 1) : 0;
        return `<tr><td class="small">${esc(dove)}</td><td class="mono">${esc(a ? codiceArticolo(a) : codiceBib(b))}</td><td>${esc((a || b).nOrdine || (b ? b.codice : ""))}</td>
        <td class="num">${euro(prima)}</td><td class="num"><b>${euro(m.prezzo)}</b>${m.sconto > 0.0001 ? `<div class="muted small">${euro(m.unitario)} con ${pctTxt(m.sconto)}</div>` : ""}</td>
        <td class="num"><span class="badge ${d > 0 ? "bad" : "ok"}">${d > 0 ? "+" : ""}${fNum.format(Math.round(d * 1000) / 10)}%</span></td>
        <td class="mono">${a && m.gmv && m.gmv !== gma ? `${esc(gma || "—")} → <b>${esc(m.gmv)}</b>` : esc(gma || "")}</td></tr>`; }).join("")}
      </tbody></table></div>
      <div class="row" style="margin-top:10px"><button class="btn primary" id="applica-oracle">Aggiorna ${cambi.length} articoli</button>
        <span class="muted small">La data di aggiornamento diventa quella di oggi.</span></div>` : ""}
    ${nonTrovati.length ? `<details style="margin-top:8px"><summary class="muted small">Codici che non sono né nel database né in biblioteca (${nonTrovati.length})</summary><p class="mono small">${nonTrovati.map(esc).join(", ")}</p></details>` : ""}`;
  $("#applica-oracle")?.addEventListener("click", applicaOracle);
}
async function applicaOracle() {
  const cambi = S.aggOracle || [];
  const caps = new Set(), forn = new Set();
  const oggi = oggiIso();
  for (const { a, b, m } of cambi) {
    if (b) { b.listino = m.prezzo; b.dataAgg = oggi; forn.add(b.fornitore); continue; }
    const art = S.perK.get(a.k); if (!art) continue;
    const auto = !num(art.listino) || Math.abs(num(art.listino) - listinoCalcolato(art)) < 0.005;
    art.listinoAcquisto = m.prezzo;
    if (m.gmv) art.gmv = m.gmv;
    if (auto) art.listino = listinoCalcolato(art);
    art.dataAgg = oggi;
    caps.add(art.capitolo);
    for (const c of propagaClone(art)) caps.add(c);
  }
  try {
    await db.scrivi([
      ...[...caps].flatMap((c) => opsCapitolo(c, S.catalogo.filter((a) => a.capitolo === c))),
      ...[...forn].flatMap((f) => opsBib(f, S.bib.filter((x) => x.fornitore === f))),
    ]);
    $("#riepilogo-oracle").innerHTML = `<div class="avviso ok">Aggiornati ${cambi.length} articoli${caps.size ? ` in ${caps.size} capitoli del database` : ""}${forn.size ? `${caps.size ? " e" : ""} in ${forn.size} listini della biblioteca` : ""}.
      Le offerte già fatte non cambiano: per riallinearle usa “Aggiorna prezzi dal database” nella pagina Righe.</div>`;
    $("#incolla-oracle").value = "";
    S.aggOracle = null;
  } catch (x) { errore(x, "Aggiornamento non riuscito"); }
}

/* ================= Biblioteca (listini fornitore) =================
   Gli articoli dei listini stanno in oracle_biblioteca, a blocchi per fornitore, e si caricano solo quando servono.
   Profili dei fornitori (colonne, precodice, famiglie di sconto → GMA) in oracle_config/fornitori;
   sconto standard di ogni GMA nell'elenco GMA (oracle_config/gmv, campo "sconto", es. "50+10"). */
function caricaBib() {
  if (S.bibAscolto) return;
  S.bibAscolto = db.ascolta(COLL.bib, (docs) => {
    S.bibChunks = docs;
    S.bib = [];
    for (const d of docs) for (const a of d.articoli || []) S.bib.push({ ...a, _chunk: d.id });
    S.bibPerK = new Map(S.bib.map((a) => [a.k, a]));
    S.bibPronta = true;
    // i risultati compaiono subito anche mentre si scrive nella ricerca
    if (S.vista === "offerta" && S.sub === "righe" && capAtt()) { disegnaPicker(capAtt()); if (S.vistaRighe === "bib") disegnaBibRighe(); }
    ridisegna();
  }, (e) => errore(e, "Errore di lettura della biblioteca"));
  ascoltatori.push(S.bibAscolto);
}
const profiloDi = (nome) => S.fornitori.find((f) => f.nome === nome) || { nome, famiglie: {} };
const testoStd = (x) => [x.sc1, x.sc2, x.sc3].filter((v) => num(v)).map((v) => pctTesto(v)).concat(num(x.aum) ? [pctTesto(-num(x.aum))] : []).join("+");
const scontoGmaTxt = (g, forn) => { const x = voceGma(g, forn) || voceGma(g); return x ? testoStd(x) || x.sconto || "" : ""; };
/* Condizioni di un articolo della biblioteca: GMA e sconto standard, confronto con quanto dice il listino. */
function condBib(a) {
  const p = profiloDi(a.fornitore);
  const gmaFam = (a.famiglia && p.famiglie?.[a.famiglia]) || p.gmaDefault || "";
  const gmaListino = a.gma || "";
  const diffGma = !!(gmaListino && gmaFam && gmaListino !== gmaFam);
  const gma = a.decisione === "listino" ? gmaListino || gmaFam : gmaFam || gmaListino;
  const stdTxt = scontoGmaTxt(gma, a.fornitore);
  const stdTerm = BIB.terminiSconto(stdTxt);
  const scL = a.scListino ?? null;
  const diffSc = scL !== null && Math.abs(1 - BIB.fattoreDi(stdTerm) - scL) > 0.0005;
  const usaListino = a.decisione === "listino" && scL !== null;
  const termini = usaListino ? [scL] : stdTerm;
  return { gma, gmaFam, gmaListino, stdTxt, scL, diffGma, diffSc, daDecidere: (diffGma || diffSc) && !a.decisione, termini,
    fonte: usaListino ? "listino" : "gma", sconto: 1 - BIB.fattoreDi(termini) };
}
const codiceBib = (a) => `${a.precodice ?? profiloDi(a.fornitore).precodice ?? ""}${a.codice}`;
function cercaBib(q, { forn = "", soloDecidere = false, escludiDb = false } = {}) {
  q = (q || "").trim().toLowerCase();
  const nelDb = escludiDb ? new Set(S.catalogo.map(codiceArticolo)) : null;
  const parole = q.split(/\s+/).filter(Boolean);
  return S.bib.filter((a) => (!forn || a.fornitore === forn) &&
    (!parole.length || parole.every((p) => `${codiceBib(a)} ${a.codice} ${a.nOrdine} ${a.descrizione} ${a.famiglia} ${a.fornitore}`.toLowerCase().includes(p))) &&
    (!nelDb || !nelDb.has(codiceBib(a))) && (!soloDecidere || condBib(a).daDecidere));
}
/* Una riga della biblioteca: nella ricerca della pagina Righe ("picker"), nella vista Biblioteca dell'offerta ("offerta")
   o nella pagina Biblioteca ("pagina"). */
function rigaBibHtml(a, modo) {
  const k = condBib(a);
  const sc = pctTxt(k.sconto) || "0%";
  const decidi = k.daDecidere
    ? `<span class="badge warn" title="${esc([k.diffSc ? `Il listino indica lo sconto ${pctTxt(k.scL) || "0%"}, il GMA ${k.gma || "—"} ha ${k.stdTxt ? k.stdTxt + "%" : "nessuno sconto standard"}` : "", k.diffGma ? `Il listino indica il GMA ${k.gmaListino}, la famiglia porta a ${k.gmaFam}` : ""].filter(Boolean).join(". "))}">da decidere</span>
       <button type="button" class="btn small" data-bib-dec="gma" data-k="${esc(a.k)}" title="Usa lo sconto standard del GMA">GMA ${esc(k.stdTxt ? k.stdTxt + "%" : "—")}</button>
       <button type="button" class="btn small" data-bib-dec="listino" data-k="${esc(a.k)}" title="Usa lo sconto indicato dal listino">listino ${esc(pctTxt(k.scL) || "0%")}</button>`
    : a.decisione ? `<button type="button" class="btn small ghost" data-bib-dec="" data-k="${esc(a.k)}" title="Annulla la scelta">${a.decisione === "listino" ? "sconto del listino" : "sconto del GMA"} ✕</button>` : "";
  const nelDb = S.perKCodice?.has(codiceBib(a));
  const azione = modo === "pagina"
    ? (nelDb ? '<span class="badge ok">nel database</span>' : `<button type="button" class="btn small" data-bib-db="${esc(a.k)}">Nel database…</button>`)
    : `<button type="button" class="btn small primary" data-bib-add="${esc(a.k)}">Aggiungi…</button>`;
  const desc = `<td class="desc"><b>${esc(a.nOrdine || a.codice)}</b>${modo !== "pagina" ? ` <span class="badge info">${esc(a.fornitore)}</span>` : ""}<div class="d">${esc(a.descrizione)}</div>${a.sottocapitolo ? `<div class="muted small">${esc(a.sottocapitolo)}</div>` : ""}</td>`;
  if (modo === "picker") return `<tr class="bib" data-bib="${esc(a.k)}"><td></td><td class="mono nowrap">${esc(codiceBib(a))}</td>${desc}<td>${esc(a.um)}</td>
    <td class="num">${euro(a.listino)}<div class="muted small">sc. ${sc}</div></td><td class="mono">${esc(k.gma)}</td><td>${decidi}</td><td class="num nowrap">${azione}</td></tr>`;
  return `<tr class="bib" data-bib="${esc(a.k)}"><td>${esc(a.fornitore)}</td><td class="mono nowrap">${esc(codiceBib(a))}</td>${desc}<td>${esc(a.um)}</td>
    <td class="num">${euro(a.listino)}</td><td class="mono small">${a.famiglia ? esc(a.famiglia) + " → " : ""}<b>${esc(k.gma || "—")}</b></td>
    <td class="num nowrap">${sc}${k.fonte === "listino" ? ' <span class="muted small">(listino)</span>' : ""}</td><td class="nowrap">${decidi} ${azione}</td></tr>`;
}
function opzioniFornitori(sel, valore) {
  const nomi = [...new Set([...S.fornitori.map((f) => f.nome), ...S.bib.map((a) => a.fornitore)])].sort();
  sel.innerHTML = `<option value="">Tutti i fornitori</option>` + nomi.map((n) => `<option ${n === valore ? "selected" : ""}>${esc(n)}</option>`).join("");
}
function disegnaBiblioteca() {
  S.perKCodice = new Set(S.catalogo.map(codiceArticolo));
  opzioniFornitori($("#b-fornitore"), $("#b-fornitore").value);
  const conta = new Map();
  for (const a of S.bib) { const x = conta.get(a.fornitore) || { n: 0, dec: 0 }; x.n++; if (condBib(a).daDecidere) x.dec++; conta.set(a.fornitore, x); }
  const nomi = [...new Set([...S.fornitori.map((f) => f.nome), ...conta.keys()])].sort();
  $("#b-fornitori").innerHTML = !S.bibPronta ? `<tr><td colspan="6" class="vuoto">Carico la biblioteca…</td></tr>` : nomi.length ? nomi.map((n) => {
    const p = profiloDi(n), x = conta.get(n) || { n: 0, dec: 0 };
    return `<tr><td><b>${esc(n)}</b></td><td class="mono">${esc(p.precodice || "")}</td><td class="num">${x.n}</td>
      <td class="num">${x.dec ? `<span class="badge warn">${x.dec}</span>` : ""}</td><td class="small">${p.aggiornatoIl ? new Date(p.aggiornatoIl).toLocaleDateString("it-IT") : ""}</td>
      <td class="nowrap"><button class="btn small" data-forn-agg="${esc(n)}">Aggiorna listino</button>
        <button class="btn small" data-forn-imp="${esc(n)}">Impostazioni</button>
        <button class="icon" data-forn-del="${esc(n)}" title="Togli il listino dalla biblioteca">✕</button></td></tr>`;
  }).join("") : `<tr><td colspan="6" class="vuoto">Nessun listino: caricane uno con “+ Carica listino”.</td></tr>`;
  const forn = $("#b-fornitore").value, q = $("#b-cerca").value, solo = $("#b-decidere").checked;
  const lista = S.bibPronta ? cercaBib(q, { forn, soloDecidere: solo }) : [];
  const vis = lista.slice(0, 300);
  $("#b-righe").innerHTML = vis.map((a) => rigaBibHtml(a, "pagina")).join("") ||
    `<tr><td colspan="8" class="vuoto">${S.bibPronta ? (S.bib.length ? "Nessun articolo trovato." : "La biblioteca è vuota.") : "Carico…"}</td></tr>`;
  $("#b-info").textContent = lista.length > vis.length ? `Mostrati ${vis.length} di ${lista.length}: affina la ricerca.` : lista.length ? `${lista.length} articoli` : "";
  const dec = lista.filter((a) => condBib(a).daDecidere);
  $("#b-bulk").innerHTML = dec.length ? `<div class="avviso">${dec.length} articoli ${forn ? `di ${esc(forn)} ` : ""}hanno nel listino uno sconto o un GMA diverso da quello impostato: decidi voce per voce oppure
    <button class="btn small" data-bib-tutti="gma">usa il GMA per tutti</button> <button class="btn small" data-bib-tutti="listino">usa il listino per tutti</button></div>` : "";
  S.bibDecVisibili = dec;
}
function disegnaBibRighe() {
  const sel = $("#rb-fornitore");
  opzioniFornitori(sel, sel.value);
  S.perKCodice = new Set(S.catalogo.map(codiceArticolo));
  const lista = S.bibPronta ? cercaBib($("#rb-cerca").value, { forn: sel.value }) : [];
  const vis = lista.slice(0, 200);
  $("#rb-righe").innerHTML = vis.map((a) => rigaBibHtml(a, "offerta")).join("") ||
    `<tr><td colspan="8" class="vuoto">${S.bibPronta ? (S.bib.length ? "Nessun articolo trovato." : "La biblioteca è vuota: carica i listini dalla pagina Biblioteca.") : "Carico la biblioteca…"}</td></tr>`;
  $("#rb-info").textContent = lista.length > vis.length ? `Mostrati ${vis.length} di ${lista.length}: affina la ricerca.` : "";
}
$("#b-fornitore").addEventListener("change", disegnaBiblioteca);
$("#b-cerca").addEventListener("input", disegnaBiblioteca);
$("#b-decidere").addEventListener("change", disegnaBiblioteca);
$("#rb-fornitore").addEventListener("change", disegnaBibRighe);
$("#rb-cerca").addEventListener("input", disegnaBibRighe);

/* Salvataggio: blocchi sotto il limite di 1 MB di Firestore, per fornitore. */
function opsBib(fornitore, articoli) {
  const blocchi = [];
  let cur = [], dim = 0;
  for (const { _chunk, ...a } of articoli) {
    const d = JSON.stringify(a).length * 1.2 + 50;
    if (cur.length && dim + d > 700000) { blocchi.push(cur); cur = []; dim = 0; }
    cur.push(a); dim += d;
  }
  if (cur.length) blocchi.push(cur);
  const base = "f_" + slug(fornitore);
  const ops = blocchi.map((b, i) => ({ tipo: "set", coll: COLL.bib, id: `${base}__${i}`, data: { fornitore, parte: i, articoli: b, aggiornatoIl: Date.now() } }));
  const nuovi = new Set(ops.map((o) => o.id));
  for (const c of S.bibChunks) if (c.fornitore === fornitore && !nuovi.has(c.id)) ops.push({ tipo: "del", coll: COLL.bib, id: c.id });
  return ops;
}
/* Riscrive solo i blocchi toccati (decisioni voce per voce). */
function salvaBloccoBib(ids) {
  const ops = [...new Set(ids)].map((id) => {
    const ch = S.bibChunks.find((c) => c.id === id);
    return ch && { tipo: "set", coll: COLL.bib, id, data: { ...ch, articoli: S.bib.filter((a) => a._chunk === id).map(({ _chunk, ...a }) => a) } };
  }).filter(Boolean).map((o) => { delete o.data.id; return o; });
  return db.scrivi(ops);
}
async function decidiBib(arts, decisione) {
  for (const a of arts) { if (decisione) a.decisione = decisione; else delete a.decisione; }
  ridisegna(true);
  if (S.vista === "offerta" && capAtt()) { disegnaPicker(capAtt()); disegnaBibRighe(); }
  try { await salvaBloccoBib(arts.map((a) => a._chunk)); } catch (x) { errore(x, "Salvataggio non riuscito"); }
}
document.addEventListener("click", (e) => {
  const d = e.target.closest("[data-bib-dec]");
  if (d) { const a = S.bibPerK.get(d.dataset.k); if (a) decidiBib([a], d.dataset.bibDec); return; }
  const t = e.target.closest("[data-bib-tutti]");
  if (t) { decidiBib(S.bibDecVisibili || [], t.dataset.bibTutti); return; }
  const ad = e.target.closest("[data-bib-add]");
  if (ad) { const a = S.bibPerK.get(ad.dataset.bibAdd); if (a) aggiungiDaBib(a); return; }
  const dbb = e.target.closest("[data-bib-db]");
  if (dbb) { const a = S.bibPerK.get(dbb.dataset.bibDb); if (a) mettiNelDatabase(a); }
});

/* Articolo del database fatto da un articolo della biblioteca. */
function articoloDaBib(a, capitolo, sottocapitolo) {
  const k = condBib(a), p = profiloDi(a.fornitore);
  const pos = k.termini.filter((x) => x >= 0), neg = k.termini.filter((x) => x < 0);
  const sc = pos.slice(0, 2);
  if (pos.length > 2) sc.push(1 - BIB.fattoreDi(pos.slice(2)));
  const art = {
    capitolo, sottocapitolo: sottocapitolo || "", rifCap: "", pos: "",
    precodice: a.precodice ?? p.precodice ?? "", codice: a.codice, nOrdine: a.nOrdine || "", descrizione: a.descrizione || "", um: a.um || "n.",
    listinoAcquisto: num(a.listino), giorni: 0, riparam: 0,
    // sconti uguali all'acquisto standard del GMA: l'articolo non li salva e segue lo standard
    ...(k.fonte === "gma" && standardGma(k.gma, (voceGma(k.gma, a.fornitore) || voceGma(k.gma))?.fornitore) ? { sc3: 0, sc4: 0, sc5: 0, aum1: 0 } : { sc3: sc[0] || 0, sc4: sc[1] || 0, sc5: sc[2] || 0, aum1: neg.length ? BIB.fattoreDi(neg) - 1 : 0 }),
    gmv: k.gma, notePre: "", notePost: "", dataAgg: a.dataAgg || oggiIso(), daBiblioteca: a.fornitore,
  };
  { const v = voceGma(k.gma, a.fornitore) || voceGma(k.gma); if (v?.fornitore) art.fornitore = v.fornitore; }
  art.listino = listinoCalcolato(art);
  const base = `${capitolo}||${codiceArticolo(art)}|${art.nOrdine}`;
  let kk = base, n = 1;
  while (S.perK.has(kk)) kk = `${base}|${++n}`;
  art.k = kk;
  art.ordine = Math.max(0, ...S.catalogo.filter((x) => x.capitolo === capitolo).map((x) => num(x.ordine))) + 1;
  return art;
}
async function scriviNelDatabase(art) {
  S.catalogo.push(art); S.perK.set(art.k, art);
  await salvaCapitolo(art.capitolo);
}
const opzCapitoliDb = (sel) => S.capitoli.map((c) => `<option ${c.nome === sel ? "selected" : ""}>${esc(c.nome)}</option>`).join("");
const sottoDi = (cap) => [...new Set(S.catalogo.filter((x) => x.capitolo === cap).map((x) => x.sottocapitolo).filter(Boolean))];
const htmlDecisione = (k) => !k.daDecidere ? "" : `<div class="avviso" style="margin-top:10px"><b>Il listino non coincide con le tue condizioni.</b>
  ${k.diffSc ? `Sconto del listino ${pctTxt(k.scL) || "0%"}, sconto standard del GMA ${esc(k.gma || "—")}: ${k.stdTxt ? esc(k.stdTxt) + "%" : "non impostato"}.` : ""}
  ${k.diffGma ? `GMA del listino ${esc(k.gmaListino)}, dalla famiglia ${esc(k.gmaFam)}.` : ""}
  <div class="row" style="margin-top:6px"><label class="row small"><input type="radio" name="ad-dec" value="gma" checked> usa il GMA</label>
  <label class="row small"><input type="radio" name="ad-dec" value="listino"> usa il listino</label></div></div>`;

/* Dalla pagina Righe: aggiunge all'offerta; ogni volta chiede se metterlo anche nel database. */
function aggiungiDaBib(a) {
  const c = capAtt(); if (!c) return toast("Apri prima un capitolo dell'offerta.");
  const k = condBib(a);
  const capDef = S.capitoli.some((x) => x.nome === dbNome(c)) ? dbNome(c) : S.capitoli[0]?.nome || "";
  dialogo({
    titolo: `Aggiungi ${a.nOrdine || a.codice}`,
    html: `<p class="muted small" style="margin-top:0">${esc(a.fornitore)} · <span class="mono">${esc(codiceBib(a))}</span> · ${esc(a.descrizione)}<br>
      Listino ${euro(a.listino)} · GMA ${esc(k.gma || "—")} · sconto di acquisto ${pctTxt(k.sconto) || "0%"}</p>
      <div class="row">
        <label class="f">Quantità<input id="ad-qta" inputmode="decimal" style="width:110px" placeholder="es. 2"></label>
        <label class="f grow">Capitolato<select id="ad-cap"><option value="">— nessuno —</option>${c.capitolati.map((n) => `<option ${n === c.capNuove ? "selected" : ""}>${esc(n)}</option>`).join("")}</select></label>
      </div>
      <div class="row" style="margin-top:8px">
        <label class="f grow">Nota pre-linea<input id="ad-pre"></label><label class="f grow">Nota post-linea<input id="ad-post"></label>
      </div>
      ${htmlDecisione(k)}
      <div class="card" style="margin:12px 0 0;padding:10px 12px">
        <label class="row"><input type="checkbox" id="ad-db"> <b>Metti l'articolo anche nel database</b> (la prossima volta lo trovi tra gli articoli normali)</label>
        <div class="row" id="ad-db-box" hidden style="margin-top:8px">
          <label class="f">Capitolo del database<select id="ad-dbcap">${opzCapitoliDb(capDef)}</select></label>
          <label class="f grow">Sottocapitolo${campoSotto("ad-sotto", a.sottocapitolo || "", [])}</label>
        </div>
      </div>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Aggiungi all'offerta", value: "ok", cls: "primary" }],
    onApri(d) {
      legaSotto("ad-sotto", d);
      const sotti = (tieni) => aggiornaSotto("ad-sotto", d, sottoDi($("#ad-dbcap", d).value), tieni === true);
      $("#ad-db", d).addEventListener("change", (e) => { $("#ad-db-box", d).hidden = !e.target.checked; });
      $("#ad-dbcap", d).addEventListener("change", sotti); sotti(true);
      setTimeout(() => $("#ad-qta", d).focus(), 50);
    },
    async onConferma(v, d) {
      const q = num($("#ad-qta", d).value);
      if (!q) { toast("Scrivi la quantità."); return false; }
      const dec = $('input[name="ad-dec"]:checked', d)?.value;
      if (dec) { a.decisione = dec; salvaBloccoBib([a._chunk]).catch((x) => errore(x)); }
      let art;
      if ($("#ad-db", d).checked) {
        art = articoloDaBib(a, $("#ad-dbcap", d).value, valSotto("ad-sotto", d));
        try { await scriviNelDatabase(art); } catch (x) { errore(x, "Non riesco a salvarlo nel database"); return false; }
      } else {
        art = articoloDaBib(a, c.nome, a.sottocapitolo);
        art.k = "bib|" + a.k; // non è nel database: la riga ne tiene la copia
      }
      const r = { ...rigaDaArticolo(art), id: db.nuovoId(), qta: q, capitolato: $("#ad-cap", d).value };
      const pre = $("#ad-pre", d).value.trim(), post = $("#ad-post", d).value.trim();
      if (pre) r.notePre = pre;
      if (post) r.notePost = post;
      c.righe.push(r);
      modificata(); disegnaOfferta();
      toast(`Aggiunta ${a.nOrdine || a.codice}${$("#ad-db", d).checked ? `, anche nel database (${art.capitolo})` : " (solo nell'offerta)"}.`, 4000);
    },
  });
}
/* Dalla pagina Biblioteca: mette l'articolo nel database. */
function mettiNelDatabase(a) {
  const k = condBib(a);
  dialogo({
    titolo: `Metti nel database ${a.nOrdine || a.codice}`,
    html: `<p class="muted small" style="margin-top:0">${esc(a.fornitore)} · <span class="mono">${esc(codiceBib(a))}</span> · ${esc(a.descrizione)}<br>
      Listino ${euro(a.listino)} · GMA ${esc(k.gma || "—")} · sconto ${pctTxt(k.sconto) || "0%"}</p>
      <div class="row"><label class="f">Capitolo<select id="md-cap">${opzCapitoliDb(S.capitoli[0]?.nome)}</select></label>
      <label class="f grow">Sottocapitolo${campoSotto("md-sotto", a.sottocapitolo || "", [])}</label></div>
      ${htmlDecisione(k)}`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Metti nel database", value: "ok", cls: "primary" }],
    onApri(d) {
      legaSotto("md-sotto", d);
      const sotti = (tieni) => aggiornaSotto("md-sotto", d, sottoDi($("#md-cap", d).value), tieni === true);
      $("#md-cap", d).addEventListener("change", sotti); sotti(true);
    },
    async onConferma(v, d) {
      if (!S.capitoli.length) { toast("Prima crea un capitolo del database."); return false; }
      const dec = $('input[name="ad-dec"]:checked', d)?.value;
      if (dec) { a.decisione = dec; salvaBloccoBib([a._chunk]).catch((x) => errore(x)); }
      const art = articoloDaBib(a, $("#md-cap", d).value, valSotto("md-sotto", d));
      try { await scriviNelDatabase(art); toast(`${a.nOrdine || a.codice} è nel database (${art.capitolo}).`); disegnaBiblioteca(); }
      catch (x) { errore(x, "Non riesco a salvarlo nel database"); return false; }
    },
  });
}

/* --- Caricamento di un listino --- */
$("#b-carica").addEventListener("click", () => { S.bibAtteso = ""; $("#b-file").click(); });
$("#b-file").addEventListener("change", (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) apriListino(f, S.bibAtteso); });
document.addEventListener("click", (e) => {
  const ag = e.target.closest("[data-forn-agg]");
  if (ag) { S.bibAtteso = ag.dataset.fornAgg; $("#b-file").click(); return; }
  const im = e.target.closest("[data-forn-imp]");
  if (im) return impostaFornitore(im.dataset.fornImp);
  const de = e.target.closest("[data-forn-del]");
  if (de) {
    const n = de.dataset.fornDel, quanti = S.bib.filter((a) => a.fornitore === n).length;
    dialogo({ titolo: `Togliere il listino ${n}?`, html: `<p>I ${quanti} articoli di ${esc(n)} escono dalla biblioteca. Database e offerte non cambiano.</p>`,
      pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Togli", value: "ok", cls: "danger" }],
      async onConferma() {
        const ops = S.bibChunks.filter((c) => c.fornitore === n).map((c) => ({ tipo: "del", coll: COLL.bib, id: c.id }));
        ops.push({ tipo: "set", coll: COLL.cfg, id: "fornitori", data: { lista: S.fornitori.filter((f) => f.nome !== n) } });
        await db.scrivi(ops);
      } });
  }
});

/* ================= Offerte dei fornitori → capitoli di articoli generici =================
   Il PDF (o Excel / CSV) dell'offerta del fornitore si legge in automatico: ogni sezione diventa un capitolo,
   ogni riga un articolo generico P000 ARTICOLO_GEN (n° di ordinazione = codice del fornitore) e le prestazioni
   un'attività tecnica PITATT. Le righe restano solo nell'offerta: non vanno nel database né in biblioteca. */
const tolTot = (t) => Math.max(1, Math.abs(t) * 0.0005);
const sommaRighe = (rr) => rr.reduce((a, r) => a + num(r.totale), 0);
function rigaDaOffertaForn(x, att, fonte, capNome) {
  const base = {
    id: db.nuovoId(), capitolo: capNome, sottocapitolo: "", rifCap: "", pos: "",
    nOrdine: x.codice || "", descrizione: x.descrizione || "", um: x.um || "n.",
    riparam: 0, sc3: 0, sc4: 0, sc5: 0, aum1: 0, notePre: "", notePost: "", capitolato: "", dataAgg: oggiIso(), daOfferta: fonte,
  };
  const r = att
    // prestazione: in Oracle PITATT.TEC; la quantità del fornitore va nei "giorni" (1 × valore complessivo, arrotondato come l'attività)
    ? { ...base, precodice: "PIT", codice: "ATT.TEC", generico: false, qta: 1, giorni: num(x.qta) || 1, aGiornate: false, listinoAcquisto: num(x.unitario), gmv: "ATT",
        notePre: x.codice || "", notePost: x.descrizione || "" }
    : { ...base, precodice: "P000", codice: "ARTICOLO_GEN", generico: true, qta: num(x.qta), giorni: 0, listinoAcquisto: num(x.unitario), gmv: "ARTGEN" };
  r.listino = listinoCalcolato(r);
  return r;
}
/* Riga generica ↔ attività (tasto destro sulle righe). */
function convertiRiga(r, inAttivita) {
  const x = isAttivita(r)
    ? { codice: r.notePre || r.nOrdine, descrizione: r.notePost || r.descrizione, um: r.um, qta: num(r.qta) * (num(r.giorni) || 1), unitario: num(r.listinoAcquisto) }
    : { codice: r.nOrdine, descrizione: r.descrizione, um: r.um, qta: num(r.qta), unitario: num(r.listinoAcquisto) * fattoreStandard(r) };
  const n = rigaDaOffertaForn(x, inAttivita, r.daOfferta || "", r.capitolo);
  return { ...n, id: r.id, capitolato: r.capitolato || "" };
}

async function leggiFileOfferta(file) {
  const nome = file.name || "";
  if (/\.pdf$/i.test(nome)) {
    const pagine = await testoPdf(file);
    const o = OFF.leggiOffertaPdf(pagine);
    if (o) return { o };
    const fogli = BIB.righeDaPdf(pagine);
    if (!fogli.length) throw new Error("Nel PDF non trovo righe con un prezzo (se è una scansione, il testo non si può leggere).");
    return { fogli };
  }
  return { fogli: await righeDaFile(file) };
}

async function importaOffertaFornitore(file) {
  if (!S.off) return;
  toast(`Leggo ${file.name}…`, 2500);
  let letto;
  try { letto = await leggiFileOfferta(file); } catch (x) { return errore(x, "File non leggibile"); }
  // tabella (Excel, CSV o PDF non riconosciuto): colonne da scegliere
  let fogli = null, mappa = {}, celle = [];
  if (!letto.o) {
    fogli = letto.fogli.filter((f) => BIB.trovaIntestazione(f.righe));
    if (!fogli.length) return toast("Non trovo la riga delle intestazioni: servono almeno le colonne codice, quantità e prezzo.", 7000);
    celle = [...new Set(fogli.flatMap((f) => BIB.trovaIntestazione(f.righe).celle.filter(Boolean)))];
    mappa = OFF.indovinaMappaOff(BIB.trovaIntestazione(fogli[0].righe).celle);
  }
  const rif = file.name.replace(/\.[^.]+$/, "");
  const st = { o: letto.o, aperte: new Set(), att: new Map(), nomi: new Map(), escluse: new Set(), note: true };
  const rileggi = () => {
    if (fogli) st.o = OFF.leggiOffertaTabella(fogli, mappa) || { sezioni: [], totaleGenerale: null, note: [] };
    st.att = new Map(); st.nomi = new Map();
    st.o.sezioni.forEach((sz, i) => { st.nomi.set(i, sz.nome); sz.righe.forEach((r, j) => st.att.set(`${i}|${j}`, OFF.sembraPrestazione(r.codice, r.descrizione))); });
  };
  rileggi();
  const opzCol = (v) => `<option value="">— nessuna —</option>` + celle.map((h) => `<option ${h === v ? "selected" : ""}>${esc(h)}</option>`).join("");
  // capitolo di destinazione: quello degli articoli generici del database, se c'è
  const capDef = S.capitoli.find((c) => /GENERIC/i.test(c.nome))?.nome || "ARTICOLI_GENERICI";
  const disegna = (d) => {
    const o = st.o, tutte = o.sezioni.flatMap((x) => x.righe), tot = sommaRighe(tutte);
    const totCfr = sommaRighe(o.sezioni.filter((x) => x.nelTotale !== false).flatMap((x) => x.righe));
    const nAtt = [...st.att.values()].filter(Boolean).length, nNo = tutte.filter((r) => r.nonTorna).length;
    const okGen = o.totaleGenerale === null ? null : Math.abs(totCfr - o.totaleGenerale) <= tolTot(o.totaleGenerale);
    $("#of-riep", d).innerHTML = !tutte.length ? `<div class="avviso">Nessuna riga letta${fogli ? ": controlla le colonne codice, quantità e prezzo" : ""}.</div>` : `
      <div class="row" style="flex-wrap:wrap;gap:6px 14px;align-items:center">
        <span><b>${tutte.length}</b> righe in <b>${o.sezioni.length}</b> ${o.sezioni.length === 1 ? "sezione" : "sezioni"}</span>
        <span>di cui <b class="att-txt">${nAtt}</b> prestazioni → PITATT</span>
        <span>totale letto <b>${euro(tot)}</b></span>
        ${okGen === null ? "" : okGen ? `<span class="badge ok">✓ uguale al totale del documento (${euro(o.totaleGenerale)})${totCfr !== tot ? " · le sezioni dopo il totale (opzioni) sono a parte" : ""}</span>` : `<span class="badge bad">✗ il documento dice ${euro(o.totaleGenerale)}</span>`}
        ${nNo ? `<span class="badge bad">${nNo} righe con q.tà × prezzo ≠ totale</span>` : ""}
      </div>`;
    $("#of-sez", d).innerHTML = o.sezioni.map((sz, i) => {
      const somma = sommaRighe(sz.righe), ok = sz.totalePdf === null ? null : Math.abs(somma - sz.totalePdf) <= tolTot(sz.totalePdf);
      const nA = sz.righe.filter((_, j) => st.att.get(`${i}|${j}`)).length, aperta = st.aperte.has(i);
      return `<tr class="of-sez${st.escluse.has(i) ? " qta-zero" : ""}"><td><input type="checkbox" data-of-inc="${i}" ${st.escluse.has(i) ? "" : "checked"} title="Importa questo sottocapitolo"></td>
        <td><input class="of-nome" data-of-nome="${i}" value="${esc(st.nomi.get(i))}"></td>
        <td class="num">${sz.righe.length}</td><td class="num">${nA || ""}</td><td class="num">${euro(somma)}</td>
        <td class="num">${sz.totalePdf === null ? '<span class="muted">—</span>' : euro(sz.totalePdf)}</td>
        <td>${ok === null ? "" : ok ? '<span class="badge ok">✓</span>' : `<span class="badge bad" title="Differenza ${euro(somma - sz.totalePdf)}">✗</span>`}</td>
        <td><button type="button" class="btn small" data-of-apri="${i}">${aperta ? "▾ chiudi" : "▸ righe"}</button></td></tr>` +
        (!aperta ? "" : `<tr><td colspan="8" style="padding:0 0 8px 28px"><div class="tbl-wrap" style="max-height:340px"><table class="of-righe">
          <thead><tr><th title="Prestazione → attività tecnica PITATT">Att.</th><th>Codice fornitore</th><th>Descrizione</th><th>U.M.</th><th class="num">Q.tà</th><th class="num">Unitario</th><th class="num">Totale</th></tr></thead><tbody>
          ${sz.righe.map((r, j) => { const a = st.att.get(`${i}|${j}`); return `<tr class="${a ? "riga-att" : ""}${r.nonTorna ? " of-ko" : ""}">
            <td><input type="checkbox" data-of-att="${i}|${j}" ${a ? "checked" : ""}></td><td class="mono small">${esc(r.codice) || '<span class="badge bad">senza codice</span>'}</td>
            <td class="small">${esc(r.descrizione.length > 220 ? r.descrizione.slice(0, 220) + "…" : r.descrizione)}</td><td>${esc(r.um)}</td>
            <td class="num">${numIt(r.qta)}</td><td class="num">${euro(r.unitario)}${r.listino && r.listino - r.unitario > 0.005 ? `<div class="muted small" title="Listino del fornitore">list. ${euro(r.listino)}</div>` : ""}</td>
            <td class="num">${euro(r.totale)}${r.nonTorna ? '<div class="small" title="Quantità × unitario non dà questo totale">≠ q.tà × unit.</div>' : ""}</td></tr>`; }).join("")}
          </tbody></table></div></td></tr>`);
    }).join("");
    $("#of-note", d).hidden = !o.note.length;
    $("#of-note-n", d).textContent = o.note.length;
    $("#of-note-lista", d).innerHTML = o.note.map((n) => `<li class="small">${esc(n)}</li>`).join("");
  };
  dialogo({
    titolo: `Importa offerta fornitore · ${file.name}`, largo: true,
    html: `<p class="muted small" style="margin-top:0">Tutto entra in <b>un capitolo</b>; ogni sezione del documento diventa un <b>sottocapitolo</b> (capitolato, “Nome Capitolato” in Oracle).
        Le righe entrano come articoli generici <b>P000 ARTICOLO_GEN</b> (n° di ordinazione = codice del fornitore, prezzo = listino acquisto); quelle con <b>Att.</b> spuntato
        come attività tecnica <b>PITATT</b>. Non vanno nel database né nella biblioteca. Quali sottocapitoli esportare lo scegli poi nelle impostazioni di esportazione.</p>
      ${fogli ? `<div style="margin:8px 0 10px"><span class="muted small" style="font-weight:600">Colonne del file</span>
        <div class="lg-mappa">${OFF.CAMPI_OFF.map((c) => `<label class="f">${esc(c.label)}${c.obbl ? " *" : ""}<select data-of-campo="${c.k}">${opzCol(mappa[c.k])}</select></label>`).join("")}</div>
        <p class="muted small" style="margin:4px 0 0">Una riga con solo un testo (senza codice e senza prezzo) apre una nuova sezione.</p></div>` : ""}
      <div class="row" style="flex-wrap:wrap;gap:8px 14px">
        <label class="f grow" style="min-width:260px">Capitolo dell'offerta (nuovo o esistente)<input id="of-cap" list="of-caps" value="${esc(capDef)}"></label>
        <label class="f grow" style="min-width:260px">Riferimento dell'offerta del fornitore<input id="of-rif" value="${esc(rif)}"></label>
      </div><datalist id="of-caps">${[...new Set([...S.off.capitoli.map((c) => c.nome), ...S.capitoli.map((c) => c.nome)])].map((n) => `<option value="${esc(n)}">`).join("")}</datalist>
      <p class="muted small" id="of-cap-stato" style="margin:4px 0 0"></p>
      <div id="of-riep" style="margin:10px 0"></div>
      <div class="tbl-wrap" style="max-height:52vh"><table><thead><tr><th></th><th>Sottocapitolo (capitolato) = sezione del documento</th><th class="num">Righe</th><th class="num">Prestaz.</th>
        <th class="num">Somma righe</th><th class="num">Totale nel documento</th><th></th><th></th></tr></thead><tbody id="of-sez"></tbody></table></div>
      <details id="of-note" style="margin-top:10px" hidden><summary class="small"><b id="of-note-n"></b> note del fornitore</summary>
        <ul class="riep" id="of-note-lista"></ul>
        <label class="row small"><input type="checkbox" id="of-note-ok" checked> aggiungile alle note dell'offerta</label></details>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Importa", value: "ok", cls: "primary" }],
    onApri(d) {
      d.addEventListener("change", (e) => {
        const t = e.target;
        if (t.dataset.ofCampo) { if (t.value) mappa[t.dataset.ofCampo] = t.value; else delete mappa[t.dataset.ofCampo]; rileggi(); }
        else if (t.dataset.ofInc !== undefined) { const i = +t.dataset.ofInc; t.checked ? st.escluse.delete(i) : st.escluse.add(i); }
        else if (t.dataset.ofAtt !== undefined) st.att.set(t.dataset.ofAtt, t.checked);
        else if (t.dataset.ofNome !== undefined) { st.nomi.set(+t.dataset.ofNome, t.value.trim()); return; }
        else if (t.id === "of-note-ok") { st.note = t.checked; return; }
        else return;
        disegna(d);
      });
      const statoCap = () => { const n = $("#of-cap", d).value.trim(), c = S.off.capitoli.find((x) => x.nome === n);
        $("#of-cap-stato", d).textContent = !n ? "Scrivi il nome del capitolo." : c ? `Capitolo già nell'offerta (${c.righe.length} righe): le nuove righe si aggiungono in fondo.` : "Nuovo capitolo dell'offerta."; };
      $("#of-cap", d).addEventListener("input", statoCap); statoCap();
      d.addEventListener("click", (e) => {
        const b = e.target.closest("[data-of-apri]"); if (!b) return;
        const i = +b.dataset.ofApri; st.aperte.has(i) ? st.aperte.delete(i) : st.aperte.add(i); disegna(d);
      });
      disegna(d);
    },
    onConferma(v, d) {
      const o = st.o, fonte = $("#of-rif", d).value.trim() || rif, nomeCap = $("#of-cap", d).value.trim();
      if (!nomeCap) { toast("Scrivi il nome del capitolo."); return false; }
      const scelte = o.sezioni.map((sz, i) => ({ sz, i })).filter((x) => !st.escluse.has(x.i) && x.sz.righe.length);
      if (!scelte.length) { toast("Nessun sottocapitolo da importare."); return false; }
      if (scelte.some((x) => !st.nomi.get(x.i))) { toast("Ogni sottocapitolo deve avere un nome."); return false; }
      let c = S.off.capitoli.find((x) => x.nome === nomeCap);
      const nuovo = !c;
      if (nuovo) { c = nuovoCapitolo(nomeCap, { imp: impOfferta(S.off) }); S.off.capitoli.push(c); }
      c.sottocapitoli = true; // i sottocapitoli vanno in Oracle come "Nome Capitolato"
      c.fonte = [...new Set([...(c.fonte ? c.fonte.split(" · ") : []), fonte])].join(" · ");
      let n = 0;
      for (const { sz, i } of scelte) {
        const cap = st.nomi.get(i);
        aggiungiCapitolato(c, cap);
        for (const [j, r] of sz.righe.entries()) { c.righe.push({ ...rigaDaOffertaForn(r, st.att.get(`${i}|${j}`), fonte, nomeCap), capitolato: cap }); n++; }
      }
      if (st.note && o.note.length) S.off.note = [S.off.note, `Note dell'offerta ${fonte}:`, ...o.note.map((x) => "- " + x)].filter(Boolean).join("\n");
      S.capAtt = nomeCap; S.sub = "righe"; S.sel = new Set();
      modificata(); disegnaTestata(); disegnaOfferta(true);
      toast(`${n} righe in ${scelte.length} sottocapitoli ${nuovo ? "nel nuovo capitolo" : "aggiunte al capitolo"} ${nomeCap}.`, 5000);
    },
  });
}
$("#sb-imp-off").addEventListener("click", () => $("#imp-off-file").click());
$("#imp-off-file").addEventListener("change", (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) { menuAperto(false); importaOffertaFornitore(f); } });

/* Testo di un PDF con le posizioni, pagina per pagina (pdf.js di Mozilla, caricato solo quando serve). */
async function testoPdf(file) {
  const pdfjs = await import("./vendor/pdfjs/pdf.min.js");
  pdfjs.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdfjs/pdf.worker.min.js", import.meta.url).href;
  const doc = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pagine = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const tc = await (await doc.getPage(p)).getTextContent();
    pagine.push(tc.items.filter((i) => i.str && i.str.trim()).map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: Math.abs(i.transform[3]) || i.height, f: i.fontName })));
  }
  return pagine;
}
async function righeDaFile(file) {
  const nome = file.name || "listino";
  if (/\.(xlsx|xlsm)$/i.test(nome)) {
    const wb = await leggiExcel(await file.arrayBuffer());
    const fogli = [];
    for (const n of wb.nomi) fogli.push({ nome: n, righe: (await wb.foglio(n)) || [] });
    return fogli;
  }
  if (/\.xls$/i.test(nome)) throw new Error("Il formato .xls (Excel 97-2003) non si legge: aprilo in Excel e salvalo come .xlsx.");
  if (/\.pdf$/i.test(nome)) {
    const fogli = BIB.righeDaPdf(await testoPdf(file));
    if (!fogli.length) throw new Error("Nel PDF non trovo righe con un prezzo (se è una scansione, il testo non si può leggere).");
    return fogli;
  }
  return [{ nome: nome.replace(/\.[^.]+$/, ""), righe: BIB.righeDaTesto(await file.text()) }];
}
async function apriListino(file, atteso = "") {
  toast(`Leggo ${file.name}…`, 2500);
  let fogli;
  try { fogli = await righeDaFile(file); } catch (x) { return errore(x, "File non leggibile"); }
  const conInt = fogli.map((f) => ({ ...f, int: BIB.trovaIntestazione(f.righe) })).filter((f) => f.int);
  if (!conInt.length) return toast("Non trovo la riga delle intestazioni: servono almeno una colonna con il codice e una con il prezzo.", 7000);
  const tutteInt = [...new Set(conInt.flatMap((f) => f.int.celle.filter(Boolean)))];
  const metel = BIB.isMetel(conInt[0].int.celle);
  // fornitore: quello richiesto, oppure un profilo che conosce già queste colonne, oppure il marchio METEL o il nome del file
  let prof = atteso ? profiloDi(atteso) : S.fornitori.find((f) => f.mappa?.codice && tutteInt.includes(f.mappa.codice) && tutteInt.includes(f.mappa.prezzo) && f.file && file.name.toLowerCase().includes(f.file));
  let nomeForn = prof?.nome || "";
  if (!nomeForn && metel) { const r = conInt[0].righe[conInt[0].int.indice + 1] || []; nomeForn = String(r[conInt[0].int.celle.indexOf("sigla marchio")] || "").trim(); prof = profiloDi(nomeForn); }
  if (!nomeForn) {
    const vuote = /^(listino|listini|prezzi|prezzo|price|list|pricelist|copia|di|del|the|upgrade|licenze|rev|tot|con|sel|lock|compressed|\d.*)$/i;
    nomeForn = file.name.replace(/\.[^.]+$/, "").split(/[_\s.-]+/).find((x) => x.length > 2 && !vuote.test(x)) || "";
  }
  const st = {
    fornitore: nomeForn, precodice: prof?.precodice ?? (metel ? nomeForn : ""),
    mappa: prof?.mappa && tutteInt.includes(prof.mappa.codice) ? { ...prof.mappa } : metel ? { ...BIB.MAPPA_METEL } : BIB.indovinaMappa(conInt[0].int.celle),
    fogli: new Set((prof?.fogli && conInt.some((f) => prof.fogli.includes(f.nome)) ? conInt.filter((f) => prof.fogli.includes(f.nome)) : conInt).map((f) => f.nome)),
    sottoFoglio: prof?.sottoFoglio ?? conInt.length > 1,
    famiglie: { ...(prof?.famiglie || {}) }, gmaDefault: prof?.gmaDefault || "", sconti: {},
  };
  const opzCol = (v) => `<option value="">— nessuna —</option>` + tutteInt.map((h) => `<option ${h === v ? "selected" : ""}>${esc(h)}</option>`).join("");
  let letti = { articoli: [], scartate: 0, perFoglio: {} };
  const leggi = () => { letti = BIB.leggiListino(conInt.filter((f) => st.fogli.has(f.nome)), st.mappa, { sottoFoglio: st.sottoFoglio }); };
  const gmaList = `<datalist id="lg-gma">${S.gmv.map((g) => `<option value="${esc(g.codice)}">${esc(g.fornitore || "")}</option>`).join("")}</datalist>`;
  dialogo({
    titolo: `Carica listino · ${file.name}`, largo: true,
    html: `<div class="row">
        <label class="f">Fornitore<input id="lg-forn" list="lg-forns" value="${esc(st.fornitore)}"><datalist id="lg-forns">${S.fornitori.map((f) => `<option value="${esc(f.nome)}">`).join("")}</datalist></label>
        <label class="f">Precodice per Oracle<input id="lg-pre" value="${esc(st.precodice)}" style="width:110px" placeholder="es. SIE"></label>
        ${metel ? '<span class="badge ok" style="align-self:flex-end">formato METEL riconosciuto</span>' : ""}
      </div>
      ${conInt.length > 1 ? `<div style="margin-top:10px"><span class="muted small" style="font-weight:600">Fogli da leggere</span>
        <div class="row" style="flex-wrap:wrap;gap:4px 14px;margin-top:4px">${conInt.map((f) => `<label class="row small"><input type="checkbox" data-lg-foglio="${esc(f.nome)}" ${st.fogli.has(f.nome) ? "checked" : ""}> ${esc(f.nome)} <span class="muted" data-lg-n="${esc(f.nome)}"></span></label>`).join("")}</div>
        <label class="row small" style="margin-top:6px"><input type="checkbox" id="lg-sottofoglio" ${st.sottoFoglio ? "checked" : ""}> il nome del foglio diventa il sottocapitolo</label></div>` : ""}
      <div style="margin-top:12px"><span class="muted small" style="font-weight:600">Colonne del listino</span>
        <div class="lg-mappa">${BIB.CAMPI.map((c) => `<label class="f">${esc(c.label)}${c.obbl ? " *" : ""}<select data-lg-campo="${c.k}">${opzCol(st.mappa[c.k])}</select></label>`).join("")}</div></div>
      <div id="lg-anteprima" style="margin-top:12px"></div>
      <div id="lg-famiglie" style="margin-top:12px"></div>${gmaList}`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Carica nella biblioteca", value: "ok", cls: "primary" }],
    onApri(d) {
      const aggiorna = () => {
        leggi();
        for (const [n, x] of Object.entries(letti.perFoglio)) { const el = $(`[data-lg-n="${CSS.escape(n)}"]`, d); if (el) el.textContent = `(${x})`; }
        const es = letti.articoli.slice(0, 6);
        $("#lg-anteprima", d).innerHTML = `<span class="muted small" style="font-weight:600">Anteprima: ${letti.articoli.length} articoli letti${letti.scartate ? `, ${letti.scartate} righe saltate (senza codice o prezzo)` : ""}</span>
          <div class="tbl-wrap" style="max-height:220px"><table><thead><tr><th>Codice Oracle</th><th>Modello</th><th>Descrizione</th><th>U.M.</th><th class="num">Listino</th><th class="num">Sconto listino</th><th>Famiglia</th><th>Sottocapitolo</th></tr></thead><tbody>
          ${es.map((a) => `<tr><td class="mono">${esc(st.precodice + a.codice)}</td><td>${esc(a.nOrdine)}</td><td class="small">${esc(a.descrizione.slice(0, 80))}</td><td>${esc(a.um)}</td><td class="num">${euro(a.listino)}</td><td class="num">${a.scListino === null ? "" : pctTxt(a.scListino)}</td><td>${esc(a.famiglia)}</td><td class="small">${esc(a.sottocapitolo)}</td></tr>`).join("") || `<tr><td colspan="8" class="vuoto">Nessun articolo: controlla le colonne codice e prezzo.</td></tr>`}
          </tbody></table></div>`;
        // famiglie di sconto → GMA (con lo sconto standard del GMA)
        const fam = new Map();
        for (const a of letti.articoli) fam.set(a.famiglia || "", (fam.get(a.famiglia || "") || 0) + 1);
        const righe = [...fam.entries()].sort((x, y) => y[1] - x[1]);
        const riga = (f, n) => {
          const g = f ? st.famiglie[f] || "" : st.gmaDefault;
          return `<tr><td>${f ? `<b class="mono">${esc(f)}</b>` : "<i>senza famiglia</i>"}</td><td class="num">${n}</td>
            <td><input class="cond" data-lg-fam="${esc(f)}" list="lg-gma" value="${esc(g)}" placeholder="GMA" style="width:130px"></td>
            <td><input class="cond" data-lg-sc="${esc(g)}" value="${esc(st.sconti[g] ?? scontoGmaTxt(g))}" placeholder="es. 50+10" style="width:110px" ${g ? "" : "disabled"}></td></tr>`;
        };
        $("#lg-famiglie", d).innerHTML = `<span class="muted small" style="font-weight:600">Famiglie di sconto → GMA</span>
          <p class="muted small" style="margin:2px 0 6px">Abbina ogni famiglia a un GMA e scrivi lo sconto standard del GMA (vale per tutti i suoi articoli; più sconti in cascata con il +).
            Se il listino indica per un articolo uno sconto diverso, lo trovi “da decidere” nella biblioteca.</p>
          <div class="tbl-wrap" style="max-height:240px"><table><thead><tr><th>Famiglia</th><th class="num">Articoli</th><th>GMA</th><th>Sconto standard %</th></tr></thead>
          <tbody>${righe.slice(0, 200).map(([f, n]) => riga(f, n)).join("")}</tbody></table></div>`;
      };
      d.addEventListener("change", (e) => {
        const t = e.target;
        if (t.id === "lg-forn") st.fornitore = t.value.trim();
        else if (t.id === "lg-pre") st.precodice = t.value.trim();
        else if (t.id === "lg-sottofoglio") st.sottoFoglio = t.checked;
        else if (t.dataset.lgFoglio !== undefined) { t.checked ? st.fogli.add(t.dataset.lgFoglio) : st.fogli.delete(t.dataset.lgFoglio); }
        else if (t.dataset.lgCampo) { if (t.value) st.mappa[t.dataset.lgCampo] = t.value; else delete st.mappa[t.dataset.lgCampo]; }
        else if (t.dataset.lgFam !== undefined) { const f = t.dataset.lgFam, g = t.value.trim(); if (f) { if (g) st.famiglie[f] = g; else delete st.famiglie[f]; } else st.gmaDefault = g; }
        else if (t.dataset.lgSc !== undefined) { st.sconti[t.dataset.lgSc] = t.value.trim(); return; }
        else return;
        aggiorna();
      });
      aggiorna();
    },
    async onConferma() {
      if (!st.fornitore) { toast("Scrivi il nome del fornitore."); return false; }
      if (!st.mappa.codice || !st.mappa.prezzo) { toast("Scegli le colonne del codice e del prezzo."); return false; }
      leggi();
      if (!letti.articoli.length) { toast("Nessun articolo letto: controlla le colonne."); return false; }
      const errSc = Object.entries(st.sconti).find(([, t]) => t && !testoValido(t));
      if (errSc) { toast(`Sconto standard non valido per ${errSc[0]}: scrivi per esempio 50 oppure 50+10.`); return false; }
      return salvaListino(st, letti, file.name);
    },
  });
}
async function salvaListino(st, letti, nomeFile) {
  const oggi = oggiIso();
  const esistenti = S.bib.filter((a) => a.fornitore === st.fornitore);
  const u = BIB.unisciListino(esistenti, letti.articoli, oggi);
  const articoli = u.articoli.map((a) => ({ ...a, fornitore: st.fornitore, precodice: st.precodice, k: `${slug(st.fornitore)}|${a.codice}` }));
  const ops = opsBib(st.fornitore, articoli);
  // profilo del fornitore: la prossima volta il listino si legge da solo
  const prof = { nome: st.fornitore, precodice: st.precodice, mappa: st.mappa, fogli: [...st.fogli], sottoFoglio: st.sottoFoglio,
    famiglie: st.famiglie, gmaDefault: st.gmaDefault, file: (nomeFile.toLowerCase().match(/[a-z]{4,}/) || [""])[0], aggiornatoIl: Date.now(), articoli: articoli.length };
  ops.push({ tipo: "set", coll: COLL.cfg, id: "fornitori", data: { lista: [...S.fornitori.filter((f) => f.nome !== st.fornitore), prof] } });
  // GMA nuovi e sconti standard
  const gm = new Map(S.gmv.map((g) => [g.codice, { ...g }]));
  const usati = new Set([st.gmaDefault, ...Object.values(st.famiglie)].filter(Boolean));
  for (const g of usati) if (!gm.has(g)) gm.set(g, { codice: g, fornitore: st.fornitore });
  for (const [g, t] of Object.entries(st.sconti)) if (g && gm.has(g)) gm.get(g).sconto = t;
  ops.push({ tipo: "set", coll: COLL.cfg, id: "gmv", data: { lista: [...gm.values()] } });
  try { await db.scrivi(ops); } catch (x) { errore(x, "Salvataggio non riuscito"); return false; }
  // articoli già nel database con un prezzo diverso: si segnalano (l'aggiornamento del database è un passo a parte)
  const perCod = new Map(articoli.map((a) => [`${a.precodice}${a.codice}`, a]));
  const diversi = S.catalogo.filter((x) => { const a = perCod.get(codiceArticolo(x)); return a && Math.abs(num(a.listino) - num(x.listinoAcquisto)) > 0.004; });
  setTimeout(() => dialogo({
    titolo: `Listino ${st.fornitore} caricato`,
    html: `<ul class="riep"><li><b>${u.nuovi}</b> articoli nuovi</li><li><b>${u.aggiornati.length}</b> con il prezzo cambiato</li><li>${u.invariati} invariati</li>
      ${u.tolti.length ? `<li><b>${u.tolti.length}</b> non più nel listino (tolti dalla biblioteca)</li>` : ""}</ul>
      ${diversi.length ? `<div class="avviso"><b>${diversi.length} articoli del tuo database</b> hanno un prezzo diverso da questo listino.
        <details><summary>vedi</summary><ul class="riep">${diversi.slice(0, 200).map((x) => `<li class="mono small">${esc(codiceArticolo(x))} ${esc(x.nOrdine)}: ${euro(x.listinoAcquisto)} → ${euro(perCod.get(codiceArticolo(x)).listino)}</li>`).join("")}</ul></details></div>` : ""}`,
  }), 300);
}
/* Impostazioni di un fornitore: precodice, GMA senza famiglia, famiglie → GMA. */
function impostaFornitore(nome) {
  const p = profiloDi(nome);
  const fam = [...new Set(S.bib.filter((a) => a.fornitore === nome).map((a) => a.famiglia).filter(Boolean))].sort();
  dialogo({
    titolo: `Impostazioni ${nome}`,
    html: `<div class="row"><label class="f">Precodice per Oracle<input id="fi-pre" value="${esc(p.precodice || "")}" style="width:110px"></label>
      <label class="f">GMA degli articoli senza famiglia<input id="fi-gma" list="fi-gmas" value="${esc(p.gmaDefault || "")}" style="width:140px"></label></div>
      ${fam.length ? `<div class="tbl-wrap" style="max-height:300px;margin-top:10px"><table><thead><tr><th>Famiglia</th><th>GMA</th></tr></thead><tbody>
        ${fam.map((f) => `<tr><td class="mono">${esc(f)}</td><td><input class="cond" data-fi-fam="${esc(f)}" list="fi-gmas" value="${esc(p.famiglie?.[f] || "")}" style="width:140px"></td></tr>`).join("")}</tbody></table></div>` : ""}
      <datalist id="fi-gmas">${S.gmv.map((g) => `<option value="${esc(g.codice)}">${esc(g.fornitore || "")}</option>`).join("")}</datalist>
      <p class="muted small">Il precodice cambia il codice Oracle degli articoli della biblioteca. Gli sconti standard si impostano in “GMA e acquisti standard”.</p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    async onConferma(v, d) {
      const famiglie = {};
      $$("[data-fi-fam]", d).forEach((x) => { if (x.value.trim()) famiglie[x.dataset.fiFam] = x.value.trim(); });
      const pre = $("#fi-pre", d).value.trim();
      const prof = { ...p, nome, precodice: pre, gmaDefault: $("#fi-gma", d).value.trim(), famiglie };
      const ops = [{ tipo: "set", coll: COLL.cfg, id: "fornitori", data: { lista: [...S.fornitori.filter((f) => f.nome !== nome), prof] } }];
      if (pre !== (p.precodice || "")) ops.push(...opsBib(nome, S.bib.filter((a) => a.fornitore === nome).map((a) => ({ ...a, precodice: pre }))));
      await db.scrivi(ops);
    },
  });
}
/* Sconto standard di ogni GMA (vale per gli articoli della biblioteca). */
$("#b-sconti").addEventListener("click", () => apriElencoGma());
/* Rinomina un GMA ovunque: elenco GMA, database, biblioteca, cartellini dei clienti, offerte, censimento.
   Se il nuovo codice esiste già, i due GMA si uniscono (restano fornitore e sconto del nuovo). */
async function rinominaGma(vecchio, nuovo, forn = "") {
  const ops = [], conta = { art: 0, bib: 0, cli: 0, off: 0 };
  const FU = (x) => String(x || "").trim().toUpperCase(), F = FU(forn);
  const questo = (g) => g.codice === vecchio && FU(g.fornitore) === F;
  const esiste = S.gmv.some((g) => g.codice === nuovo && FU(g.fornitore) === F);
  const lista = S.gmv.map((g) => ({ ...g })).filter((g) => !(esiste && questo(g))).map((g) => (questo(g) ? { ...g, codice: nuovo } : g));
  // lo stesso codice resta con altri fornitori? allora clienti, offerte e censimento (che conoscono solo il codice) non si toccano
  const altri = lista.some((g) => g.codice === vecchio);
  ops.push({ tipo: "set", coll: COLL.cfg, id: "gmv", data: { lista } });
  // database: solo gli articoli di questo fornitore
  const capCat = new Set();
  for (const a of S.catalogo) if (a.gmv === vecchio && FU(fornitoreDi(a)) === F) { a.gmv = nuovo; if (forn) a.fornitore = forn; capCat.add(a.capitolo); conta.art++; }
  for (const cap of capCat) ops.push(...opsCapitolo(cap, S.catalogo.filter((a) => a.capitolo === cap)));
  // biblioteca
  const chBib = new Set();
  for (const a of S.bib || []) if (a.gma === vecchio && (!altri || FU(a.fornitore) === F)) { a.gma = nuovo; chBib.add(a._chunk); conta.bib++; }
  for (const id of chBib) { const ch = S.bibChunks.find((c) => c.id === id); if (!ch) continue; const { id: _i, ...dati } = ch; ops.push({ tipo: "set", coll: COLL.bib, id, data: { ...dati, articoli: S.bib.filter((a) => a._chunk === id).map(({ _chunk, ...a }) => a) } }); }
  // cartellini, offerte e censimento conoscono solo il codice: si toccano solo se il vecchio codice sparisce del tutto
  if (!altri) {
  // cartellini dei clienti
  for (const c of S.clienti) {
    if (!c.sconti || !(vecchio in c.sconti)) continue;
    const { id, ...d } = structuredClone(c);
    if (!(nuovo in d.sconti)) { d.sconti[nuovo] = d.sconti[vecchio]; if (d.scontiTesto?.[vecchio]) d.scontiTesto[nuovo] = d.scontiTesto[vecchio]; }
    delete d.sconti[vecchio]; if (d.scontiTesto) delete d.scontiTesto[vecchio];
    ops.push({ tipo: "set", coll: COLL.cli, id, data: d }); conta.cli++;
  }
  // offerte (anche quella aperta)
  const sposta = (m) => { if (m && vecchio in m) { if (!(nuovo in m)) m[nuovo] = m[vecchio]; delete m[vecchio]; return true; } return false; };
  const sistema = (o) => {
    let cambiata = false;
    for (const c of o.capitoli || []) {
      for (const r of c.righe || []) if (r.gmv === vecchio) { r.gmv = nuovo; cambiata = true; }
      for (const m of [c.acq?.gma, c.ven?.ricGma, c.ven?.scGma, c.lis?.gma]) if (sposta(m)) cambiata = true;
    }
    return cambiata;
  };
  for (const o of S.offerte) {
    if (S.off && o.id === S.off.id) continue;
    const copia = structuredClone(o); const { id, ...d } = copia;
    if (sistema(copia)) { ops.push({ tipo: "set", coll: COLL.off, id, data: d }); conta.off++; }
  }
  if (S.off && sistema(S.off)) { conta.off++; modificata(); }
  // censimento
  if (S.censCfg?.righe?.some((x) => x.g === vecchio)) {
    const righe = S.censCfg.righe.filter((x) => !(esiste && x.g === vecchio && S.censCfg.righe.some((y) => y.g === nuovo))).map((x) => (x.g === vecchio ? { ...x, g: nuovo } : x));
    ops.push({ tipo: "set", coll: COLL.cfg, id: "censimento", data: { righe } });
  }
  } else if (S.off) {
    // solo le righe dell'offerta aperta di questo fornitore
    let c2 = false;
    for (const c of S.off.capitoli || []) for (const r of c.righe || []) if (r.gmv === vecchio && FU(r.fornitore) === F && F) { r.gmv = nuovo; c2 = true; }
    if (c2) { conta.off++; modificata(); }
  }
  await db.scrivi(ops);
  S.gmv = lista; impostaStandardGma(S.gmv);
  return conta;
}
function chiediRinominaGma(vecchio, forn = "") {
  const FU = (x) => String(x || "").trim().toUpperCase();
  const altri = S.gmv.filter((g) => g.codice === vecchio && FU(g.fornitore) !== FU(forn)).map((g) => g.fornitore || "senza fornitore");
  dialogo({
    titolo: `Rinomina il GMA ${vecchio}${forn ? " · " + forn : ""}`,
    html: `<label class="f">Nuovo codice<input id="rg-nuovo" value="${esc(vecchio)}" style="text-transform:uppercase"></label>
      <p class="muted small">Cambia solo il GMA <b>${esc(vecchio)}</b> di <b>${esc(forn || "senza fornitore")}</b> e i suoi articoli del database e della biblioteca.
        ${altri.length ? `Lo stesso codice esiste anche per ${esc(altri.join(", "))}: quello resta com'è, e quindi cartellini dei clienti e condizioni delle offerte (che conoscono solo il codice) non vengono toccati.`
          : "Il codice cambia anche nei cartellini dei clienti e nelle condizioni e righe di tutte le offerte."}
        Se scrivi un codice che esiste già per lo stesso fornitore, i due GMA vengono uniti.</p><p class="small" id="rg-info"></p>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Rinomina", value: "ok", cls: "primary" }],
    onApri(d) {
      const i = $("#rg-nuovo", d), info = () => { const n = i.value.trim().toUpperCase(); $("#rg-info", d).innerHTML = n && n !== vecchio && S.gmv.some((g) => g.codice === n && FU(g.fornitore) === FU(forn)) ? `<b>${esc(n)}</b> esiste già per ${esc(forn || "senza fornitore")}: ${esc(vecchio)} verrà unito a ${esc(n)}.` : ""; };
      i.addEventListener("input", info); setTimeout(() => { i.focus(); i.select(); }, 50);
    },
    async onConferma(v, d) {
      const n = $("#rg-nuovo", d).value.trim().toUpperCase();
      if (!n) { toast("Scrivi il codice."); return false; }
      if (n === vecchio) return;
      try {
        const c = await rinominaGma(vecchio, n, forn);
        toast(`GMA ${vecchio} → ${n}: ${c.art} articoli, ${c.bib} in biblioteca, ${c.cli} clienti, ${c.off} offerte aggiornati.`, 7000);
      } catch (x) { errore(x, "Rinomina non riuscita"); return false; }
      setTimeout(() => apriAcquistiStandard(), 50);
    },
  });
}

/* ================= Impostazioni › Acquisti standard =================
   Sconti di acquisto standard per GMA (Sconto 1, 2, 3 in cascata e Aumento), divisi per fornitore.
   Gli articoli senza sconti propri usano questi; quelli con sconti propri restano un'eccezione. */
function apriAcquistiStandard() {
  const lista = S.gmv.map((g) => ({ ...g }));
  // vecchio sconto standard in testo ("50+10") → campi
  for (const g of lista) if (!(num(g.sc1) || num(g.sc2) || num(g.sc3) || num(g.aum)) && g.sconto) {
    const t = BIB.terminiSconto(g.sconto); const pos = t.filter((x) => x > 0), neg = t.filter((x) => x < 0);
    [g.sc1, g.sc2, g.sc3] = [pos[0] || 0, pos[1] || 0, pos[2] || 0]; g.aum = neg.length ? BIB.fattoreDi(neg) - 1 : 0;
  }
  // articoli per GMA (fornitore + codice)
  const kv = (f, c) => `${String(f || "").trim().toUpperCase()}|${c}`;
  const conta = new Map();
  for (const a of S.catalogo) {
    if (isAttivita(a) || !a.gmv) continue;
    const v = voceGma(a.gmv, a.fornitore) || { codice: a.gmv, fornitore: fornitoreDi(a) };
    const k = kv(v.fornitore, v.codice), x = conta.get(k) || { n: 0, propri: 0 }; x.n++; if (haScontiPropri(a)) x.propri++; conta.set(k, x);
  }
  const cella = (i, g, k) => `<td><span class="pct-wrap"><input class="pct" data-as="${i}" data-k="${k}" value="${esc(pctVal(g[k]))}" style="width:72px"></span></td>`;
  const gruppi = new Map();
  lista.forEach((g, i) => { const f = (g.fornitore || "").trim().toUpperCase() || "ALTRI"; if (!gruppi.has(f)) gruppi.set(f, []); gruppi.get(f).push(i); });
  const ordG = [...gruppi.entries()].sort((a, b) => (a[0] === "ALTRI") - (b[0] === "ALTRI") || a[0].localeCompare(b[0]));
  dialogo({
    titolo: "GMA e acquisti standard", largo: true,
    html: `<p class="muted small" style="margin-top:0">Tutti i GMA, divisi per fornitore. <b>Un GMA è fornitore + codice</b>: lo stesso codice con due fornitori sono due GMA distinti,
        ognuno con i suoi sconti. Qui cambi il fornitore, rinomini o unisci un GMA (✎), ne aggiungi uno nuovo e imposti gli sconti di acquisto standard in cascata
        (Sconto 1, poi 2, poi 3) più un eventuale aumento. Gli articoli <b>senza sconti propri</b> usano questi valori; quelli con sconti propri restano un'eccezione.</p>
      <div class="row" style="margin-bottom:6px"><input id="as-cerca" type="search" class="grow" placeholder="Cerca GMA o fornitore"></div>
      <div class="tbl-wrap" style="max-height:52vh"><table><thead><tr><th>GMA</th><th>Fornitore</th><th class="num">Articoli</th><th class="num" title="Articoli con sconti propri: non seguono lo standard">Con sconti propri</th><th>Sconto 1</th><th>Sconto 2</th><th>Sconto 3</th><th>Aumento</th><th title="Codice con cui Oracle conosce questo GMA (es. SENSITRON e WAGNER in Oracle sono SIFSFI): cartellino e censimento usano quello">GMA in Oracle</th></tr></thead>
      <tbody id="as-righe">${ordG.map(([forn, idx]) => `<tr class="sotto as-forn"><td colspan="9">${esc(forn)}</td></tr>` +
        idx.sort((x, y) => lista[x].codice.localeCompare(lista[y].codice)).map((i) => { const g = lista[i], x = conta.get(kv(g.fornitore, g.codice)) || { n: 0, propri: 0 };
          const doppio = lista.filter((y) => y.codice === g.codice).length > 1;
          return `<tr data-as-riga="${esc(g.codice)}" data-forn="${esc(forn)}"><td class="mono nowrap"><b>${esc(g.codice)}</b>${doppio ? ' <span class="badge info" title="Lo stesso codice esiste anche con un altro fornitore">+ altro fornitore</span>' : ""} <button type="button" class="tg-ed" data-sg-rin="${i}" title="Rinomina il GMA di questo fornitore">✎</button></td>
            <td><input data-as-f="${i}" list="as-forns" value="${esc(g.fornitore || "")}" placeholder="fornitore" style="width:210px"></td>
            <td class="num">${x.n || ""}</td><td class="num">${x.propri || ""}</td>
            ${cella(i, g, "sc1")}${cella(i, g, "sc2")}${cella(i, g, "sc3")}${cella(i, g, "aum")}
            <td><input data-as-o="${i}" list="as-gmas" class="mono" value="${esc(g.oracle || "")}" placeholder="${esc(g.codice)}" style="width:110px" title="Lascia vuoto se in Oracle il codice è lo stesso"></td></tr>`; }).join("")).join("")}</tbody></table></div>
      <datalist id="as-gmas">${[...new Set(lista.map((g) => g.codice))].sort().map((c) => `<option value="${esc(c)}">`).join("")}</datalist>
      <div class="card row" style="margin-top:10px;gap:8px;align-items:flex-end;flex-wrap:wrap"><b style="align-self:center">+ Nuovo GMA</b>
        <label class="f">Codice<input id="as-nuovo" placeholder="es. SN07" style="width:110px"></label>
        <label class="f grow">Fornitore<input id="as-nuovo-f" list="as-forns" placeholder="scegli o scrivi"></label>
        ${["Sconto 1", "Sconto 2", "Sconto 3", "Aumento"].map((l, i) => `<label class="f">${l}<span class="pct-wrap"><input class="pct" id="as-nuovo-${i}" style="width:72px"></span></label>`).join("")}</div>
      <datalist id="as-forns">${fornitoriGma().map((f) => `<option value="${esc(f)}">`).join("")}</datalist>`,
    pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(d) {
      $("#as-cerca", d).addEventListener("input", (e) => {
        const q = e.target.value.trim().toLowerCase();
        $$("#as-righe tr[data-as-riga]", d).forEach((tr) => { const f = $("[data-as-f]", tr)?.value || ""; tr.hidden = !!q && !(tr.dataset.asRiga + " " + tr.dataset.forn + " " + f).toLowerCase().includes(q); });
        $$("#as-righe tr.as-forn", d).forEach((tr) => { let n = tr.nextElementSibling, vis = false; while (n && !n.classList.contains("as-forn")) { if (!n.hidden) vis = true; n = n.nextElementSibling; } tr.hidden = !vis; });
      });
      d.addEventListener("click", (e) => { const b = e.target.closest("[data-sg-rin]"); if (b) { const g = lista[+b.dataset.sgRin]; chiediRinominaGma(g.codice, g.fornitore || ""); } });
    },
    async onConferma(v, d) {
      for (const i of $$("[data-as]", d)) { const g = lista[+i.dataset.as]; const x = daPct(i.value); if (x) g[i.dataset.k] = x; else delete g[i.dataset.k]; }
      // fornitore cambiato: gli articoli che avevano quel fornitore scritto lo seguono
      const cambi = [];
      for (const i of $$("[data-as-o]", d)) { const g = lista[+i.dataset.asO], o = i.value.trim().toUpperCase(); if (o && o !== g.codice) g.oracle = o; else delete g.oracle; }
      for (const i of $$("[data-as-f]", d)) { const g = lista[+i.dataset.asF], nf = i.value.trim(); if ((g.fornitore || "") !== nf) { cambi.push({ c: g.codice, da: g.fornitore || "", a: nf }); g.fornitore = nf; } }
      // stesso codice + stesso fornitore due volte: si chiede conferma, al secondo "Salva" i GMA vengono uniti
      const gruppi = new Map(); for (const g of lista) { const k = kv(g.fornitore, g.codice); gruppi.set(k, [...(gruppi.get(k) || []), g]); }
      const doppi = [...gruppi.values()].filter((l) => l.length > 1);
      if (doppi.length) {
        const firma = doppi.map((l) => kv(l[0].fornitore, l[0].codice)).sort().join(";");
        if (d._forzaUnione !== firma) {
          let box = $("#as-unisci", d);
          if (!box) { box = document.createElement("div"); box.id = "as-unisci"; box.className = "avviso small"; d.querySelector("#dlg-body, .dlg-body")?.prepend(box) || d.prepend(box); }
          box.innerHTML = `<b>Questi GMA ci sarebbero due volte con lo stesso fornitore:</b> ${doppi.map((l) => `${esc(l[0].codice)} · ${esc(l[0].fornitore || "nessun fornitore")}`).join(", ")}.
            Premi di nuovo <b>Salva</b> per <b>unirli</b> in uno solo (restano gli sconti già scritti; dove mancano, si prendono dall'altro), oppure correggi il fornitore.`;
          box.scrollIntoView({ block: "nearest" });
          d._forzaUnione = firma;
          for (const c of cambi) { const g = lista.find((x) => x.codice === c.c && x.fornitore === c.a && !x._rip); if (g) { g.fornitore = c.da; g._rip = 1; } }
          lista.forEach((g) => delete g._rip);
          toast("GMA doppi: premi di nuovo Salva per unirli.", 5000);
          return false;
        }
        for (const l of doppi) {
          const tieni = l.find((g) => !cambi.some((c) => c.c === g.codice && c.a === g.fornitore)) || l[0];
          for (const g of l) if (g !== tieni) for (const k of ["sc1", "sc2", "sc3", "aum", "nome", "sconto"]) if (!tieni[k] && g[k]) tieni[k] = g[k];
          for (const g of l) if (g !== tieni) lista.splice(lista.indexOf(g), 1);
        }
      }
      const n = $("#as-nuovo", d).value.trim().toUpperCase(), nf = $("#as-nuovo-f", d).value.trim();
      if (n) {
        if (lista.some((g) => kv(g.fornitore, g.codice) === kv(nf, n))) { toast(`Il GMA ${n} di ${nf || "nessun fornitore"} c'è già.`); return false; }
        const g = { codice: n, fornitore: nf };
        ["sc1", "sc2", "sc3", "aum"].forEach((k, i) => { const x = daPct($(`#as-nuovo-${i}`, d).value); if (x) g[k] = x; });
        lista.push(g);
      }
      for (const g of lista) if (num(g.sc1) || num(g.sc2) || num(g.sc3) || num(g.aum)) delete g.sconto;
      const ops = [{ tipo: "set", coll: COLL.cfg, id: "gmv", data: { lista } }];
      if (cambi.length) {
        const caps = new Set();
        for (const c of cambi) for (const a of S.catalogo) if (a.gmv === c.c && a.fornitore && a.fornitore.trim().toUpperCase() === c.da.trim().toUpperCase()) { a.fornitore = c.a; caps.add(a.capitolo); }
        for (const c of caps) ops.push(...opsCapitolo(c, S.catalogo.filter((a) => a.capitolo === c)));
      }
      try { await db.scrivi(ops); } catch (x) { errore(x, "Salvataggio non riuscito"); return false; }
      S.gmv = lista; impostaStandardGma(S.gmv); ridisegna(true);
      toast(n ? `GMA ${n} aggiunto.` : "GMA e acquisti standard salvati.");
    },
  });
}
$("#c-acqstd").addEventListener("click", () => apriAcquistiStandard());

function apriElencoGma() { apriAcquistiStandard(); }

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
window.__app = { S, calcolaOfferta, opzCap, BIB };
