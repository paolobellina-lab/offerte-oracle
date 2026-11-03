import * as db from "./firebase.js";
import {
  TIPI, FONTI, CAMPI_SCONTO, ETICHETTE, COLONNE_ORACLE, num, isAttivita, codiceOracle, prezzoListino,
  listinoCalcolato, prezzoBase, calcolaRiga, calcolaOfferta, righeOracle, controlloCartellino, freschezza, ordinaRighe,
  riepilogoCapitolo, ricaricoPct, marginePct, ricaricoPerTarget, isGenerico, notePreRiga, notePostRiga, esclusoDaExport,
  codiceArticolo, statoPrezzo, CODICE_GENERICO,
} from "./calc.js";
import { analizzaFile, leggiAggiornamentoListino, righeDaTestoIncollato, leggiExcel, slug } from "./importa.js";
import { scriviExcel } from "./xlsx-lite.js";

/* ================= Stato ================= */
const COLL = { cat: "oracle_catalogo", cfg: "oracle_config", cli: "oracle_clienti", off: "oracle_offerte" };
const S = {
  capitoli: [], gmv: [], clienti: [], chunks: [], catalogo: [], perK: new Map(), offerte: [],
  off: null, vista: "offerte", sub: "righe", dirty: false, salvando: false, utente: null,
  limiteCat: 300, limitePick: 300, pickCap: "", pickDb: "", importato: null, aggOracle: null,
  capAtt: "", sel: new Set(), precodici: [],
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
  if (!tr || e.target.closest("input, select, a")) return;
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
    S.utente = null; S.off = null; S.offerte = []; S.catalogo = []; S.clienti = [];
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
  $("#sb-offerta").hidden = vista !== "offerta";
  for (const v of ["offerte", "offerta", "catalogo", "clienti", "importa"]) $("#v-" + v).hidden = v !== vista;
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
  });
}

/* ================= Modello dell'offerta =================
   offerta = { id, nome, cliente, note, capitoli: [capitolo], creatoIl, aggiornatoIl, creatoDa }
   capitolo = { nome, tipo, prezzi, sottocapitoli, lineaPartenza, magazzino, globali, gmv, notaNuove,
                capitolati: [nomi dei capitolati per Oracle], capNuove, righe: [riga] }
   Il capitolato (colonna "Nome Capitolato" in Oracle) si decide offerta per offerta e non c'entra con
   i sottocapitoli del database, che servono solo a organizzare gli articoli.
   riga = copia dei dati dell'articolo + { id, qta, giorni, notePre, notePost }. Le righe sono nell'ordine scelto. */
function nuovoCapitolo(nome, base = {}) {
  return {
    nome, tipo: base.tipo || "VENDITA CLIENTE", prezzi: base.prezzi || "DA ORACLE", sottocapitoli: !!base.sottocapitoli,
    lineaPartenza: base.lineaPartenza ?? 10, magazzino: base.magazzino ?? "BG",
    globali: { ...(base.globali || {}) }, gmv: structuredClone(base.gmv || {}), notaNuove: "", capitolati: [], capNuove: "", righe: [],
  };
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
      if (!c) { c = nuovoCapitolo(r.capitolo, base); o.capitoli.push(c); }
      c.righe.push(r);
    }
    for (const k of ["righe", "tipo", "prezzi", "sottocapitoli", "lineaPartenza", "magazzino", "globali", "gmv"]) delete o[k];
  }
  for (const c of o.capitoli) {
    c.globali ||= {}; c.gmv ||= {}; c.righe ||= []; c.capitolati ||= [];
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
  const ric = num(S.capitoli.find((c) => c.nome === cap.nome)?.ricarico);
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

/* ================= Elenco offerte ================= */
$("#nuova-offerta").addEventListener("click", async () => {
  const o = nuovaOffertaVuota();
  try { await db.salva(COLL.off, o.id, o); } catch (x) { return errore(x); }
  apriOfferta(o.id, false, o);
  $("#o-nome").focus();
});
$("#cerca-offerte").addEventListener("input", disegnaElencoOfferte);

function disegnaElencoOfferte() {
  const q = $("#cerca-offerte").value.trim().toLowerCase();
  const nomeCli = (id) => S.clienti.find((c) => c.id === id)?.nome || "";
  const lista = S.offerte.filter((o) => !q || (o.nome || "").toLowerCase().includes(q) || nomeCli(o.cliente).toLowerCase().includes(q));
  $("#lista-offerte").innerHTML = lista.length ? lista.map((o) => {
    const t = totaliOfferta(o);
    return `<tr class="cliccabile" data-id="${esc(o.id)}">
      <td><b>${esc(o.nome || "(senza riferimento)")}</b></td><td>${esc(nomeCli(o.cliente))}</td>
      <td class="small">${capitoliOff(o).map((c) => `<span class="badge info">${esc(c.nome)}</span>`).join(" ")}</td>
      <td class="num">${t.righe}</td><td class="num">${euro(t.venduto)}</td>
      <td class="num">${t.venduto ? pctTxt(marginePct(t.venduto, t.acquistato)) || "0%" : ""}</td>
      <td class="nowrap muted">${dataOra(o.aggiornatoIl)}</td>
      <td class="nowrap"><button class="btn small" data-dup="${esc(o.id)}">Duplica</button></td></tr>`;
  }).join("") : `<tr><td colspan="8" class="vuoto">${S.offerte.length ? "Nessuna offerta trovata." :
    S.catalogo.length ? "Ancora nessuna offerta: crea la prima con “+ Nuova offerta”." :
    "Il database articoli è vuoto: vai in “Importa e aggiorna” e carica il file Excel."}</td></tr>`;
}
$("#lista-offerte").addEventListener("click", (e) => {
  const dup = e.target.closest("[data-dup]");
  if (dup) { e.stopPropagation(); return duplica(S.offerte.find((o) => o.id === dup.dataset.dup)); }
  const tr = e.target.closest("tr[data-id]");
  if (tr) apriOfferta(tr.dataset.id);
});

async function duplica(o) {
  if (!o) return;
  const n = { ...normalizza(o), id: db.nuovoId(), nome: (o.nome || "Offerta") + " (copia)",
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

/* --- Testata dell'offerta --- */
function seg(el, valori, attuale, etichetta = (v) => v) {
  el.innerHTML = valori.map((v) => `<button type="button" data-v="${esc(String(v))}" class="${String(v) === String(attuale) ? "on" : ""}">${esc(etichetta(v))}</button>`).join("");
}
function disegnaTestata() {
  const o = S.off;
  if (!o) return;
  $("#o-nome").value = o.nome || "";
  $("#o-note").value = o.note || "";
  $("#o-cliente").innerHTML = `<option value="">— nessun cliente (niente controllo cartellino) —</option>` +
    S.clienti.map((c) => `<option value="${esc(c.id)}" ${c.id === o.cliente ? "selected" : ""}>${esc(c.nome)}</option>`).join("");
}
$("#o-nome").addEventListener("input", (e) => { S.off.nome = e.target.value; modificata(); });
$("#o-note").addEventListener("input", (e) => { S.off.note = e.target.value; modificata(); });
$("#o-cliente").addEventListener("change", (e) => { S.off.cliente = e.target.value; modificata(); disegnaOfferta(); });

/* --- Barra dei capitoli --- */
function disegnaBarraCapitoli() {
  const o = S.off;
  const tot = (c) => riepilogoCapitolo(c, opzCap(c)).venduto;
  // sigla: una lettera per parola (KNX_MDT → KM, CONTROLLO_ACCESSI → CA), oppure le prime tre lettere (FIRE → FIR)
  const sigla = (n) => { const p = n.split(/[_\s.-]+/).filter(Boolean); return (p.length > 1 ? p.map((x) => x[0]).join("") : n).slice(0, 3).toUpperCase(); };
  $("#cap-tabs").innerHTML =
    `<button type="button" data-cap="" class="sb-voce ${S.capAtt === "" ? "on" : ""}" title="Riepilogo dell'offerta"><span class="sb-ico">${ICONA_RIEP}</span><span class="sb-testo">Riepilogo</span></button>` +
    o.capitoli.map((c) => `<button type="button" data-cap="${esc(c.nome)}" class="sb-voce ${S.capAtt === c.nome ? "on" : ""}" title="${esc(c.nome)} · ${euro(tot(c))}">
      <span class="sb-sigla">${esc(sigla(c.nome))}</span><span class="sb-testo sb-cap"><span>${esc(c.nome)}</span><span class="conta">${euro(tot(c))}</span></span></button>`).join("");
  const liberi = S.capitoli.filter((c) => !o.capitoli.some((x) => x.nome === c.nome));
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

$("#cap-tabs").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-cap]"); if (!b) return;
  S.capAtt = b.dataset.cap; S.sel = new Set();
  disegnaOfferta(true);
});
$("#add-cap").addEventListener("change", (e) => {
  const nome = e.target.value; if (!nome) return;
  const modello = capAtt() || S.off.capitoli[S.off.capitoli.length - 1];
  // il nuovo capitolo parte con tipo offerta e prezzi del capitolo attuale, ma senza sconti
  S.off.capitoli.push(nuovoCapitolo(nome, modello ? { tipo: modello.tipo, prezzi: modello.prezzi, sottocapitoli: modello.sottocapitoli } : {}));
  S.capAtt = nome; S.sub = "righe"; S.sel = new Set();
  menuAperto(false);
  modificata(); disegnaOfferta(true);
});

/* --- Impostazioni del capitolo --- */
function disegnaImpostCapitolo() {
  const c = capAtt(); if (!c) return;
  seg($("#c-tipo"), TIPI, c.tipo);
  seg($("#c-prezzi"), FONTI, c.prezzi);
  seg($("#c-sotto"), ["SI", "NO"], c.sottocapitoli ? "SI" : "NO");
  $("#c-linea").value = c.lineaPartenza ?? 10;
  $("#c-magazzino").value = c.magazzino ?? "BG";
  const ricAuto = num(S.capitoli.find((x) => x.nome === c.nome)?.ricarico);
  $("#c-globali").innerHTML = [...CAMPI_SCONTO, "ric"].map((k) => `
    <label class="f" ${k === "ric" ? 'title="Solo per VENDITA CLIENTE. Vuoto = ricarico automatico del capitolo."' : ""}>${ETICHETTE[k]}${k === "ric" ? " *" : ""}
      <span class="pct-wrap"><input class="pct" data-glob="${k}" inputmode="decimal" value="${esc(pctVal(c.globali[k]))}"
        ${k === "ric" ? (c.tipo !== "VENDITA CLIENTE" ? "disabled" : 'placeholder="auto"') : ""}></span></label>`).join("") +
    `<span class="muted small" style="align-self:flex-end;padding-bottom:7px">* ricarico solo per vendita cliente; vuoto = automatico del capitolo (${pctTxt(ricAuto) || "0%"})</span>`;
  $("#c-target-box").hidden = c.tipo !== "VENDITA CLIENTE";
}
const suCap = (fn) => (e) => { const c = capAtt(); if (!c) return; fn(c, e); };
$("#c-tipo").addEventListener("click", suCap((c, e) => {
  const b = e.target.closest("button"); if (!b) return;
  c.tipo = b.dataset.v; modificata(); disegnaOfferta();
}));
$("#c-prezzi").addEventListener("click", suCap((c, e) => {
  const b = e.target.closest("button"); if (!b) return;
  c.prezzi = b.dataset.v; modificata(); disegnaOfferta();
}));
$("#c-sotto").addEventListener("click", suCap((c, e) => {
  const b = e.target.closest("button"); if (!b) return;
  c.sottocapitoli = b.dataset.v === "SI"; modificata(); disegnaOfferta();
}));
$("#c-linea").addEventListener("change", suCap((c, e) => { c.lineaPartenza = num(e.target.value) || 10; modificata(); disegnaOfferta(); }));
$("#c-magazzino").addEventListener("input", suCap((c, e) => { c.magazzino = e.target.value.trim(); modificata(); }));
$("#c-globali").addEventListener("change", suCap((c, e) => {
  const k = e.target.dataset.glob; if (!k) return;
  const v = daPct(e.target.value);
  if (v) c.globali[k] = v; else delete c.globali[k];
  e.target.value = pctVal(v);
  modificata(); ridisegnaOfferta();
}));
$("#c-target-calcola").addEventListener("click", suCap((c) => {
  const target = num($("#c-target").value);
  if (!target) return toast("Scrivi il prezzo target del capitolo.");
  const ric = ricaricoPerTarget(c, target, opzCap(c));
  if (ric === null) return toast("Non riesco a calcolarlo: servono righe senza ricarico forzato sul GMV.");
  c.globali.ric = Math.round(ric * 1e6) / 1e6;
  modificata(); disegnaOfferta();
  toast(`Ricarico impostato a ${pctTxt(ric)} per arrivare a ${euro(target)}.`, 5000);
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
    <div class="kpi kpi-main"><span>Venduto</span><span class="big">${euro(venduto)}</span></div>
    <div class="kpi"><span>Acquistato</span><span class="big">${euro(acquistato)}</span></div>
    <div class="kpi"><span>Margine</span><span class="big">${venduto ? pctTxt(marginePct(venduto, acquistato)) || "0%" : "—"}</span></div>
    <div class="kpi"><span>Ricarico</span><span class="big">${acquistato ? pctTxt(ricaricoPct(venduto, acquistato)) || "0%" : "—"}</span></div>
    <div class="kpi"><span>Capitoli · righe</span><span class="big">${o.capitoli.length} · ${righe}</span></div>
    ${o.cliente ? `<div class="kpi"><span>Cartellino ${esc(cli?.nome || "")}</span><span>${
      fuori ? `<span class="badge bad">${fuori} GMV fuori cartellino</span>` : `<span class="badge ok">nessun problema</span>`}</span></div>` : ""}`;
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
    <span class="mini-kpi"><span>Venduto</span><b>${euro(rc.venduto)}</b></span>
    <span class="mini-kpi"><span>Acquistato</span><b>${euro(rc.acquistato)}</b></span>
    <span class="mini-kpi"><span>Margine</span><b>${rc.venduto ? pctTxt(marginePct(rc.venduto, rc.acquistato)) || "0%" : "—"}</b></span>
    <span class="mini-kpi"><span>Tipo</span><b>${esc(c.tipo)} · ${esc(c.prezzi)}</b></span>`);
  $$("#subtabs button").forEach((b) => b.classList.toggle("on", b.dataset.sub === S.sub));
  for (const s of ["righe", "prezzi", "export"]) $("#s-" + s).hidden = s !== S.sub;
  $("#conta-righe").textContent = c.righe.length ? `(${c.righe.length})` : "";
  $("#conta-gmv").textContent = rc.perGmv.size ? `(${rc.perGmv.size})` : "";
  if (S.sub === "righe") { disegnaRighe(c); if (!S.noPicker) disegnaPicker(c, cambioSub); }
  else if (S.sub === "prezzi") { disegnaGmv(c, rc); disegnaRiepilogoRighe(c); }
  else if (S.sub === "export") disegnaExport(c);
}

/* --- Riepilogo dell'offerta (ex ANALISI PACCHETTO) --- */
function disegnaRiepilogo() {
  const o = S.off, cli = clienteOff();
  const tot = { venduto: 0, acquistato: 0, attV: 0, matV: 0, righe: 0 };
  const righe = o.capitoli.map((c) => {
    const r = riepilogoCapitolo(c, opzCap(c));
    const fuori = controlloCartellino(c, cli, opzCap(c)).filter((x) => x.stato === "fuori").length;
    tot.venduto += r.venduto; tot.acquistato += r.acquistato; tot.attV += r.attivita.venduto; tot.matV += r.materiale.venduto; tot.righe += r.righe;
    return `<tr><td><button class="btn small" data-vai-cap="${esc(c.nome)}">${esc(c.nome)}</button></td>
      <td class="small nowrap">${esc(c.tipo)} · ${esc(c.prezzi)}</td><td class="num">${r.righe}</td>
      <td class="num"><b>${euro(r.venduto)}</b></td><td class="num">${euro(r.acquistato)}</td>
      <td class="num">${euro(r.materiale.venduto)}</td><td class="num">${euro(r.attivita.venduto)}</td>
      <td class="num">${r.acquistato ? pctTxt(ricaricoPct(r.venduto, r.acquistato)) || "0%" : "—"}</td>
      <td class="num">${r.venduto ? pctTxt(marginePct(r.venduto, r.acquistato)) || "0%" : "—"}</td>
      <td>${!cli ? "" : fuori ? `<span class="badge bad">${fuori} fuori</span>` : '<span class="badge ok">OK</span>'}</td>
      <td class="nowrap"><button class="btn small" data-xlsx-cap="${esc(c.nome)}" ${r.righe ? "" : "disabled"}>Excel Oracle</button>
        <button class="icon" data-togli-cap="${esc(c.nome)}" title="Togli il capitolo dall'offerta">✕</button></td></tr>`;
  }).join("");
  $("#riepilogo-capitoli").innerHTML = righe ? righe + `<tr class="totale"><td>Totale offerta</td><td></td><td class="num">${tot.righe}</td>
      <td class="num">${euro(tot.venduto)}</td><td class="num">${euro(tot.acquistato)}</td><td class="num">${euro(tot.matV)}</td><td class="num">${euro(tot.attV)}</td>
      <td class="num">${tot.acquistato ? pctTxt(ricaricoPct(tot.venduto, tot.acquistato)) || "0%" : ""}</td>
      <td class="num">${tot.venduto ? pctTxt(marginePct(tot.venduto, tot.acquistato)) || "0%" : ""}</td><td></td><td></td></tr>`
    : `<tr><td colspan="11" class="vuoto">Nessun capitolo: aggiungine uno con “+ Aggiungi capitolo”.</td></tr>`;
  $("#scarica-tutti").disabled = !tot.righe;
}
$("#riepilogo-capitoli").addEventListener("click", (e) => {
  const v = e.target.closest("[data-vai-cap]");
  if (v) { S.capAtt = v.dataset.vaiCap; S.sub = "righe"; S.sel = new Set(); return disegnaOfferta(true); }
  const t = e.target.closest("[data-togli-cap]");
  if (t) return togliCapitolo(S.off.capitoli.find((c) => c.nome === t.dataset.togliCap));
  const x = e.target.closest("[data-xlsx-cap]");
  if (x) scaricaXlsx(S.off.capitoli.find((c) => c.nome === x.dataset.xlsxCap));
});
$("#scarica-tutti").addEventListener("click", () => {
  const caps = S.off.capitoli.filter((c) => calcolaOfferta(c).righe.length);
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
function disegnaCapitolati(c) {
  const conta = (n) => c.righe.filter((r) => r.capitolato === n).length;
  $("#capitolati").innerHTML = `<span class="muted small" style="font-weight:600">Capitolati per Oracle:</span>` +
    c.capitolati.map((n, i) => `<span class="chip"><button type="button" class="chip-x" data-capsu="${i}" title="Sposta prima" ${i === 0 ? "disabled" : ""}>‹</button><button type="button" class="chip-nome" data-capi="${i}" title="Rinomina">${esc(n)}</button><span class="muted">${conta(n)}</span><button type="button" class="chip-x" data-capgiu="${i}" title="Sposta dopo" ${i === c.capitolati.length - 1 ? "disabled" : ""}>›</button><button type="button" class="chip-x" data-capx="${i}" title="Elimina il capitolato">✕</button></span>`).join("") +
    `<input id="nuovo-capitolato" placeholder="nuovo capitolato, es. PIANO TERRA" style="min-width:220px"><button class="btn small" id="aggiungi-capitolato" type="button">+ Aggiungi</button>`;
  const opz = `<option value="">— capitolato —</option>` + c.capitolati.map((n) => `<option>${esc(n)}</option>`).join("");
  const prev = $("#bulk-cap").value;
  $("#bulk-cap").innerHTML = opz;
  if (c.capitolati.includes(prev)) $("#bulk-cap").value = prev;
}
function aggiungiCapitolato(c, nome) {
  nome = (nome || "").trim();
  if (!nome) return false;
  if (!c.capitolati.includes(nome)) c.capitolati.push(nome);
  return true;
}
$("#capitolati").addEventListener("click", suCap((c, e) => {
  if (e.target.id === "aggiungi-capitolato") {
    if (!aggiungiCapitolato(c, $("#nuovo-capitolato").value)) return toast("Scrivi il nome del capitolato.");
    modificata(); return disegnaOfferta();
  }
  const su = e.target.closest("[data-capsu]"), giu = e.target.closest("[data-capgiu]");
  if (su || giu) {
    const i = +(su || giu).dataset[su ? "capsu" : "capgiu"], j = su ? i - 1 : i + 1;
    if (j < 0 || j >= c.capitolati.length) return;
    [c.capitolati[i], c.capitolati[j]] = [c.capitolati[j], c.capitolati[i]];
    modificata(); return disegnaOfferta();
  }
  const x = e.target.closest("[data-capx]");
  if (x) {
    const nome = c.capitolati[+x.dataset.capx];
    c.capitolati.splice(+x.dataset.capx, 1);
    c.righe.forEach((r) => { if (r.capitolato === nome) r.capitolato = ""; });
    if (c.capNuove === nome) c.capNuove = "";
    c.esclusiExport = (c.esclusiExport || []).filter((x) => x !== nome);
    modificata(); return disegnaOfferta();
  }
  const b = e.target.closest("[data-capi]");
  if (b) {
    const i = +b.dataset.capi, vecchio = c.capitolati[i];
    dialogo({
      titolo: "Rinomina capitolato",
      html: `<label class="f">Nome che va in Oracle (Nome Capitolato)<input id="r-cap" value="${esc(vecchio)}"></label>`,
      pulsanti: [{ label: "Annulla", value: "annulla" }, { label: "Rinomina", value: "ok", cls: "primary" }],
      onConferma(v, body) {
        const nuovo = $("#r-cap", body).value.trim();
        if (!nuovo) return false;
        c.capitolati[i] = nuovo;
        c.righe.forEach((r) => { if (r.capitolato === vecchio) r.capitolato = nuovo; });
        if (c.capNuove === vecchio) c.capNuove = nuovo;
        c.esclusiExport = (c.esclusiExport || []).map((x) => (x === vecchio ? nuovo : x));
        modificata(); disegnaOfferta();
      },
    });
  }
}));
$("#capitolati").addEventListener("keydown", (e) => {
  if (e.target.id === "nuovo-capitolato" && e.key === "Enter") { e.preventDefault(); $("#aggiungi-capitolato").click(); }
});

function disegnaRighe(c) {
  disegnaCapitolati(c);
  const tb = $("#righe-offerta");
  const sel = S.sel;
  for (const id of [...sel]) if (!c.righe.some((r) => r.id === id)) sel.delete(id);
  $("#bulk-n").textContent = !sel.size ? "Nessuna riga selezionata" : sel.size === 1 ? "1 riga selezionata" : `${sel.size} righe selezionate`;
  $("#bulk").classList.toggle("attiva", sel.size > 0);
  $$("#bulk [data-bulk]").forEach((b) => { if (!b.hasAttribute("data-sempre")) b.disabled = !sel.size; });
  $("#sel-tutte").checked = c.righe.length > 0 && sel.size === c.righe.length;
  if (!c.righe.length) {
    tb.innerHTML = `<tr><td colspan="12" class="vuoto">Nessuna riga: cerca gli articoli qui sopra e scrivi le quantità.</td></tr>`;
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
    let testa = "";
    if ((r.capitolato || "") !== capPrec) {
      capPrec = r.capitolato || "";
      testa = `<tr class="gruppo" data-gruppo="${esc(capPrec)}"><td colspan="12"><input type="checkbox" class="sel-gruppo" data-selgruppo title="Seleziona tutte le righe di questo capitolato"> ${capPrec ? esc(capPrec) : '<span style="font-weight:500">senza capitolato</span>'}</td></tr>`;
    }
    return testa + `<tr data-id="${esc(r.id)}" class="${sel.has(r.id) ? "sel" : ""}${num(r.qta) ? "" : " qta-zero"}">
      <td><input type="checkbox" data-sel ${sel.has(r.id) ? "checked" : ""}></td>
      <td class="num muted small">${linee.get(r.id) ?? ""}</td>
      <td class="mono nowrap">${codiceArticolo(r) ? codiceVis(r) : '<span class="badge bad">senza codice</span>'} ${statoRiga(r)}</td>
      <td class="desc"><b>${esc(r.nOrdine)}</b>${r.capitolo !== c.nome ? ` <span class="badge info" title="Articolo preso da un altro database">${esc(r.capitolo)}</span>` : ""}<div class="d">${esc(r.descrizione)}</div></td>
      <td>${esc(r.um)}</td>
      <td class="num"><input class="qta" data-campo="qta" inputmode="decimal" value="${esc(numIt(r.qta))}">${num(x.qta) !== num(r.qta) && num(r.qta) ? `<div class="muted small nowrap" title="Quantità che va in Oracle">Oracle: ${esc(numIt(x.qta))}</div>` : ""}</td>
      <td class="num nowrap">${att ? `<input class="qta" data-campo="giorni" inputmode="decimal" value="${esc(numIt(r.giorni))}">
        <label class="giornate" title="In vendita e a listino: in Oracle quantità = giorni e prezzo = valore giornaliero (invece di un valore complessivo con quantità 1)">
          <input type="checkbox" data-campo="aGiornate" ${r.aGiornate ? "checked" : ""} ${c.tipo === "ACQUISTO" ? "disabled" : ""}> a giornate</label>` : ""}</td>
      <td class="num"><b>${euro(x.unitario)}</b></td><td class="num">${euro(num(r.qta) ? x.totale : 0)}</td>
      <td><select class="capsel" data-campo="capitolato">${opzioniCap(r.capitolato)}</select></td>
      ${isGenerico(r)
        ? `<td><span class="nota-auto" title="Articolo generico: nota pre-linea automatica = n° di ordinazione">auto: ${esc(r.nOrdine)}</span></td>
           <td><span class="nota-auto" title="Articolo generico: nota post-linea automatica = descrizione">auto: descrizione</span></td>`
        : `<td><input class="nota" data-campo="notePre" value="${esc(r.notePre)}" placeholder="nota pre-linea"></td>
           <td><input class="nota" data-campo="notePost" value="${esc(r.notePost)}" placeholder="nota post-linea"></td>`}</tr>`;
  }).join("") + `<tr class="sotto"><td colspan="8" class="num">Totale ${esc(c.nome)}</td><td class="num">${euro(tot)}</td><td colspan="3"></td></tr>`;
  applicaGruppi(tb);
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
  if (!righe.length) { tb.innerHTML = `<tr><td colspan="13" class="vuoto">Nessuna riga: aggiungile nella pagina “1 · Righe”.</td></tr>`; return; }
  const linee = lineeOracle(c), opz = opzCap(c);
  let tot = 0, capPrec = null;
  tb.innerHTML = righe.map((r) => {
    const x = calcolaRiga(r, c, opz.ricarichi);
    tot += x.totale;
    let testa = "";
    if ((r.capitolato || "") !== capPrec) {
      capPrec = r.capitolato || "";
      testa = `<tr class="gruppo" data-gruppo="${esc(capPrec)}"><td colspan="13">${capPrec ? esc(capPrec) : '<span style="font-weight:500">senza capitolato</span>'}</td></tr>`;
    }
    return testa + `<tr data-id="${esc(r.id)}">
      <td class="num muted small">${linee.get(r.id) ?? ""}</td>
      <td class="mono nowrap">${codiceArticolo(r) ? codiceVis(r) : '<span class="badge bad">senza codice</span>'} ${statoRiga(r)}</td>
      <td class="desc"><b>${esc(r.nOrdine)}</b><div class="d">${esc(r.descrizione)}</div></td>
      <td class="num">${esc(numIt(x.qta))}</td>
      <td class="num">${isGenerico(r) ? `<input class="qta prezzo" data-campo="prezzoGen" inputmode="decimal" value="${esc(numIt(r.listinoAcquisto))}" title="Prezzo del generico: se lo cambi si aggiorna anche nel database con la data di oggi">` : euro(x.base)}</td>
      <td class="num">${pctTxt(x.sc3)}</td><td class="num">${pctTxt(x.sc4)}</td><td class="num">${pctTxt(x.sc5)}</td>
      <td class="num">${pctTxt(x.aum1)}</td><td class="num">${pctTxt(x.ric)}</td>
      <td class="num"><b>${euro(x.unitario)}</b></td><td class="num">${euro(x.totale)}</td>
      <td class="mono">${esc(r.gmv)}</td></tr>`;
  }).join("") + `<tr class="sotto"><td colspan="11" class="num">Totale ${esc(c.nome)}</td><td class="num">${euro(tot)}</td><td></td></tr>`;
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
      salvaCapitolo(art.capitolo).then(() => toast(`Prezzo aggiornato anche nel database (${art.capitolo}) con la data di oggi.`))
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
/* righe del blocco sotto un'intestazione di capitolato (fino alla successiva) */
function righeDelBlocco(h) {
  const ids = [];
  for (let tr = h.nextElementSibling; tr && !tr.hasAttribute("data-gruppo") && !tr.classList.contains("sotto"); tr = tr.nextElementSibling) if (tr.dataset.id) ids.push(tr.dataset.id);
  return ids;
}
/* Selezione: casella, clic su codice/descrizione, Maiusc+clic per un intervallo, casella del capitolato. */
$("#righe-offerta").addEventListener("click", (e) => {
  const c = capAtt(); if (!c) return;
  const g = e.target.closest("[data-selgruppo]");
  if (g) {
    const ids = righeDelBlocco(g.closest("tr"));
    ids.forEach((id) => (g.checked ? S.sel.add(id) : S.sel.delete(id)));
    return disegnaOfferta();
  }
  const tr = e.target.closest("tr[data-id]"); if (!tr) return;
  const suCasella = e.target.matches("[data-sel]");
  const suTesto = !e.target.closest("input, select, button, label") && e.target.closest("td.desc, td.mono, td.num:nth-child(2)");
  if (!suCasella && !suTesto) return;
  const id = tr.dataset.id;
  if (e.shiftKey && S.ultimoSel && S.ultimoSel !== id) {
    const a = c.righe.findIndex((r) => r.id === S.ultimoSel), b = c.righe.findIndex((r) => r.id === id);
    if (a >= 0 && b >= 0) { const [x, y] = a < b ? [a, b] : [b, a]; for (let k = x; k <= y; k++) S.sel.add(c.righe[k].id); }
    window.getSelection?.().removeAllRanges();
  } else if (suCasella) {
    e.target.checked ? S.sel.add(id) : S.sel.delete(id);
  } else {
    S.sel.has(id) ? S.sel.delete(id) : S.sel.add(id);
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
$("#bulk").addEventListener("click", (e) => {
  const b = e.target.closest("[data-bulk]"); if (!b) return;
  const c = capAtt(); if (!c) return;
  const scelte = c.righe.filter((r) => S.sel.has(r.id));
  const nota = $("#bulk-nota").value.trim();
  const capSc = $("#bulk-cap").value;
  const az = b.dataset.bulk;
  if (az === "dup") {
    // le copie vanno in fondo, nello stesso ordine, con il capitolato e la nota scelti (se indicati)
    const copie = scelte.map((r) => ({ ...structuredClone(r), id: db.nuovoId(),
      ...(capSc ? { capitolato: capSc } : {}), ...(nota && !isGenerico(r) ? { notePost: nota } : {}) }));
    c.righe.push(...copie);
    S.sel = new Set(copie.map((r) => r.id));
    toast(`Duplicate ${copie.length} righe${capSc ? ` nel capitolato “${capSc}”` : ""}${nota ? ` con nota “${nota}”` : ""}: le trovi in fondo, già selezionate.`, 4500);
  } else if (az === "cap") {
    if (!capSc) return toast("Scegli il capitolato (o crealo nella barra “Capitolati per Oracle”).");
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
    if (scelte.length > 1 && !confirm(`Eliminare ${scelte.length} righe dall'offerta?`)) return;
    c.righe = c.righe.filter((r) => !S.sel.has(r.id));
    S.sel = new Set();
  } else if (az === "nessuna") S.sel = new Set();
  modificata(); disegnaOfferta();
});
$("#aggiorna-da-catalogo").addEventListener("click", () => {
  const c = capAtt(); if (!c) return;
  let n = 0, mancanti = 0;
  c.righe = c.righe.map((r) => {
    const a = S.perK.get(r.k) || S.catalogo.find((x) => x.capitolo === r.capitolo && codiceArticolo(x) === codiceArticolo(r) && x.nOrdine === r.nOrdine);
    if (!a) { mancanti++; return r; }
    n++;
    const nuovo = rigaDaArticolo(a);
    // restano quelle della riga: quantità, giorni dell'attività e note
    return { ...nuovo, id: r.id, qta: r.qta, giorni: isAttivita(r) ? r.giorni : a.giorni, notePre: r.notePre, notePost: r.notePost };
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
  if (S.pickCap !== c.nome || !dbs.includes(S.pickDb)) { S.pickCap = c.nome; S.pickDb = dbs.includes(c.nome) ? c.nome : dbs[0] || ""; reset = true; }
  $("#a-db").innerHTML = dbs.map((d) => `<option value="${esc(d)}" ${d === S.pickDb ? "selected" : ""}>${esc(d)}${d === c.nome ? " (questo capitolo)" : ""}</option>`).join("");
  $("#a-nota").value = c.notaNuove || "";
  $("#a-cap").innerHTML = `<option value="">— nessuno —</option>` + c.capitolati.map((n) => `<option ${n === c.capNuove ? "selected" : ""}>${esc(n)}</option>`).join("") +
    `<option value="__nuovo">+ nuovo capitolato…</option>`;
  const nelDb = S.catalogo.filter((a) => a.capitolo === S.pickDb);
  const sotti = [...new Set(nelDb.map((a) => a.sottocapitolo).filter(Boolean))];
  const sottoSel = $("#a-sotto");
  const prev = reset ? "" : sottoSel.value;
  sottoSel.innerHTML = `<option value="">Tutti i sottocapitoli</option>` + sotti.map((s) => `<option ${s === prev ? "selected" : ""}>${esc(s)}</option>`).join("");
  const q = $("#a-cerca").value.trim().toLowerCase();
  const solo = $("#a-solo").checked;
  const conta = contaInOfferta(c);
  const lista = nelDb.filter((a) => (!sottoSel.value || a.sottocapitolo === sottoSel.value) &&
    (!solo || conta.has(a.k)) &&
    (!q || `${codiceArticolo(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)));
  const vis = lista.slice(0, S.limitePick);
  let sotto = null, html = "";
  for (const a of vis) {
    if ((a.sottocapitolo || "") !== sotto && !sottoSel.value) {
      sotto = a.sottocapitolo || "";
      html += `<tr class="sotto" data-gruppo="${esc(sotto)}"><td colspan="8">${esc(sotto || "—")}</td></tr>`;
    }
    const att = isAttivita(a);
    html += `<tr data-k="${esc(a.k)}">
      <td class="mono nowrap">${codiceVis(a)}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b><div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td><td class="num">${euro(prezzoListino(a))}</td>
      <td class="mono">${esc(a.gmv)}</td>
      <td>${pallino(a)}</td>
      <td class="num nowrap">${att ? `<input class="qta" data-pick="giorni" inputmode="decimal" title="Giorni" value="${esc(numIt(a.giorni))}"> gg ` : ""}<input class="qta" data-pick="qta" inputmode="decimal" placeholder="q.tà"></td></tr>`;
  }
  $("#righe-picker").innerHTML = html || `<tr><td colspan="8" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Il database è vuoto: importalo dal file Excel."}</td></tr>`;
  applicaGruppi($("#righe-picker"));
  $("#picker-info").textContent = lista.length > vis.length ? `Mostrati ${vis.length} di ${lista.length}: affina la ricerca.` :
    `${lista.length} articoli. Scrivi la quantità e premi Invio o Tab: ogni volta si aggiunge una riga nuova in fondo all'offerta.`;
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
$("#a-solo").addEventListener("change", () => { const c = capAtt(); if (c) disegnaPicker(c); });
$("#a-nota").addEventListener("input", suCap((c, e) => { c.notaNuove = e.target.value; modificata(); }));
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
  if (c.notaNuove && !r.notePost && !isGenerico(r)) r.notePost = c.notaNuove;
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

/* --- GMV: peso, sconti e cartellino --- */
function disegnaGmv(c, rc) {
  const esiti = new Map(controlloCartellino(c, clienteOff(), opzCap(c)).map((x) => [x.gmv, x]));
  const html = [...rc.perGmv.entries()].sort((a, b) => b[1].venduto - a[1].venduto).map(([g, info]) => {
    const peso = rc.venduto ? info.venduto / rc.venduto : 0;
    const comuni = `<td class="num">${info.righe}</td><td class="num">${euro(info.venduto)}</td><td class="num">${euro(info.acquistato)}</td>
      <td class="num"><div class="peso"><span style="width:${Math.max(0, Math.min(100, peso * 100)).toFixed(1)}%"></span></div>${pctTxt(peso) || "0%"}</td>
      <td class="num">${info.venduto ? pctTxt(marginePct(info.venduto, info.acquistato)) || "0%" : "—"}</td>`;
    if (!g) return `<tr><td class="muted">(senza GMV)</td><td></td>${comuni}<td colspan="8" class="muted small">Senza GMV valgono gli sconti del capitolo o del database.</td></tr>`;
    const v = c.gmv[g] || {};
    const e = esiti.get(g);
    const inp = (k) => `<span class="pct-wrap"><input class="pct" data-gmv="${esc(g)}" data-k="${k}" inputmode="decimal" value="${esc(pctVal(v[k]))}"
      ${k === "ric" && c.tipo !== "VENDITA CLIENTE" ? "disabled" : ""}></span>`;
    const att = e?.tipo === "attivita";
    const badge = !e ? '<span class="muted small">—</span>' :
      `<span class="badge ${e.stato === "ok" ? "ok" : e.stato === "pari" ? "warn" : "bad"}">${esc(e.testo)}</span>`;
    return `<tr><td class="mono"><b>${esc(g)}</b></td><td>${esc(fornitore(g))}</td>${comuni}
      <td>${inp("sc3")}</td><td>${inp("sc4")}</td><td>${inp("sc5")}</td><td>${inp("aum1")}</td>
      <td>${inp("ric")}</td>
      <td class="num">${att ? euro(e.offerta) + "/g" : pctTxt(1 - info.prima.c.fattore) || "0%"}</td>
      <td class="num">${e ? (att ? euro(e.cartellino) + "/g" : pctTxt(e.cartellino) || "0%") : ""}</td><td>${badge}</td></tr>`;
  }).join("");
  $("#righe-gmv").innerHTML = html || `<tr><td colspan="15" class="vuoto">Aggiungi articoli al capitolo per vedere i loro GMV.</td></tr>`;
}
$("#righe-gmv").addEventListener("change", suCap((c, e) => {
  const inp = e.target.closest("input[data-gmv]"); if (!inp) return;
  const g = inp.dataset.gmv, k = inp.dataset.k;
  const v = daPct(inp.value);
  c.gmv[g] ||= {};
  if (v) c.gmv[g][k] = v; else delete c.gmv[g][k];
  if (!Object.keys(c.gmv[g]).length) delete c.gmv[g];
  modificata(); ridisegnaOfferta();
}));

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
    g.righe++; g.totale += calcolaRiga(r, c, opzCap(c).ricarichi).totale;
  }
  const escl = new Set(c.esclusiExport || []);
  $("#export-capitolati").innerHTML = gruppi.length > 1 || escl.size ? `<span class="muted small" style="font-weight:600">Capitolati da caricare in questa offerta Oracle:</span>` +
    gruppi.map((g) => `<label class="chip-check"><input type="checkbox" data-escl="${esc(g.nome)}" ${escl.has(g.nome) ? "" : "checked"}>
      <b>${g.nome ? esc(g.nome) : "senza capitolato"}</b> <span class="muted">${g.righe} righe · ${euro(g.totale)}</span></label>`).join("") +
    `<button class="btn small" id="export-tutti" type="button">Tutti</button>` : "";
  const { righe, avvisi } = datiOracle(c);
  const av = [];
  if (escl.size) av.push(`<div class="avviso info">Esportazione parziale: ${[...escl].map((n) => `“${esc(n || "senza capitolato")}”`).join(", ")} ${escl.size === 1 ? "escluso" : "esclusi"}. Le linee ripartono da ${num(c.lineaPartenza) || 10}.</div>`);
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
function scaricaXlsx(c) {
  if (!c) return;
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

function disegnaCatalogo() {
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
    (!q || `${codiceArticolo(a)} ${a.nOrdine} ${a.descrizione} ${a.gmv}`.toLowerCase().includes(q)));
  const ord = ordinaRighe(lista, ordineCap());
  const vis = ord.slice(0, S.limiteCat);
  let gruppo = null, html = "";
  for (const a of vis) {
    const g = (cap ? "" : a.capitolo + " · ") + (a.sottocapitolo || "—");
    if (g !== gruppo) { gruppo = g; html += `<tr class="sotto" data-gruppo="${esc(g)}"><td colspan="12">${esc(g)}</td></tr>`; }
    html += `<tr class="cliccabile" data-k="${esc(a.k)}">
      <td class="mono nowrap">${codiceVis(a)}</td>
      <td class="desc"><b>${esc(a.nOrdine)}</b><div class="d">${esc(a.descrizione)}</div></td>
      <td>${esc(a.um)}</td><td class="num">${euro(a.listinoAcquisto)}</td><td class="num">${euro(prezzoListino(a))}</td>
      <td class="num">${pctTxt(a.sc3)}</td><td class="num">${pctTxt(a.sc4)}</td><td class="num">${pctTxt(a.sc5)}</td>
      <td class="num">${pctTxt(a.aum1)}</td><td class="num">${pctTxt(a.riparam)}</td><td class="mono">${esc(a.gmv)}</td>
      <td class="nowrap">${pallino(a)}${esc(dataIt(a.dataAgg))}</td></tr>`;
  }
  $("#righe-catalogo").innerHTML = html || `<tr><td colspan="12" class="vuoto">${S.catalogo.length ? "Nessun articolo trovato." : "Catalogo vuoto: importalo da “Importa e aggiorna”."}</td></tr>`;
  applicaGruppi($("#righe-catalogo"));
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
function modificaArticolo(a) {
  const nuovo = !a;
  const capDef = $("#c-capitolo").value || S.capitoli[0]?.nome || "";
  a = a || { capitolo: capDef, sottocapitolo: $("#c-sotto").value || "", precodice: "", codice: "", nOrdine: "", descrizione: "", um: "n.", listinoAcquisto: 0, sc3: 0, sc4: 0, sc5: 0, aum1: 0, riparam: 0, giorni: 0, gmv: "", dataAgg: oggiIso() };
  const auto = !num(a.listino) || Math.abs(num(a.listino) - listinoCalcolato(a)) < 0.005;
  const sotti = [...new Set(S.catalogo.filter((x) => x.capitolo === a.capitolo).map((x) => x.sottocapitolo).filter(Boolean))];
  dialogo({
    titolo: nuovo ? "Nuovo articolo" : `${codiceArticolo(a)} · ${a.nOrdine}`,
    html: `<div class="grid-set">
      <label class="f">Capitolo<select id="m-cap">${S.capitoli.map((c) => `<option ${c.nome === a.capitolo ? "selected" : ""}>${esc(c.nome)}</option>`).join("")}</select></label>
      <label class="f wide">Sottocapitolo del database (scegline uno o scrivine uno nuovo)<input id="m-sotto" list="m-sotti" value="${esc(a.sottocapitolo || "")}"><datalist id="m-sotti">${sotti.map((s) => `<option value="${esc(s)}">`).join("")}</datalist></label>
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
    <label class="row small" style="margin-top:8px"><input type="checkbox" id="m-gen" ${isGenerico(a) ? "checked" : ""}>
      <b>Tratta come articolo generico</b> — in Oracle: ${CODICE_GENERICO}, prezzo netto, n° di ordinazione nella nota pre-linea e descrizione nella post-linea</label>
    <p class="muted small" style="margin:4px 0 0">Se cambi un prezzo o uno sconto, la data di aggiornamento diventa quella di oggi (puoi correggerla).</p>
    <label class="row small" style="margin-top:8px"><input type="checkbox" id="m-auto" ${auto ? "checked" : ""}>
      Prezzo di listino calcolato automaticamente (come la colonna K dell'Excel)</label>
    <label class="f" style="margin-top:10px">Descrizione<textarea id="m-des" rows="4">${esc(a.descrizione)}</textarea></label>
    <div class="grid-set" style="margin-top:10px" id="m-note-box">
      <label class="f wide">Note pre-linea<textarea id="m-npre" rows="2">${esc(a.notePre)}</textarea></label>
      <label class="f wide">Note post-linea<textarea id="m-npost" rows="2">${esc(a.notePost)}</textarea></label>
    </div>
    <p class="muted small" id="m-note-gen">Articolo generico (P000 ARTICOLO_GEN): in Oracle la nota pre-linea è il n° di ordinazione e la nota post-linea è la descrizione.</p>`,
    pulsanti: [...(nuovo ? [] : [{ label: "Elimina articolo", value: "del", cls: "danger" }]),
      { label: "Annulla", value: "annulla" }, { label: "Salva", value: "ok", cls: "primary" }],
    onApri(b) {
      const sync = () => {
        const li = $("#m-li", b);
        li.disabled = $("#m-auto", b).checked;
        if (li.disabled) li.value = String(Math.round(listinoCalcolato(leggi(b)) * 10000) / 10000).replace(".", ",");
      };
      const gen = () => { const g = isGenerico({ codice: $("#m-cod", b).value, generico: $("#m-gen", b).checked }); $("#m-note-box", b).hidden = g; $("#m-note-gen", b).hidden = !g; };
      // data automatica: quando cambia un prezzo o uno sconto, la data di aggiornamento diventa oggi
      const prezzi = ["#m-la", "#m-li", "#m-sc3", "#m-sc4", "#m-sc5", "#m-aum1", "#m-rp"];
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
      ${r.capitoli.map((c) => `<tr><td>${esc(c.nome)}${c.blocco0 ? ` <span class="badge info" title="Righe copiate in testa al foglio per l'offerta in corso: non entrano nel database perché ci sono già">${c.blocco0} righe del blocco 0 escluse</span>` : ""}</td><td class="num">${c.articoli}</td><td class="num">${pctTxt(c.ricarico)}${c.ricaricoDaFile ? "" : ' <span class="badge warn" title="Il capitolo non è nella tabella Capitoli dell\'Excel">predefinito</span>'}</td>
        <td class="num">${r.conQuantita.filter((a) => a.capitolo === c.nome).length || ""}</td>
        <td>${giaCat.has(c.nome) ? '<span class="badge warn">sostituisce l\'attuale</span>' : '<span class="badge info">nuovo</span>'}</td></tr>`).join("")}
      </tbody></table></div>
      <div class="stack" style="margin-top:10px">
        <label class="row"><input type="checkbox" id="i-cat" checked> Catalogo articoli (sostituisce i capitoli elencati sopra; le offerte già fatte non cambiano)</label>
        <label class="row"><input type="checkbox" id="i-cap" checked> Capitoli e ricarichi automatici, elenco GMV</label>
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
    const wb = await leggiExcel(await f.arrayBuffer());
    const nome = wb.nomi.find((n) => /AGGIORNAMENTO LISTINO$/i.test(n)) || wb.nomi[0];
    confrontaOracle(await wb.foglio(nome));
  } catch (x) { errore(x, "File non leggibile"); }
});

const CODICI_GENERICI = new Set([CODICE_GENERICO, "PITATT.TEC"]);
function confrontaOracle(righe) {
  const box = $("#riepilogo-oracle");
  let mappa;
  try { mappa = leggiAggiornamentoListino(righe); } catch (x) { box.innerHTML = `<div class="avviso bad">${esc(x.message)}</div>`; return; }
  const cambi = [];
  let uguali = 0;
  const trovati = new Set();
  for (const a of S.catalogo) {
    const cod = codiceArticolo(a);
    if (isGenerico(a) || CODICI_GENERICI.has(cod)) continue;
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
        return `<tr><td>${esc(a.capitolo)}</td><td class="mono">${esc(codiceArticolo(a))}</td><td>${esc(a.nOrdine)}</td>
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
window.__app = { S, calcolaOfferta, prezzoBase, opzCap };
