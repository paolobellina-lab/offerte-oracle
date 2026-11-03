// Importazione dal file Excel "Template_Import_Righe-OFFERTA": cataloghi, GMV, capitoli e cartellini clienti.
import { leggiExcel, dataDaSeriale } from "./xlsx-lite.js";
import { num } from "./calc.js";

// Colonne dei fogli catalogo (A=0)
const C = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9, K: 10, L: 11, M: 12, N: 13, O: 14,
  P: 15, Q: 16, R: 17, S: 18, T: 19, U: 20, W: 22, X: 23, Z: 25 };

const testo = (v) => {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return ""; // errori come #N/A
  return String(v).trim();
};
const numero = (v) => (v && typeof v === "object" ? 0 : num(v));
const gmvPulito = (v) => { const s = testo(v); return s === "-" || s === "0" ? "" : s; };

function eFoglioCatalogo(righe) {
  const h = righe?.[0] || [];
  return testo(h[C.D]).toUpperCase() === "CODICE" && testo(h[C.N]).toUpperCase().includes("LISTINO ACQUISTO");
}

/* Righe del catalogo di un foglio. Le righe con RIF CAP e senza numero di ordinazione sono le testate dei sottocapitoli. */
function leggiCatalogo(nome, righe) {
  const articoli = [];
  let sotto = "", rif = "";
  const chiavi = new Map();
  for (let i = 2; i < righe.length; i++) {
    const r = righe[i] || [];
    const rifCap = testo(r[C.F]), nOrd = testo(r[C.G]), cod = testo(r[C.D]);
    if (rifCap && !nOrd) { // testata sottocapitolo
      rif = rifCap;
      sotto = testo(r[C.H]) || testo(r[C.B]) || sotto;
      continue;
    }
    if (!cod && !nOrd) continue; // titoli o righe vuote
    const precodice = testo(r[C.C]);
    const sottocapitolo = testo(r[C.B]) && testo(r[C.B]) !== "0" ? testo(r[C.B]) : sotto;
    const base = `${nome}|${rif}|${precodice}${cod}|${nOrd}`;
    const n = (chiavi.get(base) || 0) + 1;
    chiavi.set(base, n);
    const qta = numero(r[C.J]);
    articoli.push({
      k: n > 1 ? `${base}|${n}` : base,
      capitolo: nome,
      ordine: i + 1, // riga del foglio Excel
      rifCap: rif,
      sottocapitolo,
      pos: testo(r[C.E]),
      precodice,
      codice: cod,
      nOrdine: nOrd,
      descrizione: testo(r[C.H]),
      um: testo(r[C.I]),
      listinoAcquisto: numero(r[C.N]),
      giorni: numero(r[C.O]),
      riparam: numero(r[C.P]),
      sc3: numero(r[C.Q]),
      sc4: numero(r[C.R]),
      sc5: numero(r[C.S]),
      aum1: numero(r[C.T]),
      gmv: gmvPulito(r[C.U]),
      // per gli ARTICOLO_GEN le note sono formule (=G, =H): l'app le ricava da sola, non si salvano
      notePre: cod.toUpperCase() === "ARTICOLO_GEN" || testo(r[C.W]) === "0" ? "" : testo(r[C.W]),
      notePost: cod.toUpperCase() === "ARTICOLO_GEN" || testo(r[C.X]) === "0" ? "" : testo(r[C.X]),
      dataAgg: dataDaSeriale(r[C.Z]),
      listino: numero(r[C.K]), // colonna K "Prezzo di Listino" come calcolata (o scritta) nell'Excel
      _qta: qta,
      _blocco0: rif === "0",
    });
  }
  /* Blocco 0 = zona di lavoro dell'offerta in testa al foglio (righe copiate dal database per l'offerta in corso).
     Le sue righe non entrano nel database se lo stesso articolo c'è già più sotto. */
  const chiave = (a) => `${a.precodice}${a.codice}|${a.nOrdine}`;
  const nelDb = new Map();
  for (const a of articoli) if (!a._blocco0 && !nelDb.has(chiave(a))) nelDb.set(chiave(a), a);
  const db = [], lavoro = [];
  for (const a of articoli) {
    const doppio = a._blocco0 && nelDb.get(chiave(a));
    if (doppio) { a._kCatalogo = doppio.k; lavoro.push(a); } else db.push(a);
  }
  return { articoli: db, lavoro };
}

/* Capitoli e ricarico automatico (foglio ELENCHI, tabella Capitoli). */
function leggiCapitoli(righe) {
  const out = [];
  if (!righe) return out;
  let inTabella = false;
  for (const r of righe) {
    if (!r) continue;
    if (testo(r[2]) === "Capitoli") { inTabella = true; continue; }
    if (!inTabella) continue;
    const nome = testo(r[2]);
    if (!nome || typeof r[1] !== "number" || out.some((c) => c.nome === nome)) break; // fine tabella (riga dei totali)
    out.push({ nome, ricarico: numero(r[3]) });
  }
  return out;
}

/* GMV e fornitori (TIPO OFFERTA righe 13–112, colonne B e C). */
function leggiGmv(righe) {
  const out = [];
  if (!righe) return out;
  for (let i = 12; i < Math.min(righe.length, 112); i++) {
    const r = righe[i] || [];
    const codice = testo(r[2]), fornitore = testo(r[1]);
    if (!codice || codice === "-" || /^gmv\d+$/i.test(codice)) continue;
    out.push({ codice, fornitore });
  }
  return out;
}

/* Cartellini clienti (foglio SC_CLIENTI: intestazioni GMV alla riga 10, clienti dalla riga 11). */
function leggiClienti(righe) {
  const out = [];
  if (!righe) return out;
  // la riga delle intestazioni è quella con "Colonna1 | CLIENTE | CODICE" (riga 10); la riga 1 ne ripete i nomi
  let hi = -1;
  righe.forEach((r, i) => { if (r && testo(r[1]).toUpperCase() === "CLIENTE" && testo(r[2]).toUpperCase() === "CODICE") hi = i; });
  if (hi < 0) return out;
  const head = righe[hi];
  for (let i = hi + 1; i < righe.length; i++) {
    const r = righe[i] || [];
    const nome = testo(r[1]);
    if (!nome || nome === "-") continue;
    const sconti = {};
    for (let c = 4; c < head.length; c++) {
      const g = testo(head[c]);
      if (!g || /^GMA\d+$/i.test(g)) continue;
      const v = numero(r[c]);
      if (v) sconti[g] = v;
    }
    const att = r[3] === undefined || r[3] === "" ? "" : numero(r[3]);
    out.push({ id: slug(nome), nome, codice: testo(r[2]), att, sconti });
  }
  return out;
}

export const slug = (s) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "").toLowerCase().slice(0, 80) || "x";

/* Legge tutto il file. Restituisce un riepilogo da mostrare prima di salvare. */
export async function analizzaFile(buf) {
  const wb = await leggiExcel(buf);
  const capList = leggiCapitoli(await wb.foglio("ELENCHI"));
  const capitoli = [];
  const articoli = [];
  const conQuantita = [];
  for (const nome of wb.nomi) {
    const righe = await wb.foglio(nome);
    if (!eFoglioCatalogo(righe)) continue;
    if (/^AGGIORNAMENTO/i.test(nome)) continue;
    const { articoli: art, lavoro } = leggiCatalogo(nome, righe);
    if (!art.length) continue;
    const def = capList.find((c) => c.nome === nome);
    capitoli.push({ nome, ricarico: def ? def.ricarico : 0.28, ricaricoDaFile: !!def, articoli: art.length, blocco0: lavoro.length });
    articoli.push(...art);
    // quantità: se il foglio ha righe di lavoro nel blocco 0 con quantità, l'offerta è quella; altrimenti le quantità sparse nel foglio
    const lavQ = [...lavoro, ...art.filter((a) => a._blocco0)].filter((a) => a._qta).sort((a, b) => a.ordine - b.ordine);
    conQuantita.push(...(lavQ.length ? lavQ : art.filter((a) => a._qta)));
  }
  // ordine dei capitoli: quello della tabella ELENCHI, poi gli altri
  capitoli.sort((a, b) => {
    const ia = capList.findIndex((c) => c.nome === a.nome), ib = capList.findIndex((c) => c.nome === b.nome);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  });
  capitoli.forEach((c, i) => (c.ordine = i));
  const gmv = leggiGmv(await wb.foglio("TIPO OFFERTA"));
  const noti = new Set(gmv.map((g) => g.codice));
  for (const a of articoli) if (a.gmv && !noti.has(a.gmv)) { noti.add(a.gmv); gmv.push({ codice: a.gmv, fornitore: "" }); }
  const clienti = leggiClienti(await wb.foglio("SC_CLIENTI"));
  return { capitoli, articoli, gmv, clienti, conQuantita };
}

/* Aggiornamento listino da Oracle: testo incollato dalla griglia di Oracle (con intestazioni)
   oppure foglio "AGGIORNAMENTO LISTINO". Servono le colonne "Codice Articolo Fornitore" e "Prezzo Unitario";
   "GMA" (se c'è) aggiorna anche il GMV. */
export function leggiAggiornamentoListino(righe) {
  const norm = (s) => testo(s).toLowerCase().replace(/\s+/g, " ");
  const hi = righe.findIndex((r) => r && r.some((c) => norm(c) === "codice articolo fornitore"));
  if (hi < 0) throw new Error('Non trovo la colonna "Codice Articolo Fornitore". Copia anche la riga delle intestazioni.');
  const h = righe[hi].map(norm);
  const iCod = h.indexOf("codice articolo fornitore");
  const iPrezzo = h.indexOf("prezzo unitario");
  const iGma = h.indexOf("gma");
  if (iPrezzo < 0) throw new Error('Non trovo la colonna "Prezzo Unitario".');
  const out = new Map();
  for (let i = hi + 1; i < righe.length; i++) {
    const r = righe[i] || [];
    const cod = testo(r[iCod]);
    if (!cod) continue;
    const prezzo = numero(r[iPrezzo]);
    if (!prezzo) continue;
    out.set(cod, { prezzo, gmv: iGma >= 0 ? gmvPulito(r[iGma]) : "" });
  }
  return out;
}

export function righeDaTestoIncollato(testoIncollato) {
  return testoIncollato.replace(/\r/g, "").split("\n").filter((l) => l.trim()).map((l) => l.split("\t"));
}

export { leggiExcel };
