// Lettura delle offerte dei fornitori (PDF, Excel, CSV): articoli con quantità e prezzo, divisi in sezioni.
// Non c'è intelligenza artificiale né OCR: si legge il testo scritto nel PDF con la sua posizione e si
// riconosce la tabella con delle regole. Ogni numero o è quello del PDF o non viene letto.
import { num } from "./calc.js";
import { norm, trovaIntestazione } from "./biblioteca.js";

const RE_SOLDI = /^(?:(?:€|EUR)\s*)?-?\d{1,3}(?:[.\s]\d{3})*,(?:\d{2}|\d{4,5})(?:\s*(?:€|EUR))?$|^(?:(?:€|EUR)\s*)?-?\d+,(?:\d{2}|\d{4,5})(?:\s*(?:€|EUR))?$/;
const RE_QTA = /^\d+(?:,\d+)?$/;
const soldi = (t) => num(String(t).replace(/€|EUR/g, "").replace(/\s/g, ""));
const mediana = (a) => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[Math.floor(s.length / 2)]; };
const moda = (a, arr = 1) => { const m = new Map(); for (const v of a) { const k = Math.round(v / arr) * arr; m.set(k, (m.get(k) || 0) + 1); } return [...m.entries()].sort((x, y) => y[1] - x[1])[0]?.[0]; };
const RE_TOTALE_INIZIO = /^\s*(sub\s*-?\s*)?total[ei]\b/i;
// anche "Nome sezione Totale: EUR 1.234,56" (la scritta totale subito prima dell'importo, in fondo alla riga)
const RE_TOTALE_FINE = /\btotal[ei]\s*:?\s*(?:€|EUR)?\s*-?\d{1,3}(?:\.\d{3})*,\d{2}\s*(?:€|EUR)?\s*$/i;
const RE_TOTALE = { test: (t) => RE_TOTALE_INIZIO.test(t) || RE_TOTALE_FINE.test(t) };
const RE_INTEST = /^(codice|cod\.?|articolo|prodotto|descrizione|u\.?\s?m\.?|q\.?\s?t[aà]['.]?|quantit[aà]|prezzo|unitario|totale|importo|vendita|in|euro|€|pos\.?|sconto|netto|listino|n\.|riga|voce|capitolato|azienda|marca|gdp|%|sc\.?)$/i;

/* Prestazioni / servizi: in Oracle diventano attività tecnica (PITATT). */
const RE_PRESTAZIONE = /^(prestazion|servizi|attivit|manodopera|giornat|ore\b|engineering|commissioning|messa in servizio|programmazione|assistenza|collaudo|oneri|project management)/i;
const RE_PRESTAZIONE_COD = /\b(prestazion\w*|servizi\w*|attivit\w*|manodopera|giornat\w*|engineering|commissioning|project\s+management|oneri)\b/i;
export const sembraPrestazione = (codice, descrizione) => {
  const c = String(codice || "").replace(/^[^A-Za-zÀ-ÿ]+/, "").trim(), d = String(descrizione || "").replace(/^[^A-Za-zÀ-ÿ]+/, "").trim();
  return RE_PRESTAZIONE.test(c) || RE_PRESTAZIONE_COD.test(c) || ((!c || c.length < 3) && RE_PRESTAZIONE.test(d)) || /^(commissioning|attivit[aà] di gestione commessa|oneri di configurazione)/i.test(d);
};

/* Righe di testo di una pagina: gli elementi sulla stessa altezza; i pezzi vicini si uniscono in celle. */
function lineeDi(items, p) {
  const ord = items.filter((i) => i.str.trim()).slice().sort((a, b) => b.y - a.y || a.x - b.x);
  const linee = [];
  for (const it of ord) {
    const tol = Math.max(0.8, (it.h || 8) * 0.3);
    let l = linee.find((x) => Math.abs(x.y - it.y) <= tol);
    if (!l) { l = { p, y: it.y, h: 0, items: [] }; linee.push(l); }
    l.items.push(it); l.h = Math.max(l.h, it.h || 0);
  }
  for (const l of linee) {
    l.items.sort((a, b) => a.x - b.x);
    const celle = [];
    for (const it of l.items) {
      const c = celle[celle.length - 1];
      const gap = c ? it.x - c.x2 : Infinity;
      const soloEuro = (/^\s*€\s*$/.test(it.str) && c && /\d\s*$/.test(c.t)) || (c && /^(€|EUR)$/.test(c.t.trim()) && /^\s*-?\d/.test(it.str) && gap < 40);
      const largChar = it.str.length ? it.w / it.str.length : (it.h || 8) * 0.5;
      if (c && (soloEuro || gap < Math.max(1, Math.min((it.h || 8) * 0.45, largChar * 0.8)))) { c.t += (gap > 0.5 ? " " : "") + it.str; c.x2 = Math.max(c.x2, it.x + it.w); c.y = Math.min(c.y, it.y); }
      else celle.push({ t: it.str, x: it.x, x2: it.x + it.w, y: it.y, h: it.h || 0, f: it.f || "" });
    }
    for (const c of celle) c.t = c.t.replace(/\s+/g, " ").trim();
    l.celle = celle.filter((c) => c.t);
    l.testo = l.celle.map((c) => c.t).join(" ");
  }
  return linee.filter((l) => l.celle.length);
}

/* Alcuni gestionali vanno a capo dentro i numeri ("1.812,00" e sotto "000") e scrivono il prezzo
   qualche punto più in alto della riga dell'articolo. Qui si rimettono insieme:
   1) un pezzo di sole cifre subito sotto un numero con la virgola, nella stessa colonna, si attacca al numero;
   2) un importo da solo su una riga vicina (meno di una riga di distanza) a una riga con quantità e importi
      entra in quella riga, se nella sua colonna non c'è già niente. */
function ricuciNumeri(linee) {
  const vicino = (a, b) => a.x < b.x2 + 2 && b.x < a.x2 + 2;
  for (let i = 0; i < linee.length; i++) for (const c of linee[i].celle) {
    if (!/,\d+$/.test(c.t)) continue;
    for (let j = i + 1; j < linee.length && linee[i].y - linee[j].y <= (c.h || 6) * 1.3; j++) {
      const f = linee[j].celle.find((x) => /^\d{1,4}$/.test(x.t) && vicino(c, x) && x.x >= c.x - 1);
      if (f) { c.t += f.t; c.x2 = Math.max(c.x2, f.x2); linee[j].celle = linee[j].celle.filter((x) => x !== f); break; }
    }
  }
  for (const l of linee) l.testo = l.celle.map((c) => c.t).join(" ");
  const conQta = (l) => { const k = codaNumerica(l); return !!(k && k.qta); };
  for (let i = 0; i < linee.length; i++) {
    const L = linee[i];
    if (!conQta(L)) continue;
    const h = L.h || 6;
    for (const j of [i - 1, i + 1, i - 2, i + 2]) {
      const M = linee[j];
      if (!M || M === L || Math.abs(M.y - L.y) > h * 0.95 || conQta(M)) continue;
      for (const c of M.celle.filter((x) => RE_SOLDI.test(x.t))) {
        if (L.celle.some((x) => vicino(x, c))) continue;
        if (c.x < Math.min(...L.celle.filter((x) => RE_SOLDI.test(x.t) || RE_QTA.test(x.t)).map((x) => x.x))) continue;
        L.celle.push(c); M.celle = M.celle.filter((x) => x !== c);
      }
      L.celle.sort((a, b) => a.x - b.x);
    }
  }
  for (const l of linee) l.testo = l.celle.map((c) => c.t).join(" ");
  return linee.filter((l) => l.celle.length);
}

/* Parte numerica a destra di una riga: unità di misura, quantità, prezzo unitario, totale. */
function codaNumerica(l) {
  const c = l.celle;
  // le celle "n. 1" o "1 n." si dividono in pezzi
  const pezzi = [];
  for (const x of c) {
    if (RE_SOLDI.test(x.t)) { pezzi.push({ ...x, tipo: "soldi" }); continue; }
    const parti = x.t.split(" ");
    if (parti.length > 1 && parti.length <= 3 && parti.every((t) => RE_QTA.test(t) || /^[a-zA-Z°.]{1,4}\.?$/.test(t))) {
      let xx = x.x; const passo = (x.x2 - x.x) / x.t.length;
      for (const t of parti) { pezzi.push({ t, x: xx, x2: xx + t.length * passo, y: x.y, tipo: RE_QTA.test(t) ? "num" : "testo" }); xx += (t.length + 1) * passo; }
    } else pezzi.push({ ...x, tipo: RE_QTA.test(x.t) ? "num" : "testo" });
  }
  // quantità nella prima colonna a sinistra scritta con i decimali ("1,00 C-GE300WR descrizione … EUR 1.812,00"): non è un importo
  if (pezzi.length > 2 && pezzi[0].tipo === "soldi" && pezzi[1].tipo === "testo" && pezzi.slice(2).some((x) => x.tipo === "soldi") && /^\d+,\d+$/.test(pezzi[0].t)) pezzi[0] = { ...pezzi[0], tipo: "num" };
  const soldiIdx = pezzi.map((x, i) => (x.tipo === "soldi" ? i : -1)).filter((i) => i >= 0);
  if (!soldiIdx.length) return null;
  // dopo il primo importo ci possono essere altre colonne (sconti %, codici sconto, marca…): contano solo gli importi
  const primo = soldiIdx[0];
  const valori = pezzi.slice(primo).filter((x) => x.tipo === "soldi");
  let i = primo - 1, qta = null, um = "";
  if (i >= 0 && pezzi[i].tipo === "num") { qta = pezzi[i]; i--; }
  if (qta && i >= 0 && pezzi[i].tipo === "testo" && pezzi[i].t.length <= 5 && !/\d/.test(pezzi[i].t)) { um = pezzi[i]; i--; }
  // quantità nella prima colonna a sinistra, prima del codice
  if (!qta && primo > 1 && pezzi[0].tipo === "num" && pezzi.slice(1, primo).some((x) => x.tipo === "testo"))
    return { valori, qta: pezzi[0], um: "", inizio: valori[0].x, resto: pezzi.slice(1, primo), qtaSx: true };
  return { valori, qta, um, inizio: (um || qta || valori[0]).x, resto: pezzi.slice(0, i + 1) };
}

/* Offerta in PDF: { sezioni: [{ nome, totalePdf, righe: [{ codice, descrizione, um, qta, unitario, totale }] }], totaleGenerale, note }
   Restituisce null se nel PDF non c'è una tabella riconoscibile. */
export function leggiOffertaPdf(pagine) {
  const perPagina = pagine.map((items, p) => ricuciNumeri(lineeDi(items, p)));
  // 1. righe articolo (ancore): a destra quantità e prezzi, con quantità × unitario = totale quando ci sono entrambi
  const ancore = [];
  for (const linee of perPagina) for (const l of linee) {
    if (RE_TOTALE.test(l.testo)) continue;
    const k = codaNumerica(l);
    if (!k || !k.qta) continue;
    // coppie unitario → totale (q.tà × unitario = totale): la più bassa è il netto, la più alta il listino
    const q = num(k.qta.t), v = k.valori.map((x) => soldi(x.t));
    const torna = (u, t) => Math.abs(q * u - t) <= Math.max(0.05, Math.abs(t) * 0.002);
    const coppie = [];
    for (let i = 0; i + 1 < v.length; i++) if (torna(v[i], v[i + 1])) coppie.push([v[i], v[i + 1]]);
    // tra prezzo e totale ci può essere un'altra colonna di importi (es. sconto 0,00)
    if (!coppie.length) for (let i = 0; i < v.length; i++) for (let j = i + 2; j < v.length; j++) if (v[i] && torna(v[i], v[j])) coppie.push([v[i], v[j]]);
    if (coppie.length) {
      coppie.sort((a, b) => a[0] - b[0]);
      k.unitario = coppie[0][0]; k.totale = coppie[0][1]; k.listino = coppie[coppie.length - 1][0];
    } else {
      // q.tà × unitario non torna: può essere un errore del fornitore, la riga si tiene ma si segnala
      k.unitario = v[0]; k.totale = v.length > 1 ? v[v.length - 1] : v[0] * q; k.listino = v[0];
      if (v.length > 1) l.nonTorna = true;
      // "prezzo unitario di listino" + "totale netto" (sconto già applicato): l'unitario netto è totale / q.tà
      if (v.length === 2 && q > 0 && v[1] > 0 && v[1] < q * v[0] - 0.01) { k.unitario = v[1] / q; k.totale = v[1]; k.listino = v[0]; l.nonTorna = false; }
    }
    l.k = k; ancore.push(l);
  }
  if (ancore.length < 2) return null;
  // 2. colonne: eventuale numero di riga a sinistra, inizio della parte numerica, codice e descrizione
  const primoCella = ancore.map((l) => l.k.resto[0]).filter(Boolean);
  const conPos = primoCella.filter((c) => /^\d{1,4}(\s|$)/.test(c.t));
  const xPos = conPos.length >= primoCella.length * 0.6 ? moda(conPos.map((c) => c.x), 2) : null;
  const togliPos = (celle) => {
    if (xPos === null || !celle.length || Math.abs(celle[0].x - xPos) > 3 || !/^\d{1,4}(\s|$)/.test(celle[0].t)) return { pos: null, celle };
    const m = celle[0].t.match(/^(\d{1,4})\s*(.*)$/);
    const resto = m[2] ? [{ ...celle[0], t: m[2], x: celle[0].x + (celle[0].x2 - celle[0].x) * (m[1].length + 1) / celle[0].t.length }, ...celle.slice(1)] : celle.slice(1);
    return { pos: m[1], celle: resto };
  };
  for (const l of ancore) l.k.resto = togliPos(l.k.resto).celle;
  const limDx = moda(ancore.map((l) => l.k.inizio), 2) - 1;
  const xCod0 = moda(ancore.map((l) => l.k.resto[0]?.x).filter((x) => x !== undefined), 2);
  let xDesc = moda(ancore.flatMap((l) => l.k.resto.filter((c) => c.x > xCod0 + 8).map((c) => c.x)), 2) ?? Infinity;
  // descrizioni centrate nella colonna (iniziano a x diverse): la colonna comincia subito dopo il codice più lungo
  {
    const due = ancore.filter((l) => l.k.resto.length >= 2);
    if (due.length >= ancore.length * 0.6) {
      const codX2 = Math.max(...due.map((l) => Math.max(...l.k.resto.slice(0, -1).map((c) => c.x2))));
      const descX = Math.min(...due.map((l) => l.k.resto[l.k.resto.length - 1].x));
      if (codX2 + 2 < xDesc && descX >= codX2 + 2 && xDesc - descX > 6) xDesc = codX2 + 2;
    }
  }
  // colonna del codice: a sinistra della descrizione ci possono essere più colonne (es. marca, voce di computo):
  // il codice è la più vicina alla descrizione tra quelle presenti in almeno metà degli articoli
  const cluster = new Map();
  for (const l of ancore) for (const x of new Set(l.k.resto.filter((c) => c.x < xDesc - 2).map((c) => Math.round(c.x / 4) * 4))) cluster.set(x, (cluster.get(x) || 0) + 1);
  const xCod = Math.max(...[...cluster.entries()].filter(([, n]) => n >= ancore.length * 0.5).map(([x]) => x), -Infinity);
  const inCodice = (c) => c.x < xDesc - 2 && (xCod === -Infinity || c.x >= xCod - 4);
  // altezza del testo delle descrizioni (i prezzi possono essere scritti più piccoli)
  const hb = moda(ancore.flatMap((l) => l.k.resto.filter((c) => c.x >= xDesc - 2).map((c) => c.h)), 0.5) || moda(ancore.flatMap((l) => l.items.map((i) => i.h)), 0.5) || 8;
  const pagAncore = new Set(ancore.map((l) => l.p));
  // 3. righe ripetute su più pagine (intestazioni e piè di pagina): si tolgono
  const firma = (l) => norm(l.testo).replace(/\d+/g, "#") + "|" + Math.round(l.y / 3);
  const conta = new Map();
  for (const p of pagAncore) for (const s of new Set(perPagina[p].map(firma))) conta.set(s, (conta.get(s) || 0) + 1);
  const ripetute = new Set([...conta.entries()].filter(([, n]) => n >= Math.max(2, Math.ceil(pagAncore.size / 2)) && pagAncore.size > 1).map(([s]) => s));
  // 4. sequenza delle righe utili: dalla prima pagina con articoli, sotto l'intestazione della tabella
  const seq = [];
  let fine = false;
  for (const linee of perPagina) {
    if (fine) break;
    if (!pagAncore.has(linee[0]?.p)) continue;
    let da = 0;
    const primaAncora = linee.findIndex((l) => l.k);
    const iInt = linee.findIndex((l, i) => i < primaAncora && /descrizione|description/i.test(l.testo) && !l.k);
    if (iInt >= 0) { da = iInt + 1; while (da < primaAncora && linee[da].celle.every((c) => c.t.split(" ").every((t) => RE_INTEST.test(t)))) da++; }
    for (const l of linee.slice(da)) {
      if (!l.k && ripetute.has(firma(l))) continue;
      // intestazione della tabella ripetuta sotto ogni sezione ("Q.tà Articolo Descrizione Prezzo unitario …")
      if (!l.k) { const tk = l.celle.flatMap((c) => c.t.split(" ")).filter(Boolean); if (tk.length >= 2 && tk.every((t) => RE_INTEST.test(t))) continue; }
      seq.push(l);
      if (/totale\s+generale/i.test(l.testo)) { fine = true; break; }
    }
  }
  // carattere del testo delle descrizioni, pagina per pagina (i nomi dei caratteri cambiano da una pagina all'altra):
  // una riga scritta con un altro carattere (es. grassetto) è un titolo
  const fontDesc = (l) => { const c = (l.k ? l.k.resto : l.celle).filter((x) => x.x >= xDesc - 2 && x.x < limDx); return c.length ? c[0].f : ""; };
  const fontCorpo = new Map();
  for (const p of new Set(seq.map((l) => l.p))) {
    const m = new Map();
    for (const l of seq) if (l.p === p) for (const c of (l.k ? l.k.resto : l.celle)) if (c.f && c.x >= xDesc - 2 && c.x < limDx) m.set(c.f, (m.get(c.f) || 0) + c.t.length);
    fontCorpo.set(p, [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "");
  }
  const maiuscolo = (t) => { const lettere = t.replace(/[^A-Za-zÀ-ÿ]/g, ""); return lettere.length >= 3 && lettere === lettere.toUpperCase(); };
  // tipo di ogni riga e divisione in parti codice / descrizione
  for (const l of seq) {
    let sx = l.k ? l.k.resto : l.celle.filter((c) => c.x < limDx);
    if (!l.k) { const tp = togliPos(sx); sx = tp.celle; l.pos = tp.pos; }
    l.cod = sx.filter(inCodice).map((c) => c.t).join(" ").trim();
    l.desc = sx.filter((c) => c.x >= xDesc - 2).map((c) => c.t).join(" ").trim();
    const soldiTot = l.celle.filter((c) => RE_SOLDI.test(c.t) || /\d,\d{2}\s*€?$/.test(c.t));
    if (l.k) l.tipo = "articolo";
    else if (RE_TOTALE.test(l.testo) && soldiTot.length && /\biva\b|ivato|i\.v\.a/i.test(l.testo)) l.tipo = "ignora"; // totale IVA inclusa
    else if (RE_TOTALE.test(l.testo) && soldiTot.length) {
      // il primo importo dopo la scritta "totale" (sulla stessa riga ci può essere anche il totale IVA inclusa)
      l.tipo = "totale";
      const m = l.testo.replace(/^.*?total[ei]/i, "").match(/-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}/);
      l.valore = m ? soldi(m[0]) : soldi(soldiTot[soldiTot.length - 1].t);
    }
    else if (l.h >= hb * 1.2 && !soldiTot.length && (l.desc || (!l.cod && sx.length))) l.tipo = "sezione";
    else if (l.pos && !soldiTot.length && (l.cod || l.desc)) l.tipo = "sezione"; // riga numerata senza prezzo: titolo di gruppo
    else l.tipo = "testo";
    if (l.tipo === "testo" && !l.cod && l.desc && fontCorpo.get(l.p) && fontDesc(l) && fontDesc(l) !== fontCorpo.get(l.p)) {
      l.titolo = true;
      // titolo tutto maiuscolo in un altro carattere: è una sezione (es. "ANTINTRUSIONE", "CONTROLLO ACCESSI")
      if (!soldiTot.length && maiuscolo(l.desc) && l.desc.length <= 80) { l.tipo = "sezione"; l.sezDaCarattere = true; }
    }
    if (l.tipo === "sezione" && !l.sezDaCarattere && !l.pos) l.sezDaAltezza = true;
  }
  // se le sezioni sono scritte più grandi, i titoli maiuscoli dello stesso corpo del testo sono solo sottotitoli
  if (seq.some((l) => l.sezDaAltezza)) for (const l of seq) if (l.sezDaCarattere) l.tipo = "testo";
  // passo tra le righe di una descrizione
  const passi = [];
  for (let i = 1; i < seq.length; i++) if (seq[i].p === seq[i - 1].p) { const d = seq[i - 1].y - seq[i].y; if (d > 0.5 && d < hb * 3) passi.push(d); }
  const passo = mediana(passi) || hb * 1.2;
  // 5. ogni riga di testo della colonna descrizione va all'articolo più vicino (nello stesso blocco),
  //    poi si tengono solo le righe centrate sull'articolo (la descrizione è centrata sulla riga dei prezzi)
  const blocchi = []; let blocco = [];
  for (const l of seq) {
    if (l.tipo === "sezione" || l.tipo === "totale" || (blocco.length && blocco[blocco.length - 1].p !== l.p)) { if (blocco.length) blocchi.push(blocco); blocco = []; }
    if (l.tipo === "articolo" || l.tipo === "testo") blocco.push(l);
  }
  if (blocco.length) blocchi.push(blocco);
  const scarto = 0; // la descrizione è centrata sulla riga dei prezzi
  // ogni descrizione è un gruppo di righe consecutive centrato sulla riga dei prezzi del suo articolo:
  // si cerca la divisione migliore delle righe tra gli articoli (programmazione dinamica); le righe che non
  // stanno bene con nessuno (titoli sopra un articolo) restano libere
  const P_LIBERA = 0.3, P_VUOTO = 1;
  for (const b of blocchi) {
    const art = b.filter((l) => l.tipo === "articolo");
    if (!art.length) continue;
    // codice scritto un po' più in alto o più in basso della riga dei prezzi (riga con solo il codice)
    for (const a of art) a.codRighe = a.cod ? [{ y: a.y, t: a.cod }] : [];
    for (const l of b) {
      if (l.tipo !== "testo" || !l.cod || l.usata || /^note?\b/i.test(l.cod)) continue;
      // riga con solo il codice: fino a un'altezza di testo dall'articolo; con anche la descrizione, più vicina
      const lim = l.desc ? hb * 0.8 : hb;
      let best = null;
      for (const a of art) if (Math.abs(a.y - l.y) <= lim && (!best || Math.abs(a.y - l.y) < Math.abs(best.y - l.y))) best = a;
      if (!best) continue;
      best.codRighe.push({ y: l.y, t: l.cod });
      if (l.desc) l.cod = ""; else l.usata = true; // la sua descrizione resta una riga di descrizione
    }
    for (const a of art) a.cod = a.codRighe.sort((x, y) => y.y - x.y).map((x) => x.t).join(" ").trim();
    const L = [];
    for (const l of b) {
      if (l.tipo === "articolo") { if (l.desc) L.push({ y: l.y, t: l.desc, own: art.indexOf(l) }); }
      // con qualcosa nella colonna codice è una riga a sé (NOTE, Punti gestiti…) e separa le descrizioni
      else if (l.tipo === "testo" && !l.usata && (l.cod || l.desc)) L.push({ y: l.y, t: l.desc, l, own: l.cod || l.titolo ? -2 : -1 });
    }
    L.sort((x, y) => y.y - x.y);
    // NOTE: l'etichetta è centrata sul suo paragrafo; tante righe sopra quante sotto sono la nota
    L.forEach((r, i) => {
      if (!r.l || !/^note?\b/i.test(r.l.cod)) return;
      r.l.nota = true;
      let su = i, giu = i;
      while (su > 0 && L[su - 1].own === -1 && L[su - 1].y - L[su].y <= passo * 1.45) su--;
      const span = (L[su].y - r.y) + passo * 0.6;
      while (giu < L.length - 1 && L[giu + 1].own === -1 && r.y - L[giu + 1].y <= span && L[giu].y - L[giu + 1].y <= passo * 1.45) giu++;
      for (let k = su; k <= giu; k++) { L[k].own = -2; if (L[k].l) L[k].l.nota = true; }
    });
    const n = L.length, m = art.length;
    const rotto = L.map((r, i) => (i > 0 && L[i - 1].y - r.y > passo * 1.45 ? 1 : 0));
    // descrizioni che partono dalla riga dei prezzi e continuano sotto (allineate in alto): ogni gruppo di righe
    // vicine (separato dal successivo da uno spazio più grande) contiene un solo articolo, che ne è la prima riga
    {
      const salti = L.slice(1).map((r, i) => L[i].y - r.y).filter((d) => d > 0.5);
      const soglia = salti.length ? Math.min(...salti) * 1.4 : Infinity;
      const gruppi = []; let g = [];
      L.forEach((r, i) => { if (i && L[i - 1].y - r.y > soglia) { gruppi.push(g); g = []; } g.push(r); });
      if (g.length) gruppi.push(g);
      const conArt = gruppi.filter((x) => x.some((r) => r.own >= 0));
      const buoni = conArt.filter((x) => x.filter((r) => r.own >= 0).length === 1 && x[0].own >= 0 && x.every((r) => r.own !== -2));
      if (conArt.length >= 2 && buoni.length === conArt.length && conArt.length === m && gruppi.some((x) => x.length > 1)) {
        for (const x of buoni) { const a = art[x[0].own]; a.righe = x; for (const r of x) if (r.l) r.l.usata = true; }
        continue;
      }
    }
    const costo = (j, s, e) => { // righe s..e-1 all'articolo j
      for (let i = s; i < e; i++) { if (L[i].own === -2 || (L[i].own >= 0 && L[i].own !== j)) return Infinity; if (i > s && rotto[i]) return Infinity; }
      const proprio = L.findIndex((r) => r.own === j);
      if (proprio >= 0 && (proprio < s || proprio >= e)) return Infinity;
      if (e === s) return P_VUOTO;
      const centro = (L[s].y + L[e - 1].y) / 2, d = (centro - art[j].y) / passo;
      return d * d;
    };
    const INF = Infinity;
    let f = new Array(n + 1).fill(INF); f[0] = 0;
    const sceltaS = [], sceltaDa = [];
    for (let j = 0; j < m; j++) {
      const g = f.slice(), gda = f.map((_, i) => i);
      for (let i = 1; i <= n; i++) if (L[i - 1].own < 0 && g[i - 1] + P_LIBERA < g[i]) { g[i] = g[i - 1] + P_LIBERA; gda[i] = gda[i - 1]; }
      const nf = new Array(n + 1).fill(INF), sc = new Array(n + 1).fill(-1);
      for (let s = 0; s <= n; s++) {
        if (g[s] === INF) continue;
        for (let e = s; e <= n; e++) {
          if (e - 1 > s && rotto[e - 1]) break;
          const c = g[s] + costo(j, s, e);
          if (c < nf[e]) { nf[e] = c; sc[e] = s; }
          if (e < n && (L[e].own === -2 || (L[e].own >= 0 && L[e].own !== j))) break;
        }
      }
      sceltaS.push(sc); sceltaDa.push(gda); f = nf;
    }
    // righe libere in fondo
    let fine = -1, best = INF;
    for (let e = n; e >= 0; e--) { const c = f[e] + (n - e) * P_LIBERA; if (c < best) { best = c; fine = e; } if (e > 0 && L[e - 1].own >= 0) break; }
    if (fine < 0) { for (const a of art) a.righe = a.desc ? [{ y: a.y, t: a.desc }] : []; continue; }
    for (let j = m - 1; j >= 0; j--) {
      const s = sceltaS[j][fine];
      art[j].righe = L.slice(s, fine);
      for (const r of art[j].righe) if (r.l) r.l.usata = true;
      fine = sceltaDa[j][s];
    }
  }
  // 6. sequenza finale: sezioni, articoli, totali; le righe libere (titoli, "Punti gestiti") vanno davanti
  //    alla descrizione dell'articolo che segue, le NOTE si tengono a parte
  const sezioni = [];
  let sez = null, totaleGenerale = null, sezGenerale = null;
  const note = [];
  let par = [], nota = [];
  const chiudiNota = () => { const t = nota.join(" ").trim(); if (t.length >= 25) note.push(t); nota = []; };
  const nuovaSez = (nome) => { if (sez && !sez.righe.length && !sez.totalePdf) { sez.nome = nome; return; } sez = { nome, totalePdf: null, righe: [] }; sezioni.push(sez); };
  let precSezione = null;
  for (const l of seq) {
    if (!(l.tipo === "testo" && l.nota)) chiudiNota();
    const eraSezione = precSezione; precSezione = null;
    if (l.tipo === "sezione") {
      par = [];
      const nome = [l.cod, l.desc].filter(Boolean).join(" ") || l.celle.map((c) => c.t).filter((t) => !RE_SOLDI.test(t)).join(" ");
      // titolo su due righe: si unisce
      const unita = eraSezione && sez && !sez.righe.length && !sez.totalePdf && eraSezione.p === l.p && eraSezione.y - l.y < l.h * 1.6 && Math.abs((eraSezione.celle[0]?.x ?? 0) - (l.celle[0]?.x ?? 0)) < 3;
      if (unita) sez.nome += " " + nome; else nuovaSez(nome);
      precSezione = l;
    }
    else if (l.tipo === "totale") {
      par = [];
      // "Nome sezione Totale: EUR …": il nome prima della scritta totale è quello della sezione che si chiude
      const pref = l.testo.replace(/\btotal[ei]\b.*$/i, "").replace(/[:\s]+$/, "").trim();
      if (/generale/i.test(l.testo) || /^\s*total[ei]\s*$/i.test(l.desc || "") || (!pref && RE_TOTALE_INIZIO.test(l.testo) && sezioni.length && sez?.totalePdf !== null)) { totaleGenerale = l.valore; sezGenerale = sezioni.length; }
      else if (sez) {
        sez.totalePdf = l.valore;
        if (pref && pref.length <= 120 && !RE_SOLDI.test(pref)) { if (sez.autoNome || sez.nome === "Offerta") sez.nome = pref; sez = null; }
      }
    } else if (l.tipo === "testo" && !l.usata) {
      if (l.nota) { if (l.desc) nota.push(l.desc); continue; }
      const t = [l.cod && !/^\.$/.test(l.cod) ? l.cod + (l.desc ? ":" : "") : "", l.desc].filter(Boolean).join(" ");
      if (t) par.push(t);
    } else if (l.tipo === "articolo") {
      if (!sez) nuovaSez("Offerta");
      const k = l.k;
      const testa = par.join(" ").trim(); par = [];
      const corpo = l.righe.map((r) => r.t).join(" ");
      const descr = (testa && corpo ? testa + (/-$/.test(testa) ? "" : " – ") + corpo : testa || corpo).replace(/\s+/g, " ").trim();
      sez.righe.push({
        codice: l.cod, descrizione: descr, um: k.um ? k.um.t : "n.", qta: num(k.qta.t),
        unitario: k.unitario, totale: k.totale, listino: k.listino, nonTorna: !!l.nonTorna,
      });
    }
  }
  chiudiNota();
  // un "totale" che non torna con la sua sezione ma torna con tutta l'offerta è il totale generale
  const somma = (rr) => rr.reduce((a, r) => a + r.totale, 0);
  const tutte = sezioni.flatMap((x) => x.righe), tol = (t) => Math.max(1, Math.abs(t) * 0.0005);
  for (const x of sezioni) if (x.totalePdf !== null && Math.abs(somma(x.righe) - x.totalePdf) > tol(x.totalePdf) && Math.abs(somma(tutte) - x.totalePdf) <= tol(x.totalePdf)) {
    if (totaleGenerale === null) totaleGenerale = x.totalePdf;
    x.totalePdf = null;
  }
  // il totale generale vale per le sezioni che lo precedono (dopo possono esserci le opzioni)
  const prima = sezGenerale === null ? sezioni : sezioni.slice(0, sezGenerale);
  for (const x of sezioni) x.nelTotale = prima.includes(x);
  const piene = sezioni.filter((s) => s.righe.length);
  return piene.length ? { sezioni: piene, totaleGenerale, note } : null;
}

/* Offerta da tabella (Excel, CSV, o PDF letto come tabella): servono le colonne scelte.
   mappa = { codice, descrizione, um, qta, unitario, totale, sezione } (nomi delle intestazioni normalizzati).
   Una riga senza prezzo e senza codice ma con un testo apre una nuova sezione. */
export const CAMPI_OFF = [
  { k: "codice", label: "Codice", obbl: true, cerca: ["codice prodotto", "codice", "articolo", "cod", "item", "part number", "modello"] },
  { k: "descrizione", label: "Descrizione", cerca: ["descrizione", "description", "desc"] },
  { k: "um", label: "U.M.", cerca: ["u.m.", "um", "unita", "unita di misura"] },
  { k: "qta", label: "Quantità", obbl: true, cerca: ["q.ta", "qta", "quantita", "qty", "q.ty", "quantity"] },
  { k: "unitario", label: "Prezzo unitario", obbl: true, cerca: ["netto unitario", "prezzo netto", "unitario", "prezzo unitario", "prezzo", "listino unitario", "unit price"] },
  { k: "sconto", label: "Sconto %", cerca: ["sconto", "sc.", "discount"] },
  { k: "totale", label: "Totale (controllo)", cerca: ["totale", "importo", "total"] },
  { k: "sezione", label: "Sezione", cerca: ["sez", "sezione", "capitolo", "section"] },
];
export function indovinaMappaOff(celle) {
  const m = {};
  const usate = new Set();
  for (const c of CAMPI_OFF) {
    // prima le corrispondenze esatte, poi le parziali
    let h = null;
    for (const w of c.cerca) { h = celle.find((x) => x && !usate.has(x) && x === w); if (h) break; }
    if (!h) for (const w of c.cerca) { h = celle.find((x) => x && !usate.has(x) && x.includes(w)); if (h) break; }
    if (h) { m[c.k] = h; usate.add(h); }
  }
  return m;
}
export function leggiOffertaTabella(fogli, mappa) {
  const sezioni = [];
  let sez = null;
  for (const f of fogli) {
    const int = trovaIntestazione(f.righe);
    if (!int) continue;
    const col = (k) => (mappa[k] ? int.celle.indexOf(mappa[k]) : -1);
    const C = Object.fromEntries(CAMPI_OFF.map((c) => [c.k, col(c.k)]));
    if (C.codice < 0 || C.unitario < 0) continue;
    for (const r of f.righe.slice(int.indice + 1)) {
      const v = (k) => (C[k] >= 0 ? r[C[k]] : "");
      const cod = String(v("codice") ?? "").trim(), des = String(v("descrizione") ?? "").replace(/\s+/g, " ").trim();
      const pr = v("unitario"), prezzo = typeof pr === "number" ? pr : soldi(pr);
      const haPrezzo = pr !== "" && pr !== null && pr !== undefined && Number.isFinite(prezzo) && String(pr).trim() !== "";
      if (!cod && !haPrezzo) {
        const titolo = des || String(v("sezione") ?? "").trim();
        if (titolo && !RE_TOTALE.test(titolo)) { sez = { nome: titolo, totalePdf: null, righe: [] }; sezioni.push(sez); }
        else if (titolo && RE_TOTALE.test(titolo) && sez) { const t = v("totale"); if (t !== "") sez.totalePdf = typeof t === "number" ? t : soldi(t); }
        continue;
      }
      if (!haPrezzo) continue;
      if (!sez) { sez = { nome: f.nome || "Offerta", totalePdf: null, righe: [] }; sezioni.push(sez); }
      // sconto in cascata ("50+10"); non si applica se il prezzo scelto è già netto
      let sc = 0;
      if (C.sconto >= 0 && !/nett/.test(mappa.unitario || "")) {
        const sv = v("sconto");
        const termini = typeof sv === "number" ? [sv > 1 ? sv / 100 : sv] : String(sv ?? "").replace(/%/g, "").split("+").map((x) => num(x.trim()) / 100).filter((x) => x);
        sc = 1 - termini.reduce((f, x) => f * (1 - x), 1);
      }
      const qta = num(v("qta")) || 0;
      const unitario = prezzo * (1 - sc);
      const tot = v("totale");
      const totale = tot === "" || tot === null || tot === undefined ? unitario * qta : typeof tot === "number" ? tot : soldi(tot);
      sez.righe.push({ codice: cod, descrizione: des, um: String(v("um") || "n.").trim(), qta, unitario, totale, listino: prezzo,
        nonTorna: Math.abs(unitario * qta - totale) > Math.max(0.05, Math.abs(totale) * 0.002) && Math.abs(unitario - totale) > 0.005 });
    }
  }
  const piene = sezioni.filter((s) => s.righe.length);
  return piene.length ? { sezioni: piene, totaleGenerale: null, note: [] } : null;
}
