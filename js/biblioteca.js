// Biblioteca: lettura dei listini fornitore (Excel, CSV, testo incollato) in un elenco di articoli.
// Funzioni pure sulle righe già lette (array di array): si usano nell'app e nei test.
import { num } from "./calc.js";

/* Campi che si possono leggere da un listino, con le parole che di solito ci sono nell'intestazione. */
export const CAMPI = [
  { k: "codice", label: "Codice (va in Oracle)", obbl: true, chiavi: ["codice prodotto azienda", "codice ordinazione", "codice d'ordine", "codice articolo", "codice prodotto", "order code", "sen", "codici", "codice", "code", "articolo", "art."] },
  { k: "nOrdine", label: "Modello / n° ordinazione", chiavi: ["article type", "riferimento prodotto", "modello", "tipo", "type", "riferimento"] },
  { k: "descrizione", label: "Descrizione", chiavi: ["descrizione prodotto", "descrizione breve", "short description", "descrizione", "description", "desc"] },
  { k: "prezzo", label: "Prezzo di listino", obbl: true, chiavi: ["prezzo al rivenditore", "prezzo di listino", "listino prezzi", "llp", "prezzo unitario", "prezzo listino", "listino", "list price", "prezzo", "price", "value"] },
  { k: "netto", label: "Prezzo netto (se c'è)", chiavi: ["prezzo netto", "netto", "net price", "net"] },
  { k: "sconto", label: "Sconto % (se c'è)", chiavi: ["sconto", "discount"] },
  { k: "sottocapitolo", label: "Gruppo prodotti / categoria (sottocapitolo)", chiavi: ["gruppo prodotti", "categoria", "famiglia prodotto", "linea prodotto"] },
  { k: "famiglia", label: "Famiglia di sconto / gruppo", chiavi: ["famiglia di sconto", "local price group", "price group", "gruppo sconto", "famiglia", "gruppo"] },
  { k: "gma", label: "GMA (se c'è)", chiavi: ["gma", "gmv"] },
  { k: "um", label: "Unità di misura", chiavi: ["unita di misura", "unità di misura", "u.m.", "um", "unit"] },
  { k: "molt", label: "Moltiplicatore prezzo (prezzo per N pezzi)", chiavi: ["moltiplicatore prezzo", "moltiplicatore"] },
];

export const norm = (v) => String(v ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()
  .replace(/[€()\[\]]/g, " ").replace(/\s+/g, " ").trim();
const testo = (v) => (v === null || v === undefined || typeof v === "object" ? "" : String(v).trim());

const pareCodice = (h) => /\b(codic|code|sen\b|articolo|art\.)/.test(h);
const parePrezzo = (h) => /(prezzo|price|listino|llp|value)/.test(h);

/* Riga di intestazione: una delle prime 40 con un "codice" e un "prezzo". */
export function trovaIntestazione(righe) {
  for (let i = 0; i < Math.min(righe.length, 40); i++) {
    const r = (righe[i] || []).map(norm);
    const piene = r.filter(Boolean);
    if (piene.length < 2) continue;
    if (r.some(pareCodice) && r.some(parePrezzo)) return { indice: i, celle: r };
  }
  return null;
}

export const isMetel = (celle) => ["codice prodotto azienda", "prezzo al rivenditore", "moltiplicatore prezzo"].every((x) => celle.includes(x));

/* Prima ipotesi: per ogni campo la colonna con l'intestazione più somigliante. Restituisce { campo: intestazione }. */
export function indovinaMappa(celle) {
  const mappa = {}, usate = new Set();
  // tre passate: intestazione identica, poi che comincia con la parola, poi che la contiene
  const prove = [(h, ch) => h === ch, (h, ch) => ch.length >= 3 && h.startsWith(ch), (h, ch) => ch.length >= 4 && h.includes(ch)];
  for (const prova of prove) {
    for (const c of CAMPI) {
      if (mappa[c.k]) continue;
      for (const ch of c.chiavi) {
        const h = celle.find((x) => x && !usate.has(x) && prova(x, ch) && !(c.k === "prezzo" && /nett|net\b/.test(x)));
        if (h) { mappa[c.k] = h; usate.add(h); break; }
      }
    }
  }
  return mappa;
}

/* Mappa fissa per il formato standard METEL. */
export const MAPPA_METEL = {
  codice: "codice prodotto azienda", descrizione: "descrizione prodotto", prezzo: "prezzo al rivenditore",
  famiglia: "famiglia di sconto", um: "unita di misura", molt: "moltiplicatore prezzo",
};

const UM = { each: "n.", pce: "n.", pc: "n.", pz: "n.", nr: "n.", n: "n.", "n.": "n.", pcs: "n.", lm: "m", mtr: "m", m: "m", mt: "m", kg: "kg", set: "kit", kit: "kit" };
export const unita = (v) => { const s = norm(v); return s ? UM[s] || s : "n."; };

/* Prezzo da una cella: numero, "4.725,00", "335,00 €", "1157". */
export function prezzoDa(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = testo(v).replace(/[€\s]/g, "");
  if (!/\d/.test(s) || /[a-z]/i.test(s)) return 0;
  return num(s);
}

/* Legge i fogli scelti con la mappa delle colonne.
   fogli = [{ nome, righe }], mappa = { campo: intestazione normalizzata },
   opz.sottoFoglio = il nome del foglio fa da sottocapitolo.
   Le righe con un solo testo e senza prezzo (es. "Accessories", "Centrali serie Cerberus PRO") sono titoli di sezione.
   Restituisce { articoli, scartate, perFoglio }. */
export function leggiListino(fogli, mappa, opz = {}) {
  const articoli = [], perFoglio = {};
  const visti = new Map();
  let scartate = 0;
  const colDi = (celle) => Object.fromEntries(Object.entries(mappa).filter(([, h]) => h).map(([k, h]) => [k, celle.indexOf(h)]).filter(([, i]) => i >= 0));
  for (const { nome, righe } of fogli) {
    let col = null, sezione = "", n = 0;
    for (let i = 0; i < righe.length; i++) {
      const r = righe[i] || [];
      const celle = r.map(norm);
      // intestazione (anche ripetuta più volte nel foglio, come nei listini a sezioni)
      if (mappa.codice && mappa.prezzo && celle.includes(mappa.codice) && celle.includes(mappa.prezzo)) { col = colDi(celle); continue; }
      if (!col) continue;
      const piene = r.map(testo).filter(Boolean);
      const cod = testo(r[col.codice]);
      const prezzo = prezzoDa(r[col.prezzo]);
      if (!prezzo) {
        if (piene.length === 1 && !cod.match(/^\d/) && piene[0].length < 80) sezione = piene[0]; // titolo di sezione
        else if (piene.length) scartate++;
        continue;
      }
      if (!cod) { scartate++; continue; }
      const molt = col.molt !== undefined ? num(r[col.molt]) || 1 : 1;
      const netto = col.netto !== undefined ? prezzoDa(r[col.netto]) : 0;
      let sconto = col.sconto !== undefined ? num(r[col.sconto]) : 0;
      if (sconto > 1) sconto /= 100; // 54 → 54%
      const listino = prezzo / molt;
      const a = {
        codice: cod.replace(/\s+/g, " "),
        nOrdine: col.nOrdine !== undefined ? testo(r[col.nOrdine]) : "",
        descrizione: col.descrizione !== undefined ? testo(r[col.descrizione]).replace(/\s+/g, " ") : "",
        listino: Math.round(listino * 1e6) / 1e6,
        scListino: netto ? 1 - netto / prezzo : sconto || null,
        famiglia: col.famiglia !== undefined ? testo(r[col.famiglia]) : "",
        gma: col.gma !== undefined ? testo(r[col.gma]) : "",
        um: col.um !== undefined ? unita(r[col.um]) : "n.",
        sottocapitolo: [opz.sottoFoglio ? nome : "", col.sottocapitolo !== undefined ? testo(r[col.sottocapitolo]) : sezione].filter(Boolean).join(" · "),
        _sez: !!sezione || col.sottocapitolo !== undefined,
      };
      if (a.scListino !== null) a.scListino = Math.round(a.scListino * 1e6) / 1e6;
      // lo stesso codice in più fogli (es. "All Parts" e il foglio della linea): vale il primo, ma si tiene la sezione
      const prima = visti.get(a.codice);
      if (prima) { if (!prima._sez && a._sez) { prima.sottocapitolo = a.sottocapitolo; prima._sez = true; } continue; }
      visti.set(a.codice, a);
      articoli.push(a);
      n++;
    }
    perFoglio[nome] = n;
  }
  for (const a of articoli) delete a._sez;
  return { articoli, scartate, perFoglio };
}

/* Testo incollato o CSV: tabulazione, punto e virgola o virgola. */
export function righeDaTesto(t) {
  const linee = t.replace(/\r/g, "").split("\n").filter((l) => l.trim());
  const sep = linee.some((l) => l.includes("\t")) ? "\t" : linee.slice(0, 5).every((l) => l.includes(";")) ? ";" : ",";
  return linee.map((l) => {
    if (sep !== ",") return l.split(sep).map((x) => x.replace(/^"|"$/g, ""));
    const out = []; let cur = "", q = false;
    for (const ch of l) { if (ch === '"') q = !q; else if (ch === "," && !q) { out.push(cur); cur = ""; } else cur += ch; }
    out.push(cur);
    return out;
  });
}

/* Sconti in cascata da testo ("50+10", "75") → frazioni [0.5, 0.1]. */
export const terminiSconto = (t) => String(t ?? "").replace(/%/g, "").replace(/\s+/g, "").split("+").filter((x) => /^-?\d+(?:[.,]\d+)?$/.test(x)).map((x) => num(x) / 100);
export const fattoreDi = (termini) => termini.reduce((f, x) => f * (1 - x), 1);

/* Unione con quello che c'è già in biblioteca per quel fornitore (abbinamento per codice).
   Restituisce { articoli, nuovi, aggiornati, invariati, tolti: [...] }. La decisione voce per voce resta
   se lo sconto del listino non è cambiato. */
export function unisciListino(esistenti, letti, oggi) {
  const perCod = new Map(esistenti.map((a) => [a.codice, a]));
  const articoli = [], aggiornati = [];
  let nuovi = 0, invariati = 0;
  for (const l of letti) {
    const e = perCod.get(l.codice);
    if (!e) { articoli.push({ ...l, dataAgg: oggi }); nuovi++; continue; }
    perCod.delete(l.codice);
    const cambiaPrezzo = Math.abs(num(e.listino) - l.listino) > 0.004;
    const stessoSconto = (e.scListino ?? null) === (l.scListino ?? null);
    const a = { ...e, ...l, dataAgg: cambiaPrezzo ? oggi : e.dataAgg || oggi };
    if (e.decisione && stessoSconto) a.decisione = e.decisione; else delete a.decisione;
    if (cambiaPrezzo) aggiornati.push({ codice: l.codice, da: num(e.listino), a: l.listino });
    else invariati++;
    articoli.push(a);
  }
  return { articoli, nuovi, aggiornati, invariati, tolti: [...perCod.values()] };
}
