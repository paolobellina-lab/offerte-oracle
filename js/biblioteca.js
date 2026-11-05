// Biblioteca: lettura dei listini fornitore (Excel, CSV, testo incollato) in un elenco di articoli.
// Funzioni pure sulle righe già lette (array di array): si usano nell'app e nei test.
import { num } from "./calc.js";

/* Campi che si possono leggere da un listino, con le parole che di solito ci sono nell'intestazione. */
export const CAMPI = [
  { k: "codice", label: "Codice (va in Oracle)", obbl: true, chiavi: ["codice prodotto azienda", "codice ordinazione", "codice d'ordine", "codice articolo", "codice prodotto", "order code", "sen", "codici", "codice", "code", "articolo", "art."] },
  { k: "nOrdine", label: "Modello / n° ordinazione", chiavi: ["article type", "riferimento prodotto", "modello", "tipo", "type", "riferimento", "codice prodotto"] },
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

/* ================= Listini in PDF =================
   pagine = [[{ str, x, y, w, h }]] (testo di ogni pagina con la posizione, come lo dà pdf.js).
   Ricostruisce una tabella: righe del PDF → celle → colonne (per posizione), così il resto funziona come per un Excel.
   - riga di dati = riga che contiene un prezzo (es. "4.725,00", "335,00 €");
   - le righe di testo senza prezzo (descrizioni su più righe) si attaccano alla riga di dati più vicina;
   - una riga breve che parte dalla prima colonna è un titolo di sezione (sottocapitolo);
   - intestazioni e piè di pagina che si ripetono su più pagine si ignorano. */
const RE_PREZZO = /^(?:€\s*)?\d{1,3}(?:\.\d{3})*,\d{2}\s*€?$|^(?:€\s*)?\d+[.,]\d{2}\s*€?$/;
const pareParolaIntestazione = (t) => /\b(codic|prezz|descriz|listino|riferim|unitari|articol|code|price|description)/i.test(t);
export function righeDaPdf(pagine) {
  // 1. righe di ogni pagina, con le celle (pezzi di testo vicini uniti)
  const tutte = [];
  pagine.forEach((items, p) => {
    const ord = items.filter((i) => i.str.trim()).slice().sort((a, b) => b.y - a.y || a.x - b.x);
    const linee = [];
    for (const it of ord) {
      const tol = Math.max(2, (it.h || 8) * 0.35);
      let l = linee.find((x) => Math.abs(x.y - it.y) <= tol);
      if (!l) { l = { p, y: it.y, items: [] }; linee.push(l); }
      l.items.push(it);
    }
    for (const l of linee) {
      l.items.sort((a, b) => a.x - b.x);
      const celle = [];
      for (const it of l.items) {
        const c = celle[celle.length - 1];
        const gap = c ? it.x - c.x2 : Infinity;
        const soloEuro = /^\s*€\s*$/.test(it.str) && c && /\d\s*$/.test(c.t);
        const largChar = it.str.length ? it.w / it.str.length : (it.h || 8) * 0.5;
        if (c && (soloEuro || gap < Math.max(1.5, Math.min((it.h || 8) * 0.5, largChar * 0.9)))) { c.t += (gap > 0.8 ? " " : "") + it.str; c.x2 = Math.max(c.x2, it.x + it.w); }
        else celle.push({ t: it.str, x: it.x, x2: it.x + it.w });
      }
      for (const c of celle) { c.t = c.t.replace(/\s+/g, " ").trim(); c.prezzo = RE_PREZZO.test(c.t); }
      l.celle = celle.filter((c) => c.t);
      l.testo = l.celle.map((c) => c.t).join(" ");
      if (l.celle.length) tutte.push(l);
    }
  });
  // pagine con margini diversi (pari / dispari): le x si contano dal bordo sinistro della tabella di ogni pagina
  const offset = new Map();
  for (const l of tutte) if (l.celle.some((c) => c.prezzo)) offset.set(l.p, Math.min(offset.get(l.p) ?? Infinity, l.celle[0].x));
  for (const l of tutte) { const o = offset.get(l.p) ?? 0; for (const c of l.celle) { c.x -= o; c.x2 -= o; } }
  // 2. intestazioni / piè di pagina ripetuti (stesso testo, numeri esclusi, su almeno 3 pagine)
  const chiave = (l) => l.testo.replace(/\d+/g, "#").toLowerCase();
  const pagineDi = new Map();
  for (const l of tutte) { const k = chiave(l); if (!pagineDi.has(k)) pagineDi.set(k, new Set()); pagineDi.get(k).add(l.p); }
  const ripetuta = (l) => pagine.length >= 3 && pagineDi.get(chiave(l)).size >= 3 && !l.celle.some((c) => c.prezzo);
  const linee = tutte.filter((l) => !ripetuta(l) || pareParolaIntestazione(l.testo));
  const dati = linee.filter((l) => l.celle.some((c) => c.prezzo));
  if (!dati.length) return [];
  // 3. colonne: punto di ancoraggio = inizio per i testi, fine per i numeri (allineati a destra)
  const anc = (c) => (c.prezzo ? { tipo: "n", v: c.x2 } : { tipo: "t", v: c.x });
  const colonne = [];
  for (const l of dati) for (const c of l.celle) {
    const a = anc(c);
    let col = colonne.find((k) => k.tipo === a.tipo && Math.abs(k.v - a.v) <= 14);
    if (!col) { col = { tipo: a.tipo, v: a.v, n: 0, x1: c.x, x2: c.x2, somma: 0 }; colonne.push(col); }
    col.n++; col.somma += a.v; col.v = col.somma / col.n; col.x1 = Math.min(col.x1, c.x); col.x2 = Math.max(col.x2, c.x2);
  }
  // colonne dello stesso tipo che si sovrappongono (testi centrati, numeri di larghezza diversa) sono la stessa colonna
  colonne.sort((a, b) => a.x1 - b.x1);
  for (let i = 0; i < colonne.length; i++) for (let j = i + 1; j < colonne.length; j++) {
    const a = colonne[i], b = colonne[j];
    if (!a || !b || a.tipo !== b.tipo) continue;
    const sov = Math.min(a.x2, b.x2) - Math.max(a.x1, b.x1);
    if (sov > 0.5 * Math.min(a.x2 - a.x1, b.x2 - b.x1)) {
      a.n += b.n; a.somma += b.somma; a.v = a.somma / a.n; a.x1 = Math.min(a.x1, b.x1); a.x2 = Math.max(a.x2, b.x2);
      colonne.splice(j, 1); j = i; // ricomincia: la colonna unita può toccarne altre
    }
  }
  // colonne rare (meno del 3% delle righe) si uniscono alla più vicina
  const soglia = Math.max(2, dati.length * 0.03);
  const buone = colonne.filter((k) => k.n >= soglia);
  const centro = (k) => (k.x1 + k.x2) / 2;
  buone.sort((a, b) => centro(a) - centro(b));
  const colonnaDi = (c) => {
    const a = anc(c);
    const dentro = buone.findIndex((k) => k.tipo === a.tipo && a.v >= k.x1 - 3 && a.v <= k.x2 + 3);
    if (dentro >= 0) return dentro;
    let best = 0, dist = Infinity;
    buone.forEach((k, i) => { const d = Math.abs((k.tipo === a.tipo ? k.v : centro(k)) - a.v) + (k.tipo === a.tipo ? 0 : 40); if (d < dist) { dist = d; best = i; } });
    return best;
  };
  // 4. righe di dati e righe di testo attaccate alla riga di dati più vicina (stessa pagina)
  const righe = dati.map((l) => ({ l, celle: buone.map(() => []) }));
  for (const r of righe) for (const c of r.l.celle) r.celle[colonnaDi(c)].push(c.t);
  const sezioni = [];
  const hMed = (() => { const h = tutte.flatMap((l) => l.celle.length ? [Math.abs(l.y - (tutte.find((m) => m.p === l.p && m.y < l.y)?.y ?? l.y - 12))] : []).filter((x) => x > 3 && x < 30).sort((a, b) => a - b); return h[Math.floor(h.length / 2)] || 11; })();
  const passo = hMed * 1.9; // distanza massima tra due righe dello stesso blocco di testo
  const rigaDi = new Map(righe.map((r) => [r.l, r]));
  const perPagina = new Map();
  for (const l of linee) { if (!perPagina.has(l.p)) perPagina.set(l.p, []); perPagina.get(l.p).push(l); }
  for (const [, ls] of perPagina) {
    ls.sort((a, b) => b.y - a.y);
    const tipo = ls.map((l) => {
      if (l.celle.some((c) => c.prezzo)) return "dati";
      if (pareParolaIntestazione(l.testo) && l.celle.length >= 2) return "intest";
      return "testo";
    });
    ls.forEach((l, i) => {
      if (tipo[i] !== "testo") return;
      // titolo di sezione: breve, nella prima colonna, subito sopra una riga di dati
      if (Math.abs(l.celle[0].x) < 6 && l.celle.length === 1 && l.testo.length < 80 && !/^\d/.test(l.testo) && tipo[i + 1] === "dati" && l.y - ls[i + 1].y < passo * 1.6) {
        sezioni.push({ p: l.p, y: l.y, t: l.testo }); tipo[i] = "sezione"; return;
      }
    });
    // ogni riga di testo va con la riga di dati raggiungibile senza salti (sopra o sotto), la più vicina
    ls.forEach((l, i) => {
      if (tipo[i] !== "testo") return;
      const cerca = (dir) => {
        let d = 0;
        for (let j = i + dir, prec = l; j >= 0 && j < ls.length; j += dir) {
          const gap = Math.abs(prec.y - ls[j].y);
          if (gap > passo || tipo[j] === "sezione" || tipo[j] === "intest") return null;
          d += gap;
          if (tipo[j] === "dati") return { r: rigaDi.get(ls[j]), d };
          prec = ls[j];
        }
        return null;
      };
      const su = cerca(-1), giu = cerca(1);
      const best = su && giu ? (su.d <= giu.d ? su : giu) : su || giu;
      if (!best?.r) return;
      for (const c of l.celle) (best.r.extra ||= []).push({ y: l.y, col: colonnaDi(c), t: c.t });
    });
  }
  for (const r of righe) {
    for (const e of (r.extra || []).sort((a, b) => b.y - a.y)) {
      const sopra = e.y > r.l.y;
      if (sopra) r.celle[e.col].unshift(e.t); else r.celle[e.col].push(e.t);
    }
  }
  // colonne di prezzi che non sono mai piene insieme (pagine di larghezza diversa) sono la stessa colonna
  for (let i = 0; i < buone.length; i++) for (let j = i + 1; j < buone.length; j++) {
    if (buone[i].tipo !== "n" || buone[j].tipo !== "n" || buone[j].via || buone[i].via) continue;
    if (righe.some((r) => r.celle[i].length && r.celle[j].length)) continue;
    for (const r of righe) { r.celle[i].push(...r.celle[j]); r.celle[j] = []; }
    buone[j].via = true; buone[i].x1 = Math.min(buone[i].x1, buone[j].x1); buone[i].x2 = Math.max(buone[i].x2, buone[j].x2);
  }
  const tenere = buone.map((k, i) => (k.via ? -1 : i)).filter((i) => i >= 0);
  if (tenere.length < buone.length) {
    for (const r of righe) r.celle = tenere.map((i) => r.celle[i]);
    buone.splice(0, buone.length, ...tenere.map((i) => buone[i]));
  }
  // 5. intestazione: la riga (con quelle attaccate, per le intestazioni su più righe) con più parole da intestazione
  const conteggio = (l) => l.celle.filter((c) => pareParolaIntestazione(c.t)).length;
  const cand = linee.filter((l) => !l.celle.some((c) => c.prezzo) && conteggio(l) >= 1);
  // l'intestazione della tabella principale è quella che si ripete su più pagine
  const quante = (l) => pagineDi.get(chiave(l))?.size || 1;
  const migliore = cand.slice().sort((a, b) => Math.min(conteggio(b), 2) * quante(b) - Math.min(conteggio(a), 2) * quante(a) || conteggio(b) - conteggio(a) || a.p - b.p || b.y - a.y)[0];
  const blocco = migliore ? cand.concat(linee.filter((l) => !l.celle.some((c) => c.prezzo)))
    .filter((l, i, arr) => arr.indexOf(l) === i && l.p === migliore.p && Math.abs(l.y - migliore.y) <= 12) : [];
  const nomi = buone.map((k) => {
    const t = blocco.sort((a, b) => b.y - a.y).flatMap((l) => l.celle.filter((c) => c.x2 >= k.x1 - 4 && c.x <= k.x2 + 4).map((c) => c.t));
    return [...new Set(t)].join(" ").slice(0, 40);
  });
  const usati = new Set();
  const intest = nomi.map((n, i) => {
    let h = n || `colonna ${i + 1}`;
    if (buone[i].tipo === "n" && !/prezz|listino|price|netto|unitari/i.test(h)) h = `${h} prezzo`.trim();
    while (usati.has(h)) h += " ·";
    usati.add(h);
    return h;
  });
  if (!intest.some((h) => /codic|code/i.test(h))) intest[0] = `codice (${intest[0]})`;
  // 6. righe in ordine, con i titoli di sezione al loro posto
  const out = [intest];
  const eventi = [...righe.map((r) => ({ p: r.l.p, y: r.l.y, r })), ...sezioni.map((s) => ({ p: s.p, y: s.y + 0.001, s }))]
    .sort((a, b) => a.p - b.p || b.y - a.y);
  for (const e of eventi) out.push(e.s ? [e.s.t] : e.r.celle.map((x) => x.join(" ")));
  return [{ nome: "PDF", righe: out }];
}
