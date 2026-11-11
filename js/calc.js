// Motore di calcolo: replica le formule dei fogli TIPO OFFERTA e Sheet1 dell'Excel
// "Template_Import_Righe-OFFERTA_V22". Nessuna dipendenza: si usa sia nell'app sia nei test.

export const CAMPI_SCONTO = ["sc3", "sc4", "sc5", "aum1"]; // colonne Q R S T del catalogo (lo standard dell'articolo)
// Condizioni del capitolo (pagina Prezzi) e scelte dell'output (pagina Oracle)
export const MODI_ACQ = { standard: "Standard", extra: "Extra" };
export const MODI_VEN = { ricarico: "Ricarico", listino: "Listino / sconto" };
// impostazione del capitolo: acquisto e vendita, oppure solo un listino riparametrizzato per tutti gli articoli
export const IMPOSTAZIONI = { av: "Acquisto / vendita", listino: "Listino" };
export const OUT_PREZZO = { vendita: "Vendita", acquisto: "Acquisto" };
export const OUT_FORMATO = { sconti: "Sconti", netti: "Prezzi netti" };

// Intestazioni di Sheet1, identiche all'Excel (anche lo spazio finale di "Prezzo Netto ").
export const COLONNE_ORACLE = [
  "Linea N", "Codice Articolo Fornitore", "Magazzino Spedizione", "Quantità Ordinata", "Nome Capitolato",
  "Prezzo Netto ", "Sconto 3 [%]", "Sconto 4 [%]", "Sconto 5 [%]", "Sconto 3 [€]", "Sconto 4 [€]",
  "Sconto 5 [€]", "Aumento 1 [%]", "Aumento 2 [%]", "Aumento 1 [€]", "Aumento 2 [€]",
  "Note Pre-Linea", "Note Post-Linea",
];

/* Numero da qualunque valore: accetta "0,75", "1.234,50", "75%", "-" (= 0). */
export function num(v) {
  if (v === null || v === undefined || v === "" || v === "-") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  let s = String(v).trim().replace(/[€\s]/g, "");
  const perc = s.endsWith("%");
  if (perc) s = s.slice(0, -1);
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  if (!Number.isFinite(n)) return 0;
  return perc ? n / 100 : n;
}

const arrot = (n, dec = 6) => {
  const f = 10 ** dec;
  return Math.round((n + Number.EPSILON) * f) / f;
};

export const isAttivita = (r) => String(r.codice || "").trim().toUpperCase() === "ATT.TEC";
/* Codice dell'articolo come scritto nel database (precodice + codice). */
export const codiceArticolo = (r) => `${r.precodice || ""}${r.codice || ""}`.trim();
export const CODICE_GENERICO = "P000ARTICOLO_GEN";
/* Codice che va in Oracle: per gli articoli trattati come generici è sempre P000ARTICOLO_GEN. */
export const codiceOracle = (r) => (isGenerico(r) ? CODICE_GENERICO : codiceArticolo(r));
/* Articoli generici (P000 ARTICOLO_GEN): in Oracle il codice è sempre lo stesso, quindi l'articolo vero
   va nelle note: nota pre-linea = numero di ordinazione, nota post-linea = descrizione. Per tutti gli altri
   articoli le note sono testo libero scritto a mano. */
export const isGenerico = (r) => !!r.generico || String(r.codice || "").trim().toUpperCase() === "ARTICOLO_GEN";
export const notePreRiga = (r) => (isGenerico(r) ? (r.nOrdine || "") : (r.notePre || ""));
export const notePostRiga = (r) => (isGenerico(r) ? (r.descrizione || "") : (r.notePost || ""));

/* Colonna M del catalogo: valore riparametrizzato (solo se c'è la riparametrizzazione P). */
export function riparametrizzato(r) {
  const P = num(r.riparam);
  if (!P || P === 1) return null;
  const e = scontiEffettivi(r);
  return num(r.listinoAcquisto) * (1 - e.sc3) * (1 - e.sc4) * (1 - e.sc5) * (1 - e.aum1) / (1 - P);
}

/* Colonna K del catalogo: prezzo di listino, calcolato con le formule dell'Excel.
   Articoli: IF(P<>0, M, N). Attività tecnica: ROUNDUP(M*giorni, -2). */
export function listinoCalcolato(r) {
  const M = riparametrizzato(r);
  if (isAttivita(r)) {
    const base = (M ?? num(r.listinoAcquisto)) * num(r.giorni);
    return Math.ceil(arrot(base / 100, 9)) * 100;
  }
  return M ?? num(r.listinoAcquisto);
}

/* Prezzo di listino della riga: quello salvato nel catalogo (importato dalla colonna K o modificato a mano),
   altrimenti quello calcolato con le formule. */
export function prezzoListino(r) {
  return num(r.listino) || listinoCalcolato(r);
}

/* ================= Condizioni di acquisto e di vendita =================
   Ogni capitolo dell'offerta ha:
   acq = { modo: "standard" | "extra", gen: "…", gma: { GMA: "…" } }
   ven = { modo: "ricarico" | "listino", ricGen, ricGma: {}, scGen, scGma: {} }
   out = { prezzo: "vendita" | "acquisto", formato: "sconti" | "netti" }
   Gli sconti si scrivono come testo, in percentuale:
     "75+15" = sconti in cascata al posto dello standard dell'articolo;
     "55"    = al posto dello standard;
     "+15"   = in aggiunta allo standard (es. standard 75 → 75+15);
     "0"     = nessuno sconto (prezzo pieno); "-5" = aumento del 5%.
   Vuoto = vale il livello sopra (GMA → generale → standard del database).
   Restano salvati tutti, anche quando il selettore è sull'altra scelta. */
const RE_TERMINE = /^-?\d+(?:[.,]\d+)?$/;
export function leggiSconti(t) {
  let s = String(t ?? "").replace(/%/g, "").replace(/\s+/g, "");
  if (!s) return null;
  const aggiunta = s.startsWith("+");
  if (aggiunta) s = s.slice(1);
  const parti = s.split("+");
  if (!parti.length || !parti.every((p) => RE_TERMINE.test(p))) return { errore: true };
  const fattore = parti.reduce((f, p) => f * (1 - num(p) / 100), 1);
  return { aggiunta, fattore, testo: (aggiunta ? "+" : "") + parti.join("+") };
}
/* Ricarico in percentuale: "28" = 28% al posto dell'automatico; "20+5" = 20% poi 5%;
   "+5" = in aggiunta al ricarico automatico del capitolo. Restituisce { aggiunta, fattore } o null. */
export function leggiRicSpec(t) {
  let s = String(t ?? "").replace(/%/g, "").replace(/\s+/g, "");
  if (!s) return null;
  const aggiunta = s.startsWith("+");
  if (aggiunta) s = s.slice(1);
  const parti = s.split("+");
  if (!parti.every((p) => RE_TERMINE.test(p))) return null;
  return { aggiunta, fattore: parti.reduce((f, p) => f * (1 + num(p) / 100), 1) };
}
/* Ricarico risultante (frazione) dato l'automatico, oppure null se il testo è vuoto o sbagliato. */
export function leggiRicarico(t, auto = 0) {
  const sp = leggiRicSpec(t);
  if (!sp) return null;
  return (sp.aggiunta ? (1 + num(auto)) * sp.fattore : sp.fattore) - 1;
}
export const testoValido = (t, ric = false) => !String(t ?? "").trim() || (ric ? leggiRicarico(t) !== null : !leggiSconti(t)?.errore);
/* Acquisti standard per GMA (Impostazioni › Acquisti standard): Sconto 1, 2, 3 e Aumento.
   Un articolo senza sconti propri usa quelli del suo GMA; se ne ha anche uno solo, valgono i suoi (eccezione).
   Nei dati gli sconti dell'articolo restano nei campi sc3, sc4, sc5, aum1 (Sconto 1, 2, 3, Aumento 1). */
// Un GMA è identificato da fornitore + codice: lo stesso codice con due fornitori sono due GMA diversi.
let STD_GMA = new Map(), PER_CODICE = new Map();
const chiaveGma = (forn, g) => `${String(forn || "").trim().toUpperCase()}|${g}`;
export function impostaStandardGma(lista) {
  STD_GMA = new Map(); PER_CODICE = new Map();
  for (const g of lista || []) {
    if (!PER_CODICE.has(g.codice)) PER_CODICE.set(g.codice, []);
    PER_CODICE.get(g.codice).push(g);
    if (num(g.sc1) || num(g.sc2) || num(g.sc3) || num(g.aum)) STD_GMA.set(chiaveGma(g.fornitore, g.codice), g);
  }
}
/* Voce GMA di un articolo: quella del suo fornitore; se l'articolo non ha fornitore, la prima con quel codice. */
export function voceGma(g, forn) {
  const l = PER_CODICE.get(g) || [];
  if (forn) return l.find((x) => String(x.fornitore || "").trim().toUpperCase() === String(forn).trim().toUpperCase()) || null;
  return l[0] || null;
}
/* Codice con cui Oracle conosce il GMA (es. SENSITRON → SIFSFI): per cartellino e censimento. */
export const gmaOracle = (g, forn) => (voceGma(g, forn) || voceGma(g))?.oracle || g;
export const standardGma = (g, forn) => { const v = voceGma(g, forn); return v && STD_GMA.get(chiaveGma(v.fornitore, v.codice)) ? v : null; };
export const haScontiPropri = (r) => !!(num(r.sc3) || num(r.sc4) || num(r.sc5) || num(r.aum1));
export function scontiEffettivi(r) {
  // le righe prese dall'offerta di un fornitore hanno già il prezzo netto del fornitore
  if (haScontiPropri(r) || isAttivita(r) || r.daOfferta) return { sc3: num(r.sc3), sc4: num(r.sc4), sc5: num(r.sc5), aum1: num(r.aum1), daGma: false };
  const s = standardGma(r.gmv, r.fornitore);
  return s ? { sc3: num(s.sc1), sc4: num(s.sc2), sc5: num(s.sc3), aum1: num(s.aum), daGma: true } : { sc3: 0, sc4: 0, sc5: 0, aum1: 0, daGma: false };
}
/* Sconti standard dell'articolo: i suoi o quelli del GMA. L'attività tecnica non ne ha. */
export const fattoreStandard = (r) => { if (isAttivita(r)) return 1; const e = scontiEffettivi(r); return (1 - e.sc3) * (1 - e.sc4) * (1 - e.sc5) * (1 + e.aum1); };
const applica = (sp, f0) => (sp.aggiunta ? f0 * sp.fattore : sp.fattore);
const valido = (sp) => (sp && !sp.errore ? sp : null);
/* Il valore generale vale anche per l'attività tecnica solo se è "in aggiunta" (+3):
   l'attività non ha uno standard da sostituire, quindi un "75+15" non la tocca (come nell'Excel l'extrasconto
   del capitolo valeva per tutte le righe). */
const generalePer = (att, testo) => { const sp = valido(leggiSconti(testo)); return sp && (!att || sp.aggiunta) ? sp : null; };

/* Condizioni del capitolo, sempre complete (anche per i capitoli salvati con il vecchio modello). */
export function condizioni(cap) {
  const c = cap.acq && cap.ven ? cap : migraCondizioni(cap);
  const k = {
    acq: { modo: c.acq?.modo === "extra" ? "extra" : "standard", gen: c.acq?.gen || "", gma: c.acq?.gma || {} },
    ven: { modo: c.ven?.modo === "listino" ? "listino" : "ricarico", ricGen: c.ven?.ricGen || "", ricGma: c.ven?.ricGma || {}, scGen: c.ven?.scGen || "", scGma: c.ven?.scGma || {} },
    out: { prezzo: c.out?.prezzo === "acquisto" ? "acquisto" : "vendita", formato: c.out?.formato === "netti" ? "netti" : "sconti" },
    imp: c.imp === "listino" ? "listino" : "av",
    lis: { gen: c.lis?.gen || "", gma: c.lis?.gma || {} },
  };
  if (k.imp === "listino") Object.assign(k.lis, stdListino(cap.righe || [], k));
  return k;
}
/* Listino: lo sconto di listino standard è lo sconto di acquisto del GMA con più righe nel capitolo
   (il più frequente tra le sue righe). Così gli articoli di quel GMA restano al loro listino e gli altri
   si riparametrizzano sullo stesso sconto. */
export function stdListino(righe, k) {
  const perGma = new Map();
  for (const r of righe) {
    if (!num(r.qta) || isAttivita(r) || !r.gmv) continue;
    if (!perGma.has(r.gmv)) perGma.set(r.gmv, []);
    perGma.get(r.gmv).push(r);
  }
  const [gma, rr] = [...perGma.entries()].sort((a, b) => b[1].length - a[1].length)[0] || [];
  if (!gma) return { std: 0, stdGma: "" };
  const freq = new Map();
  for (const r of rr) { const f = Math.round(fattoreAcquisto(r, k).f * 1e8) / 1e8; freq.set(f, (freq.get(f) || 0) + 1); }
  const f = [...freq.entries()].sort((a, b) => b[1] - a[1])[0][0];
  return { std: Math.round((1 - f) * 1e8) / 1e8, stdGma: gma };
}
/* Frazione → testo in percentuale ("0.125" → "12,5"). */
export const pctTesto = (v) => String(Math.round(num(v) * 1e8) / 1e6).replace(".", ",");
/* Dal vecchio modello (tipo offerta, DA ORACLE/DA EXCEL, Sconto 3/4/5, Aumento 1, ricarico) al nuovo.
   Un campo scritto nel capitolo o nel GMA sostituiva solo quello del database: se c'è lo Sconto 3
   diventa "al posto dello standard", altrimenti "in aggiunta" (Sconto 4 e 5 nel database sono quasi sempre vuoti). */
export function migraCondizioni(cap) {
  const glob = cap.globali || {}, perGmv = cap.gmv || {};
  const espr = (o) => {
    const parti = CAMPI_SCONTO.filter((k) => num(o[k])).map((k) => pctTesto(k === "aum1" ? -num(o[k]) : num(o[k])));
    if (!parti.length) return "";
    return (num(o.sc3) ? "" : "+") + parti.join("+");
  };
  const gen = espr(glob);
  const gma = {}, ricGma = {};
  for (const [g, v] of Object.entries(perGmv)) {
    const e = espr({ ...glob, ...Object.fromEntries(CAMPI_SCONTO.filter((k) => num(v[k])).map((k) => [k, v[k]])) });
    if (e && e !== gen) gma[g] = e;
    if (num(v.ric)) ricGma[g] = pctTesto(v.ric);
  }
  const tipo = cap.tipo || "VENDITA CLIENTE";
  const listino = tipo === "LISTINO" || tipo === "LISTINO COMPUTO";
  return {
    ...cap,
    acq: { modo: gen || Object.keys(gma).length ? "extra" : "standard", gen, gma },
    ven: { modo: listino ? "listino" : "ricarico", ricGen: num(glob.ric) ? pctTesto(glob.ric) : "", ricGma,
      scGen: listino ? gen.replace(/^\+/, "") || "0" : "", scGma: listino ? Object.fromEntries(Object.entries(gma).map(([g, e]) => [g, e.replace(/^\+/, "")])) : {} },
    out: { prezzo: tipo === "ACQUISTO" ? "acquisto" : "vendita", formato: cap.prezzi === "DA EXCEL" ? "netti" : "sconti" },
  };
}

/* Attività "a giornate": in vendita il prezzo dell'attività è complessivo (quantità 1);
   con questa scelta va in Oracle come valore giornaliero × numero di giorni. In acquisto è sempre così. */
export const aGiornate = (r) => isAttivita(r) && !!r.aGiornate && num(r.giorni) > 0;

/* Fattore di acquisto di una riga: standard del database o condizioni della pagina Prezzi. */
export function fattoreAcquisto(r, k) {
  const att = isAttivita(r), g = r.gmv || "", f0 = fattoreStandard(r);
  const spG = valido(leggiSconti(k.acq.gma[g])), spC = generalePer(att, k.acq.gen);
  if (spG) return { f: applica(spG, f0), fonte: "gma" };
  if (spC) return { f: applica(spC, f0), fonte: "generale" };
  return { f: f0, fonte: "standard" };
}

/* Calcolo completo di una riga. ricarichi = { NOME_CAPITOLO_DATABASE: 0.28, ... } (ricarico automatico).
   Restituisce i valori di vendita (unitario, qta, totale) e quelli di acquisto in "acq". */
export function calcolaRiga(r, cap, ricarichi = {}, k = condizioni(cap)) {
  const att = isAttivita(r), g = r.gmv || "";
  const gg = num(r.giorni), la = num(r.listinoAcquisto);
  const f0 = fattoreStandard(r);
  // acquisto: standard del database, oppure le condizioni scritte nella pagina Prezzi (del GMA o generali)
  const { f: fAcq, fonte: fonteAcq } = fattoreAcquisto(r, k);
  const acqUnit = la * fAcq; // per l'attività: al giorno
  const acqQta = att && gg > 0 ? gg * num(r.qta) : num(r.qta);
  const acq = { unitario: acqUnit, qta: acqQta, totale: acqUnit * acqQta, fattore: fAcq, fonte: fonteAcq };
  // vendita: ricarico sull'acquisto, oppure sconto sul listino
  let perPezzo, ric = null, fVen = null, baseVen, fonteVen, scLis = null;
  if (k.imp === "listino") {
    // listino riparametrizzato: il prezzo di listino è quello che, con lo sconto di listino, dà il prezzo di acquisto.
    // Standard: articoli = sconto di listino del GMA con più righe; attività = nessuno sconto (si imposta nella riga ATT
    // della pagina prezzi). La riparametrizzazione dell'articolo nel database non si usa più. Attività arrotondata alle centinaia.
    const f0L = att ? 1 : 1 - (k.lis.std || 0);
    const spG = valido(leggiSconti(k.lis.gma[g])), spC = generalePer(att, k.lis.gen);
    const fL = spG ? applica(spG, f0L) : spC ? applica(spC, f0L) : f0L;
    fonteVen = spG ? "gma" : spC ? "generale" : "standard";
    scLis = 1 - fL;
    const costo = acqUnit * (att && gg > 0 ? gg : 1);
    baseVen = costo;
    perPezzo = fL > 0 ? costo / fL : costo;
    if (att && gg > 0) perPezzo = Math.ceil(Math.round(perPezzo / 100 * 1e9) / 1e9) * 100;
  } else if (k.ven.modo === "ricarico") {
    const auto = num(ricarichi[r.capitolo]);
    const rG = leggiRicarico(k.ven.ricGma[g], auto), rC = leggiRicarico(k.ven.ricGen, auto);
    ric = rG ?? rC ?? auto;
    fonteVen = rG !== null ? "gma" : rC !== null ? "generale" : "automatico";
    baseVen = acqUnit * (att && gg > 0 ? gg : 1);
    perPezzo = baseVen * (1 + ric);
  } else {
    baseVen = att ? prezzoListino(r) : la; // attività: il prezzo di listino di tutti i giorni
    const spG = valido(leggiSconti(k.ven.scGma[g])), spC = generalePer(att, k.ven.scGen);
    fVen = spG ? applica(spG, f0) : spC ? applica(spC, f0) : f0;
    fonteVen = spG ? "gma" : spC ? "generale" : "standard";
    perPezzo = baseVen * fVen;
  }
  const giornate = aGiornate(r);
  const unitario = giornate ? perPezzo / gg : perPezzo;
  const qta = giornate ? gg * num(r.qta) : num(r.qta);
  // sconto complessivo rispetto al listino acquisto (quello che Oracle conosce): serve al cartellino
  const scontoListino = !att && la ? 1 - perPezzo / la : null;
  return { unitario, qta, totale: unitario * qta, base: baseVen, ric, fVen, fonteVen, scontoListino, acq, f0, scLis };
}

/* Valori da mandare in Oracle per una riga: vendita o acquisto, come scelto nell'output. */
export function valoriOutput(c, k) {
  return k.out.prezzo === "acquisto" ? { unitario: c.acq.unitario, qta: c.acq.qta } : { unitario: c.unitario, qta: c.qta };
}

/* Ordine delle righe: come nell'Excel, nell'ordine del catalogo, capitolo per capitolo. */
export function ordinaRighe(righe, ordineCapitoli = []) {
  const pos = (c) => { const i = ordineCapitoli.indexOf(c); return i < 0 ? 999 : i; };
  return [...righe].sort((a, b) =>
    pos(a.capitolo) - pos(b.capitolo) || (a.capitolo || "").localeCompare(b.capitolo || "") ||
    num(a.ordine) - num(b.ordine));
}

/* Righe effettive di un capitolo dell'offerta, nell'ordine scelto dall'utente:
   solo quelle con quantità diversa da 0 (come il filtro su Q.tà ≠ 0).
   "off" è un capitolo dell'offerta: { acq, ven, out, righe, ... }. */
export function calcolaOfferta(off, { ricarichi = {} } = {}) {
  const k = condizioni(off);
  const righe = (off.righe || []).filter((r) => num(r.qta) !== 0)
    .map((r) => ({ riga: r, c: calcolaRiga(r, off, ricarichi, k) }));
  const totale = righe.reduce((s, x) => s + x.c.totale, 0);
  return { righe, totale };
}

/* Tracciato Sheet1 per Oracle. Restituisce { righe: [array di 18 valori], avvisi: [...] }. */
/* Capitolati esclusi dall'esportazione (per caricarli in un'altra offerta Oracle). "" = righe senza capitolato. */
export const esclusoDaExport = (r, off) => (off.esclusiExport || []).includes(r.capitolato || "");

export function righeOracle(off, opz = {}) {
  // solo i capitolati spuntati; le linee ripartono dal n° di partenza
  const { righe } = calcolaOfferta({ ...off, righe: (off.righe || []).filter((r) => !esclusoDaExport(r, off)) }, opz);
  const k = condizioni(off);
  const start = num(off.lineaPartenza) || 10;
  const out = [], avvisi = [];
  let linea = start;
  for (const { riga: r, c } of righe) {
    const codice = codiceOracle(r);
    if (!codice) {
      avvisi.push(`Riga senza codice articolo esclusa: ${r.nOrdine || r.descrizione || "(senza descrizione)"} (${r.capitolo})`);
      continue;
    }
    const v = valoriOutput(c, k);
    const la = num(r.listinoAcquisto);
    // Prezzi netti: solo la colonna del netto. Sconti: un unico sconto in Sconto 3, calcolato sul listino
    // acquisto (il prezzo che Oracle conosce), anche negativo se il prezzo supera il listino.
    // Attività tecnica e generici vanno sempre a prezzo netto.
    let conNetto = k.out.formato === "netti" || codice === "PITATT.TEC" || isGenerico(r);
    if (!conNetto && !la) { conNetto = true; avvisi.push(`Senza listino acquisto, va a prezzo netto: ${codice} ${r.nOrdine || ""}`); }
    const prezzoNetto = conNetto ? arrot(v.unitario) : "";
    const sc = conNetto ? "" : arrot((1 - v.unitario / la) * 100);
    out.push([
      linea, codice, off.magazzino || "BG", arrot(v.qta),
      off.sottocapitoli ? (r.capitolato || "") : "", // capitolato dell'offerta, non il sottocapitolo del database
      prezzoNetto,
      sc === 0 ? "" : sc, "", "", "", "", "",
      "", "", "", "",
      notePreRiga(r) !== "0" ? notePreRiga(r).trim() : "",
      notePostRiga(r) !== "0" ? notePostRiga(r).trim() : "",
    ]);
    linea += 10;
  }
  return { righe: out, avvisi };
}

/* Controllo cartellino cliente (foglio SC_CLIENTI e colonna K di TIPO OFFERTA).
   Per ogni GMA dell'offerta confronta lo sconto complessivo con quello del cartellino del cliente;
   per l'attività tecnica confronta il prezzo a giornata con la tariffa del cliente
   (l'Excel confrontava il prezzo di tutti i giorni con la tariffa giornaliera). */
export const gmaGenerico = (g, forn) => {
  const c = String(g || "").toUpperCase().replace(/[\s_.-]/g, "");
  const f = String(forn || voceGma(g)?.fornitore || "").toUpperCase().replace(/\s+/g, " ").trim();
  return c === "ARTGEN" || f === "ARTICOLI GENERICI";
};
export function controlloCartellino(off, cliente, opz = {}) {
  if (!cliente) return [];
  const { righe } = calcolaOfferta(off, opz);
  const visti = new Map();
  for (const x of righe) {
    const g = x.riga.gmv || "";
    // GMA degli articoli generici (ART_GEN / fornitore ARTICOLI GENERICI) e righe prese da un'offerta fornitore:
    // nessuno sconto di cartellino da confrontare. Un articolo generico con un GMA vero si controlla.
    if (gmaGenerico(g, x.riga.fornitore) || (isGenerico(x.riga) && x.riga.daOfferta)) continue;
    // prestazioni prese dall'offerta di un fornitore: valore complessivo del fornitore, non una tariffa a giornata
    if (isAttivita(x.riga) && x.riga.daOfferta) continue;
    if (!g || visti.has(g)) continue; // come CERCA.VERT: vale la prima riga del GMA
    visti.set(g, x);
  }
  const esiti = [];
  for (const [gmv, x] of visti) {
    if (gmv === "ATT" || isAttivita(x.riga)) {
      const att = righe.find((y) => isAttivita(y.riga) && !y.riga.daOfferta) || x;
      const tariffa = cliente.att === undefined || cliente.att === "" ? null : num(cliente.att);
      if (tariffa === null) continue;
      // tariffa giornaliera: in LISTINO e VENDITA CLIENTE il prezzo dell'attività comprende tutti i giorni
      const gg = num(att.riga.giorni);
      const prezzo = gg > 0 && !aGiornate(att.riga) ? att.c.unitario / gg : att.c.unitario;
      const stato = tariffa > prezzo + 1e-9 ? "ok" : Math.abs(tariffa - prezzo) < 1e-9 ? "pari" : "fuori";
      esiti.push({ gmv, tipo: "attivita", offerta: prezzo, cartellino: tariffa, stato,
        testo: stato === "ok" ? "OK" : stato === "pari" ? "Valore di cartellino" : "Superiore al valore di cartellino" });
      continue;
    }
    const scontoOfferta = x.c.scontoListino ?? 0;
    // il cartellino del GMA; se non c'è, quello del codice con cui lo conosce Oracle (es. SENSITRON → SIFSFI)
    const proprio = cliente.sconti?.[gmv], og = gmaOracle(gmv, x.riga.fornitore);
    const scCart = num(proprio !== undefined && proprio !== "" && proprio !== null ? proprio : cliente.sconti?.[og]);
    const stato = scCart < scontoOfferta - 1e-9 ? "ok" : Math.abs(scCart - scontoOfferta) < 1e-9 ? "pari" : "fuori";
    esiti.push({ gmv, tipo: "sconto", offerta: scontoOfferta, cartellino: scCart, stato,
      testo: stato === "ok" ? "OK" : stato === "pari" ? "Sconto di cartellino" : "Inferiore al valore di cartellino" });
  }
  return esiti;
}

/* Colore della data di aggiornamento prezzi (legenda dei fogli catalogo). */
export function freschezza(iso, oggi = new Date()) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const mesi = (oggi - d) / (1000 * 60 * 60 * 24 * 30.44);
  if (mesi < 1) return "fresco";
  if (mesi <= 2) return "medio";
  return "vecchio";
}

/* Riepilogo economico di un capitolo (ex foglio ANALISI PACCHETTO).
   Venduto = condizioni di vendita; Acquistato = condizioni di acquisto (standard o extra). */
export function riepilogoCapitolo(cap, opz = {}) {
  const v = calcolaOfferta(cap, opz);
  const out = {
    righe: v.righe.length, venduto: 0, acquistato: 0,
    attivita: { venduto: 0, acquistato: 0 }, materiale: { venduto: 0, acquistato: 0 },
    perGmv: new Map(),
  };
  for (const x of v.righe) {
    const costo = x.c.acq.totale, vend = x.c.totale;
    out.venduto += vend; out.acquistato += costo;
    const parte = isAttivita(x.riga) ? out.attivita : out.materiale;
    parte.venduto += vend; parte.acquistato += costo;
    const g = x.riga.gmv || "";
    if (!out.perGmv.has(g)) out.perGmv.set(g, { righe: 0, venduto: 0, acquistato: 0, prima: x });
    const pg = out.perGmv.get(g);
    pg.righe++; pg.venduto += vend; pg.acquistato += costo;
  }
  return out;
}
export const ricaricoPct = (venduto, acquistato) => (acquistato ? venduto / acquistato - 1 : 0);
export const marginePct = (venduto, acquistato) => (venduto ? 1 - acquistato / venduto : 0);

/* Prezzo target (condizione a pacchetto): aggiunge alla riga Generale una cella extra (ricarico, sconto di vendita
   o sconto di listino) che porta il totale del capitolo al target. Lo standard e gli extra già scritti restano.
   Il valore si trova per bisezione sul calcolo vero (vale anche con l'arrotondamento dell'attività).
   Restituisce { campo: "ricGen" | "scGen" | "lisGen", extra (frazione), testo } oppure null. */
export function vendutaPerTarget(cap, target, opz = {}) {
  const k = condizioni(cap);
  const campo = k.imp === "listino" ? "lisGen" : k.ven.modo === "ricarico" ? "ricGen" : "scGen";
  const attuale = String((campo === "lisGen" ? k.lis.gen : k.ven[campo]) || "").trim().replace(/\+$/, "");
  const testoCon = (x) => `${attuale}+${pctTesto(x)}`;
  const totale = (x) => {
    const t = testoCon(x);
    const c = campo === "lisGen" ? { ...cap, lis: { ...k.lis, gen: t } } : { ...cap, ven: { ...k.ven, [campo]: t } };
    return calcolaOfferta(c, opz).totale;
  };
  let lo = campo === "ricGen" ? -0.95 : -3, hi = campo === "ricGen" ? 20 : 0.99;
  const tLo = totale(lo), tHi = totale(hi), t = num(target);
  if (Math.abs(tHi - tLo) < 1e-9) return null; // nessuna riga dipende dal generale
  if ((t - tLo) * (t - tHi) > 0) return null;   // target fuori portata
  const cresce = tHi > tLo;
  for (let n = 0; n < 80; n++) {
    const m = (lo + hi) / 2, v = totale(m);
    if ((v < t) === cresce) lo = m; else hi = m;
  }
  const extra = Math.round(((lo + hi) / 2) * 1e8) / 1e8;
  return { campo, extra, testo: testoCon(extra) };
}

/* Stato del prezzo di un articolo, in base alla data di aggiornamento e alla regola del suo precodice.
   regole = { PRECODICE: { tipo: "richiesta", validitaMesi } | { tipo: "listino", cambioListino: "AAAA-MM-GG" } }
   Restituisce { livello: "fresco" | "medio" | "vecchio" | "", testo }. */
export function statoPrezzo(a, regole = {}, oggi = new Date()) {
  const regola = regole[String(a.precodice || "").trim().toUpperCase()];
  const agg = a.dataAgg ? new Date(a.dataAgg) : null;
  const dataTxt = agg ? agg.toLocaleDateString("it-IT") : "";
  const mesiDa = (d) => (oggi - d) / (1000 * 60 * 60 * 24 * 30.44);
  if (regola?.tipo === "listino" && regola.cambioListino) {
    const cambio = new Date(regola.cambioListino);
    const cambioTxt = cambio.toLocaleDateString("it-IT");
    if (cambio <= oggi) {
      if (!agg || agg < cambio) return { livello: "vecchio", testo: `Listino cambiato il ${cambioTxt}: prezzo da aggiornare${dataTxt ? ` (aggiornato il ${dataTxt})` : ""}` };
      return { livello: "fresco", testo: `Prezzo aggiornato il ${dataTxt}, dopo il cambio listino del ${cambioTxt}` };
    }
    const giorni = (cambio - oggi) / 86400000;
    if (giorni <= 31) return { livello: "medio", testo: `Cambio listino previsto il ${cambioTxt}: possibile aumento` };
    return { livello: "fresco", testo: `Listino valido; prossimo cambio il ${cambioTxt}${dataTxt ? ` · aggiornato il ${dataTxt}` : ""}` };
  }
  if (!agg || Number.isNaN(agg.getTime())) return { livello: regola ? "vecchio" : "", testo: "Data di aggiornamento non indicata" };
  const mesi = mesiDa(agg);
  if (regola?.tipo === "richiesta" && num(regola.validitaMesi) > 0) {
    const v = num(regola.validitaMesi);
    if (mesi >= v) return { livello: "vecchio", testo: `Prezzo a richiesta scaduto: aggiornato il ${dataTxt}, validità ${v} mesi` };
    if (mesi >= v - 1) return { livello: "medio", testo: `Prezzo a richiesta in scadenza: aggiornato il ${dataTxt}, validità ${v} mesi` };
    return { livello: "fresco", testo: `Prezzo aggiornato il ${dataTxt}, validità ${v} mesi` };
  }
  // senza regola: legenda dell'Excel (meno di 1 mese, 1–2 mesi, oltre 2 mesi)
  return { livello: mesi < 1 ? "fresco" : mesi <= 2 ? "medio" : "vecchio", testo: `Prezzo aggiornato il ${dataTxt}` };
}
