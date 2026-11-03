// Motore di calcolo: replica le formule dei fogli TIPO OFFERTA e Sheet1 dell'Excel
// "Template_Import_Righe-OFFERTA_V22". Nessuna dipendenza: si usa sia nell'app sia nei test.

export const TIPI = ["ACQUISTO", "LISTINO", "VENDITA CLIENTE", "LISTINO COMPUTO"];
export const FONTI = ["DA ORACLE", "DA EXCEL"];
export const CAMPI_SCONTO = ["sc3", "sc4", "sc5", "aum1"]; // colonne Q R S T del catalogo
export const ETICHETTE = { sc3: "Sconto 3", sc4: "Sconto 4", sc5: "Sconto 5", aum1: "Aumento 1", ric: "Ricarico" };

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
export const isListino = (tipo) => tipo === "LISTINO" || tipo === "LISTINO COMPUTO";
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
  return num(r.listinoAcquisto) * (1 - num(r.sc3)) * (1 - num(r.sc4)) * (1 - num(r.sc5)) * (1 - num(r.aum1)) / (1 - P);
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

/* Colonna M di TIPO OFFERTA: prezzo di partenza della riga in base al tipo di offerta. */
export function prezzoBase(r, tipo) {
  if (isAttivita(r)) {
    if (tipo === "ACQUISTO") return num(r.listinoAcquisto);
    if (isListino(tipo)) return prezzoListino(r);
    if (tipo === "VENDITA CLIENTE") return num(r.listinoAcquisto) * num(r.giorni);
    return num(r.listinoAcquisto);
  }
  if (tipo === "LISTINO COMPUTO") return prezzoListino(r);
  return num(r.listinoAcquisto);
}

/* Colonna H di TIPO OFFERTA: per l'attività tecnica in ACQUISTO la quantità sono i giorni. */
export function quantita(r, tipo) {
  if (isAttivita(r) && tipo === "ACQUISTO") return num(r.giorni);
  return num(r.qta);
}

/* Colonne O P Q R di TIPO OFFERTA. Priorità:
   sconto del GMV → sconto globale → 0 per offerte a listino → 0 per l'attività tecnica → valore del catalogo. */
export function sconto(campo, r, off) {
  const perGmv = num(off.gmv?.[r.gmv]?.[campo]);
  if (perGmv) return perGmv;
  const globale = num(off.globali?.[campo]);
  if (globale) return globale;
  if (isListino(off.tipo)) return 0;
  if (isAttivita(r)) return 0;
  return num(r[campo]);
}

/* Colonna S di TIPO OFFERTA (Aumento 2): ricarico, solo per VENDITA CLIENTE.
   Ricarico del GMV → ricarico dell'offerta → ricarico automatico del capitolo. */
export function ricarico(r, off, ricaricoCapitolo = 0) {
  if (off.tipo !== "VENDITA CLIENTE") return 0;
  const perGmv = num(off.gmv?.[r.gmv]?.ric);
  if (perGmv) return perGmv;
  const globale = num(off.globali?.ric);
  if (globale) return globale;
  return num(ricaricoCapitolo);
}

/* Calcolo completo di una riga. ricarichi = { NOME_CAPITOLO: 0.28, ... } */
/* Attività "a giornate": in vendita e a listino il prezzo dell'attività è complessivo (quantità 1);
   con questa scelta va in Oracle come valore giornaliero × numero di giorni. In acquisto è sempre così. */
export const aGiornate = (r, tipo) => isAttivita(r) && !!r.aGiornate && tipo !== "ACQUISTO" && num(r.giorni) > 0;

export function calcolaRiga(r, off, ricarichi = {}) {
  const tipo = off.tipo;
  const giornate = aGiornate(r, tipo);
  const base = giornate ? prezzoBase(r, tipo) / num(r.giorni) : prezzoBase(r, tipo);
  const sc3 = sconto("sc3", r, off), sc4 = sconto("sc4", r, off), sc5 = sconto("sc5", r, off);
  const aum1 = sconto("aum1", r, off);
  const ric = ricarico(r, off, ricarichi[r.capitolo]);
  const fattore = (1 - sc3) * (1 - sc4) * (1 - sc5) * (1 + aum1) * (1 + ric);
  const unitario = base * fattore;
  const qta = giornate ? num(r.giorni) * num(r.qta) : quantita(r, tipo);
  return { base, sc3, sc4, sc5, aum1, ric, fattore, unitario, qta, totale: unitario * qta };
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
   "off" è un capitolo dell'offerta: { tipo, prezzi, globali, gmv, righe, ... }. */
export function calcolaOfferta(off, { ricarichi = {} } = {}) {
  const righe = (off.righe || []).filter((r) => num(r.qta) !== 0)
    .map((r) => ({ riga: r, c: calcolaRiga(r, off, ricarichi) }));
  const totale = righe.reduce((s, x) => s + x.c.totale, 0);
  return { righe, totale };
}

/* Tracciato Sheet1 per Oracle. Restituisce { righe: [array di 18 valori], avvisi: [...] }. */
/* Capitolati esclusi dall'esportazione (per caricarli in un'altra offerta Oracle). "" = righe senza capitolato. */
export const esclusoDaExport = (r, off) => (off.esclusiExport || []).includes(r.capitolato || "");

export function righeOracle(off, opz = {}) {
  // solo i capitolati spuntati; le linee ripartono dal n° di partenza
  const { righe } = calcolaOfferta({ ...off, righe: (off.righe || []).filter((r) => !esclusoDaExport(r, off)) }, opz);
  const daExcel = off.prezzi === "DA EXCEL";
  const start = num(off.lineaPartenza) || 10;
  const out = [], avvisi = [];
  let linea = start;
  for (const { riga: r, c } of righe) {
    const codice = codiceOracle(r);
    if (!codice) {
      avvisi.push(`Riga senza codice articolo esclusa: ${r.nOrdine || r.descrizione || "(senza descrizione)"} (${r.capitolo})`);
      continue;
    }
    // Prezzo netto = unitario finale del tipo di offerta scelto (acquisto, listino, vendita…).
    // Quando c'è il prezzo netto Oracle ignora sconti e aumenti, quindi le percentuali restano vuote.
    const conNetto = codice === "PITATT.TEC" || isGenerico(r) || daExcel;
    const prezzoNetto = conNetto ? arrot(c.unitario) : "";
    const perc = (v) => (v === 0 || conNetto ? "" : arrot(v * 100));
    out.push([
      linea, codice, off.magazzino || "BG", arrot(c.qta),
      off.sottocapitoli ? (r.capitolato || "") : "", // capitolato dell'offerta, non il sottocapitolo del database
      prezzoNetto,
      perc(c.sc3), perc(c.sc4), perc(c.sc5), "", "", "",
      perc(c.aum1), perc(c.ric), "", "",
      notePreRiga(r) !== "0" ? notePreRiga(r).trim() : "",
      notePostRiga(r) !== "0" ? notePostRiga(r).trim() : "",
    ]);
    linea += 10;
  }
  return { righe: out, avvisi };
}

/* Controllo cartellino cliente (foglio SC_CLIENTI e colonna K di TIPO OFFERTA).
   Per ogni GMV dell'offerta confronta lo sconto complessivo con quello del cartellino del cliente;
   per l'attività tecnica confronta il prezzo a giornata con la tariffa del cliente
   (l'Excel confrontava il prezzo di tutti i giorni con la tariffa giornaliera). */
export function controlloCartellino(off, cliente, opz = {}) {
  if (!cliente) return [];
  const { righe } = calcolaOfferta(off, opz);
  const visti = new Map();
  for (const x of righe) {
    const g = x.riga.gmv || "";
    if (!g || visti.has(g)) continue; // come CERCA.VERT: vale la prima riga del GMV
    visti.set(g, x);
  }
  const esiti = [];
  for (const [gmv, x] of visti) {
    if (gmv === "ATT" || isAttivita(x.riga)) {
      const att = righe.find((y) => isAttivita(y.riga)) || x;
      const tariffa = cliente.att === undefined || cliente.att === "" ? null : num(cliente.att);
      if (tariffa === null) continue;
      // tariffa giornaliera: in LISTINO e VENDITA CLIENTE il prezzo dell'attività comprende tutti i giorni
      const gg = num(att.riga.giorni);
      const prezzo = off.tipo !== "ACQUISTO" && gg > 0 && !aGiornate(att.riga, off.tipo) ? att.c.unitario / gg : att.c.unitario;
      const stato = tariffa > prezzo + 1e-9 ? "ok" : Math.abs(tariffa - prezzo) < 1e-9 ? "pari" : "fuori";
      esiti.push({ gmv, tipo: "attivita", offerta: prezzo, cartellino: tariffa, stato,
        testo: stato === "ok" ? "OK" : stato === "pari" ? "Valore di cartellino" : "Superiore al valore di cartellino" });
      continue;
    }
    const scontoOfferta = 1 - x.c.fattore;
    const scCart = num(cliente.sconti?.[gmv]);
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

/* Il costo (acquistato): le stesse righe calcolate come ACQUISTO, con sconti ed extrasconti. */
export function capitoloCosto(cap) {
  // gli extrasconti valgono sia per l'acquisto sia per la vendita
  return { ...cap, tipo: "ACQUISTO" };
}

/* Riepilogo economico di un capitolo (ex foglio ANALISI PACCHETTO).
   Venduto = totale del capitolo con il suo tipo di offerta.
   Acquistato = le stesse righe calcolate come ACQUISTO con gli sconti del database e gli extrasconti
   del fornitore (il costo vero, senza ricarico). */
export function riepilogoCapitolo(cap, opz = {}) {
  const v = calcolaOfferta(cap, opz);
  const a = calcolaOfferta(capitoloCosto(cap), opz);
  const out = {
    righe: v.righe.length, venduto: 0, acquistato: 0,
    attivita: { venduto: 0, acquistato: 0 }, materiale: { venduto: 0, acquistato: 0 },
    perGmv: new Map(),
  };
  v.righe.forEach((x, i) => {
    const costo = a.righe[i].c.totale;
    const vend = x.c.totale;
    out.venduto += vend; out.acquistato += costo;
    const parte = isAttivita(x.riga) ? out.attivita : out.materiale;
    parte.venduto += vend; parte.acquistato += costo;
    const g = x.riga.gmv || "";
    if (!out.perGmv.has(g)) out.perGmv.set(g, { righe: 0, venduto: 0, acquistato: 0, prima: x });
    const pg = out.perGmv.get(g);
    pg.righe++; pg.venduto += vend; pg.acquistato += costo;
  });
  return out;
}
export const ricaricoPct = (venduto, acquistato) => (acquistato ? venduto / acquistato - 1 : 0);
export const marginePct = (venduto, acquistato) => (venduto ? 1 - acquistato / venduto : 0);

/* Ricarico dell'offerta che porta il capitolo (VENDITA CLIENTE) al prezzo target.
   Le righe con un ricarico forzato sul GMV restano come sono. Restituisce null se non è possibile. */
export function ricaricoPerTarget(cap, target, opz = {}) {
  if (cap.tipo !== "VENDITA CLIENTE") return null;
  const gmvSenzaRic = Object.fromEntries(Object.entries(cap.gmv || {}).map(([g, v]) => [g, { ...v, ric: 0 }]));
  const costo = calcolaOfferta({ ...cap, tipo: "VENDITA CLIENTE", globali: { ...cap.globali, ric: 0 }, gmv: gmvSenzaRic },
    { ...opz, ricarichi: {} });
  let fisso = 0, variabile = 0;
  for (const x of costo.righe) {
    const ricGmv = num(cap.gmv?.[x.riga.gmv]?.ric);
    if (ricGmv) fisso += x.c.totale * (1 + ricGmv); else variabile += x.c.totale;
  }
  if (!variabile) return null;
  const ric = (num(target) - fisso) / variabile - 1;
  return Number.isFinite(ric) ? ric : null;
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
