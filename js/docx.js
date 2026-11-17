/* Informativa commerciale in Word (.docx), compilata a partire dal modello modelli/informativa.docx.
   Per ora si compilano i capitoli 6 (KNX), 7 (impianti speciali), 9 (HVAC - BMS), gli allegati e il riepilogo economico;
   il resto del documento resta com'è nel modello. Nessuna libreria esterna: zip letto/scritto qui. */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/* ---------- zip ---------- */
async function inflateRaw(bytes) {
  const ds = new DecompressionStream("deflate-raw");
  return new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(ds)).arrayBuffer());
}
async function leggiZip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(buf), dec = new TextDecoder();
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("Modello Word non valido.");
  const n = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const file = [];
  for (let k = 0; k < n; k++) {
    const metodo = dv.getUint16(p + 10, true), comp = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true), nome = dec.decode(u8.subarray(p + 46, p + 46 + nl));
    const lnl = dv.getUint16(off + 26, true), lel = dv.getUint16(off + 28, true), s = off + 30 + lnl + lel;
    const raw = u8.subarray(s, s + comp);
    file.push([nome, metodo === 0 ? raw.slice() : await inflateRaw(raw)]);
    p += 46 + nl + el + cl;
  }
  return file;
}
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
function scriviZip(file, tipo) {
  const enc = new TextEncoder(), parti = [], cen = [];
  let off = 0;
  for (const [nome, dati] of file) {
    const n = enc.encode(nome), d = typeof dati === "string" ? enc.encode(dati) : dati, crc = crc32(d);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(12, 0x21, true); lh.setUint32(14, crc, true); lh.setUint32(18, d.length, true); lh.setUint32(22, d.length, true); lh.setUint16(26, n.length, true);
    parti.push(new Uint8Array(lh.buffer), n, d);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(14, 0x21, true); ch.setUint32(16, crc, true); ch.setUint32(20, d.length, true); ch.setUint32(24, d.length, true);
    ch.setUint16(28, n.length, true); ch.setUint32(42, off, true);
    cen.push(new Uint8Array(ch.buffer), n);
    off += 30 + n.length + d.length;
  }
  const dim = cen.reduce((s, x) => s + x.length, 0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, file.length, true); e.setUint16(10, file.length, true); e.setUint32(12, dim, true); e.setUint32(16, off, true);
  return new Blob([...parti, ...cen, new Uint8Array(e.buffer)], { type: tipo });
}

/* ---------- XML Word ---------- */
const tutti = (el, nome) => Array.from(el.getElementsByTagNameNS(W, nome));
const testo = (el) => tutti(el, "t").map((t) => t.textContent).join("");
function scrivi(el, s) {
  // tutto il testo nel primo w:t, gli altri vuoti
  const ts = tutti(el, "t");
  if (!ts.length) return;
  ts[0].textContent = s; ts[0].setAttribute("xml:space", "preserve");
  for (const t of ts.slice(1)) t.textContent = "";
}
const aggiungiInFondo = (el, s) => { const ts = tutti(el, "t"); if (ts.length) ts[ts.length - 1].textContent += s; };

/* dati: {
     rif, capitoli: { 6: [voce], 7: [...], 9: [...] }  voce = { sigla, nOff, dataOff, prezzo, attivita },
   } */
export async function informativaWord(modello, dati, euro) {
  const file = await leggiZip(modello);
  const iDoc = file.findIndex(([n]) => n === "word/document.xml");
  const doc = new DOMParser().parseFromString(new TextDecoder().decode(file[iDoc][1]), "application/xml");
  const body = tutti(doc, "body")[0];
  const figli = () => Array.from(body.childNodes).filter((n) => n.nodeType === 1);
  const isP = (n) => n.localName === "p", isT = (n) => n.localName === "tbl";
  const trova = (re, da = 0) => { const f = figli(); for (let i = da; i < f.length; i++) if (isP(f[i]) && re.test(testo(f[i]).trim())) return i; return -1; };
  const rif = dati.rif || "TE";
  const nomeOff = (v) => `OF_${rif}_${v.sigla}`;
  const descrOff = (v) => `${nomeOff(v)}${v.nOff ? ` (n° ${v.nOff}${v.dataOff ? ` del ${v.dataOff}` : ""})` : ""}`;
  const capTitoli = { 6: /^Capitolo 6\./, 7: /^Capitolo 7\./, 8: /^Capitolo 8\./, 9: /^Capitolo 9\./, 10: /^Capitolo 10\./ };

  // indice: il titolo dei capitoli assenti prende [NON PRESENTE] anche nel sommario
  for (const n of [6, 7, 9]) {
    if ((dati.capitoli[n] || []).length) continue;
    for (const p of tutti(body, "p")) {
      const t = testo(p).trim();
      if (capTitoli[n].test(t) && !/NON PRESENTE/.test(t)) {
        // nel sommario il testo del titolo è il primo w:t, nel titolo vero l'ultimo
        const ts = tutti(p, "t"), sommario = ts.length > 1 && /^\d+$/.test(ts[ts.length - 1].textContent.trim());
        if (sommario) ts[0].textContent += " [NON PRESENTE]"; else aggiungiInFondo(p, " [NON PRESENTE]");
      }
    }
  }

  // blocchi "Attività Tecnica …" del modello: chiave → regola sul titolo
  const BLOCCHI = [
    ["KNX", /^Attività Tecnica KNX/i], ["FIRE", /^Attività Tecnica FIRE/i],
    ["INTR+CA", /^Attività Tecnica ANTINTRUSIONE e CONTROLLO ACCESSI/i], ["INTR", /^Attività Tecnica ANTINTRUSIONE:/i],
    ["CA", /^Attività Tecnica CONTROLLO ACCESSI/i], ["TVCC", /^Attività Tecnica TVCC/i], ["EVAC", /^Attività Tecnica EVAC/i],
    ["BMS", /^Attività tecnica BMS/i], ["HVAC", /^Attività Tecnica HVAC/i],
  ];
  const conAtt = new Set([6, 7, 9].flatMap((n) => (dati.capitoli[n] || []).filter((v) => v.attivita).map((v) => v.sigla)));
  const tieni = new Set([...conAtt].filter((k) => k !== "INTR" && k !== "CA"));
  if (conAtt.has("INTR") && conAtt.has("CA")) tieni.add("INTR+CA"); else { if (conAtt.has("INTR")) tieni.add("INTR"); if (conAtt.has("CA")) tieni.add("CA"); }
  const fineCap = trova(/^Allegati\s*:/);
  for (const [k, re] of BLOCCHI) {
    const f = figli(), i = trova(re);
    if (i < 0 || i > fineCap || tieni.has(k)) continue;
    // il blocco va dal titolo fino al prossimo titolo di attività o capitolo (righe vuote comprese)
    let j = i + 1;
    while (j < f.length && !(isP(f[j]) && /^(Attività tecnica|Capitolo \d|Allegati)/i.test(testo(f[j]).trim())) && !isT(f[j])) j++;
    for (let x = i; x < j; x++) body.removeChild(f[x]);
  }

  // tabelle dei capitoli (una per capitolo, subito dopo il titolo): righe = capitoli dell'offerta
  for (const n of [6, 7, 9]) {
    const voci = dati.capitoli[n] || [];
    let f = figli(), i = trova(capTitoli[n]);
    if (i < 0) continue;
    let t = i + 1; while (t < f.length && !isT(f[t]) && !(isP(f[t]) && /^(Capitolo \d|Attività)/i.test(testo(f[t]).trim()))) t++;
    if (!isT(f[t])) continue;
    const tbl = f[t];
    if (!voci.length) {
      // capitolo non presente: via la tabella e le righe vuote attorno
      body.removeChild(tbl);
      continue;
    }
    const righe = tutti(tbl, "tr");
    const dati0 = righe.findIndex((r) => /Riferimento Offerta/i.test(testo(r)));
    const tot = righe.findIndex((r) => /TOTALE/i.test(testo(r)));
    if (dati0 < 0 || tot < 0) continue;
    const modelloRiga = righe[dati0];
    for (const r of righe.slice(dati0, tot)) r.parentNode.removeChild(r);
    for (const v of voci) {
      const r = modelloRiga.cloneNode(true), celle = tutti(r, "tc");
      scrivi(celle[0], v.sigla);
      scrivi(celle[1], `Riferimento Offerta ${descrOff(v)}`);
      scrivi(celle[celle.length - 1], euro(v.prezzo));
      righe[tot].parentNode.insertBefore(r, righe[tot]);
    }
    const ct = tutti(righe[tot], "tc");
    scrivi(ct[ct.length - 1], euro(voci.reduce((s, v) => s + v.prezzo, 0)));
  }

  // allegati: "Allegati Capitolo N …" + una riga per offerta
  for (const n of [6, 7, 9]) {
    const voci = dati.capitoli[n] || [], f = figli(), i = trova(new RegExp(`^Allegati Capitolo ${n}\\b`));
    if (i < 0) continue;
    const riga = f[i + 1];
    if (!voci.length) { body.removeChild(f[i]); if (riga && isP(riga) && /^OF_/.test(testo(riga).trim())) body.removeChild(riga); continue; }
    if (!riga || !isP(riga)) continue;
    for (const v of voci) { const r = riga.cloneNode(true); scrivi(r, `Offerta ${descrOff(v)}`); body.insertBefore(r, riga); }
    body.removeChild(riga);
  }

  // riepilogo costi: Capitolo 6 / 7 / 9 e totale generale
  const rc = tutti(body, "tbl").find((t) => /Riepilogo Costi/i.test(testo(t)));
  if (rc) {
    let totale = 0;
    for (const r of tutti(rc, "tr")) {
      const celle = tutti(r, "tc"), nC = celle.findIndex((c) => /^Capitolo \d+$/.test(testo(c).trim()));
      const ultimo = [...celle].reverse().find((c) => /\d,\d{2}/.test(testo(c)));
      if (nC >= 0 && ultimo) {
        const n = +testo(celle[nC]).trim().split(" ")[1];
        const v = (dati.capitoli[n] || []).reduce((s, x) => s + x.prezzo, 0);
        if ([6, 7, 9].includes(n)) scrivi(ultimo, euro(v));
        totale += [6, 7, 9].includes(n) ? v : 0;
      }
      if (/TOTALE GENERALE/i.test(testo(r)) && ultimo) scrivi(ultimo, euro(totale));
    }
  }

  file[iDoc][1] = new XMLSerializer().serializeToString(doc);
  return scriviZip(file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
}
