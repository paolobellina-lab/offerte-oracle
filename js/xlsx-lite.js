// Lettura e scrittura di file Excel (.xlsx / .xlsm) senza librerie esterne.
// Lettura: si usano i valori già calcolati salvati da Excel nel file (non si ricalcolano le formule).
// Scrittura: un foglio semplice, compatibile con Excel e con l'import di Oracle.

/* ---------------- ZIP ---------------- */
async function inflateRaw(bytes) {
  const ds = new DecompressionStream("deflate-raw");
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function leggiZip(buf) {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Il file non è un Excel valido (.xlsx / .xlsm).");
  const n = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const voci = new Map();
  const dec = new TextDecoder();
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const metodo = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const nome = dec.decode(u8.subarray(p + 46, p + 46 + nl));
    voci.set(nome, { metodo, compSize, off });
    p += 46 + nl + el + cl;
  }
  return {
    has: (nome) => voci.has(nome),
    async testo(nome) {
      const v = voci.get(nome);
      if (!v) return null;
      const lnl = dv.getUint16(v.off + 26, true), lel = dv.getUint16(v.off + 28, true);
      const start = v.off + 30 + lnl + lel;
      const raw = u8.subarray(start, start + v.compSize);
      const dati = v.metodo === 0 ? raw : await inflateRaw(raw);
      return new TextDecoder().decode(dati);
    },
  };
}

/* ---------------- Lettura ---------------- */
const xml = (s) => new DOMParser().parseFromString(s, "application/xml");
const figli = (el, nome) => Array.from(el.getElementsByTagName(nome));

function colIndex(ref) {
  let n = 0;
  for (const ch of ref) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

function testoSi(si) {
  // testo di una stringa condivisa, ignorando la guida fonetica (rPh)
  let s = "";
  for (const t of figli(si, "t")) {
    if (t.parentNode && t.parentNode.nodeName === "rPh") continue;
    s += t.textContent;
  }
  return s;
}

/* Restituisce { nomi: [...], foglio(nome) → righe (array di array, indice 0 = riga 1) } */
export async function leggiExcel(buf) {
  const zip = await leggiZip(buf);
  const wb = xml(await zip.testo("xl/workbook.xml"));
  const rels = xml(await zip.testo("xl/_rels/workbook.xml.rels"));
  const target = {};
  for (const r of figli(rels, "Relationship")) target[r.getAttribute("Id")] = r.getAttribute("Target");
  const fogli = figli(wb, "sheet").map((s) => {
    const rid = s.getAttribute("r:id") || s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    let t = target[rid] || "";
    t = t.startsWith("/") ? t.slice(1) : "xl/" + t.replace(/^\.\//, "");
    return { nome: s.getAttribute("name"), percorso: t };
  });
  let condivise = [];
  if (zip.has("xl/sharedStrings.xml")) {
    condivise = figli(xml(await zip.testo("xl/sharedStrings.xml")), "si").map(testoSi);
  }
  // colori di riempimento delle celle (per riconoscere le righe colorate, es. in rosso da cancellare)
  let fondoXf = [];
  if (zip.has("xl/styles.xml")) {
    const st = xml(await zip.testo("xl/styles.xml"));
    const fills = figli(st, "fill").map((f) => { const p = f.getElementsByTagName("patternFill")[0]; if (!p || (p.getAttribute("patternType") || "none") === "none") return ""; const fg = p.getElementsByTagName("fgColor")[0]; return fg?.getAttribute("rgb") || (fg?.getAttribute("theme") ? "tema" + fg.getAttribute("theme") : fg?.getAttribute("indexed") ? "idx" + fg.getAttribute("indexed") : "pieno"); });
    const xfs = st.getElementsByTagName("cellXfs")[0];
    if (xfs) fondoXf = figli(xfs, "xf").map((x) => fills[Number(x.getAttribute("fillId") || 0)] || "");
  }
  const cache = {};
  return {
    nomi: fogli.map((f) => f.nome),
    async foglio(nome) {
      if (cache[nome]) return cache[nome];
      const f = fogli.find((x) => x.nome === nome);
      if (!f) return null;
      const doc = xml(await zip.testo(f.percorso));
      const righe = [], fondi = [];
      for (const row of figli(doc, "row")) {
        const ri = Number(row.getAttribute("r")) - 1;
        const arr = [];
        for (const c of figli(row, "c")) {
          const ci = colIndex(c.getAttribute("r") || "");
          const fondo = fondoXf[Number(c.getAttribute("s") || 0)];
          if (fondo && !(fondi[ri] || "").includes(fondo)) fondi[ri] = (fondi[ri] ? fondi[ri] + "|" : "") + fondo;
          const t = c.getAttribute("t");
          const vEl = c.getElementsByTagName("v")[0];
          let v = vEl ? vEl.textContent : null;
          if (t === "s") v = condivise[Number(v)] ?? "";
          else if (t === "inlineStr") v = figli(c, "t").map((x) => x.textContent).join("");
          else if (t === "b") v = v === "1";
          else if (t === "e") v = { errore: v };
          else if (t === "str") v = v ?? "";
          else if (v !== null) v = Number(v);
          if (v !== null && v !== "") arr[ci] = v;
        }
        righe[ri] = arr;
      }
      for (let i = 0; i < righe.length; i++) if (!righe[i]) righe[i] = [];
      righe.fondi = fondi;
      cache[nome] = righe;
      return righe;
    },
  };
}

/* Data seriale di Excel → "AAAA-MM-GG" */
export function dataDaSeriale(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const ms = Math.round((v - 25569) * 86400000);
    return new Date(ms).toISOString().slice(0, 10);
  }
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  if (typeof v === "string" && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v)) {
    const [d, m, y] = v.split("/").map(Number);
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return "";
}

/* ---------------- Scrittura ---------------- */
const CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(u8) {
  let c = 0xffffffff;
  for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipStore(files) {
  const enc = new TextEncoder();
  const parti = [], centrale = [];
  let off = 0;
  for (const [nome, contenuto] of files) {
    const n = enc.encode(nome), d = enc.encode(contenuto), crc = crc32(d);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, 0, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, d.length, true); lh.setUint32(22, d.length, true);
    lh.setUint16(26, n.length, true); lh.setUint16(28, 0, true);
    parti.push(new Uint8Array(lh.buffer), n, d);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
    ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, d.length, true); ch.setUint32(24, d.length, true);
    ch.setUint16(28, n.length, true); ch.setUint32(42, off, true);
    centrale.push(new Uint8Array(ch.buffer), n);
    off += 30 + n.length + d.length;
  }
  const dimC = centrale.reduce((s, x) => s + x.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
  e.setUint32(12, dimC, true); e.setUint32(16, off, true);
  return new Blob([...parti, ...centrale, new Uint8Array(e.buffer)],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

const escXml = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
  // caratteri di controllo non ammessi in XML
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

function nomeCol(i) {
  let s = "";
  i++;
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

/* righe: array di array (la prima è l'intestazione). Celle vuote = "" o null. */
export function scriviExcel(righe, nomeFoglio = "Sheet1") {
  const rowsXml = righe.map((r, ri) => {
    const celle = r.map((v, ci) => {
      if (v === "" || v === null || v === undefined) return "";
      const ref = nomeCol(ci) + (ri + 1);
      const stile = ri === 0 ? ' s="1"' : "";
      if (typeof v === "number" && Number.isFinite(v)) return `<c r="${ref}"${stile}><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${stile}><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
    }).join("");
    return `<row r="${ri + 1}">${celle}</row>`;
  }).join("");
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${escXml(nomeFoglio)}" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`],
    ["xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`],
  ];
  return zipStore(files);
}

/* Foglio con stili (per il computo a listino).
   def = { nome, colonne: [larghezze], righe: [{ h, celle: [{ v, f, s }] }], unioni: ["A1:B1"], stili: [{ font:{b,sz,colore}, fondo, alH, alV, a_capo, fmt, bordo }] }
   v = valore, f = formula (senza "="), s = indice dello stile in def.stili (0 = normale). */
export function scriviExcelStile(def) {
  const fmts = [], fonts = ['<font><sz val="10"/><name val="Arial"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const idx = (arr, x) => { let i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
  for (const st of def.stili.slice(1)) {
    const font = `<font>${st.font?.b ? "<b/>" : ""}<sz val="${st.font?.sz || 10}"/>${st.font?.colore ? `<color rgb="FF${st.font.colore}"/>` : ""}<name val="Arial"/></font>`;
    const fill = st.fondo ? `<fill><patternFill patternType="solid"><fgColor rgb="FF${st.fondo}"/><bgColor indexed="64"/></patternFill></fill>` : fills[0];
    const border = st.bordo ? `<border><left style="${st.bordo}"><color auto="1"/></left><right style="${st.bordo}"><color auto="1"/></right><top style="${st.bordo}"><color auto="1"/></top><bottom style="${st.bordo}"><color auto="1"/></bottom><diagonal/></border>` : borders[0];
    let numId = 0;
    if (st.fmt) numId = 164 + idx(fmts, st.fmt);
    const al = `<alignment${st.alH ? ` horizontal="${st.alH}"` : ""}${st.alV ? ` vertical="${st.alV}"` : ""}${st.a_capo ? ' wrapText="1"' : ""}/>`;
    xfs.push(`<xf numFmtId="${numId}" fontId="${idx(fonts, font)}" fillId="${idx(fills, fill)}" borderId="${idx(borders, border)}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
  }
  const stili = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${fmts.length ? `<numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${escXml(f)}"/>`).join("")}</numFmts>` : ""}<fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills><borders count="${borders.length}">${borders.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const rowsXml = def.righe.map((r, ri) => {
    const celle = (r.celle || []).map((c, ci) => {
      if (!c) return "";
      const ref = nomeCol(ci) + (ri + 1), s = c.s ? ` s="${c.s}"` : "";
      if (c.f) return `<c r="${ref}"${s}><f>${escXml(c.f)}</f></c>`;
      const v = c.v;
      if (v === "" || v === null || v === undefined) return `<c r="${ref}"${s}/>`;
      if (typeof v === "number" && Number.isFinite(v)) return `<c r="${ref}"${s}><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${s}><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
    }).join("");
    return `<row r="${ri + 1}"${r.h ? ` ht="${r.h}" customHeight="1"` : ""}>${celle}</row>`;
  }).join("");
  const cols = def.colonne.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
  const unioni = def.unioni?.length ? `<mergeCells count="${def.unioni.length}">${def.unioni.map((u) => `<mergeCell ref="${u}"/>`).join("")}</mergeCells>` : "";
  const nome = String(def.nome || "Foglio1").replace(/[\[\]:*?/\\]/g, " ").slice(0, 31);
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${escXml(nome)}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", stili],
    ["xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${cols}</cols><sheetData>${rowsXml}</sheetData>${unioni}<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/></worksheet>`],
  ];
  return zipStore(files);
}
