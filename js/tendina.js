/* Tendina chiara al posto dell'elenco nativo dei suggerimenti (<input list="...">), che nei browser
   esce scuro e diverso dagli altri menu. Il campo resta un normale input: si può sempre scrivere un valore nuovo.
   Funziona su tutti i campi con "list", anche quelli creati dopo (finestre, tabelle). */

let pop = null, campo = null, voci = [], attiva = -1;

function opzioniDi(inp) {
  const dl = document.getElementById(inp.dataset.dl);
  if (!dl) return [];
  return [...dl.querySelectorAll("option")].map((o) => ({ v: o.value, t: o.label && o.label !== o.value ? o.label : "" })).filter((o) => o.v !== "");
}

function chiudi() {
  if (pop) pop.hidden = true;
  campo = null; voci = []; attiva = -1;
}

function posiziona() {
  if (!campo || !pop || pop.hidden) return;
  const r = campo.getBoundingClientRect();
  if (!r.width) return chiudi();
  const sotto = innerHeight - r.bottom, sopra = r.top;
  const h = Math.min(pop.scrollHeight, 300);
  pop.style.left = Math.max(4, Math.min(r.left, innerWidth - Math.max(r.width, 200) - 4)) + "px";
  pop.style.minWidth = r.width + "px";
  pop.style.maxHeight = "300px";
  if (sotto < h + 8 && sopra > sotto) pop.style.top = Math.max(4, r.top - h - 2) + "px";
  else pop.style.top = r.bottom + 2 + "px";
}

function disegna(filtra) {
  const tutte = opzioniDi(campo);
  const q = filtra ? campo.value.trim().toLowerCase() : "";
  voci = q ? tutte.filter((o) => (o.v + " " + o.t).toLowerCase().includes(q)) : tutte;
  if (!voci.length) { pop.hidden = true; return; }
  const cur = campo.value.trim().toLowerCase();
  attiva = voci.findIndex((o) => o.v.toLowerCase() === cur);
  pop.innerHTML = voci.map((o, i) => `<div class="tnd-voce${i === attiva ? " on" : ""}" data-i="${i}">${esc(o.v)}${o.t ? `<span class="tnd-sub">${esc(o.t)}</span>` : ""}</div>`).join("");
  pop.hidden = false;
  posiziona();
  pop.querySelector(".on")?.scrollIntoView({ block: "nearest" });
}

const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

function scegli(i) {
  const o = voci[i]; if (!o || !campo) return;
  const inp = campo;
  inp.value = o.v;
  chiudi();
  inp.dispatchEvent(new Event("input", { bubbles: true }));
  inp.dispatchEvent(new Event("change", { bubbles: true }));
}

function evidenzia(i) {
  if (!voci.length) return;
  attiva = (i + voci.length) % voci.length;
  pop.querySelectorAll(".tnd-voce").forEach((d, j) => d.classList.toggle("on", j === attiva));
  pop.querySelector(".on")?.scrollIntoView({ block: "nearest" });
}

// trasforma <input list> in campo con tendina chiara (toglie "list" così il browser non mostra la sua)
function converti(inp) {
  if (!inp.getAttribute("list")) return;
  inp.dataset.dl = inp.getAttribute("list");
  inp.removeAttribute("list");
  inp.classList.add("tnd");
  inp.setAttribute("autocomplete", "off");
}

export function attivaTendine() {
  pop = document.createElement("div");
  pop.className = "tnd-pop"; pop.hidden = true;
  document.body.appendChild(pop);
  // dentro un <dialog> modale la tendina deve stare nel dialog per essere visibile sopra
  const ospita = () => { const d = campo?.closest("dialog"); const dove = d || document.body; if (pop.parentNode !== dove) dove.appendChild(pop); };

  document.querySelectorAll("input[list]").forEach(converti);
  new MutationObserver((ms) => {
    for (const m of ms) for (const n of m.addedNodes) {
      if (n.nodeType !== 1) continue;
      if (n.matches?.("input[list]")) converti(n);
      n.querySelectorAll?.("input[list]").forEach(converti);
    }
    for (const m of ms) if (m.type === "attributes" && m.target.matches?.("input[list]")) converti(m.target);
  }).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["list"] });

  const apri = (inp, filtra) => { campo = inp; ospita(); disegna(filtra); };
  document.addEventListener("mousedown", (e) => {
    if (pop.contains(e.target)) { e.preventDefault(); const v = e.target.closest(".tnd-voce"); if (v) scegli(+v.dataset.i); return; }
    const inp = e.target.closest?.("input.tnd");
    if (!inp) return chiudi();
    // clic sulla freccia a destra o su un campo già attivo: apre/chiude l'elenco completo
    if (campo === inp && !pop.hidden) return chiudi();
    setTimeout(() => apri(inp, false), 0);
  }, true);
  document.addEventListener("input", (e) => {
    const inp = e.target;
    if (!inp.classList?.contains("tnd") || !e.isTrusted) return;
    apri(inp, true);
  }, true);
  document.addEventListener("keydown", (e) => {
    const inp = e.target;
    if (!inp.classList?.contains("tnd")) return;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (campo !== inp || pop.hidden) return apri(inp, false);
      evidenzia(attiva + (e.key === "ArrowDown" ? 1 : -1));
    } else if (e.key === "Enter" && campo === inp && !pop.hidden && attiva >= 0) {
      e.preventDefault(); e.stopPropagation(); scegli(attiva);
    } else if (e.key === "Escape" && campo === inp && !pop.hidden) {
      e.preventDefault(); e.stopPropagation(); chiudi();
    } else if (e.key === "Tab") chiudi();
  }, true);
  document.addEventListener("focusout", (e) => { if (e.target === campo) setTimeout(() => { if (document.activeElement !== campo) chiudi(); }, 120); }, true);
  addEventListener("scroll", (e) => { if (campo && !pop.contains(e.target)) posiziona(); }, true);
  addEventListener("resize", posiziona);
}
