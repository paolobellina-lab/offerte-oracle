// Accesso ai dati: Firestore con login Google, oppure modalità prova (dati solo in questo browser)
// finché firebase-config.js non è compilato.
import { firebaseConfig } from "../firebase-config.js";

export const configurato = !String(firebaseConfig.apiKey || "").includes("INCOLLA_QUI") && !!firebaseConfig.projectId;
const FB = "https://www.gstatic.com/firebasejs/10.12.2/";

let fs = null, auth = null, db = null, authMod = null;

/* ---------- Modalità prova: archivio nel browser (IndexedDB) ---------- */
const prova = (() => {
  const ascoltatori = new Map(); // coll → Set(cb)
  let dbp;
  const apri = () => (dbp ??= new Promise((res, rej) => {
    const r = indexedDB.open("offerte-oracle-prova", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("doc");
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  }));
  const tx = async (modo, fn) => {
    const d = await apri();
    return new Promise((res, rej) => {
      const t = d.transaction("doc", modo);
      const out = fn(t.objectStore("doc"));
      t.oncomplete = () => res(out?.result ?? out);
      t.onerror = () => rej(t.error);
    });
  };
  async function elenco(coll) {
    const d = await apri();
    return new Promise((res, rej) => {
      const out = [];
      const t = d.transaction("doc", "readonly");
      const range = IDBKeyRange.bound(coll + "/", coll + "/￿");
      const c = t.objectStore("doc").openCursor(range);
      c.onsuccess = () => {
        const cur = c.result;
        if (cur) { out.push({ ...cur.value, id: cur.key.slice(coll.length + 1) }); cur.continue(); } else res(out);
      };
      c.onerror = () => rej(c.error);
    });
  }
  const avvisa = async (coll) => {
    const set = ascoltatori.get(coll);
    if (!set) return;
    const dati = await elenco(coll);
    set.forEach((cb) => cb(dati));
  };
  return {
    async scrivi(ops) {
      await tx("readwrite", (st) => {
        for (const o of ops) {
          const key = o.coll + "/" + o.id;
          if (o.tipo === "del") st.delete(key);
          else st.put(JSON.parse(JSON.stringify(o.data)), key);
        }
      });
      for (const c of new Set(ops.map((o) => o.coll))) avvisa(c);
    },
    ascolta(coll, cb) {
      if (!ascoltatori.has(coll)) ascoltatori.set(coll, new Set());
      ascoltatori.get(coll).add(cb);
      elenco(coll).then(cb);
      return () => ascoltatori.get(coll).delete(cb);
    },
  };
})();

/* ---------- Avvio e login ---------- */
export async function avvia({ onLogin, onLogout, onVerifica, onNegato, onErrore }) {
  if (!configurato) {
    onLogin({ email: "modalità prova", prova: true });
    return;
  }
  const [appMod, a, f] = await Promise.all([
    import(FB + "firebase-app.js"), import(FB + "firebase-auth.js"), import(FB + "firebase-firestore.js"),
  ]);
  authMod = a; fs = f;
  const app = appMod.initializeApp(firebaseConfig);
  auth = a.getAuth(app);
  // Cache locale: le letture ripetute non consumano la quota e l'app regge brevi cadute di rete.
  db = f.initializeFirestore(app, { localCache: f.persistentLocalCache({ tabManager: f.persistentMultipleTabManager() }) });
  // Dopo il login Google si entra nell'app solo se le regole di Firestore riconoscono l'account:
  // un account non abilitato viene disconnesso subito e non vede nessuna schermata dell'app.
  a.onAuthStateChanged(auth, async (u) => {
    if (!u) return onLogout();
    onVerifica?.(u);
    try {
      await f.getDocFromServer(f.doc(db, "oracle_config", "accesso"));
      onLogin(u);
    } catch (e) {
      const email = u.email;
      await a.signOut(auth);
      if (e.code === "permission-denied") onNegato?.(email);
      else onErrore?.(new Error("impossibile verificare l'account (" + (e.code || e.message) + "). Controlla la connessione e riprova."));
    }
  });
  try { await a.getRedirectResult(auth); } catch (e) { onErrore?.(e); }
}

export async function accediGoogle() {
  const provider = new authMod.GoogleAuthProvider();
  provider.setCustomParameters({ prompt: "select_account" });
  try {
    await authMod.signInWithPopup(auth, provider);
  } catch (x) {
    if (/popup-blocked|operation-not-supported/.test(x.code || "")) return authMod.signInWithRedirect(auth, provider);
    if (!/popup-closed|cancelled-popup/.test(x.code || "")) throw x;
  }
}

export const esci = () => (auth ? authMod.signOut(auth) : location.reload());

/* ---------- Dati ---------- */
/* Ascolta una collezione: cb riceve un array di {id, ...campi}. */
export function ascolta(coll, cb, onErrore) {
  if (!configurato) return prova.ascolta(coll, cb);
  return fs.onSnapshot(fs.collection(db, coll),
    (snap) => cb(snap.docs.map((d) => ({ ...d.data(), id: d.id }))),
    (e) => onErrore?.(e, coll));
}

/* Scritture in blocco: ops = [{tipo: "set"|"del", coll, id, data}] (spezzate a 400 per volta). */
export async function scrivi(ops) {
  if (!ops.length) return;
  if (!configurato) return prova.scrivi(ops);
  for (let i = 0; i < ops.length; i += 400) {
    const b = fs.writeBatch(db);
    for (const o of ops.slice(i, i + 400)) {
      const ref = fs.doc(db, o.coll, o.id);
      if (o.tipo === "del") b.delete(ref); else b.set(ref, pulisci(o.data));
    }
    await b.commit();
  }
}

export const salva = (coll, id, data) => scrivi([{ tipo: "set", coll, id, data }]);
export const elimina = (coll, id) => scrivi([{ tipo: "del", coll, id }]);

// Firestore non accetta undefined
const pulisci = (o) => JSON.parse(JSON.stringify(o));

export const nuovoId = () =>
  Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
