# Offerte Oracle

App web che sostituisce il file **Template_Import_Righe-OFFERTA_V22.xlsm**. Con l'app puoi:

- scegliere gli articoli dal catalogo e scrivere le quantità;
- applicare tipo offerta, sconti, ricarico e sconti per GMV;
- controllare il cartellino del cliente;
- scaricare il foglio **Sheet1** da importare in Oracle.

L'app gira nel browser, si pubblica su GitHub Pages e salva i dati su Firebase (Firestore). Si installa sul Mac e sul telefono come un'app.

## Cosa fa rispetto all'Excel

| Excel | App |
|---|---|
| Quantità scritte nei fogli catalogo (il file si "sporca") | Le quantità stanno nell'offerta; il catalogo resta pulito |
| Un capitolo per volta; cambiando capitolo si perdevano gli extra impostati | Più capitoli nella stessa offerta, ognuno con condizioni salvate e foglio Oracle proprio |
| Righe copiate a mano in testa al foglio e cancellate a fine offerta | Righe dell'offerta separate dal database, ripetibili e duplicabili per posizione |
| ANALISI PACCHETTO compilata a mano | Riepilogo automatico per capitolo (venduto, acquistato, margine) e peso dei GMA |
| Max ~486 righe, max 1000 righe per catalogo | Nessun limite pratico |
| Macro "allinea" e controllo "Elenco allineato" | Non servono: l'elenco per Oracle è sempre allineato |
| Un file per offerta | Tutte le offerte salvate, duplicabili, condivise tra i dispositivi |

**Pagina Prezzi** di ogni capitolo, con due blocchi che restano salvati tutti e due:

- **Condizioni di acquisto:** una riga Generale e una riga per ogni GMA del capitolo (quella del GMA vince).
  - La prima cella è lo **standard** del database: se la cambi vale al posto dello standard (`0` = niente standard, prezzo di listino); l'intestazione "Standard" si barra in rosso e accanto resta lo standard di prima, barrato.
  - Nelle celle accanto gli **extra** in cascata: compilata una, ne compare un'altra.
  - In una cella `15` o `-15` è uno sconto, `+5` è un aumento.
  - Una riga GMA vuota prende la riga Generale.
- **Condizioni di vendita: Ricarico / Listino-sconto.**
  - Ricarico: vendita = acquisto × (1 + ricarico). La prima cella è il ricarico automatico del capitolo (modificabile), accanto i ricarichi extra in cascata.
  - Listino-sconto: vendita = listino acquisto con lo sconto al cliente, con le stesse celle dell'acquisto (standard modificabile + extra).
  - Anche qui c'è una riga Generale e una per GMA; sotto i due riquadri c'è il riepilogo per GMA (acquistato, venduto, peso, ricarico, margine, cartellino).
  - Attività tecnica: parte dal prezzo di listino di tutti i giorni.
- **Attività tecnica:** il generale vale solo se è "in aggiunta"; per il resto cambia solo dal suo GMA.
- **Prezzo target:** calcola il ricarico generale (o lo sconto generale) per arrivare al totale voluto.

**Pagina Oracle:**

- **Prezzi da esportare:** Vendita o Acquisto.
- **Come:**
  - Prezzi netti: si compila solo "Prezzo Netto".
  - Sconti: si compila solo "Sconto 3 [%]", con un unico sconto sul listino acquisto che somma tutti gli sconti e i ricarichi. È negativo se il prezzo supera il listino.
- **Sempre a prezzo netto:** attività tecnica (PITATT.TEC) e generici (P000ARTICOLO_GEN).
- **Linea:** 10, 20, 30…
- **Codice:** PRECODICE+CODICE.
- **Magazzino:** BG.

Le offerte salvate con il vecchio modello (tipo offerta, DA ORACLE/DA EXCEL) si convertono da sole all'apertura.

Unica differenza voluta: il controllo cartellino dell'attività tecnica confronta il **prezzo a giornata** con la tariffa del cliente. L'Excel invece usava il prezzo di tutti i giorni.

## Come sono organizzati i dati (Firestore)

| Collezione | Contenuto |
|---|---|
| `oracle_catalogo` | Articoli, a blocchi per capitolo (un documento contiene centinaia di articoli: si consumano poche letture) |
| `oracle_config` | `capitoli` (nome, ricarico automatico, ordine) e `gmv` (codice, fornitore) |
| `oracle_clienti` | Ragione sociale, codice, tariffa attività tecnica, sconti di cartellino per GMV |
| `oracle_offerte` | Testata dell'offerta e i suoi capitoli; ogni capitolo ha condizioni proprie e le righe (con una copia dei dati dell'articolo) |

Le righe dell'offerta conservano una copia del prezzo e degli sconti. Così, se il catalogo cambia, le offerte già fatte non si muovono. Il pulsante **Aggiorna prezzi dal catalogo** le riallinea quando vuoi.

## 1. Prova subito, senza Firebase

Finché `firebase-config.js` contiene `INCOLLA_QUI`, l'app funziona in **modalità prova**: i dati restano solo nel browser.

Dalla cartella del progetto, apri il Terminale e lancia:

```
python3 -m http.server 8000
```

Poi apri http://localhost:8000. Vai in **Importa e aggiorna**, carica il file Excel e spunta "Crea un'offerta per ogni capitolo con le quantità già scritte nel file": ritrovi le offerte che avevi nell'Excel.

Non funziona aprendo `index.html` con doppio clic: serve un indirizzo web.

## 2. Configurazione di Firebase

Consiglio un **progetto Firebase nuovo** (es. `offerte-oracle`), separato da quello delle ville.

1. **Crea il progetto:** vai su https://console.firebase.google.com, clic su **Aggiungi progetto**, nome `offerte-oracle`. Google Analytics non serve.
2. **Firestore:** menu **Build → Firestore Database → Crea database**. Scegli la località `europe-west`, poi avvia in modalità produzione.
3. **Regole:** in Firestore → **Regole**, incolla il contenuto di `firestore.rules` e fai clic su **Pubblica**. Nell'elenco c'è la tua email. Per un collega, aggiungi la sua email e ripubblica.
4. **Accesso:** in **Build → Authentication → Inizia → Metodo di accesso**, attiva **Google**.
5. **Configurazione:**
   - in **⚙️ Impostazioni progetto → Le tue app**, aggiungi un'app **Web** (icona `</>`), senza Hosting;
   - copia l'oggetto `firebaseConfig` nel file `firebase-config.js`, al posto di quello con `INCOLLA_QUI`.

Il piano gratuito basta: il catalogo intero sono circa 15 letture per apertura.

> Se preferisci usare lo stesso progetto Firebase delle app ville, non sostituire le regole esistenti: aggiungi solo le quattro righe `match /oracle_...` dentro le regole che ci sono già.

## 3. Pubblicazione su GitHub Pages

1. **Crea il repository:** su GitHub crea un repository (es. `offerte-oracle`) e carica tutti i file di questa cartella (anche `.nojekyll`).
2. **Attiva Pages:** nel repository vai in **Settings → Pages**. Come Source scegli *Deploy from a branch*, branch `main`, cartella `/ (root)`, poi **Save**.
3. **Indirizzo:** dopo un minuto l'app è su `https://TUO-UTENTE.github.io/offerte-oracle/`.
4. **Dominio in Firebase:** in **Authentication → Settings → Domini autorizzati**, aggiungi `TUO-UTENTE.github.io`.

Il repository può essere pubblico: `firebase-config.js` non è un segreto. I dati sono protetti dal login Google e dalle regole.

**Installazione:**

- Mac con Chrome o Edge: clic sull'icona di installazione nella barra dell'indirizzo.
- Mac con Safari: **File → Aggiungi al Dock**.
- iPhone: apri il link in Safari, poi **Condividi → Aggiungi alla schermata Home**.
- Android con Chrome: menu ⋮ → **Installa app**.

## Uso

1. **Importa e aggiorna**: carica il file Excel. Si leggono database articoli (un database per foglio capitolo), capitoli con ricarico automatico, GMV e cartellini clienti.
   Le righe del **blocco 0** (quelle copiate in testa al foglio per l'offerta in corso) non entrano nel database se l'articolo c'è già più sotto.
   Spuntando l'ultima casella si crea un'offerta con le quantità del file, un capitolo per foglio.
2. **Offerte → + Nuova offerta**: riferimento e cliente, poi **+ Aggiungi capitolo** (FIRE, KNX, ANTINTRUSIONE…). Un'offerta può avere tutti i capitoli che vuoi.
   Ogni capitolo ha tre pagine: **1 · Righe** (aggiungere articoli, quantità, ordine, capitolati, note),
   **2 · Sconti e ricarichi** (tipo offerta, extrasconti e ricarico, prezzo target, sconti per GMA con peso e cartellino, riepilogo righe)
   e **3 · Esporta per Oracle** (prezzi DA ORACLE/DA EXCEL, capitolati, n° partenza righe, magazzino, foglio Oracle).
   Gli extrasconti valgono sempre sia per l'acquisto sia per la vendita.
3. **Ogni capitolo è a sé stante**: tipo offerta, prezzi (DA ORACLE / DA EXCEL), sottocapitoli, n° partenza righe, extrasconti, ricarico, sconti per GMV.
   Restano salvati: passando da un capitolo all'altro non si perde niente.
   - *Prezzo target*: in VENDITA CLIENTE scrivi il totale che vuoi ottenere e l'app calcola il ricarico del capitolo.
4. **Aggiungi articoli**: scegli il database (di solito quello del capitolo), cerca, scrivi la quantità e premi Invio o Tab: ogni volta si aggiunge una riga in fondo.
   Con **Nota post-linea per le nuove righe** (es. "posizione A_1") le righe aggiunte prendono quella nota (se l'articolo non ne ha già una nel database).
5. **Capitolati per Oracle** (colonna "Nome Capitolato"): li crei offerta per offerta, per ogni capitolo (es. "PIANO TERRA", "PRIMO PIANO"),
   dalla barra "Capitolati per Oracle" nelle Righe o con "+ nuovo capitolato…" in Aggiungi articoli, e li assegni alle righe.
   Vanno in Oracle se "Capitolati in Oracle" è SI. Non c'entrano con i **sottocapitoli del database** (Sistema, Centrale…),
   che servono solo a organizzare gli articoli e si creano dal Catalogo scrivendo un nome nuovo nella scheda dell'articolo.
   Note: per gli ARTICOLO_GEN sono automatiche (pre-linea = n° di ordinazione, post-linea = descrizione); per gli altri sono testo libero.
6. **Righe**: l'ordine è quello che vedi e va così in Oracle. Lo stesso articolo può comparire più volte.
   Seleziona delle righe per **duplicarle in fondo con una nuova nota** (es. da "posizione A_1" a "posizione A_2"), cambiare la nota, spostarle o toglierle. ↑ ↓ ⧉ ✎ su ogni riga.
7. **GMV**: per ogni GMA del capitolo vedi venduto, acquistato, **peso** sul capitolo e margine; puoi forzare sconti e ricarico e vedi il confronto con il cartellino del cliente.
8. **Esporta per Oracle**: un foglio "Sheet1" per capitolo (Excel, CSV o copia). Dal **Riepilogo** scarichi quelli di tutti i capitoli.
9. **Riepilogo** (ex ANALISI PACCHETTO): per capitolo venduto, acquistato, materiale, attività, ricarico e margine, con i totali dell'offerta.
10. **Database**: clicca un articolo per modificarlo. Se cambi un prezzo o uno sconto la data di aggiornamento diventa quella di oggi.
   - **Tratta come articolo generico**: l'articolo tiene precodice e codice veri, ma in Oracle va come P000ARTICOLO_GEN con prezzo netto,
     n° di ordinazione nella nota pre-linea e descrizione nella post-linea. Nelle righe dell'offerta il suo prezzo di acquisto si può cambiare
     direttamente: si aggiorna anche nel database con la data di oggi.
   - **Scadenze prezzi** (per precodice): "prezzi a richiesta" con validità in mesi, oppure "listino" con la data del prossimo cambio
     (avviso un mese prima, "da aggiornare" dopo il cambio). Il pallino colorato e gli avvisi in "Esporta per Oracle" seguono queste regole.
   **Aggiorna i prezzi da Oracle**: incolla la griglia di Oracle (con intestazioni) e l'app aggiorna listino acquisto, GMV e data.
11. **Copia di sicurezza**: scarichi tutti i dati in un file `.json` e puoi ripristinarli.

## File

| File | A cosa serve |
|---|---|
| `index.html`, `css/style.css` | Pagina e stile |
| `js/app.js` | Interfaccia |
| `js/calc.js` | Calcoli prezzi, tracciato Oracle, controllo cartellino (le formule dell'Excel) |
| `js/importa.js` | Lettura del file Excel |
| `js/xlsx-lite.js` | Lettura e scrittura dei file Excel, senza librerie esterne |
| `js/firebase.js` | Firestore, login Google, modalità prova |
| `firebase-config.js` | La configurazione del tuo progetto Firebase |
| `firestore.rules` | Le regole di sicurezza da incollare in Firestore |
