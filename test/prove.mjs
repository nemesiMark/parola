// Prove automatiche di Word: apre l'app in Chrome senza finestra e usa davvero ogni funzione.
// Uso:  node test/prove.mjs            (serve Google Chrome e puppeteer-core)
// puppeteer-core si cerca in PUPPETEER_DIR oppure in ~/.claude/skills/sito-film/scripts/node_modules
import { createRequire } from "node:module";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PDIR = process.env.PUPPETEER_DIR || path.join(os.homedir(), ".claude/skills/sito-film/scripts/node_modules");
const puppeteer = createRequire(path.join(PDIR, "x.js"))("puppeteer-core");
const CHROME = process.env.CHROME || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "word-prove-"));

// --- piccolo server statico ---
const TYPES = { ".html": "text/html", ".css": "text/css", ".js": "text/javascript", ".svg": "image/svg+xml", ".png": "image/png" };
const server = http.createServer((req, res) => {
  if (req.url === "/vuoto") { res.writeHead(200, { "content-type": "text/html" }); res.end("<!doctype html><title>vuoto</title><link rel=icon href=icon.svg>"); return; }
  const p =path.join(ROOT, decodeURIComponent(new URL(req.url, "http://x").pathname).replace(/\/$/, "/index.html"));
  if (!p.startsWith(ROOT) || !fs.existsSync(p)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { "content-type": TYPES[path.extname(p)] || "application/octet-stream" });
  fs.createReadStream(p).pipe(res);
}).listen(0, "127.0.0.1");
await new Promise((r) => server.once("listening", r));
const URL0 = `http://127.0.0.1:${server.address().port}/`;

// --- risultati ---
let pass = 0, fail = 0;
const failures = [];
function ok(cond, name, extra = "") {
  if (cond) { pass++; console.log("  ✓ " + name); }
  else { fail++; failures.push(name); console.log("  ✗ " + name + (extra ? "  → " + extra : "")); }
}
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `atteso ${JSON.stringify(b)}, ottenuto ${JSON.stringify(a)}`);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const browser = await puppeteer.launch({ executablePath: CHROME, headless: true, args: ["--no-first-run", "--lang=it-IT"] });
const page = await browser.newPage();
await page.setViewport({ width: 1440, height: 900 });
const errors = [];
page.on("pageerror", (e) => errors.push(e.message));
page.on("console", (m) => { if (m.type() === "error") errors.push(m.text() + " " + (m.location()?.url || "")); });
page.on("dialog", (d) => d.accept());
const cdp = await page.createCDPSession();
await cdp.send("Browser.setDownloadBehavior", { behavior: "allow", downloadPath: TMP, eventsEnabled: true });

async function boot() {
  await page.goto(URL0, { waitUntil: "load" });
  await page.waitForFunction(() => window.WordDocx && window.docx && window.mammoth && window.__word);
}
// imposta il contenuto e mette il cursore all'inizio del primo blocco
async function setDoc(html) {
  await page.evaluate((h) => {
    const ed = document.getElementById("editor");
    ed.innerHTML = h;
    ed.focus();
    const sel = getSelection();
    sel.collapse(ed.firstElementChild || ed, 0);
    ed.dispatchEvent(new Event("input", { bubbles: true }));
  }, html);
  await sleep(60);
}
// seleziona il testo «needle» (n-esima occorrenza) nell'editor
async function select(needle, nth = 0) {
  const found = await page.evaluate((needle, nth) => {
    const ed = document.getElementById("editor");
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    let k = 0;
    while (w.nextNode()) {
      let i = w.currentNode.nodeValue.indexOf(needle);
      while (i !== -1) {
        if (k++ === nth) {
          const r = document.createRange();
          r.setStart(w.currentNode, i); r.setEnd(w.currentNode, i + needle.length);
          ed.focus();
          getSelection().removeAllRanges(); getSelection().addRange(r);
          return true;
        }
        i = w.currentNode.nodeValue.indexOf(needle, i + 1);
      }
    }
    return false;
  }, needle, nth);
  await sleep(40);
  return found;
}
async function caretIn(selector, atEnd = false) {
  await page.evaluate((s, atEnd) => {
    const el = document.querySelector("#editor " + s);
    document.getElementById("editor").focus();
    const r = document.createRange();
    r.selectNodeContents(el); r.collapse(!atEnd);
    getSelection().removeAllRanges(); getSelection().addRange(r);
  }, selector, atEnd);
  await sleep(40);
}
async function tab(name) { await page.click(`.rtab[data-tab="${name}"]`); await sleep(30); }
async function click(sel) { await page.click(sel); await sleep(60); }
const html = () => page.$eval("#editor", (e) => e.innerHTML);
const ev = (fn, ...a) => page.evaluate(fn, ...a);
async function key(combo) {
  const parts = combo.split("+");
  for (const p of parts.slice(0, -1)) await page.keyboard.down(p);
  await page.keyboard.press(parts[parts.length - 1]);
  for (const p of parts.slice(0, -1).reverse()) await page.keyboard.up(p);
  await sleep(60);
}
function waitFile(ext, timeout = 5000) {
  const t0 = Date.now();
  return new Promise((res) => {
    const iv = setInterval(() => {
      const f = fs.readdirSync(TMP).find((x) => x.endsWith(ext) && !x.endsWith(".crdownload"));
      if (f || Date.now() - t0 > timeout) { clearInterval(iv); res(f ? path.join(TMP, f) : null); }
    }, 100);
  });
}

try {
  console.log("\nAvvio");
  await page.goto(URL0, { waitUntil: "load" });
  await ev(() => localStorage.clear());
  await boot();
  eq(await page.title(), "Benvenuto in Word — Word", "la scheda del browser mostra documento e «Word»");
  ok(await ev(() => document.querySelector(".brand img").naturalWidth > 0), "l'icona si carica");
  ok((await html()).includes("Benvenuto in Word"), "documento di benvenuto");
  ok(await ev(() => document.querySelector("#saveText").textContent.length > 0), "indicatore di salvataggio presente");

  console.log("\nPulizia dell'HTML (sicurezza)");
  const clean = await ev(() => __word.sanitize(`<p onclick="x()">a<script>alert(1)<\/script></p><a href="java\tscript:alert(1)">b</a><a href="https://ok.it">c</a><img src="data:text/html,xx" onerror="x()"><img src="data:image/png;base64,AA"><hr class="page-break foo"><p class="MsoNormal" id="z">d</p><svg><a href="javascript:1">s</a></svg><iframe src="x"></iframe>`));
  ok(!/script|onclick|onerror|iframe|svg|javascript/i.test(clean), "toglie script, eventi, iframe, svg e indirizzi javascript:", clean);
  ok(clean.includes('href="https://ok.it"'), "mantiene i collegamenti normali");
  ok(!clean.includes("data:text/html") && clean.includes("data:image/png"), "immagini: solo data:image");
  ok(clean.includes('class="page-break"') && !clean.includes("Mso") && !clean.includes('id="z"'), "classi: resta solo page-break");

  console.log("\nCarattere");
  await setDoc("<p>uno due tre quattro</p>");
  await select("due");
  await click('[data-cmd="bold"]');
  ok(/<(b|strong)>due<\/(b|strong)>|font-weight: ?(bold|700)[^>]*>due/.test(await html()), "grassetto sulla selezione", await html());
  ok(await ev(() => document.querySelector('[data-cmd="bold"]').classList.contains("on")), "il pulsante Grassetto risulta attivo");
  await select("tre");
  await click('[data-cmd="italic"]');
  ok(/font-style: ?italic[^>]*>tre|<i>tre<\/i>/.test(await html()), "corsivo");
  await select("quattro");
  await page.select("#fontSize", "18");
  await sleep(60);
  ok(/font-size: ?18pt[^>]*>quattro/.test(await html()), "dimensione 18 pt", await html());
  await select("uno");
  await click("#btnGrow");
  ok(/font-size: ?14pt[^>]*>uno/.test(await html()), "Ingrandisci: da 12 a 14 pt", await html());
  await select("uno");
  await click("#btnShrink");
  ok(/font-size: ?12pt[^>]*>uno/.test(await html()), "Riduci: torna a 12 pt");
  await select("quattro");
  await page.select("#fontName", "Arial");
  await sleep(60);
  ok(/font-family: ?Arial[^>]*>quattro|font-family:[^>]*Arial/.test(await html()), "cambio carattere in Arial");
  await setDoc("<p>ciao mondo. secondo pezzo</p>");
  await select("ciao mondo. secondo pezzo");
  await click('[data-pop="popCase"]'); await click('[data-case="upper"]');
  ok((await html()).includes("CIAO MONDO. SECONDO PEZZO"), "MAIUSCOLO");
  await click('[data-pop="popCase"]'); await click('[data-case="sentence"]');
  ok((await html()).includes("Ciao mondo. Secondo pezzo"), "Come una frase (selezione mantenuta)", await html());
  await click('[data-pop="popCase"]'); await click('[data-case="title"]');
  ok((await html()).includes("Ciao Mondo. Secondo Pezzo"), "Iniziali Maiuscole");
  await select("Mondo");
  await click('[data-cmd="removeFormat"]');
  ok(true, "Cancella formattazione non dà errori");

  console.log("\nParagrafo");
  await setDoc("<p>pera</p><p>banana</p><p>ananas</p>");
  await ev(() => { const ed = document.getElementById("editor"); const r = document.createRange(); r.setStart(ed.firstChild.firstChild, 0); r.setEnd(ed.lastChild.firstChild, 3); ed.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await sleep(50);
  await click("#btnSort");
  eq(await ev(() => [...document.querySelectorAll("#editor p")].map((p) => p.textContent)), ["ananas", "banana", "pera"], "Ordina A-Z sui paragrafi");
  await caretIn("p");
  await click('[data-cmd="justifyCenter"]');
  ok(/text-align: ?center/.test(await html()), "allinea al centro");
  await click('[data-pop="popSpacing"]'); await click('[data-lh="2"]');
  ok(/line-height: ?2/.test(await html()), "interlinea 2,0");
  await click('[data-pop="popBorders"]'); await click('[data-border="box"]');
  ok(/border: ?1px solid/.test(await html()), "bordo a riquadro");
  await click('[data-pop="popBorders"]'); await click('[data-border="none"]');
  ok(!/border: ?1px/.test(await html()), "nessun bordo");
  await setDoc("<p>uno</p>");
  await caretIn("p");
  await click('[data-cmd="insertUnorderedList"]');
  ok((await html()).includes("<ul><li>uno</li></ul>"), "elenco puntato", await html());
  await click('[data-cmd="insertOrderedList"]');
  ok((await html()).includes("<ol><li>uno</li></ol>"), "elenco numerato (senza involucri inutili)", await html());
  await click("#btnMarks");
  ok(await ev(() => document.getElementById("editor").classList.contains("marks") && document.getElementById("btnMarks2").classList.contains("on")), "segni di paragrafo (e pulsante gemello)");
  await click("#btnMarks");

  console.log("\nStili");
  await setDoc("<p>Titolo prova</p><p>testo</p>");
  await caretIn("p");
  await click('.style-tile[data-style="h1"]');
  ok((await html()).startsWith("<h1>Titolo prova</h1>"), "Titolo 1 dalla galleria");
  ok(await ev(() => document.querySelector('.style-tile[data-style="h1"]').classList.contains("on")), "la galleria evidenzia Titolo 1");
  await caretIn("h1", true);
  await page.keyboard.press("Enter");
  await sleep(60);
  ok(/<\/h1><p><br><\/p>/.test(await html()), "Invio alla fine del titolo crea un paragrafo Normale", await html());
  await caretIn("p:last-child");
  await key("Meta+Alt+Digit2");
  ok((await html()).includes("<h2>testo</h2>"), "⌘⌥2 applica Titolo 2", await html());
  await key("Meta+Alt+Digit0");
  ok((await html()).includes("<p>testo</p>"), "⌘⌥0 torna a Normale");

  console.log("\nTrova e sostituisci");
  await setDoc("<p>gatto uno, Gatto due</p><p>il gatto tre</p>");
  await caretIn("p");
  await key("Meta+f");
  ok(await ev(() => !document.getElementById("findbar").hidden && document.activeElement.id === "findInput"), "⌘F apre la barra Trova");
  await page.type("#findInput", "gatto");
  await sleep(50);
  eq(await page.$eval("#findCount", (e) => e.textContent), "3 risultati", "conta 3 risultati (maiuscole ignorate)");
  ok(await ev(() => CSS.highlights.get("word-find")?.size === 3), "evidenzia tutti e 3 i risultati");
  // posizione del risultato attuale: testo del paragrafo + scostamento, letto dall'evidenziazione
  const where = async () => ev(() => { const [r] = [...CSS.highlights.get("word-find-cur")]; return r.startContainer.parentElement.textContent.slice(0, 5) + ":" + r.startOffset + " " + document.getElementById("findCount").textContent; });
  const seen = [];
  for (let k = 0; k < 4; k++) { await page.keyboard.press("Enter"); await sleep(80); seen.push(await where()); }
  eq(seen, ["gatto:0 1 di 3", "gatto:11 2 di 3", "il ga:3 3 di 3", "gatto:0 1 di 3"], "Successivo avanza su ogni risultato e poi ricomincia");
  ok(await ev(() => document.activeElement.id === "findInput"), "il cursore resta nella casella Trova");
  await click("#findPrev"); await sleep(50);
  eq(await where(), "il ga:3 3 di 3", "Precedente torna all'ultimo");
  await page.keyboard.down("Shift"); await page.keyboard.press("Enter"); await page.keyboard.up("Shift"); await sleep(80);
  eq(await where(), "gatto:11 2 di 3", "⇧Invio va al precedente");
  await page.type("#replaceInput", "cane");
  await click("#replaceAll");
  ok((await html()).includes("cane uno, cane due") && (await html()).includes("il cane tre"), "Sostituisci tutto", await html());
  await ev(() => { document.getElementById("editor").focus(); document.execCommand("undo"); document.execCommand("undo"); document.execCommand("undo"); });
  await sleep(60);
  ok((await html()).includes("il gatto tre"), "⌘Z annulla le sostituzioni", await html());
  await setDoc("<p>aa bb aa</p>");
  await page.$eval("#findInput", (i) => (i.value = ""));
  await page.type("#findInput", "aa");
  await page.$eval("#replaceInput", (i) => (i.value = ""));
  await page.type("#replaceInput", "aaa");
  await click("#replaceAll");
  eq(await ev(() => document.querySelector("#editor p").textContent), "aaa bb aaa", "sostituire con un testo che contiene la ricerca non va in ciclo");
  await page.$eval("#replaceInput", (i) => (i.value = ""));
  await page.$eval("#findInput", (i) => (i.value = ""));
  await page.type("#findInput", "bb ");
  await click("#replaceOne"); await click("#replaceOne");
  eq(await ev(() => document.querySelector("#editor p").textContent), "aaa aaa", "Sostituisci con testo vuoto cancella");
  await page.focus("#findInput");
  await page.keyboard.press("Escape"); await sleep(50);
  ok(await ev(() => document.getElementById("findbar").hidden), "Esc chiude la barra Trova");
  ok(await ev(() => !CSS.highlights.has("word-find")), "chiudendo Trova spariscono le evidenziazioni");

  console.log("\nTabelle");
  await setDoc("<p>prima</p>");
  await caretIn("p", true);
  await tab("insert");
  await click("#btnTable");
  await page.hover('#tableGrid span[data-r="2"][data-c="3"]');
  eq(await page.$eval("#tableLabel", (e) => e.textContent), "2 × 3", "la griglia mostra 2 × 3");
  await click('#tableGrid span[data-r="2"][data-c="3"]');
  eq(await ev(() => { const t = document.querySelector("#editor table"); return [t.rows.length, t.rows[0].cells.length]; }), [2, 3], "tabella 2 × 3 inserita");
  ok(await ev(() => !document.getElementById("tabTable").hidden && document.getElementById("tabTable").getAttribute("aria-selected") === "true"), "compare la scheda Tabella");
  await page.keyboard.type("A1");
  await page.keyboard.press("Tab"); await page.keyboard.type("B1");
  eq(await ev(() => document.querySelector("#editor table").rows[0].cells[1].textContent), "B1", "Tab passa alla cella successiva");
  await ev(() => { const c = document.querySelector("#editor table").rows[1].cells[2]; const r = document.createRange(); r.selectNodeContents(c); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await sleep(50);
  await page.keyboard.press("Tab"); await sleep(60);
  eq(await ev(() => document.querySelector("#editor table").rows.length), 3, "Tab nell'ultima cella aggiunge una riga");
  await click('[data-table="colRight"]');
  eq(await ev(() => document.querySelector("#editor table").rows[0].cells.length), 4, "colonna a destra");
  await click('[data-table="delCol"]');
  eq(await ev(() => document.querySelector("#editor table").rows[0].cells.length), 3, "elimina colonna");
  await click('[data-table="rowAbove"]');
  eq(await ev(() => document.querySelector("#editor table").rows.length), 4, "riga sopra");
  await click('[data-table="delRow"]');
  eq(await ev(() => document.querySelector("#editor table").rows.length), 3, "elimina riga");
  await caretIn("table td");
  await click('[data-table="header"]');
  eq(await ev(() => document.querySelector("#editor table").rows[0].cells[0].tagName), "TH", "riga di intestazione");
  await click('[data-table="delTable"]');
  ok(await ev(() => !document.querySelector("#editor table")), "elimina tabella");
  ok(await ev(() => document.getElementById("tabTable").hidden), "la scheda Tabella sparisce");

  console.log("\nInserisci");
  await setDoc("<p>testo</p>");
  await caretIn("p", true);
  await tab("insert");
  await click("#btnPageBreak");
  ok((await html()).includes('<hr class="page-break">'), "interruzione di pagina");
  await sleep(250);
  ok((await page.$eval("#stats", (e) => e.textContent)).startsWith("2 pagine"), "le pagine contano l'interruzione", await page.$eval("#stats", (e) => e.textContent));
  await click("#btnDate");
  ok(/\d{1,2} [a-z]+ 20\d\d/.test(await html()), "data di oggi");
  await click('[data-pop="popSymbols"]'); await click('#symbolGrid button[title="€"]');
  ok((await html()).includes("€"), "simbolo €");
  await click("#btnHr");
  ok(/<hr>/.test(await html()), "linea orizzontale");
  await setDoc("<h1>Uno</h1><p>a</p><h2>Due</h2><p>b</p><h3>Tre</h3>");
  await caretIn("p");
  await click("#btnToc");
  eq(await ev(() => [...document.querySelectorAll("#editor .toc p")].map((p) => p.textContent)), ["Sommario", "Uno", "Due", "Tre"], "sommario con i titoli");
  await click("#btnToc");
  eq(await ev(() => document.querySelectorAll("#editor .toc").length), 1, "il sommario si aggiorna senza duplicarsi");

  console.log("\nCollegamenti");
  await setDoc("<p>vai al sito</p>");
  await select("sito");
  await key("Meta+k");
  ok(await ev(() => document.getElementById("dlgLink").open), "⌘K apre la finestra Collegamento");
  eq(await page.$eval("#linkText", (i) => i.value), "sito", "propone il testo selezionato");
  await page.type("#linkUrl", "esempio.it");
  await page.keyboard.press("Enter"); await sleep(80);
  ok((await html()).includes('<a href="https://esempio.it">sito</a>'), "collegamento con https:// aggiunto", await html());
  await caretIn("a");
  await click("#btnLink");
  ok(await ev(() => !document.getElementById("linkRemove").hidden), "su un collegamento esistente si può togliere");
  await click("#linkRemove");
  ok(!(await html()).includes("<a"), "Togli collegamento");
  await caretIn("p", true);
  await click("#btnLink");
  await page.type("#linkText", "scrivimi");
  await page.type("#linkUrl", "javascript:alert(1)");
  await page.keyboard.press("Enter"); await sleep(80);
  ok(!(await html()).includes("javascript"), "rifiuta indirizzi javascript:");
  eq(await ev(() => __word.normalizeUrl("info@esempio.it")), "mailto:info@esempio.it", "un indirizzo email diventa mailto:");

  console.log("\nImmagini");
  const png = path.join(TMP, "prova.png");
  fs.writeFileSync(png, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP8z8Dwn4GBgYGJAQoAIwYDAcV4M6MAAAAASUVORK5CYII=", "base64"));
  await setDoc("<p>foto: </p>");
  await caretIn("p", true);
  const fi = await page.$("#fileImage");
  await fi.uploadFile(png);
  await page.waitForFunction(() => document.querySelector("#editor img"), { timeout: 3000 });
  ok(true, "immagine inserita da file");
  await click("#editor img");
  ok(await ev(() => !document.getElementById("imgbar").hidden), "clic sull'immagine mostra la sua barra");
  await click('#imgbar [data-w="25"]');
  ok(/width: ?25%/.test(await html()), "larghezza 25%");
  await click('#imgbar [data-align="center"]');
  ok(/text-align: ?center/.test(await ev(() => document.querySelector("#editor img").parentElement.getAttribute("style") || "")), "immagine al centro");
  await click('#imgbar [data-w="del"]');
  ok(await ev(() => !document.querySelector("#editor img") && document.getElementById("imgbar").hidden), "elimina immagine");

  console.log("\nIncolla");
  await setDoc("<p>x</p>");
  await caretIn("p", true);
  await ev(() => {
    const dt = new DataTransfer();
    dt.setData("text/html", '<b onclick="evil()">forte</b><script>bad()<\/script>');
    dt.setData("text/plain", "forte");
    document.getElementById("editor").dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  await sleep(60);
  ok((await html()).includes("forte") && !/onclick|script/.test(await html()), "l'HTML incollato viene ripulito", await html());

  console.log("\nCopia formato");
  await tab("home");
  await setDoc('<p><b style="color: rgb(200, 0, 0);">rosso</b> normale</p>');
  await select("rosso");
  await click("#btnPainter");
  ok(await ev(() => document.body.classList.contains("painting")), "copia formato attivo");
  await ev(() => { const t = [...document.querySelectorAll("#editor p")][0].lastChild; const r = document.createRange(); r.setStart(t, 1); r.setEnd(t, 8); getSelection().removeAllRanges(); getSelection().addRange(r); document.getElementById("editor").dispatchEvent(new MouseEvent("mouseup", { bubbles: true })); });
  await sleep(150);
  ok(await ev(() => { const w = document.createTreeWalker(document.getElementById("editor"), NodeFilter.SHOW_TEXT); while (w.nextNode()) if (w.currentNode.nodeValue.includes("normale")) { const s = getComputedStyle(w.currentNode.parentElement); return parseInt(s.fontWeight) >= 600 && s.color === "rgb(200, 0, 0)"; } return false; }), "il formato (grassetto rosso) passa al testo scelto", await html());
  ok(await ev(() => !document.body.classList.contains("painting")), "copia formato si spegne da solo");

  console.log("\nLayout e visualizza");
  await tab("layout");
  await click('[data-pop="popMargins"]'); await click('[data-margins="narrow"]');
  eq(await ev(() => document.getElementById("page").style.getPropertyValue("--ml")), "1.27cm", "margini stretti");
  await click('[data-pop="popOrient"]'); await click('[data-orient="landscape"]');
  eq(await ev(() => document.getElementById("page").style.getPropertyValue("--pw")), "29.7cm", "orizzontale");
  await click('[data-pop="popSize"]'); await click('[data-size="A5"]');
  eq(await ev(() => document.getElementById("page").style.getPropertyValue("--pw")), "21cm", "A5 orizzontale = 21 cm di larghezza");
  ok(await ev(() => [...document.querySelectorAll("style")].some((s) => s.textContent.includes("@page { size: 21cm 14.8cm"))), "regola di stampa @page aggiornata");
  await click('[data-pop="popSize"]'); await click('[data-size="A4"]');
  await click('[data-pop="popOrient"]'); await click('[data-orient="portrait"]');
  await click('[data-pop="popMargins"]'); await click('[data-margins="normal"]');
  await tab("view");
  await click("#zoomPlus");
  eq(await page.$eval("#zoomOut", (e) => e.textContent), "110%", "zoom +");
  await click("#zoomMinus"); await click("#zoomMinus");
  eq(await page.$eval("#zoomOut", (e) => e.textContent), "90%", "zoom −");
  await click("#btnFit");
  ok(parseInt(await page.$eval("#zoomOut", (e) => e.textContent)) > 100, "adatta alla larghezza");
  await click("#btnZoom100");
  const themes = [];
  for (let i = 0; i < 3; i++) { await click("#btnTheme"); themes.push(await ev(() => document.documentElement.dataset.theme || "auto")); }
  eq(themes, ["light", "dark", "auto"], "tema: chiaro → scuro → auto");
  await click("#btnNav");
  await setDoc("<h1>Capitolo</h1><p>a</p><h2>Paragrafo</h2>");
  await sleep(400);
  eq(await ev(() => [...document.querySelectorAll("#navList button")].map((b) => b.textContent)), ["Capitolo", "Paragrafo"], "la Struttura elenca i titoli");
  await click("#navList li:nth-child(2) button");
  eq(await ev(() => getSelection().anchorNode.closest ? getSelection().anchorNode.closest("h2")?.textContent : getSelection().anchorNode.parentElement.closest("h2")?.textContent), "Paragrafo", "clic nella Struttura porta al titolo");
  await click("#navClose");
  await click("#btnFocus");
  ok(await ev(() => getComputedStyle(document.querySelector(".topbar")).display === "none"), "Concentrazione nasconde le barre");
  await page.keyboard.press("Escape"); await sleep(60);
  ok(await ev(() => getComputedStyle(document.querySelector(".topbar")).display !== "none"), "Esc esce dalla concentrazione");
  await page.click('.rtab[data-tab="home"]', { clickCount: 2 }); await sleep(60);
  ok(await ev(() => document.body.classList.contains("ribbon-collapsed")), "doppio clic su una scheda riduce la barra");
  await click('.rtab[data-tab="insert"]');
  ok(await ev(() => !document.body.classList.contains("ribbon-collapsed") && !document.querySelector('[data-panel="insert"]').hidden), "clic su una scheda la riapre");
  await tab("view");
  await click("#btnKeys");
  ok(await ev(() => document.getElementById("dlgKeys").open), "elenco scorciatoie");
  await page.keyboard.press("Escape"); await sleep(60);

  console.log("\nConteggio");
  eq(await ev(() => { const d = document.createElement("div"); d.textContent = "L'acqua - è 3,5 litri e-mail ok"; return __word.textStats(d).words; }), 6, "le parole non contano i trattini isolati");
  await setDoc("<p>una due tre</p>");
  await key("Meta+Shift+KeyG");
  ok(await ev(() => document.getElementById("dlgStats").open && document.getElementById("statsGrid").textContent.includes("Parole3")), "⌘⇧G apre le statistiche");
  await page.keyboard.press("Escape"); await sleep(60);
  ok((await page.$eval("#stats", (e) => e.textContent)).includes("3 parole"), "barra di stato: 3 parole");

  console.log("\nRicerca dei comandi");
  await caretIn("p");
  await key("Meta+j");
  ok(await ev(() => document.activeElement.id === "cmdInput"), "⌘J va alla ricerca dei comandi");
  await page.type("#cmdInput", "margini stretti");
  await sleep(50);
  ok((await page.$eval("#cmdResults", (e) => e.textContent)).includes("Margini stretti"), "trova «Margini stretti»");
  await page.keyboard.press("Enter"); await sleep(80);
  eq(await ev(() => document.getElementById("page").style.getPropertyValue("--ml")), "1.27cm", "Invio esegue il comando");
  await page.click("#cmdInput");
  await page.type("#cmdInput", "tabella");
  await sleep(50);
  await page.keyboard.press("Enter"); await sleep(120);
  ok(await ev(() => !document.getElementById("tablePicker").hidden), "«tabella» apre la griglia della tabella nella scheda Inserisci");
  await page.keyboard.press("Escape");
  await page.click("#cmdInput");
  await page.type("#cmdInput", "grasset");
  await sleep(50);
  eq(await ev(() => document.querySelector("#cmdResults .cr-label").textContent), "Grassetto", "accetta parole incomplete");
  await page.keyboard.press("Escape");
  await ev(() => { document.querySelector('[data-margins="normal"]').click(); });

  console.log("\nMarkdown");
  const md = await ev(() => __word.mdToHtml("# Titolo\n\nTesto **forte** e *corsivo* con [link](https://a.it) e ~~via~~.\n\n- uno\n- due\n    - due.a\n1. primo\n\n> citazione\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n---\n\nriga  \nnuova"));
  ok(md.includes("<h1>Titolo</h1>") && md.includes("<b>forte</b>") && md.includes("<i>corsivo</i>") && md.includes('<a href="https://a.it">link</a>') && md.includes("<s>via</s>"), "Markdown → testo formattato", md);
  ok(md.includes("<ul><li>uno</li><li>due<ul><li>due.a</li></ul></li></ul>"), "elenchi annidati", md);
  ok(md.includes("<blockquote><p>citazione</p></blockquote>") && md.includes("<th>A</th>") && md.includes("<hr>") && md.includes("riga<br>nuova"), "citazione, tabella, linea, a capo");
  await setDoc(md);
  const back = await ev(() => __word.htmlToMd(document.getElementById("editor")));
  ok(back.includes("# Titolo") && back.includes("**forte**") && back.includes("*corsivo*") && back.includes("[link](https://a.it)") && back.includes("~~via~~"), "testo → Markdown", back);
  ok(back.includes("- uno\n- due\n    - due.a") && back.includes("> citazione") && back.includes("| A | B |"), "Markdown: elenchi, citazione, tabella", back);
  eq(await ev((b) => __word.mdToHtml(b), back), md.replace("<p>riga<br>nuova</p>", "<p>riga<br>nuova</p>"), "andata e ritorno Markdown identica");

  console.log("\nFile .docx");
  await setDoc('<h1>Relazione</h1><p>Testo <b>grassetto</b> e <i>corsivo</i>, <a href="https://esempio.it">sito</a>.</p><ul><li>voce A</li><li>voce B</li></ul><ol><li>uno</li></ol><table><tbody><tr><th>Nome</th><th>Età</th></tr><tr><td>Ada</td><td>36</td></tr></tbody></table><blockquote>Citazione</blockquote><hr class="page-break"><p style="text-align:center">fine</p>');
  const b64 = await ev(async () => {
    const blob = await window.WordDocx.toBlob(document.getElementById("editor"), "Relazione", { width: 21, height: 29.7, top: 2.5, side: 2 });
    const buf = new Uint8Array(await blob.arrayBuffer());
    let s = ""; for (const x of buf) s += String.fromCharCode(x);
    return btoa(s);
  });
  const docxPath = path.join(TMP, "Relazione.docx");
  fs.writeFileSync(docxPath, Buffer.from(b64, "base64"));
  ok(fs.statSync(docxPath).size > 3000 && Buffer.from(b64, "base64").subarray(0, 2).toString() === "PK", "crea un .docx valido (zip)");
  const rt = await ev(async (b) => {
    const bin = atob(b), u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return (await window.mammoth.convertToHtml({ arrayBuffer: u.buffer })).value;
  }, b64);
  ok(/<h1>(<strong>)?Relazione(<\/strong>)?<\/h1>/.test(rt), "nel .docx il titolo è un Titolo 1 vero", rt.slice(0, 200));
  ok(rt.includes("<strong>grassetto</strong>") && rt.includes("<em>corsivo</em>"), "grassetto e corsivo nel .docx");
  ok(rt.includes('<a href="https://esempio.it">sito</a>'), "collegamento nel .docx");
  ok(rt.includes("<li>voce A</li>") && rt.includes("<table>") && rt.includes("Ada"), "elenchi e tabella nel .docx");
  await page.$eval("#editor", (e) => (e.innerHTML = "<p>vecchio</p>"));
  await (await page.$("#fileOpen")).uploadFile(docxPath);
  await page.waitForFunction(() => document.getElementById("editor").textContent.includes("Relazione"), { timeout: 4000 });
  eq(await page.$eval("#title", (e) => e.value), "Relazione", "Apri .docx: titolo dal nome del file");
  ok(/<h1>(<b>|<strong>)?Relazione/.test(await html()) && (await html()).includes("<table>") && (await html()).includes("<li>voce A</li>"), "Apri .docx: contenuto", (await html()).slice(0, 200));

  console.log("\nScaricare ed esportare");
  await page.$eval("#title", (e) => (e.value = "Prova export"));
  await click("#btnSave");
  const fDocx = await waitFile(".docx".replace(".docx", "export.docx"));
  ok(fDocx && fs.statSync(fDocx).size > 3000, "Scarica .docx", String(fDocx));
  for (const [id, ext, check] of [["btnExportHtml", ".html", (t) => t.includes("<title>Prova export</title>") && /<h1>(<b>|<strong>)?Relazione/.test(t)], ["btnExportMd", ".md", (t) => t.startsWith("# Relazione")], ["btnExportTxt", ".txt", (t) => t.startsWith("Relazione")]]) {
    await click("#btnExport");
    await click("#" + id);
    const f = await waitFile("export" + ext);
    ok(f && check(fs.readFileSync(f, "utf8")), "Esporta " + ext, f ? fs.readFileSync(f, "utf8").slice(0, 80) : "nessun file");
  }
  const mdFile = path.join(TMP, "appunti.md");
  fs.writeFileSync(mdFile, "# Appunti\n\n- **uno**\n- due\n");
  await (await page.$("#fileOpen")).uploadFile(mdFile);
  await page.waitForFunction(() => document.getElementById("editor").textContent.includes("Appunti"), { timeout: 3000 });
  ok((await html()).includes("<h1>Appunti</h1><ul><li><b>uno</b></li>"), "Apri .md", await html());

  console.log("\nSalvataggio automatico");
  await setDoc("<p>da ricordare</p>");
  await page.$eval("#title", (e) => { e.value = "Memo"; e.dispatchEvent(new Event("input")); });
  await sleep(800);
  eq(await page.$eval("#saveState", (e) => e.dataset.state), "saved", "lo stato torna «salvato»");
  await page.reload({ waitUntil: "load" });
  await sleep(200);
  ok((await html()).includes("da ricordare") && (await page.$eval("#title", (e) => e.value)) === "Memo", "dopo il ricaricamento il documento c'è ancora");
  // come al primo avvio dopo l'aggiornamento: ci sono solo i dati salvati da Parola
  await page.goto(URL0 + "vuoto", { waitUntil: "load" });
  await ev(() => { localStorage.clear(); localStorage.setItem("parola.documento", JSON.stringify({ title: "Vecchio", html: "<p>dal vecchio Parola</p>" })); localStorage.setItem("parola.preferenze", JSON.stringify({ theme: "dark" })); });
  await boot();
  await sleep(200);
  ok((await html()).includes("dal vecchio Parola") && (await ev(() => document.documentElement.dataset.theme)) === "dark", "recupera documento e preferenze salvati da Parola");
  await click("#btnNew");
  ok((await html()) === "<p><br></p>" && (await page.$eval("#title", (e) => e.value)) === "Documento senza titolo", "Nuovo documento");

  console.log("\nSchermo piccolo");
  await page.setViewport({ width: 390, height: 800, isMobile: true, hasTouch: true });
  await page.reload({ waitUntil: "load" });
  await sleep(200);
  ok(await ev(() => document.documentElement.scrollWidth <= innerWidth), "nessuno scorrimento orizzontale della pagina a 390 px");
  ok(await ev(() => document.getElementById("editor").getBoundingClientRect().width > 250), "il foglio occupa lo schermo");

  console.log("\nErrori nella console");
  eq(errors, [], "nessun errore JavaScript");
} catch (e) {
  fail++;
  failures.push("eccezione: " + e.message);
  console.log("  ✗ ECCEZIONE: " + e.stack);
} finally {
  await browser.close();
  server.close();
  fs.rmSync(TMP, { recursive: true, force: true });
}
console.log(`\n${pass} prove superate, ${fail} fallite`);
if (fail) { console.log("Fallite:\n - " + failures.join("\n - ")); process.exit(1); }
