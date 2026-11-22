/* Word — elaboratore di testi */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const editor = $("editor"), page = $("page"), desk = $("desk"), titleEl = $("title"), saveState = $("saveState");
  const STORE = "word.documento", PREFS = "word.preferenze";
  const OLD_STORE = "parola.documento", OLD_PREFS = "parola.preferenze"; // nomi usati quando l'app si chiamava Parola
  const IMG_MAX = 1600;
  const UNTITLED = "Documento senza titolo";
  // Presente solo dentro l'app per Mac
  const MH = window.webkit && window.webkit.messageHandlers;
  const NATIVE = MH && (MH.word || MH.parola);
  if (NATIVE) { document.documentElement.classList.add("native"); $("saveLabel").textContent = "Salva .docx"; }
  const SAVE_LABEL = NATIVE ? "Salva .docx" : "Scarica .docx";
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  const PAGE_SIZES = { A4: [21, 29.7], A5: [14.8, 21], Letter: [21.59, 27.94] };
  const MARGINS = { normal: [2.5, 2], narrow: [1.27, 1.27], moderate: [2.54, 1.91], wide: [2.54, 5.08] };
  const SAFE_URL = /^(https?:|mailto:|tel:|#)/i;
  // una parola: lettere/cifre unite da apostrofo, trattino o punto; i numeri possono avere la virgola (3,5)
  const WORD_RE = /\p{N}+(?:[.,]\p{N}+)+|[\p{L}\p{N}]+(?:['’\-.][\p{L}\p{N}]+)*/gu;

  /* ---------- Utilità ---------- */
  let toastTimer;
  function toast(msg, ms = 2600) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), ms);
  }
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const insideEditor = (n) => n && (n === editor || editor.contains(n));
  const fmtInt = (n) => n.toLocaleString("it-IT");
  const plural = (n, one, many) => `${fmtInt(n)} ${n === 1 ? one : many}`;
  const fold = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

  // Pulisce l'HTML che arriva da fuori (file aperti, incolla, archivio)
  function sanitize(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    doc.querySelectorAll("script,style,iframe,frame,frameset,object,embed,link,meta,base,form,input,button,textarea,select,svg,math,template,noscript").forEach((n) => n.remove());
    doc.querySelectorAll("*").forEach((el) => {
      for (const a of Array.from(el.attributes)) {
        const n = a.name.toLowerCase();
        const v = a.value.replace(/[\u0000- \u007f-\u009f]/g, "").toLowerCase();
        if (n.startsWith("on") || n === "srcset" || n === "formaction" || n === "action" || n === "xlink:href" || n === "id") el.removeAttribute(a.name);
        else if (n === "href" && !(SAFE_URL.test(v) || !/^[a-z][a-z0-9+.-]*:/.test(v))) el.removeAttribute(a.name);
        else if (n === "src" && /^[a-z][a-z0-9+.-]*:/.test(v) && !/^(https?:|data:image\/|blob:)/.test(v)) el.removeAttribute(a.name);
        else if (n === "contenteditable" || n === "class") {
          const keep = n === "class" && a.value.split(/\s+/).filter((c) => c === "page-break" || c === "toc").join(" ");
          if (keep) el.setAttribute("class", keep); else el.removeAttribute(a.name);
        }
      }
    });
    return doc.body.innerHTML;
  }

  function readJSON(key, fallbackKey) {
    try {
      const v = localStorage.getItem(key) ?? (fallbackKey ? localStorage.getItem(fallbackKey) : null);
      return v ? JSON.parse(v) : null;
    } catch (e) { return null; }
  }
  let prefs = readJSON(PREFS, OLD_PREFS) || {};
  function savePrefs() { try { localStorage.setItem(PREFS, JSON.stringify(prefs)); } catch (e) { /* ignoro */ } }

  /* ---------- Selezione ---------- */
  let savedRange = null;
  document.addEventListener("selectionchange", () => {
    const sel = getSelection();
    if (sel.rangeCount && insideEditor(sel.anchorNode)) {
      savedRange = sel.getRangeAt(0).cloneRange();
      refreshToolbar();
    }
  });
  function restoreSelection() {
    editor.focus({ preventScroll: true });
    if (savedRange && document.contains(savedRange.startContainer)) {
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }
  function exec(cmd, val = null, css = true) {
    restoreSelection();
    document.execCommand("styleWithCSS", false, css);
    document.execCommand(cmd, false, val);
    if (/List|formatBlock|indent|outdent/i.test(cmd)) dropUselessSpans();
    keepSelection();
    changed();
  }
  // Memorizza subito la selezione attuale (l'evento selectionchange arriva solo dopo)
  function keepSelection() {
    const sel = getSelection();
    if (sel.rangeCount && insideEditor(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange();
  }
  // Chrome a volte avvolge il testo in <span style="font-size:…"> uguale a quello che c'è già: lo tolgo
  function dropUselessSpans() {
    $$('span[style^="font-size"]', editor).forEach((s) => {
      if (s.style.length !== 1 || !s.parentElement || /​/.test(s.textContent)) return;
      const own = s.style.fontSize;
      s.style.fontSize = "";
      const same = getComputedStyle(s).fontSize;
      s.style.fontSize = own;
      if (getComputedStyle(s).fontSize === same) s.replaceWith(...s.childNodes);
    });
  }
  function setCaret(node, offset = 0) {
    const sel = getSelection();
    editor.focus({ preventScroll: true });
    sel.collapse(node, offset);
    savedRange = sel.getRangeAt(0).cloneRange();
  }
  // Paragrafi (o voci di elenco, celle…) toccati dalla selezione, i più interni.
  function selectedBlocks() {
    if (!savedRange) return [];
    const all = $$("p,h1,h2,h3,h4,h5,h6,li,blockquote,td,th", editor).filter((b) => savedRange.intersectsNode(b));
    return all.filter((b) => !all.some((o) => o !== b && b.contains(o)));
  }

  /* ---------- Struttura: niente blocchi dentro i paragrafi ---------- */
  const BLOCKS = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "UL", "OL", "TABLE", "HR", "PRE"]);
  function normalize() {
    const bad = $$("p,h1,h2,h3,h4,h5,h6", editor).filter((p) => [...p.children].some((c) => BLOCKS.has(c.tagName)));
    if (!bad.length) return;
    const sel = getSelection();
    const mark = sel.rangeCount ? { n: sel.anchorNode, o: sel.anchorOffset } : null;
    bad.reverse().forEach((p) => {
      const frag = document.createDocumentFragment();
      let cur = null;
      [...p.childNodes].forEach((c) => {
        if (c.nodeType === 1 && BLOCKS.has(c.tagName)) { cur = null; frag.appendChild(c); return; }
        if (c.nodeType === 3 && !c.nodeValue.trim() && !cur) return;
        if (!cur) { cur = document.createElement(p.tagName); frag.appendChild(cur); }
        cur.appendChild(c);
      });
      p.replaceWith(frag);
    });
    $$("table:empty, p:empty", editor).forEach((n) => (n.tagName === "P" ? (n.innerHTML = "<br>") : n.remove()));
    if (mark && document.contains(mark.n)) {
      try { sel.collapse(mark.n, Math.min(mark.o, mark.n.nodeType === 3 ? mark.n.length : mark.n.childNodes.length)); } catch (e) { /* ignoro */ }
    }
  }

  /* ---------- Impostazioni pagina ---------- */
  let pageSet = { size: "A4", orient: "portrait", margins: "normal" };
  const pageRule = document.createElement("style");
  document.head.appendChild(pageRule);
  function pageDims() {
    let [w, h] = PAGE_SIZES[pageSet.size] || PAGE_SIZES.A4;
    if (pageSet.orient === "landscape") [w, h] = [h, w];
    const [mt, ml] = MARGINS[pageSet.margins] || MARGINS.normal;
    return { w, h, mt, ml };
  }
  function applyPage() {
    const { w, h, mt, ml } = pageDims();
    page.style.setProperty("--pw", w + "cm");
    page.style.setProperty("--ph", h + "cm");
    page.style.setProperty("--mt", mt + "cm");
    page.style.setProperty("--ml", ml + "cm");
    pageRule.textContent = `@page { size: ${w}cm ${h}cm; margin: ${mt}cm ${ml}cm; }`;
    $$("[data-size]").forEach((b) => b.classList.toggle("on", b.dataset.size === pageSet.size));
    $$("[data-orient]").forEach((b) => b.classList.toggle("on", b.dataset.orient === pageSet.orient));
    $$("[data-margins]").forEach((b) => b.classList.toggle("on", b.dataset.margins === pageSet.margins));
    clearTimeout(statsTimer);
    statsTimer = setTimeout(updateStats, 50);
  }

  /* ---------- Salvataggio automatico ---------- */
  let saveTimer, statsTimer, navTimer, findTimer;
  function setSaveState(state, text) {
    saveState.dataset.state = state;
    $("saveText").textContent = text;
  }
  function changed() {
    normalize();
    setSaveState("pending", "Modifiche non salvate…");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persist, 600);
    clearTimeout(statsTimer);
    statsTimer = setTimeout(updateStats, 120);
    if (!$("navPane").hidden) { clearTimeout(navTimer); navTimer = setTimeout(buildNav, 300); }
    if (!$("findbar").hidden) { clearTimeout(findTimer); findTimer = setTimeout(() => countMatches(), 250); }
  }
  function cleanHTML() {
    return editor.innerHTML.replace(/ class="picked"/g, "");
  }
  function persist() {
    clearTimeout(saveTimer);
    try {
      localStorage.setItem(STORE, JSON.stringify({ title: titleEl.value, html: cleanHTML(), page: pageSet, t: Date.now() }));
      setSaveState("saved", "Salvato su questo dispositivo");
      sendTitle();
    } catch (e) {
      setSaveState("error", "Non salvato: documento troppo grande, scaricalo in .docx");
    }
  }
  // titolo del documento nella barra della finestra (app per Mac) e nella scheda del browser
  function sendTitle() {
    const t = titleEl.value.trim();
    document.title = t ? `${t} — Word` : "Word";
    if (NATIVE) NATIVE.postMessage({ type: "title", title: t });
  }
  function load() {
    const d = readJSON(STORE, OLD_STORE);
    if (d && typeof d.html === "string") {
      editor.innerHTML = sanitize(d.html);
      titleEl.value = d.title || UNTITLED;
      if (d.page) pageSet = { ...pageSet, ...d.page };
      setSaveState("saved", "Salvato su questo dispositivo");
      return;
    }
    editor.innerHTML = WELCOME;
    titleEl.value = "Benvenuto in Word";
    setSaveState("saved", "Pronto");
  }

  const WELCOME = `
    <h1>Benvenuto in Word</h1>
    <p>Word è un elaboratore di testi che funziona nel browser e sul Mac. Scrivi qui, formatta con le schede in alto e scarica il risultato come file <b>.docx</b>, che si apre con Microsoft Word, Pages, LibreOffice e Google Documenti.</p>
    <h2>Le schede</h2>
    <ul>
      <li><b>Home</b>: carattere, <b>grassetto</b>, <i>corsivo</i>, <u>sottolineato</u>, <span style="color:#2457d6">colori</span>, <span style="background-color:#ffe58a">evidenziatore</span>, elenchi, allineamento, interlinea, stili e trova.</li>
      <li><b>Inserisci</b>: tabelle, immagini, collegamenti, sommario, interruzioni di pagina, data e simboli.</li>
      <li><b>Layout</b>: margini, orientamento e dimensioni del foglio.</li>
      <li><b>Revisione</b>: ortografia, statistiche, trova e sostituisci.</li>
      <li><b>Visualizza</b>: struttura del documento, modalità concentrazione, zoom, tema scuro e scorciatoie.</li>
    </ul>
    <h2>Trucchi</h2>
    <ul>
      <li>Premi <b>⌘J</b> (Ctrl+J) e scrivi cosa vuoi fare: «tabella», «margini», «pdf»…</li>
      <li><b>Esporta</b> crea anche PDF, pagine web, Markdown e testo semplice.</li>
      <li>Fai doppio clic su una scheda per ridurre la barra dei comandi.</li>
    </ul>
    <p>Il documento si salva da solo su questo dispositivo mentre scrivi. Quando hai finito, premi <b>Scarica .docx</b>.</p>`;

  /* ---------- Schede ---------- */
  function showTab(name) {
    $$(".rtab").forEach((t) => t.setAttribute("aria-selected", String(t.dataset.tab === name)));
    $$(".panel").forEach((p) => (p.hidden = p.dataset.panel !== name));
    if (name !== "table") { prefs.tab = name; savePrefs(); }
  }
  function setCollapsed(on) {
    document.body.classList.toggle("ribbon-collapsed", on);
    $("btnCollapse").setAttribute("aria-pressed", String(on));
    prefs.collapsed = on; savePrefs();
    closePops();
    placeImgbar();
  }
  $$(".rtab").forEach((t) => {
    t.addEventListener("click", () => {
      if (document.body.classList.contains("ribbon-collapsed")) setCollapsed(false);
      showTab(t.dataset.tab);
    });
    t.addEventListener("dblclick", () => setCollapsed(!document.body.classList.contains("ribbon-collapsed")));
  });
  $("btnCollapse").addEventListener("click", () => setCollapsed(!document.body.classList.contains("ribbon-collapsed")));

  // I pulsanti del nastro non devono togliere la selezione dal testo
  $$(".ribbon button, .popover button, .qb, .imgbar button, .navlist").forEach((b) => b.addEventListener("mousedown", (e) => e.preventDefault()));

  /* ---------- Menu a comparsa ---------- */
  let openPop = null;
  function closePops() {
    if (!openPop) return;
    openPop.pop.hidden = true;
    openPop.btn.setAttribute("aria-expanded", "false");
    openPop = null;
    resetGrid();
  }
  function openPopFor(btn) {
    const pop = $(btn.dataset.pop);
    pop.hidden = false;
    const r = btn.getBoundingClientRect();
    pop.style.left = Math.max(8, Math.min(r.left, innerWidth - pop.offsetWidth - 8)) + "px";
    pop.style.top = Math.min(r.bottom + 6, innerHeight - pop.offsetHeight - 8) + "px";
    btn.setAttribute("aria-expanded", "true");
    openPop = { pop, btn };
  }
  $$("[data-pop]").forEach((btn) => {
    btn.setAttribute("aria-haspopup", "true");
    btn.setAttribute("aria-expanded", "false");
    btn.addEventListener("click", () => {
      const wasOpen = openPop && openPop.btn === btn;
      closePops();
      if (!wasOpen) openPopFor(btn);
    });
  });
  document.addEventListener("mousedown", (e) => {
    if (openPop && !openPop.pop.contains(e.target) && !openPop.btn.contains(e.target)) closePops();
  });
  addEventListener("resize", closePops);

  /* ---------- Home: appunti ---------- */
  $$("[data-cmd]").forEach((b) => b.addEventListener("click", () => {
    const c = b.dataset.cmd;
    if (c === "cut" || c === "copy") {
      restoreSelection();
      if (!document.execCommand(c)) toast(c === "cut" ? "Usa ⌘X per tagliare" : "Usa ⌘C per copiare");
      if (c === "cut") changed();
      return;
    }
    exec(c);
  }));

  $("btnPaste").addEventListener("click", async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          const img = it.types.find((t) => t.startsWith("image/"));
          if (img) { insertImageFile(new File([await it.getType(img)], "incollata", { type: img })); return; }
          if (it.types.includes("text/html")) { exec("insertHTML", sanitize(await (await it.getType("text/html")).text())); return; }
          if (it.types.includes("text/plain")) { exec("insertText", await (await it.getType("text/plain")).text()); return; }
        }
      } else {
        exec("insertText", await navigator.clipboard.readText());
      }
    } catch (e) {
      toast("Il browser non permette di incollare da qui: usa ⌘V (Ctrl+V)");
    }
  });

  // Copia formato
  let painterFmt = null;
  function readFormat(el) {
    const s = getComputedStyle(el);
    let bg = null;
    for (let a = el; a && a !== editor; a = a.parentElement) {
      if (BLOCKS.has(a.tagName) || a.tagName === "LI") break;
      const c = getComputedStyle(a).backgroundColor;
      if (c && c !== "rgba(0, 0, 0, 0)" && c !== "transparent") { bg = c; break; }
    }
    return {
      bold: parseInt(s.fontWeight, 10) >= 600,
      italic: s.fontStyle === "italic",
      underline: !!el.closest("u") || hasDeco(el, "underline"),
      strike: hasDeco(el, "line-through"),
      color: s.color, bg,
      font: s.fontFamily.split(",")[0].replace(/["']/g, "").trim(),
      size: Math.round(parseFloat(s.fontSize) * 0.75 * 2) / 2,
    };
  }
  function hasDeco(el, kind) {
    for (let a = el; a && a !== editor; a = a.parentElement) {
      if ((getComputedStyle(a).textDecorationLine || "").includes(kind)) return true;
      if (BLOCKS.has(a.tagName)) break;
    }
    return false;
  }
  function anchorElement() {
    let n = savedRange ? savedRange.startContainer : null;
    if (n && n.nodeType === 3) n = n.parentElement;
    return n && insideEditor(n) ? n : null;
  }
  $("btnPainter").addEventListener("click", () => {
    if (painterFmt) return stopPainter();
    const el = anchorElement();
    if (!el) return toast("Prima metti il cursore nel testo da cui copiare il formato");
    painterFmt = readFormat(el);
    document.body.classList.add("painting");
    $("btnPainter").classList.add("on");
    $("painterHint").hidden = false;
  });
  function stopPainter() {
    painterFmt = null;
    document.body.classList.remove("painting");
    $("btnPainter").classList.remove("on");
    $("painterHint").hidden = true;
  }
  editor.addEventListener("mouseup", () => setTimeout(() => {
    if (!painterFmt) return;
    const sel = getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    savedRange = sel.getRangeAt(0).cloneRange();
    const f = painterFmt;
    stopPainter();
    exec("removeFormat");
    if (f.bold) exec("bold");
    if (f.italic) exec("italic");
    if (f.underline) exec("underline");
    if (f.strike) exec("strikeThrough");
    exec("foreColor", f.color);
    if (f.bg) exec("hiliteColor", f.bg);
    exec("fontName", f.font);
    setFontSize(f.size);
  }, 0));

  /* ---------- Home: carattere ---------- */
  $("fontName").addEventListener("change", (e) => { exec("fontName", e.target.value); });
  $("fontSize").addEventListener("change", (e) => { setFontSize(e.target.value); });
  $("btnGrow").addEventListener("click", () => stepSize(1));
  $("btnShrink").addEventListener("click", () => stepSize(-1));
  function stepSize(dir) {
    const cur = parseFloat($("fontSize").value) || 12;
    const next = dir > 0 ? SIZES.find((s) => s > cur) || cur : [...SIZES].reverse().find((s) => s < cur) || cur;
    setFontSize(next);
  }

  function setFontSize(pt) {
    restoreSelection();
    const sel = getSelection();
    if (!sel.rangeCount) return;
    if (sel.isCollapsed) {
      // Nessun testo selezionato: preparo uno spazio con la nuova dimensione per ciò che verrà scritto.
      document.execCommand("insertHTML", false, `<span style="font-size:${pt}pt">​</span>`);
      const node = sel.anchorNode;
      const txt = node.nodeType === 3 ? node : node.lastChild;
      if (txt && txt.nodeType === 3) {
        const r = document.createRange();
        r.setStart(txt, txt.length); r.collapse(true);
        sel.removeAllRanges(); sel.addRange(r);
      }
    } else {
      document.execCommand("styleWithCSS", false, false);
      document.execCommand("fontSize", false, "7");
      $$('font[size="7"]', editor).forEach((f) => {
        const s = document.createElement("span");
        s.style.fontSize = pt + "pt";
        while (f.firstChild) s.appendChild(f.firstChild);
        f.replaceWith(s);
        $$("span", s).forEach((inner) => { inner.style.fontSize = ""; if (!inner.getAttribute("style")) inner.replaceWith(...inner.childNodes); });
      });
    }
    $("fontSize").value = String(pt);
    keepSelection();
    changed();
  }

  $("foreColor").addEventListener("input", (e) => { $("swText").style.background = e.target.value; exec("foreColor", e.target.value); });
  $("hiliteColor").addEventListener("input", (e) => { $("swHi").style.background = e.target.value; exec("hiliteColor", e.target.value); });

  // Maiuscole / minuscole
  $$("[data-case]").forEach((b) => b.addEventListener("click", () => { closePops(); changeCase(b.dataset.case); }));
  function changeCase(mode) {
    if (!savedRange || savedRange.collapsed) return toast("Seleziona prima il testo");
    const r = savedRange;
    const rootNode = r.commonAncestorContainer;
    const nodes = [];
    if (rootNode.nodeType === 3) nodes.push(rootNode);
    else {
      const w = document.createTreeWalker(rootNode, NodeFilter.SHOW_TEXT);
      while (w.nextNode()) if (r.intersectsNode(w.currentNode)) nodes.push(w.currentNode);
    }
    const sc = r.startContainer, so = r.startOffset, ec = r.endContainer, eo = r.endOffset;
    let capNext = true, prevLetter = false;
    for (const n of nodes) {
      const s = n === sc ? so : 0;
      const e = n === ec ? eo : n.length;
      const part = n.nodeValue.slice(s, e);
      let out = "";
      for (const ch of part) {
        const isL = /\p{L}/u.test(ch);
        if (mode === "upper") out += ch.toLocaleUpperCase("it");
        else if (mode === "lower") out += ch.toLocaleLowerCase("it");
        else if (mode === "title") out += isL && !prevLetter ? ch.toLocaleUpperCase("it") : ch.toLocaleLowerCase("it");
        else { // come una frase
          if (isL) { out += capNext ? ch.toLocaleUpperCase("it") : ch.toLocaleLowerCase("it"); capNext = false; }
          else { out += ch; if (/[.!?]/.test(ch)) capNext = true; }
        }
        prevLetter = isL || ch === "'" || ch === "’";
      }
      if (out.length === part.length) n.nodeValue = n.nodeValue.slice(0, s) + out + n.nodeValue.slice(e);
    }
    // la selezione resta sullo stesso testo
    try {
      const nr = document.createRange();
      nr.setStart(sc, so); nr.setEnd(ec, eo);
      savedRange = nr;
    } catch (e) { /* ignoro */ }
    restoreSelection();
    changed();
  }

  /* ---------- Home: paragrafo ---------- */
  $("btnSort").addEventListener("click", () => {
    const el = anchorElement();
    const coll = new Intl.Collator("it", { sensitivity: "base", numeric: true });
    const li = el && el.closest("li");
    let items;
    if (li) items = [...li.parentElement.children].filter((c) => c.tagName === "LI");
    else {
      items = selectedBlocks().filter((b) => b.parentElement === editor);
      if (items.length < 2) return toast("Seleziona almeno due paragrafi, oppure metti il cursore in un elenco");
    }
    const sorted = [...items].sort((a, b) => coll.compare(a.textContent.trim(), b.textContent.trim()));
    // reinserisco nello stesso punto
    const marker = document.createComment("ordina");
    items[items.length - 1].after(marker);
    sorted.forEach((n) => marker.before(n));
    marker.remove();
    changed();
    toast("Ordinati dalla A alla Z");
  });

  function setMarks(on) {
    editor.classList.toggle("marks", on);
    $("btnMarks").classList.toggle("on", on);
    $("btnMarks2").classList.toggle("on", on);
    prefs.marks = on; savePrefs();
  }
  $("btnMarks").addEventListener("click", () => setMarks(!editor.classList.contains("marks")));
  $("btnMarks2").addEventListener("click", () => setMarks(!editor.classList.contains("marks")));

  $$("[data-lh]").forEach((b) => b.addEventListener("click", () => {
    closePops();
    const blocks = selectedBlocks();
    if (!blocks.length) return toast("Metti il cursore nel paragrafo");
    blocks.forEach((bl) => (bl.style.lineHeight = b.dataset.lh));
    changed();
  }));

  $("shadeColor").addEventListener("input", (e) => {
    $("swShade").style.background = e.target.value;
    const blocks = selectedBlocks();
    if (!blocks.length) return toast("Metti il cursore nel paragrafo");
    blocks.forEach((bl) => (bl.style.backgroundColor = e.target.value.toLowerCase() === "#ffffff" ? "" : e.target.value));
    changed();
  });

  $$("[data-border]").forEach((b) => b.addEventListener("click", () => {
    closePops();
    const blocks = selectedBlocks();
    if (!blocks.length) return toast("Metti il cursore nel paragrafo");
    const line = "1px solid #141a26";
    blocks.forEach((bl) => {
      if (b.dataset.border === "noshade") { bl.style.backgroundColor = ""; return; }
      ["border", "borderTop", "borderBottom", "padding", "paddingTop", "paddingBottom"].forEach((p) => (bl.style[p] = ""));
      if (b.dataset.border === "bottom") { bl.style.borderBottom = line; bl.style.paddingBottom = "2px"; }
      if (b.dataset.border === "top") { bl.style.borderTop = line; bl.style.paddingTop = "2px"; }
      if (b.dataset.border === "box") { bl.style.border = line; bl.style.padding = "4px 6px"; }
    });
    changed();
  }));

  /* ---------- Home: stili ---------- */
  function applyStyle(s) {
    if (s === "em") exec("italic");
    else exec("formatBlock", `<${s}>`);
    refreshToolbar();
  }
  $$(".style-tile").forEach((t) => {
    t.addEventListener("mousedown", (e) => e.preventDefault());
    t.addEventListener("click", () => applyStyle(t.dataset.style));
  });

  /* ---------- Stato dei pulsanti ---------- */
  const STATE_CMDS = ["bold", "italic", "underline", "strikeThrough", "subscript", "superscript", "justifyLeft", "justifyCenter", "justifyRight", "justifyFull", "insertUnorderedList", "insertOrderedList"];
  function refreshToolbar() {
    STATE_CMDS.forEach((c) => {
      let on = false;
      try { on = document.queryCommandState(c); } catch (e) { /* non supportato */ }
      $$(`[data-cmd="${c}"]`).forEach((b) => { b.classList.toggle("on", on); b.setAttribute("aria-pressed", on); });
    });
    const el = anchorElement();
    if (!el) return;

    const block = el.closest("h1,h2,h3,blockquote,p,li,td,th,div");
    let tag = block && insideEditor(block) && block !== editor ? block.tagName.toLowerCase() : "p";
    const bq = el.closest("blockquote:not([style])");
    if (bq && insideEditor(bq)) tag = "blockquote";
    const italic = getComputedStyle(el).fontStyle === "italic";
    $$(".style-tile").forEach((t) => {
      const s = t.dataset.style;
      t.classList.toggle("on", s === tag || (s === "em" && italic && tag === "p"));
    });

    const st = getComputedStyle(el);
    const fam = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    const fn = $("fontName");
    if (fam && ![...fn.options].some((o) => o.value === fam)) fn.add(new Option(fam));
    fn.value = fam;
    const ptS = String(Math.round(parseFloat(st.fontSize) * 0.75 * 2) / 2);
    const fs = $("fontSize");
    if (![...fs.options].some((o) => o.value === ptS)) fs.add(new Option(ptS));
    fs.value = ptS;

    const inTable = !!el.closest("td,th");
    $("tabTable").hidden = !inTable;
    if (!inTable && $("tabTable").getAttribute("aria-selected") === "true") showTab(prefs.tab || "home");
  }

  /* ---------- Inserisci: collegamenti ---------- */
  const dlgLink = $("dlgLink");
  let linkCtx = null;
  function normalizeUrl(u) {
    let href = u.trim();
    if (!href) return "";
    if (/^[a-z][a-z0-9+.-]*:/i.test(href.replace(/[\u0000- ]/g, "")) && !SAFE_URL.test(href)) return null;
    if (!SAFE_URL.test(href)) href = (/^[^\s@/]+@[^\s@/]+\.[^\s@/]+$/.test(href) ? "mailto:" : "https://") + href;
    return href;
  }
  function addLink() {
    const range = savedRange && insideEditor(savedRange.startContainer) ? savedRange.cloneRange() : null;
    const a = anchorElement()?.closest("a");
    const existing = a && insideEditor(a) ? a : null;
    linkCtx = { range, existing };
    $("linkText").value = existing ? existing.textContent : (range && !range.collapsed ? range.toString() : "");
    $("linkUrl").value = existing ? existing.getAttribute("href") || "" : "";
    $("linkRemove").hidden = !existing;
    $("linkOk").textContent = existing ? "Aggiorna" : "Inserisci";
    dlgLink.returnValue = "";
    dlgLink.showModal();
    ($("linkText").value ? $("linkUrl") : $("linkText")).focus();
  }
  $("linkRemove").addEventListener("click", () => dlgLink.close("remove"));
  $("linkCancel").addEventListener("click", () => dlgLink.close("cancel"));
  dlgLink.addEventListener("close", () => {
    const ctx = linkCtx; linkCtx = null;
    if (!ctx) return;
    const v = dlgLink.returnValue;
    if (ctx.range) savedRange = ctx.range;
    if (v === "remove" && ctx.existing) {
      ctx.existing.replaceWith(...ctx.existing.childNodes);
      restoreSelection(); changed();
      return;
    }
    if (v !== "ok") { restoreSelection(); return; }
    const text = $("linkText").value.trim();
    const href = normalizeUrl($("linkUrl").value);
    if (href === null) return toast("Questo tipo di indirizzo non è permesso");
    if (!href) {
      if (ctx.existing) { ctx.existing.replaceWith(...ctx.existing.childNodes); changed(); }
      restoreSelection();
      return;
    }
    if (ctx.existing) {
      ctx.existing.setAttribute("href", href);
      if (text && text !== ctx.existing.textContent) ctx.existing.textContent = text;
      restoreSelection(); changed();
      return;
    }
    const selText = ctx.range && !ctx.range.collapsed ? ctx.range.toString() : "";
    if (selText && (!text || text === selText)) exec("createLink", href);
    else exec("insertHTML", `<a href="${esc(href)}">${esc(text || $("linkUrl").value.trim())}</a>`);
  });
  $("btnLink").addEventListener("click", addLink);
  editor.addEventListener("click", (e) => {
    const a = e.target.closest("a");
    if (a && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      const href = a.getAttribute("href") || "";
      if (SAFE_URL.test(href) && !href.startsWith("#")) window.open(href, "_blank", "noopener");
    }
  });

  /* ---------- Inserisci: immagini ---------- */
  $("btnImage").addEventListener("click", () => $("fileImage").click());
  $("fileImage").addEventListener("change", (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (f) insertImageFile(f);
  });
  async function insertImageFile(f) {
    if (!f.type.startsWith("image/")) return;
    const url = await shrinkImage(f);
    exec("insertHTML", `<img src="${url}" alt="" style="width:50%">`);
  }
  function shrinkImage(f) {
    return new Promise((resolve) => {
      const rd = new FileReader();
      rd.onload = () => {
        const img = new Image();
        img.onload = () => {
          const k = Math.min(1, IMG_MAX / Math.max(img.width, img.height));
          if (k === 1 && f.size < 600000) return resolve(rd.result);
          const c = document.createElement("canvas");
          c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
          c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
          resolve(c.toDataURL(f.type === "image/png" ? "image/png" : "image/jpeg", 0.85));
        };
        img.onerror = () => resolve(rd.result);
        img.src = rd.result;
      };
      rd.readAsDataURL(f);
    });
  }

  let picked = null;
  const imgbar = $("imgbar");
  function pickImage(img) {
    if (picked) picked.classList.remove("picked");
    picked = img;
    if (!img) { imgbar.hidden = true; return; }
    img.classList.add("picked");
    imgbar.hidden = false;
    placeImgbar();
  }
  function placeImgbar() {
    if (!picked) return;
    if (!document.contains(picked)) { pickImage(null); return; }
    const r = picked.getBoundingClientRect();
    imgbar.style.left = Math.max(8, Math.min(innerWidth - imgbar.offsetWidth - 8, r.left)) + "px";
    imgbar.style.top = Math.max(8, Math.min(innerHeight - 50, r.bottom + 10)) + "px";
  }
  editor.addEventListener("click", (e) => pickImage(e.target.tagName === "IMG" ? e.target : null));
  desk.addEventListener("scroll", placeImgbar, { passive: true });
  addEventListener("resize", placeImgbar);
  imgbar.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b || !picked) return;
    const { w, align } = b.dataset;
    if (w === "del") { picked.remove(); pickImage(null); }
    else if (w) { picked.style.width = w + "%"; picked.style.height = "auto"; requestAnimationFrame(placeImgbar); }
    else if (align) {
      let block = picked.parentElement.closest("p,h1,h2,h3,h4,h5,h6,li,td,th,div,blockquote");
      if (!block || block === editor || !insideEditor(block)) {
        block = document.createElement("p");
        picked.replaceWith(block);
        block.appendChild(picked);
      }
      block.style.textAlign = align === "left" ? "" : align;
      requestAnimationFrame(placeImgbar);
    }
    changed();
  });

  editor.addEventListener("drop", (e) => {
    const f = [...(e.dataTransfer?.files || [])].find((x) => x.type.startsWith("image/"));
    if (!f) return;
    e.preventDefault();
    const pos = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
    if (pos) savedRange = pos;
    insertImageFile(f);
  });
  editor.addEventListener("paste", (e) => {
    const f = [...(e.clipboardData?.files || [])].find((x) => x.type.startsWith("image/"));
    if (f) { e.preventDefault(); insertImageFile(f); return; }
    const html = e.clipboardData?.getData("text/html");
    if (html) { e.preventDefault(); exec("insertHTML", sanitize(html)); }
  });

  /* ---------- Inserisci: tabelle ---------- */
  const grid = $("tableGrid"), tLabel = $("tableLabel");
  for (let r = 1; r <= 8; r++) for (let c = 1; c <= 8; c++) {
    const s = document.createElement("span");
    s.dataset.r = r; s.dataset.c = c;
    grid.appendChild(s);
  }
  function resetGrid() {
    $$(".hot", grid).forEach((s) => s.classList.remove("hot"));
    tLabel.textContent = "Scegli righe e colonne";
  }
  grid.addEventListener("mouseover", (e) => {
    const { r, c } = e.target.dataset;
    if (!r) return;
    $$("span", grid).forEach((s) => s.classList.toggle("hot", +s.dataset.r <= +r && +s.dataset.c <= +c));
    tLabel.textContent = `${r} × ${c}`;
  });
  grid.addEventListener("mousedown", (e) => e.preventDefault());
  grid.addEventListener("click", (e) => {
    const { r, c } = e.target.dataset;
    if (!r) return;
    closePops();
    insertTable(+r, +c);
  });
  // Inserisce un blocco al livello principale del documento, dopo il paragrafo del cursore
  function insertBlocks(...nodes) {
    let n = savedRange && insideEditor(savedRange.startContainer) ? savedRange.startContainer : null;
    while (n && n.parentNode !== editor) n = n.parentNode;
    if (!n) editor.append(...nodes);
    else if (n.nodeType === 1 && n.tagName === "P" && !n.textContent.trim() && !n.querySelector("img,table")) n.replaceWith(...nodes);
    else n.after(...nodes);
  }
  function insertTable(rows, cols) {
    const table = document.createElement("table");
    const body = table.createTBody();
    for (let i = 0; i < rows; i++) {
      const tr = body.insertRow();
      for (let j = 0; j < cols; j++) tr.insertCell().innerHTML = "<br>";
    }
    const after = document.createElement("p");
    after.innerHTML = "<br>";
    insertBlocks(table, after);
    setCaret(table.rows[0].cells[0], 0);
    changed();
    refreshToolbar();
    showTab("table");
  }

  function currentCell() {
    const el = anchorElement();
    const cell = el && el.closest("td,th");
    return cell && insideEditor(cell) ? cell : null;
  }
  $$("[data-table]").forEach((b) => b.addEventListener("click", () => {
    const cell = currentCell();
    if (!cell) return toast("Metti il cursore in una cella della tabella");
    const row = cell.parentElement, table = cell.closest("table"), idx = cell.cellIndex;
    const blank = (tag = "td") => { const c = document.createElement(tag); c.innerHTML = "<br>"; return c; };
    let focusCell = cell;
    switch (b.dataset.table) {
      case "rowAbove": case "rowBelow": {
        const nr = document.createElement("tr");
        for (let i = 0; i < row.cells.length; i++) nr.appendChild(blank());
        b.dataset.table === "rowAbove" ? row.before(nr) : row.after(nr);
        focusCell = nr.cells[Math.min(idx, nr.cells.length - 1)];
        break;
      }
      case "colLeft": case "colRight":
        for (const tr of table.rows) {
          const ref = tr.cells[Math.min(idx, tr.cells.length - 1)];
          const nc = blank(ref ? ref.tagName.toLowerCase() : "td");
          if (!ref) tr.appendChild(nc);
          else b.dataset.table === "colLeft" ? ref.before(nc) : ref.after(nc);
          if (tr === row) focusCell = nc;
        }
        break;
      case "delRow": {
        const sib = row.nextElementSibling || row.previousElementSibling;
        row.remove();
        if (!table.rows.length) table.remove();
        focusCell = sib ? sib.cells[Math.min(idx, sib.cells.length - 1)] : null;
        break;
      }
      case "delCol":
        for (const tr of table.rows) if (tr.cells[idx]) tr.cells[idx].remove();
        if (!table.rows[0] || !table.rows[0].cells.length) { table.remove(); focusCell = null; }
        else focusCell = row.cells[Math.min(idx, row.cells.length - 1)];
        break;
      case "delTable": {
        const next = table.nextElementSibling;
        table.remove();
        focusCell = null;
        if (next) setCaret(next, 0);
        break;
      }
      case "header": {
        const first = table.rows[0];
        const toTh = first.cells[0].tagName === "TD";
        [...first.cells].forEach((c) => {
          const n = document.createElement(toTh ? "th" : "td");
          while (c.firstChild) n.appendChild(c.firstChild);
          if (!n.firstChild) n.innerHTML = "<br>";
          c.replaceWith(n);
          if (c === cell) focusCell = n;
        });
        break;
      }
    }
    if (focusCell && document.contains(focusCell)) setCaret(focusCell, 0);
    else { $("tabTable").hidden = true; showTab(prefs.tab || "home"); }
    changed();
  }));

  /* ---------- Inserisci: altro ---------- */
  $("btnPageBreak").addEventListener("click", insertPageBreak);
  function insertPageBreak() {
    const hr = document.createElement("hr");
    hr.className = "page-break";
    const p = document.createElement("p");
    p.innerHTML = "<br>";
    insertBlocks(hr, p);
    setCaret(p, 0);
    changed();
  }
  $("btnHr").addEventListener("click", () => exec("insertHTML", "<hr><p><br></p>"));
  $("btnDate").addEventListener("click", () => exec("insertText", new Date().toLocaleDateString("it-IT", { day: "numeric", month: "long", year: "numeric" })));
  const SYMBOLS = "€ £ $ ¥ © ® ™ § ¶ ° ± × ÷ ≠ ≈ ≤ ≥ ∞ √ π Ω α β µ ½ ¼ ¾ ← → ↑ ↓ « » “ ” ‘ ’ – — … • ✓ ✗ ★ ☆ ♥ ☎".split(" ");
  SYMBOLS.forEach((s) => {
    const b = document.createElement("button");
    b.textContent = s; b.title = s;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => { closePops(); exec("insertText", s); });
    $("symbolGrid").appendChild(b);
  });

  // Sommario con i titoli del documento
  $("btnToc").addEventListener("click", () => {
    const heads = $$("h1,h2,h3", editor).filter((h) => h.textContent.trim() && !h.closest(".toc"));
    if (!heads.length) return toast("Non ci sono titoli: usa gli stili Titolo 1, 2 e 3");
    $$(".toc", editor).forEach((t) => t.remove()); // un solo sommario: lo rigenero
    const box = document.createElement("div");
    box.className = "toc";
    const head = document.createElement("p");
    head.innerHTML = '<b style="font-size:14pt">Sommario</b>';
    box.appendChild(head);
    heads.forEach((h) => {
      const p = document.createElement("p");
      const lv = +h.tagName[1];
      p.style.marginLeft = (lv - 1) * 1.5 + "em";
      p.style.marginBottom = ".2em";
      p.textContent = h.textContent.replace(/​/g, "").trim();
      if (lv === 1) p.style.fontWeight = "700";
      box.appendChild(p);
    });
    const after = document.createElement("p");
    after.innerHTML = "<br>";
    insertBlocks(box, after);
    setCaret(after, 0);
    changed();
    toast("Sommario inserito: premilo di nuovo per aggiornarlo");
  });

  /* ---------- Layout ---------- */
  $$("[data-margins]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.margins = b.dataset.margins; applyPage(); changed(); }));
  $$("[data-orient]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.orient = b.dataset.orient; applyPage(); changed(); }));
  $$("[data-size]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.size = b.dataset.size; applyPage(); changed(); }));

  /* ---------- Revisione ---------- */
  function setSpell(on, refresh = true) {
    editor.spellcheck = on;
    $("btnSpell").classList.toggle("on", on);
    prefs.spell = on; savePrefs();
    if (refresh) {
      // forzo l'aggiornamento delle sottolineature
      pickImage(null);
      const t = editor.innerHTML; editor.innerHTML = t;
      savedRange = null;
      changed();
    }
  }
  $("btnSpell").addEventListener("click", () => setSpell(!editor.spellcheck));

  function textStats(root = editor) {
    const text = (root.innerText ?? root.textContent).replace(/​/g, "");
    return {
      words: (text.match(WORD_RE) || []).length,
      chars: text.replace(/\n/g, "").length,
      charsNoSpace: text.replace(/\s/g, "").length,
      paras: $$("p,h1,h2,h3,h4,h5,h6,li,blockquote", editor).filter((p) => p.textContent.replace(/​/g, "").trim() && !p.querySelector("p,li,h1,h2,h3")).length,
    };
  }
  // Pagine stimate misurando l'altezza del testo rispetto all'area utile del foglio
  function pageCount() {
    const ch = $("pageProbe").getBoundingClientRect().height;
    if (!ch || !editor.lastElementChild) return 1;
    const top = editor.getBoundingClientRect().top;
    const last = editor.lastElementChild.getBoundingClientRect();
    const cuts = $$("hr.page-break", editor).filter((h) => h.parentElement === editor).map((h) => h.getBoundingClientRect().top);
    const points = [top, ...cuts, Math.max(top, last.bottom)];
    let pages = 0;
    for (let i = 1; i < points.length; i++) pages += Math.max(1, Math.ceil((points[i] - points[i - 1] - 1) / ch));
    return pages;
  }
  const readMinutes = (words) => Math.max(1, Math.round(words / 200));
  function updateStats() {
    const s = textStats();
    const pages = pageCount();
    let txt = `${plural(pages, "pagina", "pagine")} · ${plural(s.words, "parola", "parole")} · ${plural(s.charsNoSpace, "carattere", "caratteri")}`;
    if (s.words) txt += ` · lettura ${readMinutes(s.words)} min`;
    $("stats").textContent = txt;
  }
  function showStats() {
    const s = textStats();
    const rows = [
      ["Pagine", fmtInt(pageCount())],
      ["Parole", fmtInt(s.words)],
      ["Caratteri (spazi esclusi)", fmtInt(s.charsNoSpace)],
      ["Caratteri (spazi inclusi)", fmtInt(s.chars)],
      ["Paragrafi", fmtInt(s.paras)],
      ["Tempo di lettura", s.words ? `circa ${readMinutes(s.words)} min` : "—"],
    ];
    if (savedRange && !savedRange.collapsed) {
      const div = document.createElement("div");
      div.appendChild(savedRange.cloneContents());
      const words = (div.textContent.match(WORD_RE) || []).length;
      rows.unshift(["Parole selezionate", fmtInt(words)]);
    }
    $("statsGrid").innerHTML = rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join("");
    $("dlgStats").showModal();
  }
  $("btnCount").addEventListener("click", showStats);

  /* ---------- Trova e sostituisci ---------- */
  const findbar = $("findbar"), findInput = $("findInput"), replInput = $("replaceInput"), findCount = $("findCount");
  function openFind(replace) {
    findbar.hidden = false;
    const s = getSelection().toString();
    if (s && s.length < 80 && !s.includes("\n")) findInput.value = s;
    if (replace && findInput.value) { replInput.focus(); replInput.select(); }
    else { findInput.focus(); findInput.select(); }
    countMatches();
  }
  function closeFind() {
    findbar.hidden = true;
    clearHits();
    restoreSelection();
  }
  $("btnFind").addEventListener("click", () => openFind(false));
  $("btnFind2").addEventListener("click", () => openFind(false));
  $("btnReplace").addEventListener("click", () => openFind(true));
  $("btnReplace2").addEventListener("click", () => openFind(true));
  $("findClose").addEventListener("click", closeFind);

  function textNodes() {
    const w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    const out = [];
    while (w.nextNode()) out.push(w.currentNode);
    return out;
  }
  function allMatches(q) {
    if (!q) return [];
    const ql = q.toLocaleLowerCase(), res = [];
    for (const n of textNodes()) {
      const t = n.nodeValue.toLocaleLowerCase();
      let i = t.indexOf(ql);
      while (i !== -1) { res.push({ node: n, start: i, end: i + q.length }); i = t.indexOf(ql, i + q.length); }
    }
    return res;
  }
  function matchRange(m) {
    const r = document.createRange();
    r.setStart(m.node, m.start); r.setEnd(m.node, m.end);
    return r;
  }
  // Evidenzia tutti i risultati (e più forte quello attuale) senza toccare il documento
  const HL = window.Highlight && window.CSS && CSS.highlights;
  function paintHits(ms, cur) {
    if (!HL) return;
    if (ms.length) CSS.highlights.set("word-find", new Highlight(...ms.map(matchRange)));
    else CSS.highlights.delete("word-find");
    if (cur >= 0 && ms[cur]) CSS.highlights.set("word-find-cur", new Highlight(matchRange(ms[cur])));
    else CSS.highlights.delete("word-find-cur");
  }
  function clearHits() { if (HL) { CSS.highlights.delete("word-find"); CSS.highlights.delete("word-find-cur"); } }
  function countMatches(cur) {
    const ms = allMatches(findInput.value), n = ms.length;
    if (cur === undefined) {
      // il risultato attuale è quello già selezionato, se c'è
      const r = savedRange;
      cur = r ? ms.findIndex((m) => r.startContainer === m.node && r.endContainer === m.node && r.startOffset === m.start && r.endOffset === m.end) : -1;
    }
    findCount.textContent = !findInput.value ? "" : !n ? "Nessun risultato" : cur >= 0 ? `${cur + 1} di ${n}` : (n === 1 ? "1 risultato" : `${n} risultati`);
    paintHits(ms, cur);
    return n;
  }
  function selectMatch(m) {
    const r = matchRange(m);
    const sel = getSelection();
    editor.focus({ preventScroll: true });
    sel.removeAllRanges(); sel.addRange(r);
    savedRange = r.cloneRange();
    const el = m.node.parentElement;
    const er = el.getBoundingClientRect(), dr = desk.getBoundingClientRect();
    if (er.top < dr.top + 40 || er.bottom > dr.bottom - 40) el.scrollIntoView({ block: "center", behavior: "smooth" });
  }
  function findStep(dir = 1) {
    const ms = allMatches(findInput.value);
    if (!ms.length) { countMatches(); return false; }
    let idx = dir > 0 ? 0 : ms.length - 1;
    let wrapped = true;
    if (savedRange && insideEditor(savedRange.startContainer)) {
      if (dir > 0) {
        // il primo risultato che comincia dopo la fine della selezione attuale
        const i = ms.findIndex((m) => matchRange(m).compareBoundaryPoints(Range.END_TO_START, savedRange) >= 0);
        if (i !== -1) { idx = i; wrapped = false; }
      } else {
        // l'ultimo risultato che finisce prima dell'inizio della selezione attuale
        let i = -1;
        ms.forEach((m, k) => { if (matchRange(m).compareBoundaryPoints(Range.START_TO_END, savedRange) <= 0) i = k; });
        if (i !== -1) { idx = i; wrapped = false; }
      }
    } else wrapped = false;
    if (wrapped && ms.length > 1) toast(dir > 0 ? "Ricomincio dall'inizio del documento" : "Ricomincio dalla fine del documento", 1600);
    selectMatch(ms[idx]);
    countMatches(idx);
    return true;
  }
  findInput.addEventListener("input", () => countMatches());
  findInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); findInput.focus(); }
    if (e.key === "Escape") { e.preventDefault(); closeFind(); }
  });
  replInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); $("replaceOne").click(); replInput.focus(); }
    if (e.key === "Escape") { e.preventDefault(); closeFind(); }
  });
  $("findNext").addEventListener("click", () => { findStep(1); findInput.focus(); });
  $("findPrev").addEventListener("click", () => { findStep(-1); findInput.focus(); });
  // Sostituisce il testo selezionato passando dall'editor, così ⌘Z lo può annullare
  function replaceRange(r, text) {
    const sel = getSelection();
    editor.focus({ preventScroll: true });
    sel.removeAllRanges(); sel.addRange(r);
    if (text) document.execCommand("insertText", false, text);
    else document.execCommand("delete");
  }
  $("replaceOne").addEventListener("click", () => {
    const q = findInput.value;
    if (!q) return;
    const cur = savedRange ? savedRange.toString() : "";
    if (cur.toLocaleLowerCase() !== q.toLocaleLowerCase()) { findStep(1); return; }
    replaceRange(savedRange, replInput.value);
    savedRange = getSelection().getRangeAt(0).cloneRange();
    changed();
    if (!findStep(1)) toast("Fatto: non ci sono altre occorrenze");
  });
  $("replaceAll").addEventListener("click", () => {
    const q = findInput.value;
    if (!q) return;
    const ms = allMatches(q);
    const ql = q.toLocaleLowerCase();
    let n = 0;
    // dall'ultima alla prima, così le posizioni delle precedenti restano valide
    for (let i = ms.length - 1; i >= 0; i--) {
      const m = ms[i];
      if (!document.contains(m.node) || m.node.nodeValue.slice(m.start, m.end).toLocaleLowerCase() !== ql) continue;
      replaceRange(matchRange(m), replInput.value);
      n++;
    }
    if (getSelection().rangeCount) savedRange = getSelection().getRangeAt(0).cloneRange();
    toast(n ? `Sostituite ${fmtInt(n)} occorrenze` : "Nessuna occorrenza trovata");
    countMatches();
    changed();
  });

  /* ---------- Visualizza ---------- */
  function setZoom(pct) {
    pct = Math.max(50, Math.min(200, Math.round(pct)));
    page.style.setProperty("--zoom", pct / 100);
    $("zoom").value = pct;
    $("zoomOut").textContent = pct + "%";
    prefs.zoom = pct; savePrefs();
    placeImgbar();
  }
  $("zoom").addEventListener("input", (e) => setZoom(+e.target.value));
  $("zoomMinus").addEventListener("click", () => setZoom(Math.ceil(+$("zoom").value / 10) * 10 - 10));
  $("zoomPlus").addEventListener("click", () => setZoom(Math.floor(+$("zoom").value / 10) * 10 + 10));
  $("zoomOut").addEventListener("click", () => setZoom(100));
  $("btnZoom100").addEventListener("click", () => setZoom(100));
  $("btnFit").addEventListener("click", () => {
    const pw = pageDims().w * 37.795;
    setZoom((((desk.clientWidth - 48) / pw) * 100));
  });

  const THEMES = ["auto", "light", "dark"], THEME_LABEL = { auto: "Tema: auto", light: "Tema: chiaro", dark: "Tema: scuro" };
  function setTheme(t) {
    if (!THEMES.includes(t)) t = "auto";
    if (t === "auto") delete document.documentElement.dataset.theme;
    else document.documentElement.dataset.theme = t;
    $("themeLabel").textContent = THEME_LABEL[t];
    prefs.theme = t; savePrefs();
    if (NATIVE) NATIVE.postMessage({ type: "theme", theme: t });
  }
  $("btnTheme").addEventListener("click", () => setTheme(THEMES[(THEMES.indexOf(prefs.theme || "auto") + 1) % 3]));
  $("btnFull").addEventListener("click", () => {
    if (document.fullscreenElement) document.exitFullscreen();
    else document.documentElement.requestFullscreen?.().catch(() => toast("Schermo intero non disponibile"));
  });
  $("btnKeys").addEventListener("click", () => $("dlgKeys").showModal());

  // Concentrazione: resta solo la pagina
  function setFocus(on) {
    document.body.classList.toggle("focus", on);
    $("focusExit").hidden = !on;
    $("btnFocus").classList.toggle("on", on);
    closePops();
    pickImage(null);
    if (on) { restoreSelection(); toast("Modalità concentrazione: premi Esc per uscire", 2000); }
  }
  $("btnFocus").addEventListener("click", () => setFocus(!document.body.classList.contains("focus")));
  $("focusExit").addEventListener("click", () => setFocus(false));

  // Struttura: elenco dei titoli per spostarsi nel documento
  const navPane = $("navPane"), navList = $("navList");
  let navHeads = [];
  function setNav(on) {
    navPane.hidden = !on;
    $("btnNav").classList.toggle("on", on);
    prefs.nav = on; savePrefs();
    if (on) buildNav();
    placeImgbar();
  }
  function buildNav() {
    navHeads = $$("h1,h2,h3", editor).filter((h) => h.textContent.replace(/​/g, "").trim());
    navList.innerHTML = "";
    navHeads.forEach((h) => {
      const li = document.createElement("li");
      li.className = "lv" + h.tagName[1];
      const b = document.createElement("button");
      b.textContent = h.textContent.replace(/​/g, "").trim();
      b.title = b.textContent;
      b.addEventListener("click", () => {
        h.scrollIntoView({ block: "start", behavior: "smooth" });
        setCaret(h, 0);
      });
      li.appendChild(b);
      navList.appendChild(li);
    });
    $("navEmpty").hidden = navHeads.length > 0;
    navSpy();
  }
  function navSpy() {
    if (navPane.hidden || !navHeads.length) return;
    const limit = desk.getBoundingClientRect().top + 90;
    let cur = 0;
    navHeads.forEach((h, i) => { if (h.getBoundingClientRect().top <= limit) cur = i; });
    $$("button", navList).forEach((b, i) => b.classList.toggle("current", i === cur));
  }
  desk.addEventListener("scroll", navSpy, { passive: true });
  $("btnNav").addEventListener("click", () => setNav(navPane.hidden));
  $("navClose").addEventListener("click", () => setNav(false));

  /* ---------- Ricerca dei comandi (⌘J) ---------- */
  const cmdInput = $("cmdInput"), cmdResults = $("cmdResults");
  const TAB_NAMES = { home: "Home", insert: "Inserisci", layout: "Layout", review: "Revisione", view: "Visualizza", table: "Tabella" };
  let cmdIndex = null, cmdHits = [], cmdSel = 0;
  function buildCmdIndex() {
    const seen = new Set(), list = [];
    const els = $$(".file-actions button, .ribbon .panel button, .ribbon .panel label.tb, .popover.menu button, .quick .qb");
    els.forEach((el) => {
      let label = el.dataset.label || el.getAttribute("title") || el.getAttribute("aria-label") || el.textContent;
      label = label.replace(/\s+/g, " ").trim();
      const keyM = label.match(/\(([^)]*[⌘⌥⇧][^)]*)\)\s*$/);
      const key = keyM ? keyM[1] : "";
      label = label.replace(/\s*\([^)]*\)\s*$/, "").replace(/:.*$/, (m) => (el.dataset.label ? m : "")).trim();
      if (!label || seen.has(label.toLowerCase())) return;
      seen.add(label.toLowerCase());
      const panel = el.closest(".panel");
      const where = panel ? TAB_NAMES[panel.dataset.panel] : el.closest(".file-actions, #popExport") ? "File" : el.closest(".quick") ? "Modifica" : "Menu";
      const icon = el.querySelector("svg use")?.getAttribute("href") || "";
      list.push({ el, label, key, where, icon, hay: fold(label + " " + where + " " + (el.getAttribute("title") || "")) });
    });
    return list;
  }
  function searchCmd(q) {
    cmdIndex = cmdIndex || buildCmdIndex();
    const words = fold(q).split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    return cmdIndex
      .filter((c) => words.every((w) => c.hay.includes(w)))
      .map((c) => ({ c, score: (fold(c.label).startsWith(words[0]) ? 0 : 1) + (c.where === "Tabella" ? 0.5 : 0) }))
      .sort((a, b) => a.score - b.score || a.c.label.localeCompare(b.c.label, "it"))
      .slice(0, 9).map((x) => x.c);
  }
  function renderCmd() {
    const q = cmdInput.value;
    cmdHits = searchCmd(q);
    cmdSel = 0;
    if (!q.trim()) { closeCmd(false); return; }
    cmdResults.hidden = false;
    cmdInput.setAttribute("aria-expanded", "true");
    if (!cmdHits.length) { cmdResults.innerHTML = `<div class="cr-empty">Nessun comando per «${esc(q)}»</div>`; return; }
    cmdResults.innerHTML = cmdHits.map((c, i) => `
      <button role="option" data-i="${i}" aria-selected="${i === 0}">
        ${c.icon ? `<svg><use href="${esc(c.icon)}"/></svg>` : `<svg></svg>`}
        <span class="cr-label">${esc(c.label)}</span>
        <span class="cr-where">${esc(c.key || c.where)}</span>
      </button>`).join("");
  }
  function moveCmd(d) {
    if (!cmdHits.length) return;
    cmdSel = (cmdSel + d + cmdHits.length) % cmdHits.length;
    $$("button", cmdResults).forEach((b, i) => b.setAttribute("aria-selected", String(i === cmdSel)));
    $$("button", cmdResults)[cmdSel].scrollIntoView({ block: "nearest" });
  }
  function closeCmd(clear = true) {
    cmdResults.hidden = true;
    cmdInput.setAttribute("aria-expanded", "false");
    if (clear) cmdInput.value = "";
  }
  function runCmd(c) {
    closeCmd();
    cmdInput.blur();
    const panel = c.el.closest(".panel");
    if (panel) {
      if (document.body.classList.contains("ribbon-collapsed")) setCollapsed(false);
      showTab(panel.dataset.panel);
    }
    if (c.el.dataset.pop) { closePops(); requestAnimationFrame(() => openPopFor(c.el)); return; }
    if (c.el.matches("label.tb")) { restoreSelection(); c.el.querySelector("input").click(); return; }
    c.el.click();
  }
  cmdInput.addEventListener("input", renderCmd);
  cmdInput.addEventListener("focus", () => { if (cmdInput.value.trim()) renderCmd(); });
  cmdInput.addEventListener("blur", () => setTimeout(() => closeCmd(false), 150));
  cmdInput.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); moveCmd(1); }
    else if (e.key === "ArrowUp") { e.preventDefault(); moveCmd(-1); }
    else if (e.key === "Enter") { e.preventDefault(); if (cmdHits[cmdSel]) runCmd(cmdHits[cmdSel]); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeCmd(); restoreSelection(); }
  });
  cmdResults.addEventListener("mousedown", (e) => e.preventDefault());
  cmdResults.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-i]");
    if (b) runCmd(cmdHits[+b.dataset.i]);
  });
  function focusCmd() {
    if (getComputedStyle($("cmdSearch")).display === "none" || document.body.classList.contains("focus")) {
      toast("La ricerca dei comandi è nella barra in alto (finestra più larga)");
      return;
    }
    cmdInput.focus(); cmdInput.select();
  }

  /* ---------- File: nuovo, apri, scarica, esporta ---------- */
  function fileBase() {
    return (titleEl.value.trim() || "Documento").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "-").slice(0, 120);
  }
  function blobToBase64(blob) {
    return new Promise((res, rej) => {
      const rd = new FileReader();
      rd.onload = () => res(String(rd.result).split(",")[1] || "");
      rd.onerror = rej;
      rd.readAsDataURL(blob);
    });
  }
  async function deliver(blob, fileName, okMsg) {
    if (NATIVE) {
      // App per Mac: il file lo salva l'app, con la finestra «Salva» di macOS
      NATIVE.postMessage({ type: "save", name: fileName, data: await blobToBase64(blob) });
      return;
    }
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    toast(okMsg || `Scaricato «${fileName}»`);
  }

  $("btnNew").addEventListener("click", () => {
    if (editor.textContent.trim() && !confirm("Iniziare un nuovo documento? Quello attuale verrà sostituito: se ti serve, scaricalo prima in .docx.")) return;
    editor.innerHTML = "<p><br></p>";
    titleEl.value = UNTITLED;
    pageSet = { size: "A4", orient: "portrait", margins: "normal" };
    applyPage();
    pickImage(null);
    setCaret(editor.firstChild, 0);
    changed();
    desk.scrollTop = 0;
  });

  $("btnOpen").addEventListener("click", () => $("fileOpen").click());
  $("fileOpen").addEventListener("change", (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (f) openFile(f);
  });
  async function openFile(f) {
    if (editor.textContent.trim() && !confirm(`Aprire «${f.name}»? Il documento attuale verrà sostituito.`)) return;
    try {
      let html;
      const name = f.name.toLowerCase();
      if (name.endsWith(".docx")) {
        if (!window.mammoth) throw new Error("libreria non caricata");
        html = (await window.mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() })).value;
      } else if (name.endsWith(".txt")) {
        html = (await f.text()).split(/\r?\n/).map((l) => `<p>${esc(l) || "<br>"}</p>`).join("");
      } else if (/\.(md|markdown)$/.test(name)) {
        html = mdToHtml(await f.text());
      } else {
        const text = await f.text();
        const doc = new DOMParser().parseFromString(text, "text/html");
        html = doc.body ? doc.body.innerHTML : text;
      }
      editor.innerHTML = sanitize(html) || "<p><br></p>";
      if (!editor.firstElementChild) editor.innerHTML = `<p>${editor.innerHTML || "<br>"}</p>`;
      titleEl.value = f.name.replace(/\.[^.]+$/, "");
      pickImage(null);
      savedRange = null;
      changed();
      desk.scrollTop = 0;
      if (!navPane.hidden) buildNav();
      toast(`Aperto «${f.name}»`);
    } catch (err) {
      toast("Non riesco ad aprire questo file: " + err.message);
    }
  }

  // Prepara l'editor «pulito» per l'esportazione (senza zoom, segni, immagine scelta)
  async function withCleanEditor(fn) {
    pickImage(null);
    const marks = editor.classList.contains("marks");
    editor.classList.remove("marks");
    const z = page.style.getPropertyValue("--zoom");
    page.style.setProperty("--zoom", "1");
    try { return await fn(); }
    finally {
      page.style.setProperty("--zoom", z || "1");
      editor.classList.toggle("marks", marks);
    }
  }

  let saving = false;
  async function saveDocx() {
    if (saving) return;
    if (!window.WordDocx || !window.docx) return toast("Attendi un attimo: sto ancora caricando i componenti.");
    saving = true;
    const btn = $("btnSave");
    btn.disabled = true; $("saveLabel").textContent = "Preparo il file…";
    try {
      await withCleanEditor(async () => {
        const { w, h, mt, ml } = pageDims();
        const blob = await window.WordDocx.toBlob(editor, titleEl.value, { width: Math.min(w, h), height: Math.max(w, h), landscape: pageSet.orient === "landscape", top: mt, side: ml });
        await deliver(blob, fileBase() + ".docx", "File .docx scaricato");
      });
      persist();
    } catch (err) {
      console.error(err);
      toast("Errore durante il salvataggio: " + err.message);
    } finally {
      saving = false;
      btn.disabled = false; $("saveLabel").textContent = SAVE_LABEL;
    }
  }
  $("btnSave").addEventListener("click", saveDocx);

  $("btnPdf").addEventListener("click", () => {
    closePops();
    pickImage(null);
    if (NATIVE) { NATIVE.postMessage({ type: "print", title: titleEl.value || "Documento" }); return; }
    const t = document.title;
    document.title = titleEl.value || "Documento";
    toast("Nella finestra di stampa scegli «Salva come PDF»");
    setTimeout(() => { print(); document.title = t; }, 300);
  });

  function exportHtml() {
    const { w, mt, ml } = pageDims();
    const css = `body{margin:0;background:#eef0f4;font-family:-apple-system,system-ui,sans-serif}
main{box-sizing:border-box;max-width:${w}cm;margin:32px auto;background:#fff;padding:${mt}cm ${ml}cm;box-shadow:0 10px 40px -12px rgba(0,0,0,.25);font:12pt/1.5 Georgia,serif;color:#141a26}
p{margin:0 0 .6em}h1{font-size:22pt;margin:.4em 0;line-height:1.2;color:#17306b}h2{font-size:16pt;margin:.9em 0 .35em;color:#1f3f86}h3{font-size:13pt;margin:.8em 0 .3em;color:#2a4a92}
blockquote{margin:.6em 0;padding:.1em 0 .1em 1em;border-left:3px solid #2457d6;color:#4b5468;font-style:italic}a{color:#2457d6}img{max-width:100%;height:auto}
table{border-collapse:collapse;width:100%;margin:.4em 0 .8em}td,th{border:1px solid #b6bfd0;padding:4px 8px;vertical-align:top}th{background:#eef2fb;text-align:left}
hr{border:0;border-top:1px solid #cfd5e1;margin:1em 0}hr.page-break{break-after:page;border-top-style:dashed}.toc{border-top:1px solid #cfd5e1;border-bottom:1px solid #cfd5e1;padding:.6em 0;margin:.4em 0 1em}
@media print{body{background:none}main{margin:0;box-shadow:none;padding:0}hr.page-break{border:0;margin:0}}
@media (max-width:${w}cm){main{margin:0;padding:24px 20px}}`;
    const html = `<!doctype html>\n<html lang="it">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<meta name="generator" content="Word">\n<title>${esc(titleEl.value || "Documento")}</title>\n<style>\n${css}\n</style>\n</head>\n<body>\n<main>\n${cleanHTML()}\n</main>\n</body>\n</html>\n`;
    return new Blob([html], { type: "text/html;charset=utf-8" });
  }
  $("btnExportHtml").addEventListener("click", () => { closePops(); deliver(exportHtml(), fileBase() + ".html"); });
  $("btnExportTxt").addEventListener("click", () => {
    closePops();
    const text = editor.innerText.replace(/​/g, "").replace(/\n{3,}/g, "\n\n").trim() + "\n";
    deliver(new Blob([text], { type: "text/plain;charset=utf-8" }), fileBase() + ".txt");
  });
  $("btnExportMd").addEventListener("click", () => {
    closePops();
    deliver(new Blob([htmlToMd(editor)], { type: "text/markdown;charset=utf-8" }), fileBase() + ".md");
  });

  /* ---------- Markdown: da documento a testo .md ---------- */
  const mdEsc = (s) => s.replace(/([\\`*_[\]<>])/g, "\\$1");
  function mdInline(nodes, ctx = {}) {
    const runs = [];
    const walk = (n) => {
      if (n.nodeType === 3) {
        const text = n.nodeValue.replace(/​/g, "").replace(/[\r\n\t]+/g, " ");
        if (!text) return;
        const el = n.parentElement, s = getComputedStyle(el);
        const a = el.closest("a");
        runs.push({
          text,
          b: !ctx.heading && parseInt(s.fontWeight, 10) >= 600,
          i: !ctx.quote && s.fontStyle === "italic",
          s: hasDeco(el, "line-through"),
          href: a && insideEditor(a) ? a.getAttribute("href") || "" : "",
        });
        return;
      }
      if (n.nodeType !== 1) return;
      if (n.tagName === "BR") { if (n.nextSibling) runs.push({ br: true }); return; }
      if (n.tagName === "IMG") { runs.push({ img: n.getAttribute("src") || "", alt: n.getAttribute("alt") || "" }); return; }
      n.childNodes.forEach(walk);
    };
    nodes.forEach(walk);
    const same = (x, y) => x && y && !x.br && !y.br && !x.img && !y.img && x.b === y.b && x.i === y.i && x.s === y.s && x.href === y.href;
    const merged = [];
    runs.forEach((r) => { const last = merged[merged.length - 1]; if (same(last, r)) last.text += r.text; else merged.push({ ...r }); });
    const fmt = (r) => {
      if (r.br) return "  \n";
      if (r.img) return `![${mdEsc(r.alt)}](${r.img})`;
      const m = r.text.match(/^(\s*)([\s\S]*?)(\s*)$/);
      let core = mdEsc(m[2]);
      if (!core) return r.text;
      if (r.s) core = `~~${core}~~`;
      if (r.i) core = `*${core}*`;
      if (r.b) core = `**${core}**`;
      if (r.href) core = `[${core}](${r.href.replace(/[()\s]/g, encodeURIComponent)})`;
      return m[1] + core + m[3];
    };
    return merged.map(fmt).join("").replace(/ {2,}(?!\n)/g, " ").trim();
  }
  function mdBlocks(container, out, ctx = {}) {
    let buf = [];
    const flush = () => { const t = mdInline(buf, ctx); if (t) out.push(t); buf = []; };
    container.childNodes.forEach((n) => {
      if (!(n.nodeType === 1 && BLOCKS.has(n.tagName))) { buf.push(n); return; }
      flush();
      const t = n.tagName;
      if (/^H[1-6]$/.test(t)) { const txt = mdInline([...n.childNodes], { ...ctx, heading: true }); if (txt) out.push("#".repeat(+t[1]) + " " + txt); }
      else if (t === "UL" || t === "OL") out.push(mdList(n, 0, ctx).join("\n"));
      else if (t === "TABLE") out.push(mdTable(n));
      else if (t === "HR") out.push("---");
      else if (t === "PRE") out.push("```\n" + n.textContent.replace(/\n$/, "") + "\n```");
      else if (t === "BLOCKQUOTE") {
        const sub = [];
        mdBlocks(n, sub, { ...ctx, quote: !n.getAttribute("style") });
        if (sub.length) out.push(sub.map((b) => b.split("\n").map((l) => (l ? "> " + l : ">")).join("\n")).join("\n>\n"));
      } else if ([...n.children].some((c) => BLOCKS.has(c.tagName))) mdBlocks(n, out, ctx);
      else { const txt = mdInline([...n.childNodes], ctx); if (txt) out.push(txt); }
    });
    flush();
    return out;
  }
  function mdList(list, depth, ctx) {
    const ordered = list.tagName === "OL";
    const lines = [];
    let k = 1;
    [...list.children].forEach((li) => {
      if (li.tagName !== "LI") return;
      const inline = [], nested = [];
      li.childNodes.forEach((c) => {
        if (c.nodeType === 1 && (c.tagName === "UL" || c.tagName === "OL")) nested.push(c);
        else if (c.nodeType === 1 && BLOCKS.has(c.tagName)) inline.push(...c.childNodes);
        else inline.push(c);
      });
      lines.push("    ".repeat(depth) + (ordered ? `${k++}. ` : "- ") + mdInline(inline, ctx));
      nested.forEach((l) => lines.push(...mdList(l, depth + 1, ctx)));
    });
    return lines;
  }
  function mdTable(table) {
    const rows = [...table.rows].map((tr) => [...tr.cells].map((c) => {
      const parts = mdBlocks(c, [], { heading: c.tagName === "TH" }); // nelle intestazioni il grassetto è già implicito
      return parts.join(" ").replace(/\n/g, " ").replace(/\|/g, "\\|") || " ";
    }));
    if (!rows.length) return "";
    const cols = Math.max(...rows.map((r) => r.length));
    const line = (r) => "| " + Array.from({ length: cols }, (_, i) => r[i] || " ").join(" | ") + " |";
    return [line(rows[0]), "| " + Array(cols).fill("---").join(" | ") + " |", ...rows.slice(1).map(line)].join("\n");
  }
  function htmlToMd(root) {
    return mdBlocks(root, []).join("\n\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }

  /* ---------- Markdown: da testo .md a documento ---------- */
  function mdInlineToHtml(s) {
    const codes = [];
    s = s.replace(/`([^`]+)`/g, (_, c) => { codes.push(c); return `\u0000${codes.length - 1}\u0000`; });
    const escapes = [];
    s = s.replace(/\\([\\`*_[\]{}()#+\-.!<>~|])/g, (_, c) => { escapes.push(c); return `\u0001${escapes.length - 1}\u0001`; });
    s = esc(s);
    s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)/g, (_, a, u) => `<img alt="${a}" src="${u}">`);
    s = s.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)/g, (_, t, u) => `<a href="${u}">${t}</a>`);
    s = s.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, '<a href="$1">$1</a>');
    s = s.replace(/(\*\*|__)(?=\S)([\s\S]*?\S)\1/g, "<b>$2</b>");
    s = s.replace(/(^|[^*\w])\*(?=\S)([^*]*?\S)\*(?!\*)/g, "$1<i>$2</i>");
    s = s.replace(/(^|[^_\w])_(?=\S)([^_]*?\S)_(?![_\w])/g, "$1<i>$2</i>");
    s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, "<s>$1</s>");
    s = s.replace(/\u0001(\d+)\u0001/g, (_, i) => esc(escapes[+i]));
    s = s.replace(/\u0000(\d+)\u0000/g, (_, i) => `<span style="font-family:'Courier New'">${esc(codes[+i])}</span>`);
    return s;
  }
  function mdToHtml(md) {
    const lines = md.replace(/\r\n?/g, "\n").replace(/\t/g, "    ").split("\n");
    const out = [];
    const isList = (l) => /^\s*([-*+]|\d+[.)])\s+/.test(l);
    const isHr = (l) => /^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(l);
    const isTableSep = (l) => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
    const cells = (l) => l.trim().replace(/^\||\|$/g, "").replace(/\\\|/g, "\u0002").split("|").map((c) => c.trim().replace(/\u0002/g, "\\|"));
    let i = 0;
    while (i < lines.length) {
      const l = lines[i];
      if (!l.trim()) { i++; continue; }
      let m;
      if ((m = l.match(/^\s{0,3}```/))) {
        const code = [];
        i++;
        while (i < lines.length && !/^\s{0,3}```/.test(lines[i])) code.push(lines[i++]);
        i++;
        out.push(`<p style="font-family:'Courier New'">${code.map((c) => esc(c) || "").join("<br>") || "<br>"}</p>`);
      } else if ((m = l.match(/^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/))) {
        const n = Math.min(m[1].length, 3);
        out.push(`<h${n}>${mdInlineToHtml(m[2])}</h${n}>`);
        i++;
      } else if (isHr(l)) {
        out.push("<hr>"); i++;
      } else if (/^\s{0,3}>/.test(l)) {
        const q = [];
        while (i < lines.length && /^\s{0,3}>/.test(lines[i])) q.push(lines[i++].replace(/^\s{0,3}>\s?/, ""));
        out.push(`<blockquote>${mdToHtml(q.join("\n"))}</blockquote>`);
      } else if (l.includes("|") && i + 1 < lines.length && isTableSep(lines[i + 1])) {
        const head = cells(l);
        i += 2;
        const body = [];
        while (i < lines.length && lines[i].includes("|") && lines[i].trim()) body.push(cells(lines[i++]));
        const row = (cs, tag) => `<tr>${cs.map((c) => `<${tag}>${mdInlineToHtml(c) || "<br>"}</${tag}>`).join("")}</tr>`;
        out.push(`<table><tbody>${row(head, "th")}${body.map((r) => row(r, "td")).join("")}</tbody></table>`);
      } else if (isList(l)) {
        const items = [];
        while (i < lines.length && (isList(lines[i]) || (lines[i].trim() && /^\s{2,}/.test(lines[i]) && items.length))) {
          const lm = lines[i].match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
          if (lm) items.push({ indent: lm[1].length, tag: /\d/.test(lm[2]) ? "ol" : "ul", text: lm[3] });
          else items[items.length - 1].text += " " + lines[i].trim();
          i++;
        }
        let html = "";
        const stack = [];
        items.forEach((it) => {
          while (stack.length && it.indent < stack[stack.length - 1].indent) html += `</li></${stack.pop().tag}>`;
          const top = stack[stack.length - 1];
          if (!top || it.indent > top.indent) { html += `<${it.tag}><li>`; stack.push(it); }
          else if (it.tag !== top.tag) { html += `</li></${top.tag}><${it.tag}><li>`; stack[stack.length - 1] = it; } // da puntato a numerato: nuovo elenco
          else html += "</li><li>";
          html += mdInlineToHtml(it.text);
        });
        while (stack.length) html += `</li></${stack.pop().tag}>`;
        out.push(html);
      } else {
        const para = [];
        while (i < lines.length && lines[i].trim() && !isList(lines[i]) && !isHr(lines[i]) && !/^\s{0,3}(#{1,6}\s|>|```)/.test(lines[i])) para.push(lines[i++]);
        out.push(`<p>${para.map((p, k) => mdInlineToHtml(p.trim()) + (k < para.length - 1 ? (/ {2,}$/.test(p) ? "<br>" : " ") : "")).join("")}</p>`);
      }
    }
    return out.join("");
  }

  /* ---------- Scorciatoie da tastiera ---------- */
  document.addEventListener("keydown", (e) => {
    if (document.querySelector("dialog[open]")) return; // le finestre gestiscono i loro tasti
    const mod = e.metaKey || e.ctrlKey;
    if (!mod) {
      if (e.key === "Escape") {
        if (openPop) closePops();
        else if (painterFmt) stopPainter();
        else if (!findbar.hidden) closeFind();
        else if (document.body.classList.contains("focus")) setFocus(false);
        else if (picked) pickImage(null);
      }
      if ((e.key === "Delete" || e.key === "Backspace") && picked && document.activeElement === editor) {
        e.preventDefault(); picked.remove(); pickImage(null); changed();
      }
      return;
    }
    const k = e.key.toLowerCase(), code = e.code;
    const inEditor = insideEditor(document.activeElement);
    const stop = () => { e.preventDefault(); e.stopPropagation(); };
    if (k === "s" && !e.shiftKey) { stop(); saveDocx(); }
    else if (k === "o" && !e.shiftKey) { stop(); $("btnOpen").click(); }
    else if (k === "p" && !e.shiftKey) { stop(); $("btnPdf").click(); }
    else if (k === "j" && !e.shiftKey) { stop(); focusCmd(); }
    else if (code === "Slash" || k === "/") { stop(); $("dlgKeys").showModal(); }
    else if (k === "k" && inEditor) { stop(); addLink(); }
    else if (k === "f" && !e.altKey && !e.shiftKey) { stop(); openFind(false); }
    else if (code === "KeyH" && e.altKey) { stop(); openFind(true); }
    else if (code === "KeyG" && e.shiftKey) { stop(); showStats(); }
    else if (e.altKey && /^Digit[0-3]$/.test(code) && inEditor) { stop(); applyStyle(code === "Digit0" ? "p" : "h" + code.slice(-1)); }
    else if (e.key === "Enter" && inEditor && !e.shiftKey && !e.altKey) { stop(); insertPageBreak(); }
    else if (inEditor && (e.key === ">" || (e.shiftKey && code === "Period"))) { stop(); stepSize(1); }
    else if (inEditor && (e.key === "<" || (e.shiftKey && code === "Comma"))) { stop(); stepSize(-1); }
    else if (NATIVE && (k === "=" || k === "+")) { stop(); $("zoomPlus").click(); }
    else if (NATIVE && k === "-") { stop(); $("zoomMinus").click(); }
    else if (NATIVE && k === "0") { stop(); setZoom(100); }
  });
  editor.addEventListener("keydown", (e) => {
    if (e.metaKey || e.ctrlKey) return;
    // Invio alla fine di un titolo: il paragrafo successivo torna «Normale»
    if (e.key === "Enter" && !e.shiftKey) {
      const sel = getSelection();
      const h = anchorElement()?.closest("h1,h2,h3,h4,h5,h6");
      if (h && sel.isCollapsed) {
        const rest = document.createRange();
        rest.setStart(sel.anchorNode, sel.anchorOffset);
        rest.setEndAfter(h.lastChild || h);
        if (!rest.toString().replace(/​/g, "").trim()) {
          e.preventDefault();
          const p = document.createElement("p");
          p.innerHTML = "<br>";
          h.after(p);
          setCaret(p, 0);
          changed();
        }
      }
      return;
    }
    if (e.key !== "Tab") return;
    e.preventDefault();
    const cell = currentCell();
    if (cell) {
      // Tab nelle tabelle: cella successiva (o nuova riga alla fine)
      const cellsAll = $$("td,th", cell.closest("table"));
      let i = cellsAll.indexOf(cell) + (e.shiftKey ? -1 : 1);
      if (i >= cellsAll.length) { document.querySelector('[data-table="rowBelow"]').click(); return; }
      if (i >= 0) setCaret(cellsAll[i], 0);
      return;
    }
    if (anchorElement()?.closest("li")) exec(e.shiftKey ? "outdent" : "indent");
    else if (!e.shiftKey) exec("insertText", " ");
  });

  /* ---------- Varie ---------- */
  editor.addEventListener("input", changed);
  titleEl.addEventListener("input", changed);
  titleEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); restoreSelection(); } });
  titleEl.addEventListener("blur", () => { if (!titleEl.value.trim()) { titleEl.value = UNTITLED; changed(); } });

  /* ---------- Collegamento con l'app per Mac ---------- */
  window.word = {
    toast,
    command(cmd) {
      const map = { new: "btnNew", open: "btnOpen", save: "btnSave", print: "btnPdf", find: "btnFind", replace: "btnReplace", count: "btnCount", keys: "btnKeys", nav: "btnNav", focus: "btnFocus", html: "btnExportHtml", md: "btnExportMd", txt: "btnExportTxt" };
      if (map[cmd]) $(map[cmd]).click();
      else if (cmd === "search") focusCmd();
      else if (cmd === "zoomIn") $("zoomPlus").click();
      else if (cmd === "zoomOut") $("zoomMinus").click();
      else if (cmd === "zoomReset") setZoom(100);
    },
    openBase64(name, b64) {
      const bin = atob(b64), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      openFile(new File([bytes], name));
    },
    saved(name) { persist(); toast(`Salvato in «${name}»`); },
  };
  window.parola = window.word; // vecchio nome, per compatibilità
  // Solo per le prove automatiche
  window.__word = { mdToHtml, htmlToMd, sanitize, textStats, pageCount, normalizeUrl, exportHtml };

  /* ---------- Avvio ---------- */
  document.execCommand("defaultParagraphSeparator", false, "p");
  load();
  sendTitle();
  applyPage();
  setTheme(prefs.theme || "auto");
  showTab(prefs.tab && prefs.tab !== "table" ? prefs.tab : "home");
  if (prefs.collapsed) setCollapsed(true);
  if (prefs.marks) setMarks(true);
  setSpell(prefs.spell !== false, false);
  if (prefs.zoom && prefs.zoom !== 100) setZoom(prefs.zoom);
  if (prefs.nav) setNav(true);
  if (!/Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent)) {
    $$("kbd").forEach((k) => { k.textContent = k.textContent.replace("⌘", "Ctrl").replace("⌥", "Alt").replace("⇧", "Maiusc"); });
    $$("[title]").forEach((el) => { el.title = el.title.replace(/⌘/g, "Ctrl+").replace(/⌥/g, "Alt+").replace(/⇧/g, "Maiusc+"); });
    cmdIndex = null;
  }
  updateStats();
  // le pagine dipendono dall'impaginazione: ricalcolo quando cambia la finestra o arrivano i caratteri
  const restats = () => { clearTimeout(statsTimer); statsTimer = setTimeout(updateStats, 200); };
  addEventListener("resize", restats);
  document.fonts?.ready.then(restats);
  addEventListener("beforeunload", persist);
  addEventListener("pagehide", persist);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") persist(); });
})();
