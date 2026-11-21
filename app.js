/* Parola — editor di testi */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const editor = $("editor"), page = $("page"), titleEl = $("title"), saveState = $("saveState");
  const STORE = "parola.documento", PREFS = "parola.preferenze";
  const IMG_MAX = 1600;
  // Presente solo dentro l'app per Mac
  const NATIVE = window.webkit && window.webkit.messageHandlers && window.webkit.messageHandlers.parola;
  if (NATIVE) { document.documentElement.classList.add("native"); document.getElementById("btnSave").textContent = "Salva .docx"; }
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  const PAGE_SIZES = { A4: [21, 29.7], A5: [14.8, 21], Letter: [21.59, 27.94] };
  const MARGINS = { normal: [2.5, 2], narrow: [1.27, 1.27], moderate: [2.54, 1.91], wide: [2.54, 5.08] };

  /* ---------- Utilità ---------- */
  let toastTimer;
  function toast(msg, ms = 2600) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), ms);
  }
  const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const insideEditor = (n) => n && (n === editor || editor.contains(n));

  function sanitize(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    doc.querySelectorAll("script,style,iframe,object,embed,link,meta,form,input,button,textarea,select").forEach((n) => n.remove());
    doc.querySelectorAll("*").forEach((el) => {
      for (const a of Array.from(el.attributes)) {
        const n = a.name.toLowerCase(), v = a.value.trim().toLowerCase();
        if (n.startsWith("on")) el.removeAttribute(a.name);
        else if ((n === "href" || n === "src") && (v.startsWith("javascript:") || (n === "src" && v.startsWith("data:") && !v.startsWith("data:image/")))) el.removeAttribute(a.name);
        else if (n === "contenteditable" || (n === "class" && !/page-break/.test(a.value))) el.removeAttribute(a.name);
      }
    });
    return doc.body.innerHTML;
  }

  let prefs = {};
  try { prefs = JSON.parse(localStorage.getItem(PREFS) || "{}") || {}; } catch (e) { prefs = {}; }
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
    if (savedRange) {
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }
  function exec(cmd, val = null, css = true) {
    restoreSelection();
    document.execCommand("styleWithCSS", false, css);
    document.execCommand(cmd, false, val);
    changed();
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
  function applyPage() {
    let [w, h] = PAGE_SIZES[pageSet.size] || PAGE_SIZES.A4;
    if (pageSet.orient === "landscape") [w, h] = [h, w];
    const [mt, ml] = MARGINS[pageSet.margins] || MARGINS.normal;
    page.style.setProperty("--pw", w + "cm");
    page.style.setProperty("--ph", h + "cm");
    page.style.setProperty("--mt", mt + "cm");
    page.style.setProperty("--ml", ml + "cm");
    pageRule.textContent = `@page { size: ${w}cm ${h}cm; margin: ${mt}cm ${ml}cm; }`;
    $$("[data-size]").forEach((b) => b.classList.toggle("on", b.dataset.size === pageSet.size));
    $$("[data-orient]").forEach((b) => b.classList.toggle("on", b.dataset.orient === pageSet.orient));
    $$("[data-margins]").forEach((b) => b.classList.toggle("on", b.dataset.margins === pageSet.margins));
  }

  /* ---------- Salvataggio automatico ---------- */
  let saveTimer;
  function changed() {
    normalize();
    saveState.textContent = "Modifiche non salvate…";
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persist, 600);
    updateStats();
  }
  function persist() {
    try {
      localStorage.setItem(STORE, JSON.stringify({ title: titleEl.value, html: editor.innerHTML.replace(/ class="picked"/g, ""), page: pageSet, t: Date.now() }));
      saveState.textContent = "Salvato su questo dispositivo";
    } catch (e) {
      saveState.textContent = "Non salvato: documento troppo grande, scaricalo in .docx";
    }
  }
  function load() {
    try {
      const d = JSON.parse(localStorage.getItem(STORE) || "null");
      if (d && typeof d.html === "string") {
        editor.innerHTML = sanitize(d.html);
        titleEl.value = d.title || "Documento senza titolo";
        if (d.page) pageSet = { ...pageSet, ...d.page };
        saveState.textContent = "Salvato su questo dispositivo";
        return;
      }
    } catch (e) { /* archivio non disponibile */ }
    editor.innerHTML = WELCOME;
    titleEl.value = "Benvenuto in Parola";
  }

  const WELCOME = `
    <h1>Benvenuto in Parola</h1>
    <p>Parola è un editor di testi che funziona nel browser. Scrivi qui, formatta con le schede in alto e scarica il risultato come file <b>.docx</b>, che si apre in Word, Pages, LibreOffice e Google Documenti.</p>
    <h2>Le schede</h2>
    <ul>
      <li><b>Home</b>: carattere, <b>grassetto</b>, <i>corsivo</i>, <u>sottolineato</u>, <span style="color:#a8482a">colori</span>, <span style="background-color:#ffe58a">evidenziatore</span>, elenchi, allineamento, interlinea, stili e trova.</li>
      <li><b>Inserisci</b>: tabelle, immagini, collegamenti, interruzioni di pagina, data e simboli.</li>
      <li><b>Layout</b>: margini, orientamento e dimensioni del foglio.</li>
      <li><b>Revisione</b> e <b>Visualizza</b>: ortografia, conteggio parole, zoom, segni di paragrafo e tema scuro.</li>
    </ul>
    <p>Il documento si salva da solo su questo dispositivo mentre scrivi. Quando hai finito, premi <b>Scarica .docx</b>.</p>`;

  /* ---------- Schede ---------- */
  function showTab(name) {
    $$(".rtab").forEach((t) => t.setAttribute("aria-selected", String(t.dataset.tab === name)));
    $$(".panel").forEach((p) => (p.hidden = p.dataset.panel !== name));
    if (name !== "table") { prefs.tab = name; savePrefs(); }
  }
  $$(".rtab").forEach((t) => t.addEventListener("click", () => showTab(t.dataset.tab)));

  // I pulsanti del nastro non devono togliere la selezione dal testo
  $$(".ribbon button, .popover button, .qb, .imgbar button").forEach((b) => b.addEventListener("mousedown", (e) => e.preventDefault()));

  /* ---------- Menu a comparsa ---------- */
  let openPop = null;
  function closePops() {
    if (!openPop) return;
    openPop.pop.hidden = true;
    openPop.btn.setAttribute("aria-expanded", "false");
    openPop = null;
    resetGrid();
  }
  $$("[data-pop]").forEach((btn) => {
    btn.setAttribute("aria-haspopup", "true");
    btn.setAttribute("aria-expanded", "false");
    btn.addEventListener("click", () => {
      const pop = $(btn.dataset.pop);
      const wasOpen = openPop && openPop.pop === pop;
      closePops();
      if (wasOpen) return;
      pop.hidden = false;
      const r = btn.getBoundingClientRect();
      pop.style.left = Math.max(8, Math.min(r.left, innerWidth - pop.offsetWidth - 8)) + "px";
      pop.style.top = r.bottom + 4 + "px";
      btn.setAttribute("aria-expanded", "true");
      openPop = { pop, btn };
    });
  });
  document.addEventListener("mousedown", (e) => {
    if (openPop && !openPop.pop.contains(e.target) && !openPop.btn.contains(e.target)) closePops();
  });

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
      const c = getComputedStyle(a).backgroundColor;
      if (BLOCKS.has(a.tagName) || a.tagName === "LI") break;
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
    let capNext = true, prevLetter = false;
    for (const n of nodes) {
      const s = n === r.startContainer ? r.startOffset : 0;
      const e = n === r.endContainer ? r.endOffset : n.length;
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

  function toggleMarks() {
    const on = !editor.classList.contains("marks");
    editor.classList.toggle("marks", on);
    $("btnMarks").classList.toggle("on", on);
    $("btnMarks2").classList.toggle("on", on);
    prefs.marks = on; savePrefs();
  }
  $("btnMarks").addEventListener("click", toggleMarks);
  $("btnMarks2").addEventListener("click", toggleMarks);

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
    blocks.forEach((bl) => (bl.style.backgroundColor = e.target.value.toLowerCase() === "#ffffff" ? "" : e.target.value));
    changed();
  });

  $$("[data-border]").forEach((b) => b.addEventListener("click", () => {
    closePops();
    const blocks = selectedBlocks();
    if (!blocks.length) return toast("Metti il cursore nel paragrafo");
    const line = "1px solid #1f1b16";
    blocks.forEach((bl) => {
      ["border", "borderTop", "borderBottom", "padding", "paddingTop", "paddingBottom"].forEach((p) => (bl.style[p] = ""));
      if (b.dataset.border === "noshade") { bl.style.backgroundColor = ""; return; }
      if (b.dataset.border === "bottom") { bl.style.borderBottom = line; bl.style.paddingBottom = "2px"; }
      if (b.dataset.border === "top") { bl.style.borderTop = line; bl.style.paddingTop = "2px"; }
      if (b.dataset.border === "box") { bl.style.border = line; bl.style.padding = "4px 6px"; }
    });
    changed();
  }));

  /* ---------- Home: stili ---------- */
  $$(".style-tile").forEach((t) => {
    t.addEventListener("mousedown", (e) => e.preventDefault());
    t.addEventListener("click", () => {
      const s = t.dataset.style;
      if (s === "em") exec("italic");
      else exec("formatBlock", `<${s}>`);
    });
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
    let tag = block && insideEditor(block) ? block.tagName.toLowerCase() : "p";
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
    if (![...fn.options].some((o) => o.value === fam)) fn.add(new Option(fam));
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
  function addLink() {
    restoreSelection();
    const existing = anchorElement()?.closest("a");
    const url = prompt("Indirizzo del collegamento (lascia vuoto per toglierlo):", existing ? existing.getAttribute("href") : "https://");
    if (url === null) return;
    if (!url.trim() || url.trim() === "https://") { exec("unlink"); return; }
    let href = url.trim();
    if (/^javascript:/i.test(href)) return;
    if (!/^(https?:|mailto:|tel:|#)/i.test(href)) href = "https://" + href;
    if (getSelection().isCollapsed) exec("insertHTML", `<a href="${esc(href)}">${esc(url.trim())}</a>`);
    else exec("createLink", href);
  }
  $("btnLink").addEventListener("click", addLink);
  editor.addEventListener("click", (e) => {
    const a = e.target.closest("a");
    if (a && (e.metaKey || e.ctrlKey)) window.open(a.href, "_blank", "noopener");
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
    const r = picked.getBoundingClientRect();
    imgbar.style.left = Math.max(8, Math.min(innerWidth - imgbar.offsetWidth - 8, r.left)) + "px";
    imgbar.style.top = Math.max(8, Math.min(innerHeight - 50, r.bottom + 8)) + "px";
  }
  editor.addEventListener("click", (e) => pickImage(e.target.tagName === "IMG" ? e.target : null));
  $("desk").addEventListener("scroll", placeImgbar);
  addEventListener("resize", placeImgbar);
  imgbar.addEventListener("click", (e) => {
    const w = e.target.dataset.w;
    if (!w || !picked) return;
    if (w === "del") { picked.remove(); pickImage(null); }
    else { picked.style.width = w + "%"; picked.style.height = "auto"; placeImgbar(); }
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
  function insertTable(rows, cols) {
    restoreSelection();
    const table = document.createElement("table");
    const body = table.createTBody();
    for (let i = 0; i < rows; i++) {
      const tr = body.insertRow();
      for (let j = 0; j < cols; j++) tr.insertCell().innerHTML = "<br>";
    }
    const after = document.createElement("p");
    after.innerHTML = "<br>";
    let n = savedRange ? savedRange.startContainer : null;
    while (n && n.parentNode !== editor) n = n.parentNode;
    if (!n) editor.append(table, after);
    else if (n.nodeType === 1 && n.tagName === "P" && !n.textContent.trim() && !n.querySelector("img,table")) n.replaceWith(table, after);
    else n.after(table, after);
    const sel = getSelection();
    sel.collapse(table.rows[0].cells[0], 0);
    savedRange = sel.getRangeAt(0).cloneRange();
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
    switch (b.dataset.table) {
      case "rowAbove": case "rowBelow": {
        const nr = document.createElement("tr");
        for (let i = 0; i < row.cells.length; i++) nr.appendChild(blank());
        b.dataset.table === "rowAbove" ? row.before(nr) : row.after(nr);
        break;
      }
      case "colLeft": case "colRight":
        for (const tr of table.rows) {
          const ref = tr.cells[Math.min(idx, tr.cells.length - 1)];
          const nc = blank(ref ? ref.tagName.toLowerCase() : "td");
          if (!ref) tr.appendChild(nc);
          else b.dataset.table === "colLeft" ? ref.before(nc) : ref.after(nc);
        }
        break;
      case "delRow":
        row.remove();
        if (!table.rows.length) table.remove();
        break;
      case "delCol":
        for (const tr of table.rows) if (tr.cells[idx]) tr.cells[idx].remove();
        if (!table.rows[0] || !table.rows[0].cells.length) table.remove();
        break;
      case "delTable":
        table.remove();
        break;
      case "header": {
        const first = table.rows[0];
        const toTh = first.cells[0].tagName === "TD";
        [...first.cells].forEach((c) => {
          const n = document.createElement(toTh ? "th" : "td");
          while (c.firstChild) n.appendChild(c.firstChild);
          if (!n.firstChild) n.innerHTML = "<br>";
          c.replaceWith(n);
        });
        break;
      }
    }
    if (!document.contains(cell)) { $("tabTable").hidden = true; showTab(prefs.tab || "home"); }
    changed();
  }));

  /* ---------- Inserisci: altro ---------- */
  $("btnPageBreak").addEventListener("click", () => exec("insertHTML", '<hr class="page-break"><p><br></p>'));
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

  /* ---------- Layout ---------- */
  $$("[data-margins]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.margins = b.dataset.margins; applyPage(); changed(); }));
  $$("[data-orient]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.orient = b.dataset.orient; applyPage(); changed(); }));
  $$("[data-size]").forEach((b) => b.addEventListener("click", () => { closePops(); pageSet.size = b.dataset.size; applyPage(); changed(); }));

  /* ---------- Revisione ---------- */
  function setSpell(on) {
    editor.spellcheck = on;
    $("btnSpell").classList.toggle("on", on);
    prefs.spell = on; savePrefs();
    // forzo l'aggiornamento delle sottolineature
    const t = editor.innerHTML; editor.innerHTML = t;
  }
  $("btnSpell").addEventListener("click", () => setSpell(!editor.spellcheck));
  $("btnCount").addEventListener("click", () => {
    const s = textStats();
    toast(`${s.words} parole · ${s.chars} caratteri (spazi inclusi) · ${s.charsNoSpace} senza spazi · ${s.paras} paragrafi`, 5000);
  });

  /* ---------- Trova e sostituisci ---------- */
  const findbar = $("findbar"), findInput = $("findInput"), replInput = $("replaceInput"), findCount = $("findCount");
  function openFind(replace) {
    findbar.hidden = false;
    const s = getSelection().toString();
    if (s && s.length < 80 && !s.includes("\n")) findInput.value = s;
    (replace && findInput.value ? replInput : findInput).focus();
    findInput.select();
    countMatches();
  }
  $("btnFind").addEventListener("click", () => openFind(false));
  $("btnFind2").addEventListener("click", () => openFind(false));
  $("btnReplace").addEventListener("click", () => openFind(true));
  $("findClose").addEventListener("click", () => { findbar.hidden = true; restoreSelection(); });

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
  function countMatches() {
    const n = allMatches(findInput.value).length;
    findCount.textContent = findInput.value ? (n === 1 ? "1 risultato" : `${n} risultati`) : "";
    return n;
  }
  function findNext() {
    const ms = allMatches(findInput.value);
    countMatches();
    if (!ms.length) return false;
    let next = ms[0];
    if (savedRange) {
      for (const m of ms) {
        const r = document.createRange();
        r.setStart(m.node, m.start);
        if (r.compareBoundaryPoints(Range.START_TO_END, savedRange) >= 0) { next = m; break; }
      }
    }
    const r = document.createRange();
    r.setStart(next.node, next.start); r.setEnd(next.node, next.end);
    const sel = getSelection();
    editor.focus({ preventScroll: true });
    sel.removeAllRanges(); sel.addRange(r);
    savedRange = r.cloneRange();
    next.node.parentElement.scrollIntoView({ block: "center", behavior: "smooth" });
    return true;
  }
  findInput.addEventListener("input", countMatches);
  findInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); findNext(); findInput.focus(); } if (e.key === "Escape") $("findClose").click(); });
  replInput.addEventListener("keydown", (e) => { if (e.key === "Escape") $("findClose").click(); });
  $("findNext").addEventListener("click", () => { findNext(); findInput.focus(); });
  $("replaceOne").addEventListener("click", () => {
    const q = findInput.value;
    if (!q) return;
    const cur = savedRange ? savedRange.toString() : "";
    if (cur.toLocaleLowerCase() !== q.toLocaleLowerCase()) { findNext(); return; }
    exec("insertText", replInput.value);
    findNext();
  });
  $("replaceAll").addEventListener("click", () => {
    const q = findInput.value;
    if (!q) return;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    let n = 0;
    for (const t of textNodes()) {
      const v = t.nodeValue.replace(re, () => { n++; return replInput.value; });
      if (v !== t.nodeValue) t.nodeValue = v;
    }
    toast(n ? `Sostituite ${n} occorrenze` : "Nessuna occorrenza trovata");
    countMatches();
    changed();
  });

  /* ---------- Visualizza ---------- */
  function setZoom(pct) {
    pct = Math.max(50, Math.min(200, Math.round(pct)));
    page.style.setProperty("--zoom", pct / 100);
    $("zoom").value = pct;
    $("zoomOut").textContent = pct + "%";
    placeImgbar();
  }
  $("zoom").addEventListener("input", (e) => setZoom(+e.target.value));
  $("btnZoom100").addEventListener("click", () => setZoom(100));
  $("btnFit").addEventListener("click", () => {
    const pw = parseFloat(getComputedStyle(page).getPropertyValue("--pw")) * 37.795;
    setZoom(((($("desk").clientWidth - 40) / pw) * 100));
  });

  const THEMES = ["auto", "light", "dark"], THEME_LABEL = { auto: "Tema: auto", light: "Tema: chiaro", dark: "Tema: scuro" };
  function setTheme(t) {
    if (t === "auto") delete document.documentElement.dataset.theme;
    else document.documentElement.dataset.theme = t;
    $("themeLabel").textContent = THEME_LABEL[t];
    prefs.theme = t; savePrefs();
  }
  $("btnTheme").addEventListener("click", () => setTheme(THEMES[(THEMES.indexOf(prefs.theme || "auto") + 1) % 3]));
  $("btnFull").addEventListener("click", () => {
    if (document.fullscreenElement) document.exitFullscreen();
    else document.documentElement.requestFullscreen?.().catch(() => toast("Schermo intero non disponibile"));
  });

  /* ---------- File: nuovo, apri, scarica, PDF ---------- */
  $("btnNew").addEventListener("click", () => {
    if (editor.textContent.trim() && !confirm("Iniziare un nuovo documento? Quello attuale verrà sostituito: se ti serve, scaricalo prima in .docx.")) return;
    editor.innerHTML = "<p><br></p>";
    titleEl.value = "Documento senza titolo";
    pageSet = { size: "A4", orient: "portrait", margins: "normal" };
    applyPage();
    pickImage(null);
    changed();
    editor.focus();
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
      } else {
        html = await f.text();
      }
      editor.innerHTML = sanitize(html) || "<p><br></p>";
      titleEl.value = f.name.replace(/\.[^.]+$/, "");
      pickImage(null);
      changed();
      $("desk").scrollTop = 0;
      toast(`Aperto «${f.name}»`);
    } catch (err) {
      toast("Non riesco ad aprire questo file: " + err.message);
    }
  }

  async function saveDocx() {
    if (!window.ParolaDocx || !window.docx) return toast("Attendi un attimo: sto ancora caricando i componenti.");
    const btn = $("btnSave");
    btn.disabled = true; btn.textContent = "Preparo il file…";
    pickImage(null);
    const marks = editor.classList.contains("marks");
    editor.classList.remove("marks");
    const z = page.style.getPropertyValue("--zoom");
    page.style.setProperty("--zoom", "1");
    try {
      let [w, h] = PAGE_SIZES[pageSet.size] || PAGE_SIZES.A4;
      const [mt, ml] = MARGINS[pageSet.margins] || MARGINS.normal;
      const blob = await window.ParolaDocx.toBlob(editor, titleEl.value, { width: w, height: h, landscape: pageSet.orient === "landscape", top: mt, side: ml });
      const fileName = (titleEl.value.trim() || "Documento").replace(/[\\/:*?"<>|]+/g, "-") + ".docx";
      if (NATIVE) {
        // App per Mac: il file lo salva l'app, con la finestra «Salva» di macOS
        const b64 = await new Promise((res, rej) => {
          const rd = new FileReader();
          rd.onload = () => res(String(rd.result).split(",")[1]);
          rd.onerror = rej;
          rd.readAsDataURL(blob);
        });
        NATIVE.postMessage({ type: "save", name: fileName, data: b64 });
      } else {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = fileName;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        toast("File .docx scaricato");
      }
    } catch (err) {
      console.error(err);
      toast("Errore durante il salvataggio: " + err.message);
    } finally {
      page.style.setProperty("--zoom", z || "1");
      editor.classList.toggle("marks", marks);
      btn.disabled = false; btn.textContent = NATIVE ? "Salva .docx" : "Scarica .docx";
    }
  }
  $("btnSave").addEventListener("click", saveDocx);

  $("btnPdf").addEventListener("click", () => {
    pickImage(null);
    if (NATIVE) { NATIVE.postMessage({ type: "print", title: titleEl.value || "Documento" }); return; }
    const t = document.title;
    document.title = titleEl.value || "Documento";
    toast("Nella finestra di stampa scegli «Salva come PDF»");
    setTimeout(() => { print(); document.title = t; }, 300);
  });

  /* ---------- Scorciatoie da tastiera ---------- */
  document.addEventListener("keydown", (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (!mod) {
      if (e.key === "Escape") { closePops(); if (painterFmt) stopPainter(); if (!findbar.hidden) $("findClose").click(); }
      if ((e.key === "Delete" || e.key === "Backspace") && picked && document.activeElement === editor) {
        e.preventDefault(); picked.remove(); pickImage(null); changed();
      }
      return;
    }
    const k = e.key.toLowerCase();
    if (k === "s") { e.preventDefault(); saveDocx(); }
    else if (k === "o") { e.preventDefault(); $("btnOpen").click(); }
    else if (k === "k" && insideEditor(document.activeElement)) { e.preventDefault(); addLink(); }
    else if (k === "f") { e.preventDefault(); openFind(false); }
    else if (k === "h" && e.altKey) { e.preventDefault(); openFind(true); }
  });
  editor.addEventListener("keydown", (e) => {
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
          sel.collapse(p, 0);
          changed();
        }
      }
      return;
    }
    if (e.key !== "Tab") return;
    e.preventDefault();
    if (anchorElement()?.closest("li")) exec(e.shiftKey ? "outdent" : "indent");
    else if (!e.shiftKey) exec("insertText", " ");
  });

  /* ---------- Varie ---------- */
  editor.addEventListener("input", changed);
  titleEl.addEventListener("input", changed);
  titleEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); editor.focus(); } });

  function textStats() {
    const text = editor.innerText.replace(/​/g, "");
    return {
      words: (text.match(/[\p{L}\p{N}'’-]+/gu) || []).length,
      chars: text.replace(/\n/g, "").length,
      charsNoSpace: text.replace(/\s/g, "").length,
      paras: $$("p,h1,h2,h3,li,blockquote", editor).filter((p) => p.textContent.trim()).length,
    };
  }
  function updateStats() {
    const s = textStats();
    const pages = Math.max(1, Math.ceil(s.words / 450) + $$("hr.page-break", editor).length);
    $("stats").textContent = `${s.words.toLocaleString("it-IT")} ${s.words === 1 ? "parola" : "parole"} · ${s.charsNoSpace.toLocaleString("it-IT")} caratteri · circa ${pages} ${pages === 1 ? "pagina" : "pagine"}`;
  }

  /* ---------- Collegamento con l'app per Mac ---------- */
  window.parola = {
    toast,
    command(cmd) {
      const map = { new: "btnNew", open: "btnOpen", save: "btnSave", print: "btnPdf", find: "btnFind", replace: "btnReplace" };
      if (map[cmd]) $(map[cmd]).click();
    },
    openBase64(name, b64) {
      const bin = atob(b64), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      openFile(new File([bytes], name));
    },
  };

  /* ---------- Avvio ---------- */
  document.execCommand("defaultParagraphSeparator", false, "p");
  load();
  applyPage();
  setTheme(prefs.theme || "auto");
  showTab(prefs.tab || "home");
  if (prefs.marks) toggleMarks();
  editor.spellcheck = prefs.spell !== false;
  $("btnSpell").classList.toggle("on", editor.spellcheck);
  updateStats();
  addEventListener("beforeunload", () => { clearTimeout(saveTimer); persist(); });
})();
