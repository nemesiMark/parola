/* Word — conversione del documento (HTML dell'editor) in file .docx */
(function () {
  "use strict";

  const BLOCK = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "UL", "OL", "LI",
    "TABLE", "HR", "PRE", "FIGURE", "SECTION", "ARTICLE", "HEADER", "FOOTER"]);
  const PX_TWIP = 15;          // 1px = 0.75pt = 15 twip
  const MAX_IMG_W = 620;        // larghezza utile della pagina A4 con i margini, in px

  let D, root, olCounter, imgBytes;

  const isBlock = (n) => n.nodeType === 1 && BLOCK.has(n.tagName);
  const hasBlockChild = (el) => Array.from(el.childNodes).some(isBlock);
  const cs = (el) => getComputedStyle(el);

  function hex(color) {
    const m = color && color.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const p = m[1].split(",").map((s) => parseFloat(s));
    if (p.length === 4 && p[3] === 0) return null;
    return p.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  const firstFont = (f) => (f || "").split(",")[0].replace(/["']/g, "").trim() || "Georgia";

  function blockAncestor(el) {
    while (el && el !== root && !isBlock(el)) el = el.parentElement;
    return el;
  }

  function decorations(el) {
    let underline = false, strike = false, shade = null;
    const stop = blockAncestor(el);
    for (let a = el; a && a !== root; a = a.parentElement) {
      const s = cs(a);
      const line = s.textDecorationLine || s.textDecoration || "";
      if (line.includes("underline") || a.tagName === "U") underline = true;
      if (line.includes("line-through") || a.tagName === "S" || a.tagName === "STRIKE" || a.tagName === "DEL") strike = true;
      if (!shade && a !== stop && !isBlock(a)) shade = hex(s.backgroundColor);
      if (a === stop) break;
    }
    return { underline, strike, shade };
  }

  function runProps(el) {
    const s = cs(el);
    const d = decorations(el);
    const o = {
      font: firstFont(s.fontFamily),
      size: Math.max(2, Math.round(parseFloat(s.fontSize) * 0.75 * 2)),
      bold: parseInt(s.fontWeight, 10) >= 600,
      italics: s.fontStyle === "italic",
      color: hex(s.color) || "000000",
    };
    if (d.underline) o.underline = {};
    if (d.strike) o.strike = true;
    if (d.shade) o.shading = { type: D.ShadingType.CLEAR, fill: d.shade, color: "auto" };
    const va = s.verticalAlign;
    if (va === "super") o.superScript = true;
    if (va === "sub") o.subScript = true;
    return o;
  }

  function dataUrlBytes(src) {
    const b64 = src.split(",")[1] || "";
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // Le immagini vengono salvate nel .docx come PNG: converto tutto in PNG.
  function imageToPng(img) {
    return new Promise((resolve) => {
      const done = () => {
        try {
          const c = document.createElement("canvas");
          c.width = img.naturalWidth; c.height = img.naturalHeight;
          c.getContext("2d").drawImage(img, 0, 0);
          resolve(dataUrlBytes(c.toDataURL("image/png")));
        } catch (e) { resolve(null); }
      };
      if (img.complete && img.naturalWidth) done();
      else { img.addEventListener("load", done, { once: true }); img.addEventListener("error", () => resolve(null), { once: true }); }
    });
  }

  function imageRun(img) {
    const r = img.getBoundingClientRect();
    let w = r.width || img.naturalWidth || 200, h = r.height || img.naturalHeight || 150;
    const maxW = root.clientWidth || MAX_IMG_W; // larghezza utile del foglio attuale
    if (w > maxW) { h = h * maxW / w; w = maxW; }
    const bytes = imgBytes.get(img);
    return bytes ? new D.ImageRun({ data: bytes, transformation: { width: w, height: h } }) : new D.TextRun("");
  }

  function runsFrom(node, out) {
    if (node.nodeType === 3) {
      let t = node.nodeValue.replace(/​/g, "");
      const pre = node.parentElement && cs(node.parentElement).whiteSpace.startsWith("pre");
      if (!pre) t = t.replace(/[\r\n\t]+/g, " ");
      if (!t) return;
      out.push(new D.TextRun({ text: t, ...runProps(node.parentElement) }));
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = node.tagName;
    if (tag === "BR" && !node.nextSibling && (isBlock(node.parentElement) || /^T[DH]$/.test(node.parentElement.tagName))) return; // <br> finale di un paragrafo vuoto
    if (tag === "BR") { out.push(new D.TextRun({ text: "", break: 1 })); return; }
    if (tag === "IMG") { out.push(imageRun(node)); return; }
    if (tag === "SCRIPT" || tag === "STYLE") return;
    if (tag === "A" && node.getAttribute("href")) {
      const kids = [];
      node.childNodes.forEach((c) => runsFrom(c, kids));
      if (kids.length) out.push(new D.ExternalHyperlink({ children: kids, link: node.getAttribute("href") }));
      return;
    }
    node.childNodes.forEach((c) => runsFrom(c, out));
  }

  function alignment(s) {
    switch (s.textAlign) {
      case "center": return D.AlignmentType.CENTER;
      case "right": case "end": return D.AlignmentType.RIGHT;
      case "justify": return D.AlignmentType.JUSTIFIED;
      default: return D.AlignmentType.LEFT;
    }
  }

  function paragraph(styleEl, nodes, extra = {}) {
    const runs = [];
    nodes.forEach((n) => runsFrom(n, runs));
    const s = cs(styleEl);
    const fontPx = parseFloat(s.fontSize) || 16;
    const lh = parseFloat(s.lineHeight);
    const o = {
      children: runs,
      alignment: alignment(s),
      spacing: {
        before: Math.round((parseFloat(s.marginTop) || 0) * PX_TWIP),
        after: Math.round((parseFloat(s.marginBottom) || 0) * PX_TWIP),
        line: isNaN(lh) ? 276 : Math.round((lh / fontPx) * 240),
      },
      ...extra,
    };
    const h = { H1: "HEADING_1", H2: "HEADING_2", H3: "HEADING_3", H4: "HEADING_4", H5: "HEADING_5", H6: "HEADING_6" }[styleEl.tagName];
    if (h) o.heading = D.HeadingLevel[h];

    if (!o.numbering && styleEl !== root) {
      const left = styleEl.getBoundingClientRect().left - root.getBoundingClientRect().left;
      if (left > 2) o.indent = { left: Math.round(left * PX_TWIP) };
    }
    // sfondo e bordi propri del paragrafo
    if (styleEl !== root && !/^T[DH]$/.test(styleEl.tagName)) {
      const bg = hex(s.backgroundColor);
      if (bg) o.shading = { type: D.ShadingType.CLEAR, fill: bg, color: "auto" };
      const border = {};
      for (const side of ["top", "bottom", "left", "right"]) {
        const cap = side[0].toUpperCase() + side.slice(1);
        const w = parseFloat(s["border" + cap + "Width"]);
        if (w > 0 && s["border" + cap + "Style"] !== "none") {
          border[side] = { style: D.BorderStyle.SINGLE, size: Math.max(4, Math.round(w * 6)), color: hex(s["border" + cap + "Color"]) || "000000", space: 2 };
        }
      }
      if (Object.keys(border).length) o.border = border;
    }
    const bq = styleEl.closest && styleEl.closest("blockquote");
    if (bq && root.contains(bq) && !bq.getAttribute("style")) {
      o.border = { ...(o.border || {}), left: { style: D.BorderStyle.SINGLE, size: 18, color: hex(cs(bq).borderLeftColor) || "2457D6", space: 8 } };
    }
    return new D.Paragraph(o);
  }

  function listItems(list, level, out) {
    const ordered = list.tagName === "OL";
    const instance = ordered ? ++olCounter : 0;
    Array.from(list.children).forEach((li) => {
      if (li.tagName !== "LI") { out.push(...blocksFrom(li)); return; }
      const inline = [], nested = [];
      li.childNodes.forEach((c) => {
        if (c.nodeType === 1 && (c.tagName === "UL" || c.tagName === "OL")) nested.push(c);
        else if (isBlock(c)) inline.push(...c.childNodes); // <li><p>…</p></li> come prodotto da Word
        else inline.push(c);
      });
      out.push(paragraph(li, inline, {
        numbering: { reference: ordered ? "num" : "bul", level: Math.min(level, 8), instance },
      }));
      nested.forEach((n) => listItems(n, level + 1, out));
    });
  }

  function table(el) {
    const rows = Array.from(el.rows).map((tr) => new D.TableRow({
      children: Array.from(tr.cells).map((td) => {
        let kids = blocksFrom(td).filter((k) => !(k instanceof D.Table));
        if (!kids.length) kids = [new D.Paragraph("")];
        const o = { children: kids, margins: { top: 60, bottom: 60, left: 110, right: 110 } };
        if (td.colSpan > 1) o.columnSpan = td.colSpan;
        if (td.rowSpan > 1) o.rowSpan = td.rowSpan;
        const bg = hex(cs(td).backgroundColor);
        if (bg) o.shading = { type: D.ShadingType.CLEAR, fill: bg, color: "auto" };
        return new D.TableCell(o);
      }),
    }));
    if (!rows.length) return [];
    return [new D.Table({ rows, width: { size: 100, type: D.WidthType.PERCENTAGE } }), new D.Paragraph("")];
  }

  function blocksFrom(container) {
    const out = [];
    let buf = [];
    const flush = () => {
      const meaningful = buf.some((n) => n.nodeType === 1 || n.nodeValue.replace(/[\s​]/g, ""));
      if (meaningful) out.push(paragraph(container, buf));
      buf = [];
    };
    container.childNodes.forEach((n) => {
      if (!isBlock(n)) { buf.push(n); return; }
      flush();
      const t = n.tagName;
      if (t === "UL" || t === "OL") listItems(n, 0, out);
      else if (t === "TABLE") out.push(...table(n));
      else if (t === "HR") {
        if (n.classList.contains("page-break")) out.push(new D.Paragraph({ children: [new D.PageBreak()] }));
        else out.push(new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: "CFD5E1", space: 1 } } }));
      } else if (hasBlockChild(n)) out.push(...blocksFrom(n));
      else out.push(paragraph(n, Array.from(n.childNodes)));
    });
    flush();
    return out;
  }

  async function prepareImages(editor) {
    const map = new Map();
    for (const img of editor.querySelectorAll("img")) {
      const bytes = await imageToPng(img);
      if (bytes) map.set(img, bytes);
    }
    return map;
  }

  function numberingConfig() {
    const bullets = ["•", "◦", "▪"];
    const numFmt = [D.LevelFormat.DECIMAL, D.LevelFormat.LOWER_LETTER, D.LevelFormat.LOWER_ROMAN];
    const lv = (l, format, text) => ({
      level: l, format, text, alignment: D.AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    });
    return [
      { reference: "bul", levels: Array.from({ length: 9 }, (_, l) => lv(l, D.LevelFormat.BULLET, bullets[l % 3])) },
      { reference: "num", levels: Array.from({ length: 9 }, (_, l) => lv(l, numFmt[l % 3], `%${l + 1}.`)) },
    ];
  }

  async function toBlob(editor, title, pg = {}) {
    const cm = (v) => Math.round(v * 567);
    D = window.docx;
    if (!D) throw new Error("Libreria docx non caricata");
    root = editor;
    olCounter = 0;

    imgBytes = await prepareImages(editor);
    const children = blocksFrom(editor);

    const doc = new D.Document({
      creator: "Word",
      title: title || "Documento",
      styles: { default: { document: { run: { font: "Georgia", size: 24 } } } },
      numbering: { config: numberingConfig() },
      sections: [{
        properties: {
          page: {
            size: {
              width: cm(pg.width || 21), height: cm(pg.height || 29.7),
              orientation: pg.landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT,
            },
            margin: { top: cm(pg.top || 2.5), bottom: cm(pg.top || 2.5), left: cm(pg.side || 2), right: cm(pg.side || 2) },
          },
        },
        children: children.length ? children : [new D.Paragraph("")],
      }],
    });
    return D.Packer.toBlob(doc);
  }

  window.WordDocx = { toBlob };
})();
