// docx WRITE: rebuild a .docx archive from edited HTML, preserving every untouched part.
// Pure HTML -> XML (body, header/footer, comments, reactions, replies, page margins); the
// read half lives in ./read.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { toHex6, fontSizeToHalfPt, firstFontFamily } from "../../core/util";
import type { PageGeometry } from "../../core/types";
import { W, R, PKG, REL_HYPERLINK, NS_DECLS, FMT0, HL_BY_HEX, JC_BY_ALIGN } from "./shared";
import type { Fmt } from "./shared";

// ---------------------------------------------------------------------------
// HTML -> .docx
// ---------------------------------------------------------------------------

interface DocxCtx {
  doc: Document;
  rels: Document | null;
  relsAdded: boolean;
  nextRid: number;
  nextRevId: number; // next w:ins/w:del revision id
  listNumId: string | null;
  files: Record<string, Uint8Array>; // the archive, so new media/content-types can be added
}

function addHyperlinkRel(ctx: DocxCtx, target: string): string | null {
  if (!ctx.rels) return null;
  const id = `rId${ctx.nextRid++}`;
  const rel = ctx.rels.createElementNS(PKG, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", REL_HYPERLINK);
  rel.setAttribute("Target", target);
  rel.setAttribute("TargetMode", "External");
  ctx.rels.documentElement.appendChild(rel);
  ctx.relsAdded = true;
  return id;
}

const DATA_URL_EXT: Record<string, string> = {
  "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/bmp": "bmp", "image/webp": "webp",
};
const A_NS = NS_DECLS["xmlns:a"]!;
const WP_NS = NS_DECLS["xmlns:wp"]!;
const PIC_NS = NS_DECLS["xmlns:pic"]!;

/** Ensure [Content_Types].xml declares a default for the extension. */
function ensureContentType(files: Record<string, Uint8Array>, ext: string, mime: string): void {
  const key = "[Content_Types].xml";
  const xml = files[key];
  if (!xml) return;
  const doc = new DOMParser().parseFromString(strFromU8(xml), "application/xml");
  const has = Array.from(doc.getElementsByTagName("Default")).some((d) => (d.getAttribute("Extension") ?? "").toLowerCase() === ext);
  if (has) return;
  const def = doc.createElementNS(doc.documentElement!.namespaceURI, "Default");
  def.setAttribute("Extension", ext);
  def.setAttribute("ContentType", mime);
  doc.documentElement!.insertBefore(def, doc.documentElement!.firstChild);
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Embed a data: URL image as a new media part + relationship; return a w:drawing run. */
function buildImageDrawing(ctx: DocxCtx, src: string, widthPx: number, heightPx: number): Element | null {
  if (!ctx.rels) return null;
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(src);
  if (!m) return null;
  const mime = m[1]!;
  const ext = DATA_URL_EXT[mime];
  if (!ext || !m[2]) return null;
  let bin: string;
  try {
    bin = atob(m[3]!);
  } catch {
    return null;
  }
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  // unique media name
  let n = 1;
  while (ctx.files[`word/media/omni-image${n}.${ext}`]) n++;
  const name = `media/omni-image${n}.${ext}`;
  ctx.files[`word/${name}`] = bytes;
  ensureContentType(ctx.files, ext, mime);
  const rid = `rId${ctx.nextRid++}`;
  const rel = ctx.rels.createElementNS(PKG, "Relationship");
  rel.setAttribute("Id", rid);
  rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
  rel.setAttribute("Target", name);
  ctx.rels.documentElement.appendChild(rel);
  ctx.relsAdded = true;

  const cx = Math.max(1, Math.round((widthPx || 200) * 9525));
  const cy = Math.max(1, Math.round((heightPx || 200) * 9525));
  const ce = (ns: string, name: string) => ctx.doc.createElementNS(ns, name);
  const r = ce(W, "w:r");
  const drawing = ce(W, "w:drawing");
  const inline = ce(WP_NS, "wp:inline");
  const extent = ce(WP_NS, "wp:extent");
  extent.setAttribute("cx", String(cx));
  extent.setAttribute("cy", String(cy));
  const docPr = ce(WP_NS, "wp:docPr");
  docPr.setAttribute("id", String(ctx.nextRid));
  docPr.setAttribute("name", `Image ${ctx.nextRid}`);
  const graphic = ce(A_NS, "a:graphic");
  const gData = ce(A_NS, "a:graphicData");
  gData.setAttribute("uri", PIC_NS);
  const pic = ce(PIC_NS, "pic:pic");
  const nvPicPr = ce(PIC_NS, "pic:nvPicPr");
  const cNvPr = ce(PIC_NS, "pic:cNvPr");
  cNvPr.setAttribute("id", "0");
  cNvPr.setAttribute("name", name);
  nvPicPr.append(cNvPr, ce(PIC_NS, "pic:cNvPicPr"));
  const blipFill = ce(PIC_NS, "pic:blipFill");
  const blip = ce(A_NS, "a:blip");
  blip.setAttributeNS(R, "r:embed", rid);
  const stretch = ce(A_NS, "a:stretch");
  stretch.appendChild(ce(A_NS, "a:fillRect"));
  blipFill.append(blip, stretch);
  const spPr = ce(PIC_NS, "pic:spPr");
  const xfrm = ce(A_NS, "a:xfrm");
  const off = ce(A_NS, "a:off");
  off.setAttribute("x", "0");
  off.setAttribute("y", "0");
  const ext2 = ce(A_NS, "a:ext");
  ext2.setAttribute("cx", String(cx));
  ext2.setAttribute("cy", String(cy));
  xfrm.append(off, ext2);
  const geom = ce(A_NS, "a:prstGeom");
  geom.setAttribute("prst", "rect");
  geom.appendChild(ce(A_NS, "a:avLst"));
  spPr.append(xfrm, geom);
  pic.append(nvPicPr, blipFill, spPr);
  gData.appendChild(pic);
  graphic.appendChild(gData);
  inline.append(extent, docPr, graphic);
  drawing.appendChild(inline);
  r.appendChild(drawing);
  return r;
}

const fmtHasProps = (f: Fmt): boolean => !!(f.b || f.i || f.u || f.strike || f.vertAlign || f.color || f.highlight || f.shading || f.sizeHalfPt || f.font);

/** Append the property elements for a Fmt to a w:rPr, in OOXML schema order. */
function fillRPr(ctx: DocxCtx, rPr: Element, f: Fmt): void {
  const flag = (tag: string) => rPr.appendChild(ctx.doc.createElementNS(W, tag));
  const valEl = (tag: string, val: string) => {
    const el = ctx.doc.createElementNS(W, tag);
    el.setAttributeNS(W, "w:val", val);
    rPr.appendChild(el);
  };
  if (f.font) {
    const rf = ctx.doc.createElementNS(W, "w:rFonts");
    rf.setAttributeNS(W, "w:ascii", f.font);
    rf.setAttributeNS(W, "w:hAnsi", f.font);
    rPr.appendChild(rf);
  }
  if (f.b) flag("w:b");
  if (f.i) flag("w:i");
  if (f.strike) flag("w:strike");
  if (f.color) valEl("w:color", f.color);
  if (f.sizeHalfPt) {
    valEl("w:sz", String(f.sizeHalfPt));
    valEl("w:szCs", String(f.sizeHalfPt));
  }
  if (f.highlight) valEl("w:highlight", f.highlight);
  if (f.shading) {
    const shd = ctx.doc.createElementNS(W, "w:shd");
    shd.setAttributeNS(W, "w:val", "clear");
    shd.setAttributeNS(W, "w:color", "auto");
    shd.setAttributeNS(W, "w:fill", f.shading);
    rPr.appendChild(shd);
  }
  if (f.u) valEl("w:u", "single");
  if (f.vertAlign) valEl("w:vertAlign", f.vertAlign === "super" ? "superscript" : "subscript");
}

interface RprChange {
  old: Fmt;
  author: string;
  date: string;
}
function makeRun(ctx: DocxCtx, text: string, f: Fmt, del = false, change?: RprChange): Element {
  const r = ctx.doc.createElementNS(W, "w:r");
  if (fmtHasProps(f) || change) {
    const rPr = ctx.doc.createElementNS(W, "w:rPr");
    fillRPr(ctx, rPr, f);
    if (change) {
      // A tracked formatting change: record the previous properties in w:rPrChange.
      const ch = ctx.doc.createElementNS(W, "w:rPrChange");
      ch.setAttributeNS(W, "w:id", String(ctx.nextRevId++));
      ch.setAttributeNS(W, "w:author", change.author || "Author");
      if (change.date) ch.setAttributeNS(W, "w:date", change.date);
      const oldRPr = ctx.doc.createElementNS(W, "w:rPr");
      fillRPr(ctx, oldRPr, change.old);
      ch.appendChild(oldRPr);
      rPr.appendChild(ch);
    }
    r.appendChild(rPr);
  }
  const t = ctx.doc.createElementNS(W, del ? "w:delText" : "w:t");
  t.setAttribute("xml:space", "preserve");
  t.textContent = text;
  r.appendChild(t);
  return r;
}

/** A run holding a manual page break (w:br w:type="page"). */
function pageBreakRun(ctx: DocxCtx): Element {
  const r = ctx.doc.createElementNS(W, "w:r");
  const br = ctx.doc.createElementNS(W, "w:br");
  br.setAttributeNS(W, "w:type", "page");
  r.appendChild(br);
  return r;
}

/** Re-parse a passthrough fragment (stored in data-docx-xml) into the output document. */
function importPassthrough(ctx: DocxCtx, xml: string): Element | null {
  try {
    const frag = new DOMParser().parseFromString(xml, "application/xml");
    if (frag.getElementsByTagName("parsererror").length || !frag.documentElement) return null;
    return ctx.doc.importNode(frag.documentElement, true) as Element;
  } catch {
    return null;
  }
}

function appendInline(ctx: DocxCtx, node: Node, parent: Element, f: Fmt, del = false, change?: RprChange): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      const txt = child.textContent ?? "";
      if (txt) parent.appendChild(makeRun(ctx, txt, f, del, change));
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    // Tracked formatting change: record the previous run properties via rPrChange.
    if (tag === "span" && el.classList.contains("docx-rpr-change")) {
      let old: Fmt = FMT0;
      try {
        old = { ...FMT0, ...(JSON.parse(el.getAttribute("data-old") || "{}") as Partial<Fmt>) };
      } catch {
        /* keep default */
      }
      appendInline(ctx, el, parent, f, del, { old, author: el.getAttribute("data-rev-author") || "Author", date: el.getAttribute("data-rev-date") || "" });
      continue;
    }
    // Tracked changes: <ins>/<del> -> w:ins/w:del (del runs use w:delText).
    if ((tag === "ins" || tag === "del") && el.classList.contains(`docx-${tag}`)) {
      const w = ctx.doc.createElementNS(W, tag === "del" ? "w:del" : "w:ins");
      w.setAttributeNS(W, "w:id", String(ctx.nextRevId++));
      w.setAttributeNS(W, "w:author", el.getAttribute("data-author") || "Author");
      const d = el.getAttribute("data-date");
      if (d) w.setAttributeNS(W, "w:date", d);
      appendInline(ctx, el, w, f, del || tag === "del", change);
      parent.appendChild(w);
      continue;
    }
    const stash = el.getAttribute("data-docx-xml");
    if (stash) {
      const node2 = importPassthrough(ctx, stash);
      if (node2) parent.appendChild(node2);
      continue;
    }
    const pb = el.getAttribute("data-docx-pagebreak");
    if (pb) {
      if (pb !== "auto") parent.appendChild(pageBreakRun(ctx));
      continue; // auto markers are display-only; Word recreates them
    }
    if (tag === "br") {
      const r = ctx.doc.createElementNS(W, "w:r");
      r.appendChild(ctx.doc.createElementNS(W, "w:br"));
      parent.appendChild(r);
      continue;
    }
    if (tag === "a") {
      const id = addHyperlinkRel(ctx, el.getAttribute("href") ?? "");
      if (id) {
        const link = ctx.doc.createElementNS(W, "w:hyperlink");
        link.setAttributeNS(R, "r:id", id);
        appendInline(ctx, el, link, f, del, change);
        parent.appendChild(link);
      } else {
        appendInline(ctx, el, parent, f, del, change);
      }
      continue;
    }
    if (tag === "img") {
      // A newly inserted image (existing ones carry data-docx-xml, handled above).
      const src = el.getAttribute("src") ?? "";
      const w = Number(el.getAttribute("width")) || (el as HTMLImageElement).naturalWidth || 0;
      const h = Number(el.getAttribute("height")) || (el as HTMLImageElement).naturalHeight || 0;
      if (src.startsWith("data:")) {
        const run = buildImageDrawing(ctx, src, w, h);
        if (run) parent.appendChild(run);
      }
      continue;
    }
    const deco = `${el.style.textDecoration || ""} ${el.style.textDecorationLine || ""}`;
    const bgHex = toHex6(el.style.backgroundColor);
    // A background colour maps to a named highlight when it matches one exactly, else to
    // arbitrary run shading (w:shd) so any colour from the picker round-trips.
    let highlight = f.highlight;
    let shading = f.shading;
    if (bgHex) {
      const name = HL_BY_HEX.get(bgHex);
      highlight = name;
      shading = name ? undefined : bgHex;
    }
    const next: Fmt = {
      b: f.b || tag === "strong" || tag === "b" || /(^|;)\s*font-weight\s*:\s*(bold|[6-9]00)/.test(el.style.cssText),
      i: f.i || tag === "em" || tag === "i" || el.style.fontStyle === "italic",
      u: f.u || tag === "u" || /underline/.test(deco),
      strike: f.strike || tag === "s" || tag === "strike" || tag === "del" || /line-through/.test(deco),
      vertAlign: f.vertAlign ?? (tag === "sup" || /vertical-align:\s*super/.test(el.style.cssText) ? "super" : tag === "sub" || /vertical-align:\s*sub/.test(el.style.cssText) ? "sub" : undefined),
      color: toHex6(el.style.color) ?? f.color,
      highlight,
      shading,
      sizeHalfPt: fontSizeToHalfPt(el.style.fontSize) ?? f.sizeHalfPt,
      font: firstFontFamily(el.style.fontFamily) ?? f.font,
    };
    appendInline(ctx, el, parent, next, del, change);
  }
}

function makeParagraph(ctx: DocxCtx, src: HTMLElement, opts: { heading?: number; list?: boolean }): Element {
  const p = ctx.doc.createElementNS(W, "w:p");
  const jc = JC_BY_ALIGN[src.style.textAlign || ""];
  const indentPx = parseFloat(src.style.marginLeft) || 0;
  const lineHeight = parseFloat(src.style.lineHeight) || 0; // unitless multiple
  const revPara = src.getAttribute("data-rev-para"); // "ins" | "del" paragraph-mark revision
  if (opts.heading || (opts.list && ctx.listNumId) || jc || revPara || indentPx > 0 || lineHeight > 0) {
    const pPr = ctx.doc.createElementNS(W, "w:pPr");
    if (opts.heading) {
      const st = ctx.doc.createElementNS(W, "w:pStyle");
      st.setAttributeNS(W, "w:val", `Heading${opts.heading}`);
      pPr.appendChild(st);
    }
    if (opts.list && ctx.listNumId) {
      const numPr = ctx.doc.createElementNS(W, "w:numPr");
      const ilvl = ctx.doc.createElementNS(W, "w:ilvl");
      ilvl.setAttributeNS(W, "w:val", "0");
      const numId = ctx.doc.createElementNS(W, "w:numId");
      numId.setAttributeNS(W, "w:val", ctx.listNumId);
      numPr.append(ilvl, numId);
      pPr.appendChild(numPr);
    }
    // schema order: w:spacing and w:ind come before w:jc
    if (lineHeight > 0) {
      const sp = ctx.doc.createElementNS(W, "w:spacing");
      sp.setAttributeNS(W, "w:line", String(Math.round(lineHeight * 240)));
      sp.setAttributeNS(W, "w:lineRule", "auto");
      pPr.appendChild(sp);
    }
    if (indentPx > 0) {
      const ind = ctx.doc.createElementNS(W, "w:ind");
      ind.setAttributeNS(W, "w:left", String(Math.round(indentPx * 15)));
      pPr.appendChild(ind);
    }
    if (jc) {
      const j = ctx.doc.createElementNS(W, "w:jc");
      j.setAttributeNS(W, "w:val", jc);
      pPr.appendChild(j);
    }
    if (revPara === "ins" || revPara === "del") {
      const rPr = ctx.doc.createElementNS(W, "w:rPr");
      const rev = ctx.doc.createElementNS(W, revPara === "del" ? "w:del" : "w:ins");
      rev.setAttributeNS(W, "w:id", String(ctx.nextRevId++));
      rev.setAttributeNS(W, "w:author", src.getAttribute("data-rev-author") || "Author");
      const d = src.getAttribute("data-rev-date");
      if (d) rev.setAttributeNS(W, "w:date", d);
      rPr.appendChild(rev);
      pPr.appendChild(rPr);
    }
    p.appendChild(pPr);
  }
  appendInline(ctx, src, p, FMT0);
  return p;
}

/** Rebuild a w:tbl from its preserved skeleton, replacing each (non-vMerge-continue) cell's
    blocks with the edited content from the matching .docx-cell. Structure, properties and
    spans are kept from the skeleton; only cell content changes. */
function rebuildTable(ctx: DocxCtx, tableEl: HTMLElement, stash: string): Element | null {
  const tbl = importPassthrough(ctx, stash);
  if (!tbl || tbl.tagName !== "w:tbl") return tbl;
  const cells = Array.from(tableEl.querySelectorAll(".docx-cell"));
  let i = 0;
  for (const tr of Array.from(tbl.children)) {
    if (tr.tagName !== "w:tr") continue;
    for (const tc of Array.from(tr.children)) {
      if (tc.tagName !== "w:tc") continue;
      const tcPr = tc.getElementsByTagName("w:tcPr")[0];
      const vm = tcPr?.getElementsByTagName("w:vMerge")[0];
      if (vm && (vm.getAttribute("w:val") ?? "continue") === "continue") continue; // continuation cell, untouched
      const cellEl = cells[i++];
      if (!cellEl) continue;
      for (const ch of Array.from(tc.children)) if (ch.tagName !== "w:tcPr") tc.removeChild(ch);
      for (const node of Array.from(cellEl.childNodes)) appendBlock(ctx, tc, node);
      if (!Array.from(tc.children).some((ch) => ch.tagName === "w:p" || ch.tagName === "w:tbl")) {
        tc.appendChild(ctx.doc.createElementNS(W, "w:p")); // a cell needs at least one block
      }
    }
  }
  return tbl;
}

const colSpanOf = (td: HTMLTableCellElement): number => td.colSpan || 1;

/** Build a w:tbl from a table in the DOM (a newly inserted or a structurally-edited one).
    Preserved tblPr/tblGrid/tcPr (data-attributes from the read) are reused so styling
    survives a structural edit; missing ones get sensible defaults. colspan -> w:gridSpan. */
function buildNewTable(ctx: DocxCtx, tableEl: HTMLElement): Element {
  const rows = Array.from((tableEl as HTMLTableElement).rows);
  const gridCols = Math.max(1, ...rows.map((tr) => Array.from(tr.cells).reduce((s, td) => s + colSpanOf(td), 0)));
  const colW = Math.round(9000 / gridCols);
  const importEl = (xml: string | null, tag: string): Element | null => {
    if (!xml) return null;
    const el = importPassthrough(ctx, xml);
    return el && el.tagName === tag ? el : null;
  };
  const tbl = ctx.doc.createElementNS(W, "w:tbl");

  let tblPr = importEl(tableEl.getAttribute("data-docx-tblpr"), "w:tblPr");
  if (!tblPr) {
    tblPr = ctx.doc.createElementNS(W, "w:tblPr");
    const tblW = ctx.doc.createElementNS(W, "w:tblW");
    tblW.setAttributeNS(W, "w:w", "0");
    tblW.setAttributeNS(W, "w:type", "auto");
    tblPr.appendChild(tblW);
    const borders = ctx.doc.createElementNS(W, "w:tblBorders");
    for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
      const b = ctx.doc.createElementNS(W, `w:${side}`);
      b.setAttributeNS(W, "w:val", "single");
      b.setAttributeNS(W, "w:sz", "4");
      b.setAttributeNS(W, "w:space", "0");
      b.setAttributeNS(W, "w:color", "auto");
      borders.appendChild(b);
    }
    tblPr.appendChild(borders);
  }
  tbl.appendChild(tblPr);

  let grid = importEl(tableEl.getAttribute("data-docx-tblgrid"), "w:tblGrid");
  if (!grid || grid.getElementsByTagName("w:gridCol").length !== gridCols) {
    grid = ctx.doc.createElementNS(W, "w:tblGrid");
    for (let i = 0; i < gridCols; i++) {
      const gc = ctx.doc.createElementNS(W, "w:gridCol");
      gc.setAttributeNS(W, "w:w", String(colW));
      grid.appendChild(gc);
    }
  }
  tbl.appendChild(grid);

  // Walk the DOM grid, tracking columns covered by a rowspan from above so a vertical merge
  // is written as w:vMerge restart + continue cells in the rows below.
  const setSpan = (tcPr: Element, tag: string, val?: string): void => {
    const old = tcPr.getElementsByTagName(tag)[0];
    if (old) old.parentNode?.removeChild(old);
    const el = ctx.doc.createElementNS(W, tag);
    if (val !== undefined) el.setAttributeNS(W, "w:val", val);
    tcPr.insertBefore(el, tcPr.firstChild);
  };
  const occupied: ({ rows: number; span: number } | null)[] = [];
  for (const tr of rows) {
    const wtr = ctx.doc.createElementNS(W, "w:tr");
    const domCells = Array.from(tr.cells);
    let di = 0;
    let col = 0;
    while (col < gridCols) {
      const occ = occupied[col];
      if (occ && occ.rows > 0) {
        // covered by a rowspan started above -> a vMerge continuation cell
        const tc = ctx.doc.createElementNS(W, "w:tc");
        const tcPr = ctx.doc.createElementNS(W, "w:tcPr");
        if (occ.span > 1) setSpan(tcPr, "w:gridSpan", String(occ.span));
        setSpan(tcPr, "w:vMerge");
        tc.appendChild(tcPr);
        tc.appendChild(ctx.doc.createElementNS(W, "w:p"));
        wtr.appendChild(tc);
        occ.rows--;
        col += occ.span;
        continue;
      }
      const td = domCells[di++];
      if (!td) break;
      const colspan = colSpanOf(td);
      const rowspan = (td as HTMLTableCellElement).rowSpan || 1;
      const tc = ctx.doc.createElementNS(W, "w:tc");
      let tcPr = importEl(td.getAttribute("data-docx-tcpr"), "w:tcPr");
      if (!tcPr) {
        tcPr = ctx.doc.createElementNS(W, "w:tcPr");
        const tcW = ctx.doc.createElementNS(W, "w:tcW");
        tcW.setAttributeNS(W, "w:w", String(colW * colspan));
        tcW.setAttributeNS(W, "w:type", "dxa");
        tcPr.appendChild(tcW);
      }
      // remove any preserved span markers, then set from the current DOM (gridSpan before vMerge)
      for (const t of ["w:gridSpan", "w:vMerge"]) {
        const old = tcPr.getElementsByTagName(t)[0];
        if (old) old.parentNode?.removeChild(old);
      }
      if (rowspan > 1) setSpan(tcPr, "w:vMerge", "restart");
      if (colspan > 1) setSpan(tcPr, "w:gridSpan", String(colspan));
      tc.appendChild(tcPr);
      const cell = (td.querySelector(".docx-cell") as HTMLElement) ?? td;
      for (const node of Array.from(cell.childNodes)) appendBlock(ctx, tc, node);
      if (!Array.from(tc.children).some((ch) => ch.tagName === "w:p" || ch.tagName === "w:tbl")) {
        tc.appendChild(ctx.doc.createElementNS(W, "w:p"));
      }
      wtr.appendChild(tc);
      if (rowspan > 1) occupied[col] = { rows: rowspan - 1, span: colspan };
      col += colspan;
    }
    tbl.appendChild(wtr);
  }
  return tbl;
}

function appendBlock(ctx: DocxCtx, body: Element, node: Node): void {
  if (node.nodeType === 3) {
    if (!(node.textContent ?? "").trim()) return;
    const p = ctx.doc.createElementNS(W, "w:p");
    p.appendChild(makeRun(ctx, node.textContent ?? "", FMT0));
    body.appendChild(p);
    return;
  }
  if (node.nodeType !== 1) return;
  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  if (el.classList.contains("docx-table")) {
    // Editable table: rebuild from the preserved skeleton if it came from the file, or
    // build a fresh w:tbl from the DOM if it was inserted in the editor.
    const stash = el.getAttribute("data-docx-xml");
    const t = stash ? rebuildTable(ctx, el, stash) : buildNewTable(ctx, el);
    if (t) body.appendChild(t);
    return;
  }
  const stash = el.getAttribute("data-docx-xml");
  if (stash) {
    // A passthrough block (a preserved element we do not model): re-emit it verbatim.
    const node2 = importPassthrough(ctx, stash);
    if (node2) body.appendChild(node2);
    return;
  }
  const pb = el.getAttribute("data-docx-pagebreak");
  if (pb) {
    if (pb !== "auto") {
      const p = ctx.doc.createElementNS(W, "w:p");
      p.appendChild(pageBreakRun(ctx));
      body.appendChild(p);
    }
    return;
  }
  if (tag === "ul" || tag === "ol") {
    for (const li of Array.from(el.children)) {
      if (li.tagName.toLowerCase() !== "li") continue;
      body.appendChild(makeParagraph(ctx, li as HTMLElement, { list: true }));
    }
    return;
  }
  const hm = /^h([1-6])$/.exec(tag);
  body.appendChild(makeParagraph(ctx, el, hm ? { heading: Number(hm[1]) } : {}));
}

function firstNumId(file: Uint8Array | undefined): string | null {
  if (!file) return null;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  const num = doc.getElementsByTagName("w:num")[0];
  return num?.getAttributeNS(W, "numId") ?? num?.getAttribute("w:numId") ?? null;
}

const relsPathFor = (partPath: string): string => {
  const slash = partPath.lastIndexOf("/");
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
};

/** Rebuild one part (document.xml or a header/footer) in-place from its edited HTML. */
function rebuildPart(files: Record<string, Uint8Array>, partPath: string, html: string): void {
  const xml = files[partPath];
  if (!xml) return;
  const doc = new DOMParser().parseFromString(strFromU8(xml), "application/xml");
  const isBody = /document\.xml$/.test(partPath);
  const container = isBody ? doc.getElementsByTagName("w:body")[0] : doc.documentElement;
  if (!container) {
    if (isBody) throw new Error("not a .docx: w:body missing");
    return;
  }
  // For the body, keep the trailing section properties (page setup).
  const last = container.lastElementChild;
  const sectPr = isBody && last && last.tagName === "w:sectPr" ? last : null;
  while (container.firstChild) container.removeChild(container.firstChild);

  const relsPath = relsPathFor(partPath);
  const relsXml = files[relsPath];
  // Revision ids must be unique; continue past the highest already in the part.
  let maxRev = 0;
  for (const r of [...Array.from(doc.getElementsByTagName("w:ins")), ...Array.from(doc.getElementsByTagName("w:del"))]) {
    const n = Number(r.getAttribute("w:id"));
    if (Number.isFinite(n)) maxRev = Math.max(maxRev, n);
  }
  const ctx: DocxCtx = {
    doc,
    rels: relsXml ? new DOMParser().parseFromString(strFromU8(relsXml), "application/xml") : null,
    relsAdded: false,
    nextRid: 1,
    nextRevId: maxRev + 1,
    listNumId: firstNumId(files["word/numbering.xml"]),
    files,
  };
  if (ctx.rels) {
    let max = 0;
    for (const r of Array.from(ctx.rels.getElementsByTagName("Relationship"))) {
      const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
      if (m) max = Math.max(max, Number(m[1]));
    }
    ctx.nextRid = max + 1;
  }

  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  for (const node of Array.from(htmlDoc.body.childNodes)) appendBlock(ctx, container, node);
  if (!container.firstChild) container.appendChild(doc.createElementNS(W, "w:p"));
  if (sectPr) container.appendChild(sectPr);

  files[partPath] = strToU8(new XMLSerializer().serializeToString(doc));
  if (ctx.relsAdded && ctx.rels) {
    files[relsPath] = strToU8(new XMLSerializer().serializeToString(ctx.rels));
  }
}

const CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

/** Add an Override for one part to [Content_Types].xml (idempotent). */
function ensureOverride(files: Record<string, Uint8Array>, partName: string, mime: string): void {
  const key = "[Content_Types].xml";
  const xml = files[key];
  if (!xml) return;
  const doc = new DOMParser().parseFromString(strFromU8(xml), "application/xml");
  const has = Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === partName);
  if (has) return;
  const ov = doc.createElementNS(doc.documentElement!.namespaceURI, "Override");
  ov.setAttribute("PartName", partName);
  ov.setAttribute("ContentType", mime);
  doc.documentElement!.appendChild(ov);
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Create a new header/footer part from edited HTML and wire it into the body section:
    a new word/<kind>N.xml, a relationship in document.xml.rels, a w:headerReference /
    w:footerReference in the section properties, and a content-type override. */
function createHeaderFooterPart(files: Record<string, Uint8Array>, kind: "header" | "footer", html: string): void {
  if (!files["word/document.xml"]) return;
  const isHeader = kind === "header";
  let n = 1;
  while (files[`word/${kind}${n}.xml`]) n++;
  const partPath = `word/${kind}${n}.xml`;

  // Build the part from the edited HTML, reusing the body block conversion.
  const decls = Object.entries(NS_DECLS).map(([k, v]) => `${k}="${v}"`).join(" ");
  const partDoc = new DOMParser().parseFromString(`<${isHeader ? "w:hdr" : "w:ftr"} ${decls}></${isHeader ? "w:hdr" : "w:ftr"}>`, "application/xml");
  const container = partDoc.documentElement!;
  const relsDoc = new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  const ctx: DocxCtx = { doc: partDoc, rels: relsDoc, relsAdded: false, nextRid: 1, nextRevId: 1, listNumId: firstNumId(files["word/numbering.xml"]), files };
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  for (const node of Array.from(htmlDoc.body.childNodes)) appendBlock(ctx, container, node);
  if (!container.firstChild) container.appendChild(partDoc.createElementNS(W, "w:p"));
  files[partPath] = strToU8(new XMLSerializer().serializeToString(partDoc));
  if (ctx.relsAdded) files[`word/_rels/${kind}${n}.xml.rels`] = strToU8(new XMLSerializer().serializeToString(relsDoc));
  ensureOverride(files, `/${partPath}`, isHeader ? CT_HEADER : CT_FOOTER);

  // Relationship from the document to the new part.
  const relsPath = "word/_rels/document.xml.rels";
  const dRels = files[relsPath]
    ? new DOMParser().parseFromString(strFromU8(files[relsPath]!), "application/xml")
    : new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  let maxR = 0;
  for (const r of Array.from(dRels.getElementsByTagName("Relationship"))) {
    const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
    if (m) maxR = Math.max(maxR, Number(m[1]));
  }
  const rid = `rId${maxR + 1}`;
  const rel = dRels.createElementNS(PKG, "Relationship");
  rel.setAttribute("Id", rid);
  rel.setAttribute("Type", isHeader ? REL_HEADER : REL_FOOTER);
  rel.setAttribute("Target", `${kind}${n}.xml`);
  dRels.documentElement!.appendChild(rel);
  files[relsPath] = strToU8(new XMLSerializer().serializeToString(dRels));

  // Reference it from the section properties (header/footer refs lead the sectPr sequence).
  const bodyDoc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]!), "application/xml");
  const sectPr = Array.from(bodyDoc.getElementsByTagName("w:sectPr")).pop();
  if (sectPr) {
    const ref = bodyDoc.createElementNS(W, isHeader ? "w:headerReference" : "w:footerReference");
    ref.setAttributeNS(W, "w:type", "default");
    ref.setAttributeNS(R, "r:id", rid);
    sectPr.insertBefore(ref, sectPr.firstChild);
    files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(bodyDoc));
  }
}

const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/** Ensure comments.xml is declared in [Content_Types].xml and related from document.xml. */
function ensureCommentsRegistered(files: Record<string, Uint8Array>): void {
  const ctKey = "[Content_Types].xml";
  if (files[ctKey]) {
    const doc = new DOMParser().parseFromString(strFromU8(files[ctKey]), "application/xml");
    const has = Array.from(doc.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === "/word/comments.xml");
    if (!has) {
      const ov = doc.createElementNS(doc.documentElement!.namespaceURI, "Override");
      ov.setAttribute("PartName", "/word/comments.xml");
      ov.setAttribute("ContentType", CT_COMMENTS);
      doc.documentElement!.appendChild(ov);
      files[ctKey] = strToU8(new XMLSerializer().serializeToString(doc));
    }
  }
  const relsKey = "word/_rels/document.xml.rels";
  if (files[relsKey]) {
    const doc = new DOMParser().parseFromString(strFromU8(files[relsKey]), "application/xml");
    const has = Array.from(doc.getElementsByTagName("Relationship")).some((r) => r.getAttribute("Type") === REL_COMMENTS);
    if (!has) {
      let max = 0;
      for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
        const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
        if (m) max = Math.max(max, Number(m[1]));
      }
      const rel = doc.createElementNS(PKG, "Relationship");
      rel.setAttribute("Id", `rId${max + 1}`);
      rel.setAttribute("Type", REL_COMMENTS);
      rel.setAttribute("Target", "comments.xml");
      doc.documentElement!.appendChild(rel);
      files[relsKey] = strToU8(new XMLSerializer().serializeToString(doc));
    }
  }
}

/** Append any newly-added comments (markers carry the text) to word/comments.xml. */
function applyNewComments(files: Record<string, Uint8Array>, html: string): void {
  const dom = new DOMParser().parseFromString(html || "<p></p>", "text/html");
  const fresh = Array.from(dom.querySelectorAll("[data-comment-new]"));
  if (!fresh.length) return;
  ensureCommentsRegistered(files);
  const existing = files["word/comments.xml"];
  const doc = existing
    ? new DOMParser().parseFromString(strFromU8(existing), "application/xml")
    : new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W}"></w:comments>`, "application/xml");
  const root = doc.getElementsByTagName("w:comments")[0] ?? doc.documentElement!;
  for (const n of fresh) {
    const c = doc.createElementNS(W, "w:comment");
    c.setAttributeNS(W, "w:id", n.getAttribute("data-comment-id") ?? "0");
    c.setAttributeNS(W, "w:author", n.getAttribute("data-comment-author") || "Author");
    const date = n.getAttribute("data-comment-date");
    if (date) c.setAttributeNS(W, "w:date", date);
    const p = doc.createElementNS(W, "w:p");
    const paraId = n.getAttribute("data-comment-paraid");
    if (paraId) p.setAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "w14:paraId", paraId);
    const r = doc.createElementNS(W, "w:r");
    const t = doc.createElementNS(W, "w:t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = n.getAttribute("data-comment-text") ?? "";
    r.appendChild(t);
    p.appendChild(r);
    c.appendChild(p);
    root.appendChild(c);
  }
  files["word/comments.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

export interface ReactionEdit {
  commentId: string;
  emoji: string;
  person: string;
  date: string;
}

/** Append emoji reactions to existing comments (in the Google-Docs plain-text form). */
function applyReactions(files: Record<string, Uint8Array>, reactions: ReactionEdit[]): void {
  if (!reactions.length || !files["word/comments.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["word/comments.xml"]), "application/xml");
  const byId = new Map<string, Element>();
  for (const c of Array.from(doc.getElementsByTagName("w:comment"))) {
    const id = c.getAttribute("w:id");
    if (id) byId.set(id, c);
  }
  for (const re of reactions) {
    const c = byId.get(re.commentId);
    if (!c) continue;
    const p = doc.createElementNS(W, "w:p");
    const r = doc.createElementNS(W, "w:r");
    const tx = doc.createElementNS(W, "w:t");
    tx.setAttribute("xml:space", "preserve");
    tx.textContent = `${re.person} a réagi avec ${re.emoji} à ${re.date}`;
    r.appendChild(tx);
    p.appendChild(r);
    c.appendChild(p);
  }
  files["word/comments.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const REL_COMMENTS_EXT = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const CT_COMMENTS_EXT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";

export interface ReplyEdit {
  id: string;
  paraId: string;
  parentParaId: string;
  author: string;
  date: string;
  text: string;
}

/** Get (or create + register) the commentsExtended.xml document. */
function getCommentsExtDoc(files: Record<string, Uint8Array>): Document {
  const existing = files["word/commentsExtended.xml"];
  if (existing) return new DOMParser().parseFromString(strFromU8(existing), "application/xml");
  const ct = files["[Content_Types].xml"];
  if (ct) {
    const d = new DOMParser().parseFromString(strFromU8(ct), "application/xml");
    if (!Array.from(d.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === "/word/commentsExtended.xml")) {
      const ov = d.createElementNS(d.documentElement!.namespaceURI, "Override");
      ov.setAttribute("PartName", "/word/commentsExtended.xml");
      ov.setAttribute("ContentType", CT_COMMENTS_EXT);
      d.documentElement!.appendChild(ov);
      files["[Content_Types].xml"] = strToU8(new XMLSerializer().serializeToString(d));
    }
  }
  const relsKey = "word/_rels/document.xml.rels";
  if (files[relsKey]) {
    const d = new DOMParser().parseFromString(strFromU8(files[relsKey]), "application/xml");
    if (!Array.from(d.getElementsByTagName("Relationship")).some((r) => r.getAttribute("Type") === REL_COMMENTS_EXT)) {
      let max = 0;
      for (const r of Array.from(d.getElementsByTagName("Relationship"))) {
        const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
        if (m) max = Math.max(max, Number(m[1]));
      }
      const rel = d.createElementNS(PKG, "Relationship");
      rel.setAttribute("Id", `rId${max + 1}`);
      rel.setAttribute("Type", REL_COMMENTS_EXT);
      rel.setAttribute("Target", "commentsExtended.xml");
      d.documentElement!.appendChild(rel);
      files[relsKey] = strToU8(new XMLSerializer().serializeToString(d));
    }
  }
  return new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="${W15}"></w15:commentsEx>`, "application/xml");
}

/** Add reply comments to comments.xml and thread them in commentsExtended.xml. */
function applyReplies(files: Record<string, Uint8Array>, replies: ReplyEdit[]): void {
  if (!replies.length) return;
  ensureCommentsRegistered(files);
  const cdoc = files["word/comments.xml"]
    ? new DOMParser().parseFromString(strFromU8(files["word/comments.xml"]), "application/xml")
    : new DOMParser().parseFromString(`<?xml version="1.0"?><w:comments xmlns:w="${W}"></w:comments>`, "application/xml");
  const croot = cdoc.getElementsByTagName("w:comments")[0] ?? cdoc.documentElement!;
  const edoc = getCommentsExtDoc(files);
  const eroot = edoc.getElementsByTagName("w15:commentsEx")[0] ?? edoc.documentElement!;
  for (const re of replies) {
    const c = cdoc.createElementNS(W, "w:comment");
    c.setAttributeNS(W, "w:id", re.id);
    c.setAttributeNS(W, "w:author", re.author);
    if (re.date) c.setAttributeNS(W, "w:date", re.date);
    const p = cdoc.createElementNS(W, "w:p");
    p.setAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "w14:paraId", re.paraId);
    const r = cdoc.createElementNS(W, "w:r");
    const tx = cdoc.createElementNS(W, "w:t");
    tx.setAttribute("xml:space", "preserve");
    tx.textContent = re.text;
    r.appendChild(tx);
    p.appendChild(r);
    c.appendChild(p);
    croot.appendChild(c);
    const ex = edoc.createElementNS(W15, "w15:commentEx");
    ex.setAttributeNS(W15, "w15:paraId", re.paraId);
    if (re.parentParaId) ex.setAttributeNS(W15, "w15:paraIdParent", re.parentParaId);
    ex.setAttributeNS(W15, "w15:done", "0");
    eroot.appendChild(ex);
  }
  files["word/comments.xml"] = strToU8(new XMLSerializer().serializeToString(cdoc));
  files["word/commentsExtended.xml"] = strToU8(new XMLSerializer().serializeToString(edoc));
}

/** Set w15:done for resolved threads in commentsExtended.xml. */
function applyDone(files: Record<string, Uint8Array>, done: Map<string, boolean>): void {
  if (!done.size) return;
  const edoc = getCommentsExtDoc(files);
  const eroot = edoc.getElementsByTagName("w15:commentsEx")[0] ?? edoc.documentElement!;
  const byPara = new Map<string, Element>();
  for (const ex of Array.from(edoc.getElementsByTagName("w15:commentEx"))) {
    const pid = ex.getAttribute("w15:paraId");
    if (pid) byPara.set(pid, ex);
  }
  for (const [paraId, value] of done) {
    if (!paraId) continue;
    let ex = byPara.get(paraId);
    if (!ex) {
      ex = edoc.createElementNS(W15, "w15:commentEx");
      ex.setAttributeNS(W15, "w15:paraId", paraId);
      eroot.appendChild(ex);
    }
    ex.setAttributeNS(W15, "w15:done", value ? "1" : "0");
  }
  files["word/commentsExtended.xml"] = strToU8(new XMLSerializer().serializeToString(edoc));
}

/** Remove deleted comments from comments.xml and commentsExtended.xml. */
function applyDeletedComments(files: Record<string, Uint8Array>, ids: string[]): void {
  if (!ids.length || !files["word/comments.xml"]) return;
  const cdoc = new DOMParser().parseFromString(strFromU8(files["word/comments.xml"]), "application/xml");
  const idSet = new Set(ids);
  const goneParaIds = new Set<string>();
  for (const c of Array.from(cdoc.getElementsByTagName("w:comment"))) {
    if (idSet.has(c.getAttribute("w:id") ?? "")) {
      for (const p of Array.from(c.getElementsByTagName("w:p"))) {
        const pid = p.getAttribute("w14:paraId");
        if (pid) goneParaIds.add(pid);
      }
      c.parentNode?.removeChild(c);
    }
  }
  files["word/comments.xml"] = strToU8(new XMLSerializer().serializeToString(cdoc));
  const ext = files["word/commentsExtended.xml"];
  if (ext && goneParaIds.size) {
    const edoc = new DOMParser().parseFromString(strFromU8(ext), "application/xml");
    for (const ex of Array.from(edoc.getElementsByTagName("w15:commentEx"))) {
      if (goneParaIds.has(ex.getAttribute("w15:paraId") ?? "")) ex.parentNode?.removeChild(ex);
    }
    files["word/commentsExtended.xml"] = strToU8(new XMLSerializer().serializeToString(edoc));
  }
}

/**
 * Rebuild a .docx from edited HTML, preserving every other part of the archive. Pass
 * `parts` to also write back edited header/footer parts, and `opts` for comment edits.
 */
/** Update the body section's w:pgMar from edited margins (px -> twips). */
function applyPageMargins(files: Record<string, Uint8Array>, geometry: PageGeometry): void {
  const raw = files["word/document.xml"];
  if (!raw) return;
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const sectPr = Array.from(doc.getElementsByTagName("w:sectPr")).pop();
  if (!sectPr) return;
  let pgMar = sectPr.getElementsByTagName("w:pgMar")[0];
  if (!pgMar) {
    pgMar = doc.createElementNS(W, "w:pgMar");
    sectPr.appendChild(pgMar);
  }
  const tw = (px: number): string => String(Math.round(px * 15));
  pgMar.setAttributeNS(W, "w:top", tw(geometry.margin.top));
  pgMar.setAttributeNS(W, "w:right", tw(geometry.margin.right));
  pgMar.setAttributeNS(W, "w:bottom", tw(geometry.margin.bottom));
  pgMar.setAttributeNS(W, "w:left", tw(geometry.margin.left));
  files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

export function htmlToDocx(
  html: string,
  original: Uint8Array,
  parts?: { path: string; html: string }[],
  opts?: { reactions?: ReactionEdit[]; replies?: ReplyEdit[]; done?: Map<string, boolean>; deletedComments?: string[]; pageGeometry?: PageGeometry },
): Uint8Array {
  const files = unzipSync(original);
  rebuildPart(files, "word/document.xml", html);
  for (const p of parts ?? []) {
    // "header" / "footer" are sentinels for a band created in-editor (no existing part).
    if (p.path === "header" || p.path === "footer") createHeaderFooterPart(files, p.path, p.html);
    else rebuildPart(files, p.path, p.html);
  }
  applyNewComments(files, html);
  applyReactions(files, opts?.reactions ?? []);
  applyReplies(files, opts?.replies ?? []);
  applyDone(files, opts?.done ?? new Map());
  applyDeletedComments(files, opts?.deletedComments ?? []);
  if (opts?.pageGeometry) applyPageMargins(files, opts.pageGeometry);
  return zipSync(files);
}

