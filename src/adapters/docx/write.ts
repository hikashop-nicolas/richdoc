// docx WRITE: rebuild a .docx archive from edited HTML, preserving every untouched part.
// Pure HTML -> XML (body, header/footer, comments, reactions, replies, page margins); the
// read half lives in ./read.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { toHex6, fontSizeToHalfPt, firstFontFamily, imageLayoutFromEl, blockBorders, parseCssBorder } from "../../core/util";
import type { ImageLayout, NewStyle, Note, PageGeometry } from "../../core/types";
import { W, R, PKG, REL_HYPERLINK, NS_DECLS, FMT0, HL_BY_HEX, JC_BY_ALIGN } from "./shared";
import type { Fmt } from "./shared";
import { EMU_PER_PX, makeContainer } from "./image-layout";
import { mathmlToOmml } from "./omml";

// ---------------------------------------------------------------------------
// HTML -> .docx
// ---------------------------------------------------------------------------

interface DocxCtx {
  doc: Document;
  rels: Document | null;
  relsAdded: boolean;
  nextRid: number;
  nextRevId: number; // next w:ins/w:del revision id
  listIds: { bullet: string; ordered: string } | null; // resolved/created list numbering ids (lazy)
  orderedBaseUsed?: boolean; // the base ordered numId has been claimed by the first plain list
  files: Record<string, Uint8Array>; // the archive, so new media/content-types can be added
  sectionBandHtml?: Map<string, string>; // section header/footer key -> HTML, for minting new parts
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

/** Embed a data: URL image as a new media part + relationship; return a w:drawing run.
 *  `layout` null = inline; otherwise the run is anchored (floating) with the given wrap. */
function buildImageDrawing(ctx: DocxCtx, src: string, widthPx: number, heightPx: number, layout: ImageLayout | null, alt: string): Element | null {
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

  const cx = Math.max(1, Math.round((widthPx || 200) * EMU_PER_PX));
  const cy = Math.max(1, Math.round((heightPx || 200) * EMU_PER_PX));
  const ce = (ns: string, name: string) => ctx.doc.createElementNS(ns, name);
  const docPr = ce(WP_NS, "wp:docPr");
  docPr.setAttribute("id", String(ctx.nextRid));
  docPr.setAttribute("name", `Image ${ctx.nextRid}`);
  if (alt) docPr.setAttribute("descr", alt);
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
  const r = ce(W, "w:r");
  const drawing = ce(W, "w:drawing");
  drawing.appendChild(makeContainer(ctx.doc, graphic, docPr, cx, cy, layout));
  r.appendChild(drawing);
  return r;
}

/** Serialize an edited <img> back to a run. A preserved image (data-docx-xml carrying its
 *  original drawing) is rebuilt around its own a:graphic so the blip relationship survives,
 *  picking up the editor's current size + wrap; a brand-new image embeds fresh media. Returns
 *  null when there is nothing to emit (caller falls back to verbatim passthrough). */
function buildImageRun(ctx: DocxCtx, img: HTMLElement): Element | null {
  const layout = imageLayoutFromEl(img);
  const w = Number(img.getAttribute("width")) || (img as HTMLImageElement).naturalWidth || 0;
  const h = Number(img.getAttribute("height")) || (img as HTMLImageElement).naturalHeight || 0;
  const alt = img.getAttribute("alt") || "";
  const stash = img.getAttribute("data-docx-xml");
  if (stash) {
    const run = importPassthrough(ctx, stash);
    const drawing = run?.getElementsByTagName("w:drawing")[0];
    const graphic = drawing?.getElementsByTagName("a:graphic")[0];
    // Only DrawingML images are rebuilt; VML (w:pict) / OLE keep their original markup.
    if (run && drawing && graphic) {
      let docPr = drawing.getElementsByTagName("wp:docPr")[0];
      if (!docPr) {
        docPr = ctx.doc.createElementNS(WP_NS, "wp:docPr");
        docPr.setAttribute("id", String(ctx.nextRid++));
        docPr.setAttribute("name", "Image");
      }
      if (alt) docPr.setAttribute("descr", alt); else docPr.removeAttribute("descr");
      const old = drawing.getElementsByTagName("wp:extent")[0];
      const cx = w ? Math.max(1, Math.round(w * EMU_PER_PX)) : Number(old?.getAttribute("cx")) || 1;
      const cy = h ? Math.max(1, Math.round(h * EMU_PER_PX)) : Number(old?.getAttribute("cy")) || 1;
      const container = drawing.firstElementChild; // wp:inline or wp:anchor
      drawing.replaceChild(makeContainer(ctx.doc, graphic, docPr, cx, cy, layout), container!);
      return run;
    }
    return run; // non-DrawingML: re-emit verbatim
  }
  const src = img.getAttribute("src") ?? "";
  return src.startsWith("data:") ? buildImageDrawing(ctx, src, w, h, layout, alt) : null;
}

const fmtHasProps = (f: Fmt): boolean => !!(f.b || f.i || f.u || f.strike || f.vertAlign || f.color || f.highlight || f.shading || f.sizeHalfPt || f.font || f.cStyle);

/** Append the property elements for a Fmt to a w:rPr, in OOXML schema order. */
function fillRPr(ctx: DocxCtx, rPr: Element, f: Fmt): void {
  const flag = (tag: string) => rPr.appendChild(ctx.doc.createElementNS(W, tag));
  const valEl = (tag: string, val: string) => {
    const el = ctx.doc.createElementNS(W, tag);
    el.setAttributeNS(W, "w:val", val);
    rPr.appendChild(el);
  };
  if (f.cStyle) valEl("w:rStyle", f.cStyle); // first in rPr schema order
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
    if (tag === "ul" || tag === "ol" || tag === "li") continue; // nested lists are emitted as block paragraphs, not inline
    // A field (page number / count / caption sequence): a w:fldSimple whose cached result is shown.
    if (el.classList.contains("docx-field")) {
      const k = el.getAttribute("data-field") || "PAGE";
      const instr = k === "seq" ? ` SEQ ${el.getAttribute("data-seq") || "Figure"} \\* ARABIC ` : ` ${k} `;
      const fld = ctx.doc.createElementNS(W, "w:fldSimple");
      fld.setAttributeNS(W, "w:instr", instr);
      fld.appendChild(makeRun(ctx, el.textContent || "", f, del, change));
      parent.appendChild(fld);
      continue;
    }
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
    // An image (new or preserved) is rebuilt so size + wrap edits round-trip; this must run
    // before the generic passthrough below, which would otherwise re-emit the stash verbatim.
    if (tag === "img") {
      const run = buildImageRun(ctx, el);
      if (run) parent.appendChild(run);
      continue;
    }
    if (tag === "span" && el.classList.contains("docx-eq")) {
      // An equation: original OMML verbatim if unedited (data-docx-xml), else converted from its MathML.
      const stash0 = el.getAttribute("data-docx-xml");
      const math = el.querySelector("math");
      const xml = stash0 ?? (math ? mathmlToOmml(math) : null);
      if (xml) { const node = importPassthrough(ctx, xml); if (node) parent.appendChild(node); }
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
    if (el.getAttribute("data-docx-tab")) {
      const r = ctx.doc.createElementNS(W, "w:r");
      r.appendChild(ctx.doc.createElementNS(W, "w:tab"));
      parent.appendChild(r);
      continue;
    }
    if (tag === "br") {
      const r = ctx.doc.createElementNS(W, "w:r");
      r.appendChild(ctx.doc.createElementNS(W, "w:br"));
      parent.appendChild(r);
      continue;
    }
    if (tag === "sup" && el.classList.contains("docx-fnref")) {
      // Footnote / endnote reference -> a w:footnoteReference / w:endnoteReference run.
      const kind = el.getAttribute("data-fn-kind") === "endnote" ? "endnote" : "footnote";
      const run = ctx.doc.createElementNS(W, "w:r");
      const rPr = ctx.doc.createElementNS(W, "w:rPr");
      const va = ctx.doc.createElementNS(W, "w:vertAlign");
      va.setAttributeNS(W, "w:val", "superscript");
      rPr.appendChild(va);
      run.appendChild(rPr);
      const ref = ctx.doc.createElementNS(W, kind === "endnote" ? "w:endnoteReference" : "w:footnoteReference");
      ref.setAttributeNS(W, "w:id", el.getAttribute("data-fn-id") || "0");
      run.appendChild(ref);
      parent.appendChild(run);
      continue;
    }
    if (tag === "ruby") {
      // Furigana: <ruby>base<rt>reading</rt></ruby> -> w:ruby (rubyPr, w:rt, w:rubyBase).
      const ruby = ctx.doc.createElementNS(W, "w:ruby");
      const prStash = el.getAttribute("data-docx-rubypr");
      const pr = prStash ? importPassthrough(ctx, prStash) : null;
      ruby.appendChild(pr && pr.tagName === "w:rubyPr" ? pr : ctx.doc.createElementNS(W, "w:rubyPr"));
      const rtEl = Array.from(el.children).find((c) => c.tagName.toLowerCase() === "rt") as HTMLElement | undefined;
      const rt = ctx.doc.createElementNS(W, "w:rt");
      if (rtEl) appendInline(ctx, rtEl, rt, f, del, change);
      const base = ctx.doc.createElementNS(W, "w:rubyBase");
      const baseClone = el.cloneNode(true) as HTMLElement;
      for (const r of Array.from(baseClone.children)) if (r.tagName.toLowerCase() === "rt") r.remove();
      appendInline(ctx, baseClone, base, f, del, change);
      ruby.append(rt, base);
      parent.appendChild(ruby);
      continue;
    }
    if (tag === "a" && el.classList.contains("docx-bookmark")) {
      const start = ctx.doc.createElementNS(W, "w:bookmarkStart");
      start.setAttributeNS(W, "w:id", el.getAttribute("data-rdoc-bm-id") || String(ctx.nextRevId++));
      start.setAttributeNS(W, "w:name", el.getAttribute("data-rdoc-bm") || "");
      parent.appendChild(start);
      continue;
    }
    if (tag === "a" && el.classList.contains("docx-bookmark-end")) {
      const end = ctx.doc.createElementNS(W, "w:bookmarkEnd");
      end.setAttributeNS(W, "w:id", el.getAttribute("data-rdoc-bm-id") || "0");
      parent.appendChild(end);
      continue;
    }
    if (tag === "a" && el.classList.contains("docx-xref")) {
      // Cross-reference -> a REF / PAGEREF simple field; \p adds the above/below relative position.
      const xfmt = el.getAttribute("data-rdoc-xref-fmt");
      const kind = xfmt === "page" ? "PAGEREF" : "REF";
      const sw = xfmt === "direction" ? "\\p " : "";
      const fld = ctx.doc.createElementNS(W, "w:fldSimple");
      fld.setAttributeNS(W, "w:instr", ` ${kind} ${el.getAttribute("data-rdoc-xref") || ""} ${sw}\\h `);
      const run = ctx.doc.createElementNS(W, "w:r");
      const t = ctx.doc.createElementNS(W, "w:t");
      t.setAttribute("xml:space", "preserve");
      t.textContent = el.textContent ?? "";
      run.appendChild(t);
      fld.appendChild(run);
      parent.appendChild(fld);
      continue;
    }
    if (tag === "a") {
      const href = el.getAttribute("href") ?? "";
      if (href.startsWith("#")) {
        // An internal link to a bookmark: w:hyperlink w:anchor (no relationship).
        const link = ctx.doc.createElementNS(W, "w:hyperlink");
        link.setAttributeNS(W, "w:anchor", href.slice(1));
        appendInline(ctx, el, link, f, del, change);
        parent.appendChild(link);
        continue;
      }
      const id = addHyperlinkRel(ctx, href);
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
      cStyle: el.getAttribute("data-rdoc-cstyle") || f.cStyle,
    };
    appendInline(ctx, el, parent, next, del, change);
  }
}

function makeParagraph(ctx: DocxCtx, src: HTMLElement, opts: { heading?: number; listLevel?: number; listNumId?: string }): Element {
  const p = ctx.doc.createElementNS(W, "w:p");
  const jc = JC_BY_ALIGN[src.style.textAlign || ""];
  const indentPx = parseFloat(src.style.marginLeft) || 0;
  const lineHeight = parseFloat(src.style.lineHeight) || 0; // unitless multiple
  // Space before/after: only when set inline (margin-top/bottom present), so 0 round-trips too.
  const hasBefore = src.style.marginTop !== "";
  const hasAfter = src.style.marginBottom !== "";
  const beforePx = parseFloat(src.style.marginTop) || 0;
  const afterPx = parseFloat(src.style.marginBottom) || 0;
  const revPara = src.getAttribute("data-rev-para"); // "ins" | "del" paragraph-mark revision
  // A non-heading named style; a caption paragraph defaults to the built-in "Caption" style.
  const namedStyle = !opts.heading ? src.getAttribute("data-rdoc-style") || (src.hasAttribute("data-rdoc-caption") ? "Caption" : null) : null;
  const sectXml = src.getAttribute("data-docx-sectpr"); // a mid-document section break to re-emit
  const secGeom = src.getAttribute("data-rdoc-secbreak"); // this section's geometry (JSON), for rendering
  // Regenerate the sectPr from the JSON only when the section was edited / inserted; an untouched
  // section passes its original sectPr through byte-for-byte.
  const regenSect = !!secGeom && (src.getAttribute("data-rdoc-secedited") === "1" || !sectXml);
  const tabStops = src.getAttribute("data-rdoc-tabstops"); // custom tab stops (JSON), schema-ordered before spacing
  const shadeHex = toHex6(src.style.backgroundColor); // paragraph shading -> w:shd
  const borders = blockBorders(src); // paragraph borders -> w:pBdr
  if (opts.heading || namedStyle || opts.listNumId || jc || revPara || sectXml || regenSect || tabStops || shadeHex || borders.length || indentPx > 0 || lineHeight > 0 || hasBefore || hasAfter) {
    const pPr = ctx.doc.createElementNS(W, "w:pPr");
    if (opts.heading || namedStyle) {
      const st = ctx.doc.createElementNS(W, "w:pStyle");
      st.setAttributeNS(W, "w:val", opts.heading ? `Heading${opts.heading}` : namedStyle!);
      pPr.appendChild(st);
    }
    if (opts.listNumId) {
      const numPr = ctx.doc.createElementNS(W, "w:numPr");
      const ilvl = ctx.doc.createElementNS(W, "w:ilvl");
      ilvl.setAttributeNS(W, "w:val", String(opts.listLevel ?? 0));
      const numId = ctx.doc.createElementNS(W, "w:numId");
      numId.setAttributeNS(W, "w:val", opts.listNumId);
      numPr.append(ilvl, numId);
      pPr.appendChild(numPr);
    }
    // w:pBdr then w:shd sit after numPr and before tabs/spacing in the pPr schema order.
    if (borders.length) {
      const pBdr = ctx.doc.createElementNS(W, "w:pBdr");
      for (const b of borders) {
        const e = ctx.doc.createElementNS(W, `w:${b.side}`);
        e.setAttributeNS(W, "w:val", DOCX_BORDER_VAL[b.style] ?? "single");
        e.setAttributeNS(W, "w:sz", String(Math.max(2, Math.round(b.px * 6)))); // px -> eighths of a point
        e.setAttributeNS(W, "w:space", "1");
        e.setAttributeNS(W, "w:color", b.hex);
        pBdr.appendChild(e);
      }
      pPr.appendChild(pBdr);
    }
    if (shadeHex) {
      const shd = ctx.doc.createElementNS(W, "w:shd");
      shd.setAttributeNS(W, "w:val", "clear");
      shd.setAttributeNS(W, "w:color", "auto");
      shd.setAttributeNS(W, "w:fill", shadeHex);
      pPr.appendChild(shd);
    }
    if (tabStops) {
      try {
        const stops = JSON.parse(tabStops) as { pos: number; val?: string; leader?: string }[];
        if (Array.isArray(stops) && stops.length) {
          const tabs = ctx.doc.createElementNS(W, "w:tabs");
          for (const s of stops) {
            const tb = ctx.doc.createElementNS(W, "w:tab");
            tb.setAttributeNS(W, "w:val", s.val || "left");
            tb.setAttributeNS(W, "w:pos", String(Math.round((s.pos || 0) * 15)));
            if (s.leader) tb.setAttributeNS(W, "w:leader", s.leader);
            tabs.appendChild(tb);
          }
          pPr.appendChild(tabs);
        }
      } catch { /* malformed: skip */ }
    }
    // schema order: w:spacing and w:ind come before w:jc; line + before/after share one element
    if (lineHeight > 0 || hasBefore || hasAfter) {
      const sp = ctx.doc.createElementNS(W, "w:spacing");
      if (hasBefore) sp.setAttributeNS(W, "w:before", String(Math.round(beforePx * 15)));
      if (hasAfter) sp.setAttributeNS(W, "w:after", String(Math.round(afterPx * 15)));
      if (lineHeight > 0) {
        sp.setAttributeNS(W, "w:line", String(Math.round(lineHeight * 240)));
        sp.setAttributeNS(W, "w:lineRule", "auto");
      }
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
    // The section break is the last child of w:pPr (schema order: after rPr). When the section's
    // geometry was edited / inserted in-editor, regenerate the sectPr from it (merging onto the
    // preserved original); otherwise re-emit the stashed original untouched.
    if (regenSect) {
      const sect = buildSectPr(ctx, secGeom!, sectXml, src.getAttribute("data-rdoc-secheaderkey"), src.getAttribute("data-rdoc-secfooterkey"));
      if (sect) pPr.appendChild(sect);
    } else if (sectXml) {
      const sect = importPassthrough(ctx, sectXml);
      if (sect) pPr.appendChild(sect);
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

const DOCX_BORDER_VAL: Record<string, string> = { solid: "single", dashed: "dashed", dotted: "dotted", double: "double" };

/** The spec of an edited cell side, read from data-rdoc-b<side> = "<w>px <style> <#color>". */
function sideSpec(td: HTMLTableCellElement, side: string): { w: number; style: string; color: string } | null {
  const v = td.getAttribute(`data-rdoc-b${side}`);
  if (!v) return null;
  const m = v.match(/^([\d.]+)px\s+(\w+)\s+(#[0-9a-fA-F]{3,8})/i);
  return m ? { w: parseFloat(m[1]!), style: m[2]!.toLowerCase(), color: m[3]! } : { w: 1, style: "solid", color: "#000000" };
}

/** Write per-cell borders (w:tcBorders) from the picker state. On sides carry their style,
    width (px -> eighths of a point) and colour; off sides become w:nil so they override the
    table grid. */
function applyCellBorders(ctx: DocxCtx, tcPr: Element, td: HTMLTableCellElement): void {
  if (!td.classList.contains("rdoc-bordered")) return;
  const old = tcPr.getElementsByTagName("w:tcBorders")[0];
  if (old) old.parentNode?.removeChild(old);
  const tb = ctx.doc.createElementNS(W, "w:tcBorders");
  for (const [side, tag] of [["t", "w:top"], ["l", "w:left"], ["b", "w:bottom"], ["r", "w:right"]] as const) {
    const b = ctx.doc.createElementNS(W, tag);
    const spec = sideSpec(td, side);
    if (spec) {
      b.setAttributeNS(W, "w:val", DOCX_BORDER_VAL[spec.style] ?? "single");
      b.setAttributeNS(W, "w:sz", String(Math.max(2, Math.round(spec.w * 6)))); // px -> 1/8 pt
      b.setAttributeNS(W, "w:space", "0");
      b.setAttributeNS(W, "w:color", spec.color.replace("#", ""));
    } else {
      b.setAttributeNS(W, "w:val", "nil");
    }
    tb.appendChild(b);
  }
  tcPr.appendChild(tb);
}

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

  // Table indent (dragging the table's outer-left edge) -> w:tblInd.
  const indentPx = parseFloat((tableEl as HTMLElement).style.marginLeft) || 0;
  const oldInd = tblPr.getElementsByTagName("w:tblInd")[0];
  if (oldInd) oldInd.parentNode?.removeChild(oldInd);
  if (indentPx > 0) {
    const ind = ctx.doc.createElementNS(W, "w:tblInd");
    ind.setAttributeNS(W, "w:w", String(Math.round(indentPx * 15)));
    ind.setAttributeNS(W, "w:type", "dxa");
    tblPr.appendChild(ind);
  }

  // Column widths from a <colgroup> (px) take precedence: it only exists once a column was
  // dragged to resize, so untouched tables keep their preserved grid.
  const cgEl = tableEl.querySelector(":scope > colgroup");
  const colWidths: number[] = cgEl ? Array.from(cgEl.children).map((c) => parseFloat((c as HTMLElement).style.width) || 0) : [];
  const useCols = colWidths.length === gridCols && colWidths.every((w) => w > 0);
  if (useCols) {
    // Pin the table width to the column sum and use fixed layout so Word honours the widths.
    const sum = colWidths.reduce((a, b) => a + b, 0);
    const oldW = tblPr.getElementsByTagName("w:tblW")[0];
    if (oldW) oldW.parentNode?.removeChild(oldW);
    const tblW = ctx.doc.createElementNS(W, "w:tblW");
    tblW.setAttributeNS(W, "w:w", String(Math.round(sum * 15)));
    tblW.setAttributeNS(W, "w:type", "dxa");
    tblPr.insertBefore(tblW, tblPr.firstChild);
    let lay = tblPr.getElementsByTagName("w:tblLayout")[0];
    if (!lay) {
      lay = ctx.doc.createElementNS(W, "w:tblLayout");
      tblPr.appendChild(lay);
    }
    lay.setAttributeNS(W, "w:type", "fixed");
  }

  let grid = importEl(tableEl.getAttribute("data-docx-tblgrid"), "w:tblGrid");
  if (useCols) {
    grid = ctx.doc.createElementNS(W, "w:tblGrid");
    for (const wpx of colWidths) {
      const gc = ctx.doc.createElementNS(W, "w:gridCol");
      gc.setAttributeNS(W, "w:w", String(Math.round(wpx * 15))); // px -> twips (96dpi)
      grid.appendChild(gc);
    }
  } else if (!grid || grid.getElementsByTagName("w:gridCol").length !== gridCols) {
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
      if (useCols) {
        const oldW = tcPr.getElementsByTagName("w:tcW")[0];
        if (oldW) oldW.parentNode?.removeChild(oldW);
        const tcW = ctx.doc.createElementNS(W, "w:tcW");
        const wsum = colWidths.slice(col, col + colspan).reduce((a, b) => a + b, 0);
        tcW.setAttributeNS(W, "w:w", String(Math.round(wsum * 15)));
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
      applyCellBorders(ctx, tcPr, td);
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
    const rh = parseFloat((tr as HTMLElement).style.height);
    if (rh > 0) {
      const trPr = ctx.doc.createElementNS(W, "w:trPr");
      const th = ctx.doc.createElementNS(W, "w:trHeight");
      th.setAttributeNS(W, "w:val", String(Math.round(rh * 15))); // px -> twips
      th.setAttributeNS(W, "w:hRule", "atLeast");
      trPr.appendChild(th);
      wtr.insertBefore(trPr, wtr.firstChild); // w:trPr must precede the cells
    }
    tbl.appendChild(wtr);
  }
  return tbl;
}

/** An <ol>'s requested first number (the `start` attribute), or 1. */
const listStartOf = (listEl: Element): number => {
  const n = parseInt(listEl.getAttribute("start") || "1", 10);
  return Number.isFinite(n) && n > 1 ? n : 1;
};

/** A numId for an ordered list that restarts at `start`. The first plain (start 1) list reuses
    the base ordered numId; every other ordered list gets its own w:num (with a startOverride
    when start > 1) so separate lists number independently instead of sharing one sequence. */
function newOrderedNumId(ctx: DocxCtx, start: number): string {
  const ids = ensureListNumbering(ctx);
  if (start <= 1 && !ctx.orderedBaseUsed) {
    ctx.orderedBaseUsed = true;
    return ids.ordered;
  }
  const numDoc = new DOMParser().parseFromString(strFromU8(ctx.files["word/numbering.xml"]!), "application/xml");
  const numIdOf = (n: Element) => n.getAttributeNS(W, "numId") ?? n.getAttribute("w:numId");
  const nums = Array.from(numDoc.getElementsByTagName("w:num"));
  const abs = nums.find((n) => numIdOf(n) === ids.ordered)?.getElementsByTagName("w:abstractNumId")[0]?.getAttribute("w:val") ?? "0";
  const next = Math.max(0, ...nums.map((n) => Number(numIdOf(n))).filter((n) => Number.isFinite(n))) + 1;
  const override = start > 1 ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride>` : "";
  const frag = new DOMParser().parseFromString(`<w:num xmlns:w="${W}" w:numId="${next}"><w:abstractNumId w:val="${abs}"/>${override}</w:num>`, "application/xml").documentElement!;
  numDoc.documentElement!.appendChild(numDoc.importNode(frag, true));
  ctx.files["word/numbering.xml"] = strToU8(new XMLSerializer().serializeToString(numDoc));
  return String(next);
}

/** Flatten a (possibly nested) <ul>/<ol> into list paragraphs, carrying the nesting depth as
    w:ilvl and a numId per ordered list (so separate lists restart); recurses into nested lists.
    A nested ordered list under an ordered ancestor reuses its numId (one sequence across levels). */
function appendList(ctx: DocxCtx, body: Element, listEl: Element, depth: number, orderedNumId: string | null): void {
  const ordered = listEl.tagName.toLowerCase() === "ol";
  const numId = ordered ? (orderedNumId ?? newOrderedNumId(ctx, listStartOf(listEl))) : ensureListNumbering(ctx).bullet;
  for (const li of Array.from(listEl.children)) {
    if (li.tagName.toLowerCase() !== "li") continue;
    body.appendChild(makeParagraph(ctx, li as HTMLElement, { listLevel: Math.min(8, depth), listNumId: numId }));
    for (const child of Array.from(li.children)) {
      const ct = child.tagName.toLowerCase();
      if (ct === "ul" || ct === "ol") appendList(ctx, body, child, depth + 1, ordered ? numId : null);
    }
  }
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
  if (el.classList.contains("docx-field-toc")) {
    appendTOC(ctx, body, el);
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
    appendList(ctx, body, el, 0, null);
    return;
  }
  const hm = /^h([1-6])$/.exec(tag);
  body.appendChild(makeParagraph(ctx, el, hm ? { heading: Number(hm[1]) } : {}));
}

/** Emit a table of contents as a complex TOC field: the begin/instr/separate lead the first
    entry paragraph, the cached entries follow, and the end closes the last. Word can update it. */
function appendTOC(ctx: DocxCtx, body: Element, el: HTMLElement): void {
  const fldChar = (type: string): Element => {
    const r = ctx.doc.createElementNS(W, "w:r");
    const fc = ctx.doc.createElementNS(W, "w:fldChar");
    fc.setAttributeNS(W, "w:fldCharType", type);
    r.appendChild(fc);
    return r;
  };
  const instr = (): Element => {
    const r = ctx.doc.createElementNS(W, "w:r");
    const it = ctx.doc.createElementNS(W, "w:instrText");
    it.setAttribute("xml:space", "preserve");
    it.textContent = ' TOC \\o "1-3" \\h \\z \\u ';
    r.appendChild(it);
    return r;
  };
  const rows = Array.from(el.querySelectorAll<HTMLElement>(".docx-field-toc-row"));
  if (!rows.length) {
    const p = ctx.doc.createElementNS(W, "w:p");
    p.append(fldChar("begin"), instr(), fldChar("separate"), fldChar("end"));
    body.appendChild(p);
    return;
  }
  rows.forEach((row, i) => {
    const p = ctx.doc.createElementNS(W, "w:p");
    const lvl = /toc-h([1-3])/.exec(row.className)?.[1] ?? "1";
    const pPr = ctx.doc.createElementNS(W, "w:pPr");
    const st = ctx.doc.createElementNS(W, "w:pStyle");
    st.setAttributeNS(W, "w:val", `TOC${lvl}`);
    pPr.appendChild(st);
    p.appendChild(pPr);
    if (i === 0) p.append(fldChar("begin"), instr(), fldChar("separate"));
    p.appendChild(makeRun(ctx, row.querySelector(".docx-field-toc-text")?.textContent ?? "", FMT0));
    const tab = ctx.doc.createElementNS(W, "w:r");
    tab.appendChild(ctx.doc.createElementNS(W, "w:tab"));
    p.appendChild(tab);
    const page = row.querySelector(".docx-field-toc-page")?.textContent ?? "";
    if (page) p.appendChild(makeRun(ctx, page, FMT0));
    if (i === rows.length - 1) p.appendChild(fldChar("end"));
    body.appendChild(p);
  });
}

// A 9-level abstractNum: bullets (•/◦/▪) or numbers (decimal/letter/roman), 0.5in indent steps.
function abstractNumXml(absId: number, bullet: boolean): string {
  let lvls = "";
  for (let l = 0; l < 9; l++) {
    const left = 720 * (l + 1);
    const fmt = bullet ? "bullet" : ["decimal", "lowerLetter", "lowerRoman"][l % 3];
    const text = bullet ? ["•", "◦", "▪"][l % 3] : `%${l + 1}.`;
    lvls += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
  }
  return `<w:abstractNum w:abstractNumId="${absId}">${lvls}</w:abstractNum>`;
}

/** Register word/numbering.xml in [Content_Types].xml and document.xml.rels (when newly created). */
function registerNumberingPart(files: Record<string, Uint8Array>): void {
  const ct = files["[Content_Types].xml"];
  if (ct) {
    const d = new DOMParser().parseFromString(strFromU8(ct), "application/xml");
    if (!Array.from(d.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === "/word/numbering.xml")) {
      const ov = d.createElementNS(d.documentElement!.namespaceURI, "Override");
      ov.setAttribute("PartName", "/word/numbering.xml");
      ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
      d.documentElement!.appendChild(ov);
      files["[Content_Types].xml"] = strToU8(new XMLSerializer().serializeToString(d));
    }
  }
  const key = "word/_rels/document.xml.rels";
  const d = files[key]
    ? new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml")
    : new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  if (!Array.from(d.getElementsByTagName("Relationship")).some((r) => (r.getAttribute("Type") ?? "").endsWith("/numbering"))) {
    let max = 0;
    for (const r of Array.from(d.getElementsByTagName("Relationship"))) {
      const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
      if (m) max = Math.max(max, Number(m[1]));
    }
    const rel = d.createElementNS(PKG, "Relationship");
    rel.setAttribute("Id", `rId${max + 1}`);
    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
    rel.setAttribute("Target", "numbering.xml");
    d.documentElement!.appendChild(rel);
    files[key] = strToU8(new XMLSerializer().serializeToString(d));
  }
}

/** Resolve a bullet and an ordered numId, reusing existing ones or adding them to (or creating)
    word/numbering.xml. Cached on the context so it runs once per save. */
function ensureListNumbering(ctx: DocxCtx): { bullet: string; ordered: string } {
  if (ctx.listIds) return ctx.listIds;
  const files = ctx.files;
  const existing = files["word/numbering.xml"];
  let numDoc = existing ? new DOMParser().parseFromString(strFromU8(existing), "application/xml") : null;
  let bullet: string | undefined;
  let ordered: string | undefined;
  if (numDoc) {
    const absFmt = new Map<string, string>();
    for (const an of Array.from(numDoc.getElementsByTagName("w:abstractNum"))) {
      const id = an.getAttributeNS(W, "abstractNumId") ?? an.getAttribute("w:abstractNumId");
      if (id) absFmt.set(id, an.getElementsByTagName("w:numFmt")[0]?.getAttribute("w:val") ?? "");
    }
    for (const num of Array.from(numDoc.getElementsByTagName("w:num"))) {
      const id = num.getAttributeNS(W, "numId") ?? num.getAttribute("w:numId");
      const fmt = absFmt.get(num.getElementsByTagName("w:abstractNumId")[0]?.getAttribute("w:val") ?? "");
      if (!id) continue;
      if (fmt === "bullet") bullet ??= id;
      else if (fmt && fmt !== "none") ordered ??= id;
    }
  }
  if (bullet && ordered) return (ctx.listIds = { bullet, ordered });
  // Create what is missing (and the part itself if absent).
  const created = !numDoc;
  if (!numDoc) numDoc = new DOMParser().parseFromString(`<w:numbering xmlns:w="${W}"></w:numbering>`, "application/xml");
  const root = numDoc.documentElement!;
  const ids = (tag: string, attr: string) => Array.from(numDoc!.getElementsByTagName(tag)).map((e) => Number(e.getAttribute(attr))).filter((n) => Number.isFinite(n));
  let nextAbs = Math.max(-1, ...ids("w:abstractNum", "w:abstractNumId")) + 1;
  let nextNum = Math.max(0, ...ids("w:num", "w:numId")) + 1;
  const importFrag = (xml: string): Element => numDoc!.importNode(new DOMParser().parseFromString(`<r xmlns:w="${W}">${xml}</r>`, "application/xml").documentElement!.firstElementChild!, true) as Element;
  const addList = (isBullet: boolean): string => {
    const absId = nextAbs++;
    const numId = String(nextNum++);
    root.insertBefore(importFrag(abstractNumXml(absId, isBullet)), numDoc!.getElementsByTagName("w:num")[0] ?? null);
    root.appendChild(importFrag(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`));
    return numId;
  };
  if (!bullet) bullet = addList(true);
  if (!ordered) ordered = addList(false);
  files["word/numbering.xml"] = strToU8(new XMLSerializer().serializeToString(numDoc));
  if (created) registerNumberingPart(files);
  return (ctx.listIds = { bullet, ordered });
}

const relsPathFor = (partPath: string): string => {
  const slash = partPath.lastIndexOf("/");
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
};

/** Rewrite the normal notes in word/footnotes.xml / endnotes.xml from the edited bodies, keeping
    the separator / continuation-separator notes (id <= 0 or a w:type). Updates an existing part
    only (creating one from scratch is for the insert-authoring step). */
const CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

/** Add a relationship from document.xml to a target part (idempotent by target). */
function addDocRel(files: Record<string, Uint8Array>, type: string, target: string): void {
  const key = "word/_rels/document.xml.rels";
  const doc = files[key]
    ? new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml")
    : new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  if (Array.from(doc.getElementsByTagName("Relationship")).some((r) => r.getAttribute("Target") === target)) return;
  let max = 0;
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
    const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
    if (m) max = Math.max(max, Number(m[1]));
  }
  const rel = doc.createElementNS(PKG, "Relationship");
  rel.setAttribute("Id", `rId${max + 1}`);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement!.appendChild(rel);
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
}

function rebuildNotes(files: Record<string, Uint8Array>, part: "footnotes" | "endnotes", kind: "footnote" | "endnote", notes: Note[]): void {
  const key = `word/${part}.xml`;
  if (!files[key]) {
    if (!notes.some((n) => n.kind === kind)) return; // nothing of this kind to write
    // Mint the part with the standard separator notes, register it, and relate it from the body.
    const root = part === "footnotes" ? "w:footnotes" : "w:endnotes";
    const tag = part === "footnotes" ? "w:footnote" : "w:endnote";
    const seps = `<${tag} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></${tag}><${tag} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></${tag}>`;
    files[key] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} xmlns:w="${W}">${seps}</${root}>`);
    ensureOverride(files, `/${key}`, part === "footnotes" ? CT_FOOTNOTES : CT_ENDNOTES);
    addDocRel(files, part === "footnotes" ? REL_FOOTNOTES : REL_ENDNOTES, `${part}.xml`);
  }
  const doc = new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml");
  const root = doc.documentElement;
  if (!root) return;
  const tag = part === "footnotes" ? "w:footnote" : "w:endnote";
  const kept = Array.from(root.getElementsByTagName(tag)).filter((n) => n.getAttribute("w:type") || Number(n.getAttribute("w:id")) <= 0);
  while (root.firstChild) root.removeChild(root.firstChild);
  for (const k of kept) root.appendChild(k);
  const relsKey = `word/_rels/${part}.xml.rels`;
  const rels = files[relsKey] ? new DOMParser().parseFromString(strFromU8(files[relsKey]!), "application/xml") : null;
  for (const n of notes.filter((x) => x.kind === kind)) {
    const note = doc.createElementNS(W, tag);
    note.setAttributeNS(W, "w:id", n.id);
    const ctx: DocxCtx = { doc, rels, relsAdded: false, nextRid: 1, nextRevId: 1, listIds: null, files };
    const htmlDoc = new DOMParser().parseFromString(n.html || "<p><br></p>", "text/html");
    for (const node of Array.from(htmlDoc.body.childNodes)) appendBlock(ctx, note, node);
    if (!note.firstChild) note.appendChild(doc.createElementNS(W, "w:p"));
    root.appendChild(note);
  }
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Rebuild one part (document.xml or a header/footer) in-place from its edited HTML. For the body,
    bandHtml lets buildSectPr mint per-section header/footer parts. */
function rebuildPart(files: Record<string, Uint8Array>, partPath: string, html: string, bandHtml?: Map<string, string>): void {
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
    listIds: null,
    files,
    sectionBandHtml: bandHtml,
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

/** Mint a new word/<kind>N.xml from edited HTML, register its relationship + content-type, and
    return the document relationship id (no section reference is added; callers wire that). */
function mintPartFile(files: Record<string, Uint8Array>, kind: "header" | "footer", html: string): string {
  const isHeader = kind === "header";
  let n = 1;
  while (files[`word/${kind}${n}.xml`]) n++;
  const partPath = `word/${kind}${n}.xml`;
  const decls = Object.entries(NS_DECLS).map(([k, v]) => `${k}="${v}"`).join(" ");
  const partDoc = new DOMParser().parseFromString(`<${isHeader ? "w:hdr" : "w:ftr"} ${decls}></${isHeader ? "w:hdr" : "w:ftr"}>`, "application/xml");
  const container = partDoc.documentElement!;
  const relsDoc = new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  const ctx: DocxCtx = { doc: partDoc, rels: relsDoc, relsAdded: false, nextRid: 1, nextRevId: 1, listIds: null, files };
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  for (const node of Array.from(htmlDoc.body.childNodes)) appendBlock(ctx, container, node);
  if (!container.firstChild) container.appendChild(partDoc.createElementNS(W, "w:p"));
  files[partPath] = strToU8(new XMLSerializer().serializeToString(partDoc));
  if (ctx.relsAdded) files[`word/_rels/${kind}${n}.xml.rels`] = strToU8(new XMLSerializer().serializeToString(relsDoc));
  ensureOverride(files, `/${partPath}`, isHeader ? CT_HEADER : CT_FOOTER);

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
  return rid;
}

/** Create a new default header/footer part and reference it from the body section. */
function createHeaderFooterPart(files: Record<string, Uint8Array>, kind: "header" | "footer", html: string): void {
  if (!files["word/document.xml"]) return;
  const rid = mintPartFile(files, kind, html);
  const bodyDoc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]!), "application/xml");
  const sectPr = Array.from(bodyDoc.getElementsByTagName("w:sectPr")).pop();
  if (sectPr) {
    const ref = bodyDoc.createElementNS(W, kind === "header" ? "w:headerReference" : "w:footerReference");
    ref.setAttributeNS(W, "w:type", "default");
    ref.setAttributeNS(R, "r:id", rid);
    sectPr.insertBefore(ref, sectPr.firstChild);
    files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(bodyDoc));
  }
}

/** Apply the document-level header/footer variants + their on/off flags to the trailing section
    and settings.xml. `variants` holds the edited HTML for any UI-created (sentinel) variant part. */
function applyHFVariants(files: Record<string, Uint8Array>, geometry: PageGeometry, variants: Map<string, string>): void {
  if (!files["word/document.xml"]) return;
  const flagOf = (v: "first" | "even") => (v === "first" ? !!geometry.titlePage : !!geometry.evenOdd);
  // Mint any sentinel variant parts first (each adds a document relationship); collect their rids.
  const minted = new Map<string, string>(); // `${role}:${variant}` -> rId
  for (const variant of ["first", "even"] as const) {
    if (!flagOf(variant)) continue;
    for (const role of ["header", "footer"] as const) {
      const html = variants.get(`${role}:${variant}`);
      if (html != null) minted.set(`${role}:${variant}`, mintPartFile(files, role, html));
    }
  }
  const doc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]!), "application/xml");
  const sectPr = Array.from(doc.getElementsByTagName("w:sectPr")).pop();
  if (!sectPr) return;
  // Different first page: the w:titlePg flag in the trailing sectPr.
  const tp = sectPr.getElementsByTagName("w:titlePg")[0];
  if (geometry.titlePage && !tp) sectPr.appendChild(doc.createElementNS(W, "w:titlePg"));
  else if (!geometry.titlePage && tp) tp.remove();
  for (const variant of ["first", "even"] as const) {
    for (const role of ["header", "footer"] as const) {
      const refTag = role === "header" ? "w:headerReference" : "w:footerReference";
      const existing = Array.from(sectPr.getElementsByTagName(refTag)).filter((r) => r.getAttribute("w:type") === variant);
      const rid = minted.get(`${role}:${variant}`);
      if (rid) { for (const e of existing) e.remove(); const ref = doc.createElementNS(W, refTag); ref.setAttributeNS(W, "w:type", variant); ref.setAttributeNS(R, "r:id", rid); sectPr.insertBefore(ref, sectPr.firstChild); }
      else if (!flagOf(variant)) for (const e of existing) e.remove(); // flag off: drop the typed reference (part orphaned, harmless)
      // flag on without a minted part: an existing typed reference rides along untouched.
    }
  }
  files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
  setEvenAndOddHeaders(files, !!geometry.evenOdd);
}

/** Set or clear the document-level w:evenAndOddHeaders flag in word/settings.xml (minting the part,
    its content-type and the document relationship when it is turned on and the part is absent). */
function setEvenAndOddHeaders(files: Record<string, Uint8Array>, on: boolean): void {
  const key = "word/settings.xml";
  if (!files[key]) {
    if (!on) return;
    files[key] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`);
    ensureOverride(files, "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
    const relsPath = "word/_rels/document.xml.rels";
    const dRels = files[relsPath]
      ? new DOMParser().parseFromString(strFromU8(files[relsPath]!), "application/xml")
      : new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
    if (!Array.from(dRels.getElementsByTagName("Relationship")).some((r) => r.getAttribute("Target") === "settings.xml")) {
      let maxR = 0;
      for (const r of Array.from(dRels.getElementsByTagName("Relationship"))) { const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? ""); if (m) maxR = Math.max(maxR, Number(m[1])); }
      const rel = dRels.createElementNS(PKG, "Relationship");
      rel.setAttribute("Id", `rId${maxR + 1}`);
      rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
      rel.setAttribute("Target", "settings.xml");
      dRels.documentElement!.appendChild(rel);
      files[relsPath] = strToU8(new XMLSerializer().serializeToString(dRels));
    }
    return;
  }
  const doc = new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml");
  const root = doc.documentElement;
  if (!root) return;
  const cur = root.getElementsByTagName("w:evenAndOddHeaders")[0];
  if (on && !cur) root.insertBefore(doc.createElementNS(W, "w:evenAndOddHeaders"), root.firstChild);
  else if (!on && cur) cur.remove();
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
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
/** Set a w:sectPr's page size, orientation, margins and columns (px -> twips), in place. Used for
    both the body-level section (Page setup) and an in-paragraph section break (per-section setup). */
function setSectPrGeom(doc: Document, sectPr: Element, g: { w: number; h: number; mt: number; mr: number; mb: number; ml: number; cols?: number; colGap?: number; vertical?: boolean; rtl?: boolean }): void {
  const tw = (px: number): string => String(Math.round(px * 15));
  const child = (tag: string): Element => {
    let e = sectPr.getElementsByTagName(tag)[0];
    if (!e) { e = doc.createElementNS(W, tag); sectPr.appendChild(e); }
    return e;
  };
  const dropTag = (tag: string) => { for (const e of Array.from(sectPr.getElementsByTagName(tag))) e.parentNode!.removeChild(e); };
  // Page size + orientation (Word marks landscape explicitly when width > height).
  const pgSz = child("w:pgSz");
  pgSz.setAttributeNS(W, "w:w", tw(g.w));
  pgSz.setAttributeNS(W, "w:h", tw(g.h));
  if (g.w > g.h) pgSz.setAttributeNS(W, "w:orient", "landscape");
  else { pgSz.removeAttributeNS(W, "orient"); pgSz.removeAttribute("w:orient"); }
  // Margins.
  const pgMar = child("w:pgMar");
  pgMar.setAttributeNS(W, "w:top", tw(g.mt));
  pgMar.setAttributeNS(W, "w:right", tw(g.mr));
  pgMar.setAttributeNS(W, "w:bottom", tw(g.mb));
  pgMar.setAttributeNS(W, "w:left", tw(g.ml));
  // Columns: write w:cols @num + equal-width @space; drop to single column when columns <= 1.
  const n = g.cols && g.cols > 1 ? g.cols : 1;
  const cols = child("w:cols");
  cols.setAttributeNS(W, "w:num", String(n));
  cols.setAttributeNS(W, "w:space", tw(g.colGap ?? 36));
  cols.setAttributeNS(W, "w:equalWidth", "1");
  // Writing direction: vertical tategaki (w:textDirection tbRl) / horizontal RTL (w:bidi).
  dropTag("w:textDirection");
  dropTag("w:bidi");
  if (g.vertical) child("w:textDirection").setAttributeNS(W, "w:val", "tbRl");
  else if (g.rtl) child("w:bidi");
}

/** Mint a fresh header/footer part (word/<kind>N.xml) from HTML + a content-type override + a
    relationship in document.xml.rels (via ctx), returning the new relationship id. */
function mintHFPart(ctx: DocxCtx, kind: "header" | "footer", html: string): string | null {
  const isHeader = kind === "header";
  let n = 1;
  while (ctx.files[`word/${kind}${n}.xml`]) n++;
  const partPath = `word/${kind}${n}.xml`;
  const decls = Object.entries(NS_DECLS).map(([k, v]) => `${k}="${v}"`).join(" ");
  const partDoc = new DOMParser().parseFromString(`<${isHeader ? "w:hdr" : "w:ftr"} ${decls}></${isHeader ? "w:hdr" : "w:ftr"}>`, "application/xml");
  const container = partDoc.documentElement!;
  const relsDoc = new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  const pctx: DocxCtx = { doc: partDoc, rels: relsDoc, relsAdded: false, nextRid: 1, nextRevId: 1, listIds: null, files: ctx.files };
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  for (const node of Array.from(htmlDoc.body.childNodes)) appendBlock(pctx, container, node);
  if (!container.firstChild) container.appendChild(partDoc.createElementNS(W, "w:p"));
  ctx.files[partPath] = strToU8(new XMLSerializer().serializeToString(partDoc));
  if (pctx.relsAdded) ctx.files[`word/_rels/${kind}${n}.xml.rels`] = strToU8(new XMLSerializer().serializeToString(relsDoc));
  ensureOverride(ctx.files, `/${partPath}`, isHeader ? CT_HEADER : CT_FOOTER);
  if (!ctx.rels) return null;
  const rid = `rId${ctx.nextRid++}`;
  const rel = ctx.rels.createElementNS(PKG, "Relationship");
  rel.setAttribute("Id", rid);
  rel.setAttribute("Type", isHeader ? REL_HEADER : REL_FOOTER);
  rel.setAttribute("Target", `${kind}${n}.xml`);
  ctx.rels.documentElement!.appendChild(rel);
  ctx.relsAdded = true;
  return rid;
}

/** Reconcile a section's default header/footer reference with its link state: a "new..." key mints
    a part and points at it; an existing key leaves the (stashed) ref in place; no key removes the
    default ref so the section links to the previous one. */
function applyHFRef(ctx: DocxCtx, sectPr: Element, role: "header" | "footer", key: string | null): void {
  const refTag = role === "header" ? "w:headerReference" : "w:footerReference";
  const removeDefault = () => {
    for (const r of Array.from(sectPr.getElementsByTagName(refTag))) {
      const ty = r.getAttributeNS(W, "type") ?? r.getAttribute("w:type") ?? "default";
      if (ty === "default") r.parentNode!.removeChild(r);
    }
  };
  if (!key) { removeDefault(); return; }
  if (!key.startsWith("new")) return; // existing part: its ref rides along on the stashed sectPr
  const rid = mintHFPart(ctx, role, ctx.sectionBandHtml?.get(key) ?? "<p><br></p>");
  if (!rid) return;
  removeDefault();
  const ref = ctx.doc.createElementNS(W, refTag);
  ref.setAttributeNS(W, "w:type", "default");
  ref.setAttributeNS(R, "r:id", rid);
  sectPr.insertBefore(ref, sectPr.firstChild);
}

/** Build a w:sectPr for an in-paragraph section break from the edited geometry JSON
    (data-rdoc-secbreak), merging onto the preserved original sectPr when there is one (so header
    refs, page borders, etc. survive) or a fresh next-page break otherwise. headerKey/footerKey
    carry the section's header/footer link state. */
function buildSectPr(ctx: DocxCtx, geomJson: string, stashedXml: string | null, headerKey: string | null, footerKey: string | null): Element | null {
  let g: { w: number; h: number; mt: number; mr: number; mb: number; ml: number; cols?: number; colGap?: number };
  try { g = JSON.parse(geomJson); } catch { return null; }
  let sectPr = stashedXml ? importPassthrough(ctx, stashedXml) : null;
  if (!sectPr || sectPr.tagName !== "w:sectPr") {
    sectPr = ctx.doc.createElementNS(W, "w:sectPr");
    const type = ctx.doc.createElementNS(W, "w:type");
    type.setAttributeNS(W, "w:val", "nextPage");
    sectPr.appendChild(type);
  }
  setSectPrGeom(ctx.doc, sectPr, g);
  applyHFRef(ctx, sectPr, "header", headerKey);
  applyHFRef(ctx, sectPr, "footer", footerKey);
  return sectPr;
}

/** Update the body section's page geometry (size, orientation, margins, columns) from the edited
    geometry. px -> twips. Only the trailing (body-level) sectPr is touched; in-paragraph section
    breaks are regenerated from their own data-rdoc-secbreak in makeParagraph. */
function applyPageMargins(files: Record<string, Uint8Array>, geometry: PageGeometry): void {
  const raw = files["word/document.xml"];
  if (!raw) return;
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const sectPr = Array.from(doc.getElementsByTagName("w:sectPr")).pop();
  if (!sectPr) return;
  setSectPrGeom(doc, sectPr, { w: geometry.widthPx, h: geometry.heightPx, mt: geometry.margin.top, mr: geometry.margin.right, mb: geometry.margin.bottom, ml: geometry.margin.left, cols: geometry.columns, colGap: geometry.columnGapPx, vertical: geometry.vertical, rtl: geometry.rtl });
  files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Add user-authored styles to word/styles.xml (creating + registering the part if absent),
    translating the CSS-like props back to a w:style's pPr/rPr. */
function addNewStyles(files: Record<string, Uint8Array>, styles: NewStyle[]): void {
  const key = "word/styles.xml";
  const created = !files[key];
  const doc = files[key]
    ? new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml")
    : new DOMParser().parseFromString(`<w:styles xmlns:w="${W}"></w:styles>`, "application/xml");
  const root = doc.documentElement!;
  const el = (tag: string): Element => doc.createElementNS(W, tag);
  const valEl = (parent: Element, tag: string, val: string) => {
    const e = el(tag);
    e.setAttributeNS(W, "w:val", val);
    parent.appendChild(e);
  };
  const shd = (parent: Element, fill: string) => {
    const e = el("w:shd");
    e.setAttributeNS(W, "w:val", "clear");
    e.setAttributeNS(W, "w:color", "auto");
    e.setAttributeNS(W, "w:fill", fill);
    parent.appendChild(e);
  };
  // Find an existing w:style by id so editing replaces its formatting (keeping name/basedOn).
  const existingById = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("w:style"))) {
    const id = st.getAttributeNS(W, "styleId") ?? st.getAttribute("w:styleId");
    if (id) existingById.set(id, st);
  }
  // The style properties the dialog owns; on edit only these are re-derived, so the long tail
  // (tabs, keepNext, outline level, small caps, ...) the dialog does not model is preserved.
  const PPR_MODELED = new Set(["w:spacing", "w:ind", "w:jc", "w:shd", "w:pBdr"]);
  const SIDES = ["top", "right", "bottom", "left"] as const;
  const RPR_MODELED = new Set(["w:rFonts", "w:b", "w:i", "w:u", "w:strike", "w:color", "w:sz", "w:shd"]);
  const directChild = (parent: Element, tag: string): Element | undefined => Array.from(parent.children).find((c) => c.tagName === tag);
  for (const s of styles) {
    const c = s.css;
    const prev = existingById.get(s.id);
    const st = prev ?? el("w:style");
    if (!prev) {
      st.setAttributeNS(W, "w:type", s.kind === "paragraph" ? "paragraph" : "character");
      st.setAttributeNS(W, "w:styleId", s.id);
      valEl(st, "w:name", s.name);
    }
    if (s.kind === "paragraph") {
      const hasBorder = SIDES.some((s) => parseCssBorder(c[`border-${s}`]));
      const hasPara = !!(c["margin-top"] || c["margin-bottom"] || c["line-height"] || c["margin-left"] || JC_BY_ALIGN[c["text-align"] ?? ""] || /^[0-9a-f]{6}$/i.test((c["background-color"] ?? "").replace(/^#/, "")) || hasBorder);
      let pPr = directChild(st, "w:pPr");
      if (pPr || hasPara) {
        if (!pPr) { pPr = el("w:pPr"); st.insertBefore(pPr, st.firstChild); }
        for (const ch of Array.from(pPr.children)) if (PPR_MODELED.has(ch.tagName)) pPr.removeChild(ch); // re-derive only the modeled props
        if (c["margin-top"] || c["margin-bottom"] || c["line-height"]) {
          const sp = el("w:spacing");
          if (c["margin-top"]) sp.setAttributeNS(W, "w:before", String(Math.round(parseFloat(c["margin-top"]) * 15)));
          if (c["margin-bottom"]) sp.setAttributeNS(W, "w:after", String(Math.round(parseFloat(c["margin-bottom"]) * 15)));
          if (c["line-height"]) {
            sp.setAttributeNS(W, "w:line", String(Math.round(parseFloat(c["line-height"]) * 240)));
            sp.setAttributeNS(W, "w:lineRule", "auto");
          }
          pPr.appendChild(sp);
        }
        if (c["margin-left"]) {
          const ind = el("w:ind");
          ind.setAttributeNS(W, "w:left", String(Math.round(parseFloat(c["margin-left"]) * 15)));
          pPr.appendChild(ind);
        }
        const jc = JC_BY_ALIGN[c["text-align"] ?? ""];
        if (jc) valEl(pPr, "w:jc", jc);
        // paragraph borders (w:pBdr) come before w:shd in the pPr schema
        let pBdr: Element | undefined;
        for (const side of SIDES) {
          const b = parseCssBorder(c[`border-${side}`]);
          if (!b) continue;
          if (!pBdr) { pBdr = el("w:pBdr"); pPr.appendChild(pBdr); }
          const e = el(`w:${side}`);
          e.setAttributeNS(W, "w:val", DOCX_BORDER_VAL[b.style] ?? "single");
          e.setAttributeNS(W, "w:sz", String(Math.max(2, Math.round(b.px * 6))));
          e.setAttributeNS(W, "w:space", "1");
          e.setAttributeNS(W, "w:color", b.hex);
          pBdr.appendChild(e);
        }
        const pbg = (c["background-color"] ?? "").replace(/^#/, "");
        if (/^[0-9a-f]{6}$/i.test(pbg)) shd(pPr, pbg); // paragraph shading
        if (!pPr.childNodes.length) st.removeChild(pPr); // edited down to nothing
      }
    }
    {
      let rPr = directChild(st, "w:rPr");
      if (!rPr) { rPr = el("w:rPr"); st.appendChild(rPr); }
      for (const ch of Array.from(rPr.children)) if (RPR_MODELED.has(ch.tagName)) rPr.removeChild(ch); // re-derive only the modeled props
      const font = (c["font-family"] ?? "").replace(/['"]/g, "").split(",")[0]?.trim();
      if (font) {
        const rf = el("w:rFonts");
        rf.setAttributeNS(W, "w:ascii", font);
        rf.setAttributeNS(W, "w:hAnsi", font);
        rPr.appendChild(rf);
      }
      if (/bold|[6-9]00/.test(c["font-weight"] ?? "")) rPr.appendChild(el("w:b"));
      if (c["font-style"] === "italic") rPr.appendChild(el("w:i"));
      if (/underline/.test(c["text-decoration"] ?? "")) valEl(rPr, "w:u", "single");
      if (/line-through/.test(c["text-decoration"] ?? "")) rPr.appendChild(el("w:strike"));
      const color = (c["color"] ?? "").replace(/^#/, "");
      if (/^[0-9a-f]{6}$/i.test(color)) valEl(rPr, "w:color", color);
      if (c["font-size"]) valEl(rPr, "w:sz", String(Math.round(parseFloat(c["font-size"]) * 2)));
      if (s.kind === "character") {
        const cbg = (c["background-color"] ?? "").replace(/^#/, "");
        if (/^[0-9a-f]{6}$/i.test(cbg)) shd(rPr, cbg); // run shading (text background)
      }
      if (!rPr.childNodes.length) st.removeChild(rPr);
    }
    if (!prev) root.appendChild(st);
  }
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
  if (created) registerStylesPart(files);
}

/** Declare word/styles.xml in [Content_Types].xml and the document rels (when newly created). */
function registerStylesPart(files: Record<string, Uint8Array>): void {
  const ct = files["[Content_Types].xml"];
  if (ct) {
    const d = new DOMParser().parseFromString(strFromU8(ct), "application/xml");
    if (!Array.from(d.getElementsByTagName("Override")).some((o) => o.getAttribute("PartName") === "/word/styles.xml")) {
      const ov = d.createElementNS(d.documentElement!.namespaceURI, "Override");
      ov.setAttribute("PartName", "/word/styles.xml");
      ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
      d.documentElement!.appendChild(ov);
      files["[Content_Types].xml"] = strToU8(new XMLSerializer().serializeToString(d));
    }
  }
  const key = "word/_rels/document.xml.rels";
  const d = files[key]
    ? new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml")
    : new DOMParser().parseFromString(`<Relationships xmlns="${PKG}"></Relationships>`, "application/xml");
  if (!Array.from(d.getElementsByTagName("Relationship")).some((r) => (r.getAttribute("Type") ?? "").endsWith("/styles"))) {
    let max = 0;
    for (const r of Array.from(d.getElementsByTagName("Relationship"))) {
      const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
      if (m) max = Math.max(max, Number(m[1]));
    }
    const rel = d.createElementNS(PKG, "Relationship");
    rel.setAttribute("Id", `rId${max + 1}`);
    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    rel.setAttribute("Target", "styles.xml");
    d.documentElement!.appendChild(rel);
    files[key] = strToU8(new XMLSerializer().serializeToString(d));
  }
}

export function htmlToDocx(
  html: string,
  original: Uint8Array,
  parts?: { path: string; html: string }[],
  opts?: { reactions?: ReactionEdit[]; replies?: ReplyEdit[]; done?: Map<string, boolean>; deletedComments?: string[]; pageGeometry?: PageGeometry; newStyles?: NewStyle[]; notes?: Note[] },
): Uint8Array {
  const files = unzipSync(original);
  // Per-section header/footer HTML by key, so buildSectPr can mint a new part for an unlinked
  // section while rebuilding the body.
  const bandHtml = new Map<string, string>();
  for (const p of parts ?? []) bandHtml.set(p.path, p.html);
  rebuildPart(files, "word/document.xml", html, bandHtml);
  if (opts?.newStyles?.length) addNewStyles(files, opts.newStyles);
  // First/even variant parts created in-editor arrive as "header:first" / "footer:even" sentinels;
  // hold them out of the generic loop and mint them with their typed reference in applyHFVariants.
  const hfVariants = new Map<string, string>();
  for (const p of parts ?? []) {
    if (/^new(header|footer):/.test(p.path)) continue; // an unlinked section's part: minted in buildSectPr
    if (/^(header|footer):(first|even)$/.test(p.path)) { hfVariants.set(p.path, p.html); continue; }
    // "header" / "footer" are sentinels for a band created in-editor (no existing part).
    if (p.path === "header" || p.path === "footer") createHeaderFooterPart(files, p.path, p.html);
    else rebuildPart(files, p.path, p.html);
  }
  if (opts?.notes) { rebuildNotes(files, "footnotes", "footnote", opts.notes); rebuildNotes(files, "endnotes", "endnote", opts.notes); }
  applyNewComments(files, html);
  applyReactions(files, opts?.reactions ?? []);
  applyReplies(files, opts?.replies ?? []);
  applyDone(files, opts?.done ?? new Map());
  applyDeletedComments(files, opts?.deletedComments ?? []);
  if (opts?.pageGeometry) { applyPageMargins(files, opts.pageGeometry); applyHFVariants(files, opts.pageGeometry, hfVariants); }
  return zipSync(files);
}

