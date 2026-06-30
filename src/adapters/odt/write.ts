// odt WRITE: rebuild an .odt archive from edited HTML, preserving every untouched part.
// Pure HTML -> XML (body, styles, header/footer, comments, tracked changes, page margins);
// the read half lives in ./read.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { firstFontFamily, fontSizeToHalfPt, toHex6, imageLayoutFromEl, blockBorders, parseCssBorder } from "../../core/util";
import type { BlockBorderSide } from "../../core/util";
import type { NewStyle, Note, PageBorder, PageGeometry } from "../../core/types";
import { NS, fmtKey, FMT0, ODF_ALIGN, importPassthrough, IMG_MIME } from "./shared";
import type { Fmt } from "./shared";
import { applyFrameLayout } from "./image-layout";

const pxToCm = (px: number): string => `${Math.round((px / (96 / 2.54)) * 1000) / 1000}cm`;

/** Parse the data-rdoc-tabstops JSON into a tab-stop list, or undefined. */
const parseTabStops = (s: string | null): { pos: number; val?: string; leader?: string }[] | undefined => {
  if (!s) return undefined;
  try {
    const a = JSON.parse(s);
    return Array.isArray(a) && a.length ? a : undefined;
  } catch {
    return undefined;
  }
};

/** Ensure <office:automatic-styles> exists, returning it. The new node is placed before
   `beforeTag` (office:body in content.xml, office:master-styles in styles.xml). */
function ensureAutoStyles(doc: Document, beforeTag = "office:body"): Element {
  let auto = doc.getElementsByTagName("office:automatic-styles")[0];
  if (auto) return auto;
  auto = doc.createElementNS(NS.office, "office:automatic-styles");
  const before = doc.getElementsByTagName(beforeTag)[0];
  doc.documentElement.insertBefore(auto, before ?? null);
  return auto;
}

/** Create (once) a text style for a run-formatting combo and return its name. */
function styleFor(doc: Document, auto: Element, created: Map<string, string>, f: Fmt): string | null {
  const key = fmtKey(f);
  if (!key) return null;
  const parent = f.cStyle; // direct formatting layered over a named character style
  const ckey = parent ? `${key}|${parent}` : key;
  const existing = created.get(ckey);
  if (existing) return existing;
  const name = `OT_t${created.size}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "text");
  if (parent) st.setAttributeNS(NS.style, "style:parent-style-name", parent);
  const tp = doc.createElementNS(NS.style, "style:text-properties");
  if (f.b) tp.setAttributeNS(NS.fo, "fo:font-weight", "bold");
  if (f.i) tp.setAttributeNS(NS.fo, "fo:font-style", "italic");
  if (f.u) {
    tp.setAttributeNS(NS.style, "style:text-underline-style", "solid");
    tp.setAttributeNS(NS.style, "style:text-underline-width", "auto");
    tp.setAttributeNS(NS.style, "style:text-underline-color", "font-color");
  }
  if (f.strike) {
    tp.setAttributeNS(NS.style, "style:text-line-through-style", "solid");
    tp.setAttributeNS(NS.style, "style:text-line-through-type", "single");
  }
  if (f.vertAlign) tp.setAttributeNS(NS.style, "style:text-position", f.vertAlign === "super" ? "super 58%" : "sub 58%");
  if (f.color) tp.setAttributeNS(NS.fo, "fo:color", `#${f.color}`);
  if (f.bg) tp.setAttributeNS(NS.fo, "fo:background-color", `#${f.bg}`);
  if (f.font) {
    tp.setAttributeNS(NS.fo, "fo:font-family", f.font);
    tp.setAttributeNS(NS.style, "style:font-name", f.font);
  }
  if (f.sizePt) tp.setAttributeNS(NS.fo, "fo:font-size", `${f.sizePt}pt`);
  st.appendChild(tp);
  auto.appendChild(st);
  created.set(ckey, name);
  return name;
}

// Build a style:tab-stops element from a tab-stops array (px positions), or null if empty.
function buildOdtTabStops(doc: Document, stops: { pos: number; val?: string; leader?: string }[] | undefined): Element | null {
  if (!stops || !stops.length) return null;
  const ts = doc.createElementNS(NS.style, "style:tab-stops");
  const ODT_TYPE: Record<string, string> = { left: "left", center: "center", right: "right", decimal: "char" };
  for (const s of stops) {
    const tb = doc.createElementNS(NS.style, "style:tab-stop");
    tb.setAttributeNS(NS.style, "style:position", pxToCm(s.pos || 0));
    const type = ODT_TYPE[s.val || "left"] ?? "left";
    if (type !== "left") tb.setAttributeNS(NS.style, "style:type", type);
    if (type === "char") tb.setAttributeNS(NS.style, "style:char", ".");
    if (s.leader) {
      tb.setAttributeNS(NS.style, "style:leader-style", "dotted");
      tb.setAttributeNS(NS.style, "style:leader-text", ".");
    }
    ts.appendChild(tb);
  }
  return ts;
}

/** Create (once) a paragraph style for an alignment and return its name. The breakBefore /
    breakAfter / masterPage fields carry a section break (a new page, optionally with a
    different page master) so editing a paragraph does not drop it. */
function paraStyleFor(doc: Document, auto: Element, created: Map<string, string>, p: { align?: string; indentPx?: number; lineHeight?: number; spaceBeforePx?: number; spaceAfterPx?: number; parent?: string; breakBefore?: string; breakAfter?: string; masterPage?: string; tabStops?: { pos: number; val?: string; leader?: string }[]; shading?: string; borders?: BlockBorderSide[] }): string | null {
  const a = p.align ? ODF_ALIGN[p.align] : undefined;
  const align = a && a !== "left" ? a : undefined;
  const indentPx = p.indentPx && p.indentPx > 0 ? Math.round(p.indentPx) : undefined;
  const lineHeight = p.lineHeight && p.lineHeight > 0 ? p.lineHeight : undefined;
  const before = p.spaceBeforePx; // may be 0 (explicit "no space"); undefined = not set
  const after = p.spaceAfterPx;
  const { breakBefore, breakAfter, masterPage } = p;
  const tabStops = p.tabStops && p.tabStops.length ? p.tabStops : undefined;
  const shading = p.shading || undefined; // 6-hex (no #) paragraph shading
  const borders = p.borders && p.borders.length ? p.borders : undefined;
  // With no direct formatting and no section break / tabs, the caller references the named style.
  if (!align && !indentPx && !lineHeight && before === undefined && after === undefined && !breakBefore && !breakAfter && !masterPage && !tabStops && !shading && !borders) return null;
  const key = `p_${align ?? ""}_${indentPx ?? ""}_${lineHeight ?? ""}_${before ?? ""}_${after ?? ""}_${p.parent ?? ""}_${breakBefore ?? ""}_${breakAfter ?? ""}_${masterPage ?? ""}_${tabStops ? JSON.stringify(tabStops) : ""}_${shading ?? ""}_${borders ? borders.map((b) => `${b.side}${b.px}${b.style}${b.hex}`).join(",") : ""}`;
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_p${created.size}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "paragraph");
  if (p.parent) st.setAttributeNS(NS.style, "style:parent-style-name", p.parent); // direct formatting over a named style
  if (masterPage) st.setAttributeNS(NS.style, "style:master-page-name", masterPage); // a section break to a new page master
  const pp = doc.createElementNS(NS.style, "style:paragraph-properties");
  if (align) pp.setAttributeNS(NS.fo, "fo:text-align", align === "right" ? "end" : align === "center" ? "center" : "justify");
  if (indentPx) pp.setAttributeNS(NS.fo, "fo:margin-left", `${Math.round((indentPx / (96 / 2.54)) * 1000) / 1000}cm`);
  if (lineHeight) pp.setAttributeNS(NS.fo, "fo:line-height", `${Math.round(lineHeight * 100)}%`);
  if (before !== undefined) pp.setAttributeNS(NS.fo, "fo:margin-top", pxToCm(before));
  if (after !== undefined) pp.setAttributeNS(NS.fo, "fo:margin-bottom", pxToCm(after));
  if (breakBefore) pp.setAttributeNS(NS.fo, "fo:break-before", breakBefore);
  if (breakAfter) pp.setAttributeNS(NS.fo, "fo:break-after", breakAfter);
  if (shading) pp.setAttributeNS(NS.fo, "fo:background-color", `#${shading}`);
  for (const b of borders ?? []) pp.setAttributeNS(NS.fo, `fo:border-${b.side}`, `${pxToCm(b.px)} ${b.style} #${b.hex.toLowerCase()}`);
  const ts = buildOdtTabStops(doc, tabStops);
  if (ts) pp.appendChild(ts);
  st.appendChild(pp);
  auto.appendChild(st);
  created.set(key, name);
  return name;
}

interface RefMeta {
  author: string;
  date: string;
  text: string;
  resolved: boolean;
  paraId: string;
}
interface OdfCtx {
  doc: Document;
  auto: Element;
  created: Map<string, string>;
  files: Record<string, Uint8Array>; // the archive, so embedded images can be added
  pics: { path: string; mime: string }[]; // images added this run, for the manifest
  objs: { dir: string }[]; // embedded formula objects added this run, for the manifest
  refMeta: Map<string, RefMeta>; // comment id -> metadata, gathered from the body refs
  rangedIds: Set<string>; // comment ids that wrap a text range (vs a point comment)
  done: Map<string, boolean>; // resolve state keyed by comment paraId
  changes: { id: string; type: "insertion" | "deletion"; author: string; date: string; deleted?: string }[]; // tracked changes to emit
  notesById?: Map<string, Note>; // footnote/endnote bodies, to rebuild inline text:note from refs
}

/** Build an office:annotation element for a comment id from its gathered metadata. */
function makeAnnotation(ctx: OdfCtx, id: string): Element {
  const m = ctx.refMeta.get(id);
  const an = ctx.doc.createElementNS(NS.office, "office:annotation");
  an.setAttributeNS(NS.office, "office:name", id);
  const resolved = (m && ctx.done.get(m.paraId)) ?? m?.resolved ?? false;
  if (resolved) an.setAttributeNS(NS.loext, "loext:resolved", "true");
  if (m?.author) {
    const cr = ctx.doc.createElementNS(NS.dc, "dc:creator");
    cr.textContent = m.author;
    an.appendChild(cr);
  }
  if (m?.date) {
    const dt = ctx.doc.createElementNS(NS.dc, "dc:date");
    dt.textContent = m.date;
    an.appendChild(dt);
  }
  for (const line of (m?.text ?? "").split("\n")) {
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    if (line) p.textContent = line;
    an.appendChild(p);
  }
  return an;
}

/** Set (or clear) a draw:frame's alt text via its svg:desc child. */
function setFrameAlt(doc: Document, frame: Element, alt: string): void {
  for (const tag of ["svg:desc", "svg:title"]) {
    const old = frame.getElementsByTagName(tag)[0];
    if (old) old.remove();
  }
  if (alt) {
    const desc = doc.createElementNS(NS.svg, "svg:desc");
    desc.textContent = alt;
    frame.appendChild(desc);
  }
}

/** Serialize an edited <img> back to a draw:frame. A preserved image (data-odt-xml carrying its
 *  original frame) is rebuilt around its own draw:image so the picture part is reused, picking up
 *  the editor's current size + wrap; a brand-new image embeds fresh bytes. */
function buildImageFrame(img: HTMLElement, ctx: OdfCtx): Element | null {
  const layout = imageLayoutFromEl(img);
  const wPx = parseFloat(img.getAttribute("width") ?? "") || undefined;
  const hPx = parseFloat(img.getAttribute("height") ?? "") || undefined;
  const alt = img.getAttribute("alt") || "";
  const stash = img.getAttribute("data-odt-xml");
  if (stash) {
    const frame = importPassthrough(ctx.doc, stash);
    if (frame && frame.getElementsByTagName("draw:image")[0]) {
      if (wPx) frame.setAttributeNS(NS.svg, "svg:width", pxToCm(wPx));
      if (hPx) frame.setAttributeNS(NS.svg, "svg:height", pxToCm(hPx));
      setFrameAlt(ctx.doc, frame, alt);
      applyFrameLayout(ctx.doc, frame, layout, ctx.auto, ctx.created);
      return frame;
    }
    return frame; // not a resolvable image frame: re-emit verbatim
  }
  const m = /^data:([^;]+);base64,(.+)$/.exec(img.getAttribute("src") ?? "");
  if (!m) return null; // not an embeddable image (e.g. an external URL): drop it
  const mime = m[1]!;
  const bin = atob(m[2]!);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const ext = Object.entries(IMG_MIME).find(([, v]) => v === mime)?.[0] ?? (mime.split("/")[1] || "png");
  const idx = ctx.pics.length;
  const path = `Pictures/ot_img${idx}.${ext}`;
  ctx.files[path] = bytes;
  ctx.pics.push({ path, mime });

  const frame = ctx.doc.createElementNS(NS.draw, "draw:frame");
  frame.setAttributeNS(NS.draw, "draw:name", `Image${idx + 1}`);
  if (wPx) frame.setAttributeNS(NS.svg, "svg:width", pxToCm(wPx));
  if (hPx) frame.setAttributeNS(NS.svg, "svg:height", pxToCm(hPx));
  const image = ctx.doc.createElementNS(NS.draw, "draw:image");
  image.setAttributeNS(NS.xlink, "xlink:href", path);
  image.setAttributeNS(NS.xlink, "xlink:type", "simple");
  image.setAttributeNS(NS.xlink, "xlink:show", "embed");
  image.setAttributeNS(NS.xlink, "xlink:actuate", "onLoad");
  frame.appendChild(image);
  setFrameAlt(ctx.doc, frame, alt);
  applyFrameLayout(ctx.doc, frame, layout, ctx.auto, ctx.created); // sets as-char when inline
  return frame;
}

/** Register newly embedded images in META-INF/manifest.xml. */
function addManifestEntries(files: Record<string, Uint8Array>, pics: { path: string; mime: string }[]): void {
  if (!pics.length || !files["META-INF/manifest.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["META-INF/manifest.xml"]), "application/xml");
  const root = doc.getElementsByTagName("manifest:manifest")[0] ?? doc.documentElement;
  if (!root) return;
  const have = new Set(Array.from(doc.getElementsByTagName("manifest:file-entry")).map((e) => e.getAttribute("manifest:full-path")));
  for (const p of pics) {
    if (have.has(p.path)) continue;
    const e = doc.createElementNS(NS.manifest, "manifest:file-entry");
    e.setAttributeNS(NS.manifest, "manifest:full-path", p.path);
    e.setAttributeNS(NS.manifest, "manifest:media-type", p.mime);
    root.appendChild(e);
  }
  files["META-INF/manifest.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Register newly embedded formula objects (the directory + its content.xml) in the manifest. */
function addObjectManifestEntries(files: Record<string, Uint8Array>, objs: { dir: string }[]): void {
  if (!objs.length || !files["META-INF/manifest.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["META-INF/manifest.xml"]), "application/xml");
  const root = doc.getElementsByTagName("manifest:manifest")[0] ?? doc.documentElement;
  if (!root) return;
  const have = new Set(Array.from(doc.getElementsByTagName("manifest:file-entry")).map((e) => e.getAttribute("manifest:full-path")));
  const add = (full: string, media: string): void => {
    if (have.has(full)) return;
    const e = doc.createElementNS(NS.manifest, "manifest:file-entry");
    e.setAttributeNS(NS.manifest, "manifest:full-path", full);
    e.setAttributeNS(NS.manifest, "manifest:media-type", media);
    root.appendChild(e);
  };
  for (const o of objs) {
    add(`${o.dir}/`, "application/vnd.oasis.opendocument.formula");
    add(`${o.dir}/content.xml`, "text/xml");
  }
  files["META-INF/manifest.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Serialize an equation span back to a draw:frame. An untouched imported equation (data-odt-xml
 *  carrying its original frame) is re-emitted verbatim, reusing the Object sub-document already in
 *  the archive; a new or edited one writes a fresh formula sub-document and references it. */
function buildEquationFrame(span: HTMLElement, ctx: OdfCtx): Element | null {
  const stash = span.getAttribute("data-odt-xml");
  if (stash) {
    const frame = importPassthrough(ctx.doc, stash);
    if (frame) return frame;
  }
  const math = span.querySelector("math");
  if (!math) return null;
  const clone = math.cloneNode(true) as Element;
  if (clone.namespaceURI !== NS.math) clone.setAttribute("xmlns", NS.math); // ensure the MathML namespace is declared
  const dir = `Formula_rdoc${ctx.objs.length}`;
  ctx.files[`${dir}/content.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?>\n${new XMLSerializer().serializeToString(clone)}`);
  ctx.objs.push({ dir });
  const frame = ctx.doc.createElementNS(NS.draw, "draw:frame");
  frame.setAttributeNS(NS.draw, "draw:name", `Formula${ctx.objs.length}`);
  frame.setAttributeNS(NS.text, "text:anchor-type", "as-char");
  const obj = ctx.doc.createElementNS(NS.draw, "draw:object");
  obj.setAttributeNS(NS.xlink, "xlink:href", `./${dir}`);
  obj.setAttributeNS(NS.xlink, "xlink:type", "simple");
  obj.setAttributeNS(NS.xlink, "xlink:show", "embed");
  obj.setAttributeNS(NS.xlink, "xlink:actuate", "onLoad");
  frame.appendChild(obj);
  return frame;
}

/** Append the inline content of an HTML node to an ODF block element. */
function htmlInlineToOdf(node: Node, parent: Element, f: Fmt, ctx: OdfCtx): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      const txt = child.textContent ?? "";
      if (!txt) continue;
      // Direct formatting -> an automatic style (parented to the named char style if any);
      // a bare named char style -> referenced directly; nothing -> a plain text node.
      const auto = fmtKey(f) ? styleFor(ctx.doc, ctx.auto, ctx.created, f) : null;
      const styleName = auto ?? f.cStyle ?? null;
      if (!styleName) {
        parent.appendChild(ctx.doc.createTextNode(txt));
      } else {
        const span = ctx.doc.createElementNS(NS.text, "text:span");
        span.setAttributeNS(NS.text, "text:style-name", styleName);
        span.appendChild(ctx.doc.createTextNode(txt));
        parent.appendChild(span);
      }
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    if (tag === "ul" || tag === "ol" || tag === "li") continue; // nested lists handled by htmlListToOdf, not inline
    // An image (new or preserved) is rebuilt so size + wrap edits round-trip; this must run
    // before the generic passthrough below, which would otherwise re-emit the frame verbatim.
    if (tag === "img") {
      const frame = buildImageFrame(el, ctx);
      if (frame) parent.appendChild(frame);
      continue;
    }
    // An equation -> a draw:frame wrapping an embedded formula object; must run before the generic
    // passthrough so an edited equation rebuilds its MathML instead of re-emitting the old frame.
    if (tag === "span" && el.classList.contains("docx-eq")) {
      const frame = buildEquationFrame(el, ctx);
      if (frame) parent.appendChild(frame);
      continue;
    }
    const stash = el.getAttribute("data-odt-xml");
    if (stash) {
      const node2 = importPassthrough(ctx.doc, stash);
      if (node2) parent.appendChild(node2);
      continue;
    }
    if (el.getAttribute("data-docx-tab")) {
      parent.appendChild(ctx.doc.createElementNS(NS.text, "text:tab"));
      continue;
    }
    if (tag === "br") {
      parent.appendChild(ctx.doc.createElementNS(NS.text, "text:line-break"));
      continue;
    }
    if (tag === "sup" && el.classList.contains("docx-fnref")) {
      // Footnote / endnote reference -> an inline text:note (citation + body from the note store).
      const id = el.getAttribute("data-fn-id") || "";
      const kind = el.getAttribute("data-fn-kind") === "endnote" ? "endnote" : "footnote";
      const note = ctx.doc.createElementNS(NS.text, "text:note");
      note.setAttributeNS(NS.text, "text:note-class", kind);
      note.setAttributeNS(NS.text, "text:id", id);
      const cite = ctx.doc.createElementNS(NS.text, "text:note-citation");
      cite.appendChild(ctx.doc.createTextNode(el.textContent || ""));
      const body = ctx.doc.createElementNS(NS.text, "text:note-body");
      const noteHtml = ctx.notesById?.get(id)?.html || "<p><br></p>";
      const htmlDoc = new DOMParser().parseFromString(noteHtml, "text/html");
      for (const node of Array.from(htmlDoc.body.childNodes)) {
        const b = htmlBlockToOdf(node, ctx);
        if (b) body.appendChild(b);
      }
      if (!body.firstChild) body.appendChild(ctx.doc.createElementNS(NS.text, "text:p"));
      note.append(cite, body);
      parent.appendChild(note);
      continue;
    }
    if (tag === "ruby") {
      // Furigana: <ruby>base<rt>reading</rt></ruby> -> text:ruby (ruby-base + ruby-text).
      const ruby = ctx.doc.createElementNS(NS.text, "text:ruby");
      const rtEl = Array.from(el.children).find((c) => c.tagName.toLowerCase() === "rt") as HTMLElement | undefined;
      const baseEl = ctx.doc.createElementNS(NS.text, "text:ruby-base");
      const baseClone = el.cloneNode(true) as HTMLElement;
      for (const r of Array.from(baseClone.children)) if (r.tagName.toLowerCase() === "rt") r.remove();
      htmlInlineToOdf(baseClone, baseEl, f, ctx);
      const textEl = ctx.doc.createElementNS(NS.text, "text:ruby-text");
      textEl.appendChild(ctx.doc.createTextNode(rtEl?.textContent ?? ""));
      ruby.append(baseEl, textEl);
      parent.appendChild(ruby);
      continue;
    }
    if (el.classList.contains("docx-field")) {
      const k = el.getAttribute("data-field");
      if (k === "PAGE") {
        const e = ctx.doc.createElementNS(NS.text, "text:page-number");
        e.setAttributeNS(NS.text, "text:select-page", "current");
        e.appendChild(ctx.doc.createTextNode(el.textContent || "1"));
        parent.appendChild(e);
      } else if (k === "NUMPAGES") {
        const e = ctx.doc.createElementNS(NS.text, "text:page-count");
        e.appendChild(ctx.doc.createTextNode(el.textContent || "1"));
        parent.appendChild(e);
      } else if (k === "seq") {
        // A caption number: an auto-incrementing text:sequence keyed by its name.
        const name = el.getAttribute("data-seq") || "Figure";
        const e = ctx.doc.createElementNS(NS.text, "text:sequence");
        e.setAttributeNS(NS.text, "text:name", name);
        e.setAttributeNS(NS.text, "text:formula", `ooow:${name}+1`);
        e.setAttributeNS(NS.style, "style:num-format", "1");
        e.appendChild(ctx.doc.createTextNode(el.textContent || "1"));
        parent.appendChild(e);
      } else if (k === "DATE" || k === "TIME" || k === "AUTHOR" || k === "FILENAME") {
        // Information fields: the cached snapshot is the element's text content.
        const tag = k === "DATE" ? "text:date" : k === "TIME" ? "text:time" : k === "AUTHOR" ? "text:author-name" : "text:file-name";
        const e = ctx.doc.createElementNS(NS.text, tag);
        if (k === "FILENAME") e.setAttributeNS(NS.text, "text:display", "name-and-extension");
        e.appendChild(ctx.doc.createTextNode(el.textContent || ""));
        parent.appendChild(e);
      }
      continue;
    }
    if (el.classList.contains("docx-comment")) {
      // commented range: annotation at the start, the text, then annotation-end
      const id = el.getAttribute("data-comment-id") ?? "";
      parent.appendChild(makeAnnotation(ctx, id));
      htmlInlineToOdf(el, parent, f, ctx);
      const end = ctx.doc.createElementNS(NS.office, "office:annotation-end");
      end.setAttributeNS(NS.office, "office:name", id);
      parent.appendChild(end);
      continue;
    }
    if (el.classList.contains("docx-comment-ref")) {
      // metadata carrier: emit a point annotation only if it has no wrapped range
      const id = el.getAttribute("data-comment-id") ?? "";
      if (!ctx.rangedIds.has(id)) parent.appendChild(makeAnnotation(ctx, id));
      continue;
    }
    if (el.classList.contains("docx-ins")) {
      // tracked insertion: change-start ... text ... change-end, region built later
      const id = `ct${ctx.changes.length + 1}`;
      ctx.changes.push({ id, type: "insertion", author: el.getAttribute("data-author") ?? "", date: el.getAttribute("data-date") ?? "" });
      const start = ctx.doc.createElementNS(NS.text, "text:change-start");
      start.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(start);
      htmlInlineToOdf(el, parent, f, ctx);
      const end = ctx.doc.createElementNS(NS.text, "text:change-end");
      end.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(end);
      continue;
    }
    if (el.classList.contains("docx-del")) {
      // tracked deletion: a point marker in the body; the removed text lives in the region
      const id = `ct${ctx.changes.length + 1}`;
      ctx.changes.push({ id, type: "deletion", author: el.getAttribute("data-author") ?? "", date: el.getAttribute("data-date") ?? "", deleted: el.textContent ?? "" });
      const mark = ctx.doc.createElementNS(NS.text, "text:change");
      mark.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(mark);
      continue;
    }
    if (el.classList.contains("docx-cmark")) continue; // empty new-comment marker: nothing to emit
    if (el.classList.contains("docx-bookmark")) {
      const name = el.getAttribute("data-rdoc-bm") ?? "";
      if (name) { const b = ctx.doc.createElementNS(NS.text, "text:bookmark-start"); b.setAttributeNS(NS.text, "text:name", name); parent.appendChild(b); }
      continue;
    }
    if (el.classList.contains("docx-bookmark-end")) {
      const name = el.getAttribute("data-rdoc-bm-end") ?? "";
      if (name) { const b = ctx.doc.createElementNS(NS.text, "text:bookmark-end"); b.setAttributeNS(NS.text, "text:name", name); parent.appendChild(b); }
      continue;
    }
    if (el.classList.contains("docx-xref")) {
      // A cross-reference to a bookmark: text:bookmark-ref with the recomputed text as fallback content.
      const name = el.getAttribute("data-rdoc-xref") ?? "";
      const xfmt = el.getAttribute("data-rdoc-xref-fmt");
      const ref = ctx.doc.createElementNS(NS.text, "text:bookmark-ref");
      ref.setAttributeNS(NS.text, "text:reference-format", xfmt === "page" ? "page" : xfmt === "direction" ? "direction" : "text");
      ref.setAttributeNS(NS.text, "text:ref-name", name);
      ref.appendChild(ctx.doc.createTextNode(el.textContent || ""));
      parent.appendChild(ref);
      continue;
    }
    if (tag === "a") {
      const a = ctx.doc.createElementNS(NS.text, "text:a");
      a.setAttributeNS(NS.xlink, "xlink:href", el.getAttribute("href") ?? "");
      htmlInlineToOdf(el, a, f, ctx);
      parent.appendChild(a);
      continue;
    }
    const hp = fontSizeToHalfPt(el.style.fontSize);
    const next: Fmt = {
      b: f.b || tag === "strong" || tag === "b" || /(^|;)\s*font-weight\s*:\s*(bold|[6-9]00)/.test(el.style.cssText),
      i: f.i || tag === "em" || tag === "i" || el.style.fontStyle === "italic",
      u: f.u || tag === "u" || /underline/.test(el.style.textDecoration || el.style.textDecorationLine || ""),
      strike: f.strike || tag === "s" || tag === "strike" || tag === "del" || /line-through/.test(el.style.textDecoration || el.style.textDecorationLine || ""),
      vertAlign: f.vertAlign ?? (tag === "sup" || /vertical-align:\s*super/.test(el.style.cssText) ? "super" : tag === "sub" || /vertical-align:\s*sub/.test(el.style.cssText) ? "sub" : undefined),
      color: toHex6(el.style.color) ?? f.color,
      bg: toHex6(el.style.backgroundColor) ?? f.bg,
      font: firstFontFamily(el.style.fontFamily) ?? f.font,
      sizePt: hp ? hp / 2 : f.sizePt,
      cStyle: el.getAttribute("data-rdoc-cstyle") || f.cStyle,
    };
    htmlInlineToOdf(el, parent, next, ctx);
  }
}

/** A 10-level list style, all-number or all-bullet, cached on the context. Each text:list
    element carries the style matching its own tag, so nesting ol/ul in any order round-trips. */
function listStyleFor(ctx: OdfCtx, ordered: boolean, start = 1): string {
  const key = ordered ? `list:ordered:${start}` : "list:bullet";
  const cached = ctx.created.get(key);
  if (cached) return cached;
  const name = `OT_L${ordered ? "O" : "B"}${start > 1 ? start : ""}`;
  const ls = ctx.doc.createElementNS(NS.text, "text:list-style");
  ls.setAttributeNS(NS.style, "style:name", name);
  for (let l = 1; l <= 10; l++) {
    const lvl = ctx.doc.createElementNS(NS.text, ordered ? "text:list-level-style-number" : "text:list-level-style-bullet");
    lvl.setAttributeNS(NS.text, "text:level", String(l));
    if (ordered) {
      lvl.setAttributeNS(NS.style, "style:num-format", "1");
      lvl.setAttributeNS(NS.style, "style:num-suffix", ".");
      if (l === 1 && start > 1) lvl.setAttributeNS(NS.text, "text:start-value", String(start)); // restart/continue at N
    } else {
      lvl.setAttributeNS(NS.text, "text:bullet-char", ["•", "◦", "▪"][(l - 1) % 3]);
    }
    const lp = ctx.doc.createElementNS(NS.style, "style:list-level-properties");
    lp.setAttributeNS(NS.text, "text:list-level-position-and-space-mode", "label-alignment");
    const la = ctx.doc.createElementNS(NS.style, "style:list-level-label-alignment");
    la.setAttributeNS(NS.text, "text:label-followed-by", "listtab");
    la.setAttributeNS(NS.fo, "fo:margin-left", `${l * 0.635}cm`);
    la.setAttributeNS(NS.fo, "fo:text-indent", "-0.635cm");
    lp.appendChild(la);
    lvl.appendChild(lp);
    ls.appendChild(lvl);
  }
  ctx.auto.appendChild(ls);
  ctx.created.set(key, name);
  return name;
}

function htmlListToOdf(el: HTMLElement, ctx: OdfCtx): Element {
  const ordered = el.tagName.toLowerCase() === "ol";
  const start = ordered ? Math.max(1, parseInt(el.getAttribute("start") || "1", 10) || 1) : 1;
  const list = ctx.doc.createElementNS(NS.text, "text:list");
  list.setAttributeNS(NS.text, "text:style-name", listStyleFor(ctx, ordered, start));
  for (const li of Array.from(el.children)) {
    if (li.tagName.toLowerCase() !== "li") continue;
    const item = ctx.doc.createElementNS(NS.text, "text:list-item");
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    htmlInlineToOdf(li, p, FMT0, ctx);
    item.appendChild(p);
    for (const nested of Array.from(li.children)) {
      const nt = nested.tagName.toLowerCase();
      if (nt === "ul" || nt === "ol") item.appendChild(htmlListToOdf(nested as HTMLElement, ctx));
    }
    list.appendChild(item);
  }
  return list;
}

/** Rebuild a table:table from its preserved skeleton, replacing each (non-covered) cell's
    content with the edited content from the matching .docx-cell; structure/spans preserved. */
function rebuildOdtTable(tableEl: HTMLElement, stash: string, ctx: OdfCtx): Element | null {
  const tbl = importPassthrough(ctx.doc, stash);
  if (!tbl || tbl.tagName !== "table:table") return tbl;
  const cells = Array.from(tableEl.querySelectorAll(".docx-cell"));
  let i = 0;
  for (const tr of Array.from(tbl.children)) {
    if (tr.tagName !== "table:table-row") continue;
    for (const tc of Array.from(tr.children)) {
      if (tc.tagName !== "table:table-cell") continue; // covered cells stay as-is
      const cellEl = cells[i++];
      if (!cellEl) continue;
      while (tc.firstChild) tc.removeChild(tc.firstChild);
      for (const node of Array.from(cellEl.childNodes)) {
        const b = htmlBlockToOdf(node, ctx);
        if (b) tc.appendChild(b);
      }
      if (!tc.firstChild) tc.appendChild(ctx.doc.createElementNS(NS.text, "text:p"));
    }
  }
  return tbl;
}

let odtTableSeq = 0;
let odtCellBorderSeq = 0;
let odtColSeq = 0;
let odtRowSeq = 0;
let odtTableIndentSeq = 0;
let odtTocSeq = 0;

/** A table of contents as text:table-of-content carrying the cached entries; ODF apps update it. */
function buildOdtToc(el: HTMLElement, ctx: OdfCtx): Element {
  const toc = ctx.doc.createElementNS(NS.text, "text:table-of-content");
  toc.setAttributeNS(NS.text, "text:name", `TOC${++odtTocSeq}`);
  const source = ctx.doc.createElementNS(NS.text, "text:table-of-content-source");
  source.setAttributeNS(NS.text, "text:outline-level", "3");
  toc.appendChild(source);
  const idx = ctx.doc.createElementNS(NS.text, "text:index-body");
  const title = el.querySelector(".docx-field-toc-title")?.textContent;
  if (title) {
    const tp = ctx.doc.createElementNS(NS.text, "text:p");
    tp.appendChild(ctx.doc.createTextNode(title));
    idx.appendChild(tp);
  }
  for (const row of Array.from(el.querySelectorAll(".docx-field-toc-row"))) {
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    p.appendChild(ctx.doc.createTextNode(row.querySelector(".docx-field-toc-text")?.textContent ?? ""));
    const page = row.querySelector(".docx-field-toc-page")?.textContent ?? "";
    if (page) {
      p.appendChild(ctx.doc.createElementNS(NS.text, "text:tab"));
      p.appendChild(ctx.doc.createTextNode(page));
    }
    idx.appendChild(p);
  }
  toc.appendChild(idx);
  return toc;
}

function findTableStyle(ctx: OdfCtx, name: string): Element | undefined {
  for (const st of Array.from(ctx.doc.getElementsByTagName("style:style")))
    if (st.getAttribute("style:name") === name && st.getAttribute("style:family") === "table") return st;
  return undefined;
}
/** A table style carrying a left indent (fo:margin-left), cloning the preserved table style so
    its other properties (width, borders) survive. Cached per (base style + indent). */
function tableIndentStyleFor(ctx: OdfCtx, px: number, baseName: string | null): string {
  const key = `tind:${baseName ?? ""}|${Math.round(px)}`;
  const cached = ctx.created.get(key);
  if (cached) return cached;
  const name = `OTtbl${++odtTableIndentSeq}`;
  const base = baseName ? findTableStyle(ctx, baseName) : undefined;
  let st: Element;
  let props: Element;
  if (base) {
    st = base.cloneNode(true) as Element;
    st.setAttributeNS(NS.style, "style:name", name);
    props = st.getElementsByTagName("style:table-properties")[0] ?? st.appendChild(ctx.doc.createElementNS(NS.style, "style:table-properties"));
  } else {
    st = ctx.doc.createElementNS(NS.style, "style:style");
    st.setAttributeNS(NS.style, "style:name", name);
    st.setAttributeNS(NS.style, "style:family", "table");
    props = st.appendChild(ctx.doc.createElementNS(NS.style, "style:table-properties"));
  }
  props.setAttributeNS(NS.fo, "fo:margin-left", pxToCm(px));
  props.setAttributeNS(NS.table, "table:align", "left");
  ctx.auto.appendChild(st);
  ctx.created.set(key, name);
  return name;
}

/** A table-column automatic style carrying a resized column width (px). Cached per width. */
function colStyleFor(ctx: OdfCtx, px: number): string {
  const key = `colw:${px}`;
  const cached = ctx.created.get(key);
  if (cached) return cached;
  const name = `OTcol${++odtColSeq}`;
  const st = ctx.doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "table-column");
  const props = ctx.doc.createElementNS(NS.style, "style:table-column-properties");
  props.setAttributeNS(NS.style, "style:column-width", pxToCm(px));
  st.appendChild(props);
  ctx.auto.appendChild(st);
  ctx.created.set(key, name);
  return name;
}
/** A table-row automatic style carrying a resized row height (px). Cached per height. */
function rowStyleFor(ctx: OdfCtx, px: number): string {
  const key = `rowh:${px}`;
  const cached = ctx.created.get(key);
  if (cached) return cached;
  const name = `OTrow${++odtRowSeq}`;
  const st = ctx.doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "table-row");
  const props = ctx.doc.createElementNS(NS.style, "style:table-row-properties");
  props.setAttributeNS(NS.style, "style:min-row-height", pxToCm(px));
  st.appendChild(props);
  ctx.auto.appendChild(st);
  ctx.created.set(key, name);
  return name;
}

// fo:border accepts the CSS-style keywords directly (solid/dashed/dotted/double).
function odtBorderValue(td: HTMLTableCellElement, side: string): string {
  const v = td.getAttribute(`data-rdoc-b${side}`);
  if (!v) return "none";
  const m = v.match(/^([\d.]+)px\s+(\w+)\s+(#[0-9a-fA-F]{3,8})/i);
  if (!m) return "0.018cm solid #000000";
  return `${pxToCm(parseFloat(m[1]!))} ${m[2]!.toLowerCase()} ${m[3]}`;
}
function findCellStyle(ctx: OdfCtx, name: string): Element | undefined {
  for (const st of Array.from(ctx.doc.getElementsByTagName("style:style")))
    if (st.getAttribute("style:name") === name && st.getAttribute("style:family") === "table-cell") return st;
  return undefined;
}
/** A table-cell automatic style carrying the cell's borders. The preserved cell style (if any)
    is cloned so its other properties (background, padding) survive, then its borders are
    replaced with the per-side spec. Cached per (base style + border spec). */
function cellBorderStyleFor(ctx: OdfCtx, td: HTMLTableCellElement): string | null {
  if (!td.classList.contains("rdoc-bordered")) return null;
  const sides = [["t", "border-top"], ["r", "border-right"], ["b", "border-bottom"], ["l", "border-left"]] as const;
  const spec = sides.map(([s]) => odtBorderValue(td, s));
  const baseName = td.getAttribute("data-odt-cellstyle") ?? "";
  const key = "cellb:" + baseName + "|" + spec.join("|");
  const cached = ctx.created.get(key);
  if (cached) return cached;
  const name = `OTc${++odtCellBorderSeq}`;
  const base = baseName ? findCellStyle(ctx, baseName) : undefined;
  let st: Element;
  let props: Element;
  if (base) {
    st = base.cloneNode(true) as Element;
    st.setAttributeNS(NS.style, "style:name", name);
    props = st.getElementsByTagName("style:table-cell-properties")[0] ?? st.appendChild(ctx.doc.createElementNS(NS.style, "style:table-cell-properties"));
    for (const a of ["border", "border-top", "border-right", "border-bottom", "border-left"]) props.removeAttributeNS(NS.fo, a);
  } else {
    st = ctx.doc.createElementNS(NS.style, "style:style");
    st.setAttributeNS(NS.style, "style:name", name);
    st.setAttributeNS(NS.style, "style:family", "table-cell");
    props = st.appendChild(ctx.doc.createElementNS(NS.style, "style:table-cell-properties"));
  }
  sides.forEach(([, attr], i) => props.setAttributeNS(NS.fo, `fo:${attr}`, spec[i]!));
  ctx.auto.appendChild(st);
  ctx.created.set(key, name);
  return name;
}
/** Build a fresh table:table from a table inserted in the editor (no skeleton): one column
    declaration per column, one cell per <td> with its edited content. No spans. */
function buildNewOdtTable(tableEl: HTMLElement, ctx: OdfCtx): Element {
  const rows = Array.from((tableEl as HTMLTableElement).rows);
  const gridCols = Math.max(1, ...rows.map((tr) => Array.from(tr.cells).reduce((s, td) => s + ((td as HTMLTableCellElement).colSpan || 1), 0)));
  const tbl = ctx.doc.createElementNS(NS.table, "table:table");
  tbl.setAttributeNS(NS.table, "table:name", `Table${++odtTableSeq}`);
  const tstyle = tableEl.getAttribute("data-odt-tablestyle");
  // Table indent (dragging the outer-left edge) -> a table style with fo:margin-left.
  const indentPx = parseFloat((tableEl as HTMLElement).style.marginLeft) || 0;
  const styleName = indentPx > 0 ? tableIndentStyleFor(ctx, indentPx, tstyle) : tstyle;
  if (styleName) tbl.setAttributeNS(NS.table, "table:style-name", styleName);
  // Resized column widths from a <colgroup> (px) take precedence over the preserved columns.
  const cgEl = tableEl.querySelector(":scope > colgroup");
  const colWidths: number[] = cgEl ? Array.from(cgEl.children).map((c) => parseFloat((c as HTMLElement).style.width) || 0) : [];
  // Column declarations: reuse the preserved ones when the count still matches, else default.
  let cols: { s: string; r: string }[] = [];
  try {
    cols = JSON.parse(tableEl.getAttribute("data-odt-cols") ?? "[]");
  } catch {
    cols = [];
  }
  const colTotal = cols.reduce((s, c) => s + (Number(c.r) || 1), 0);
  if (colWidths.length === gridCols && colWidths.every((w) => w > 0)) {
    for (const wpx of colWidths) {
      const col = ctx.doc.createElementNS(NS.table, "table:table-column");
      col.setAttributeNS(NS.table, "table:style-name", colStyleFor(ctx, wpx));
      tbl.appendChild(col);
    }
  } else if (cols.length && colTotal === gridCols) {
    for (const c of cols) {
      const col = ctx.doc.createElementNS(NS.table, "table:table-column");
      if (c.s) col.setAttributeNS(NS.table, "table:style-name", c.s);
      if (Number(c.r) > 1) col.setAttributeNS(NS.table, "table:number-columns-repeated", String(c.r));
      tbl.appendChild(col);
    }
  } else {
    const col = ctx.doc.createElementNS(NS.table, "table:table-column");
    col.setAttributeNS(NS.table, "table:number-columns-repeated", String(gridCols));
    tbl.appendChild(col);
  }
  // Walk the grid so a vertical merge writes number-rows-spanned + covered cells below.
  const covered = (): Element => ctx.doc.createElementNS(NS.table, "table:covered-table-cell");
  const occupied: ({ rows: number; span: number } | null)[] = [];
  for (const tr of rows) {
    const wtr = ctx.doc.createElementNS(NS.table, "table:table-row");
    const rh = parseFloat((tr as HTMLElement).style.height);
    if (rh > 0) wtr.setAttributeNS(NS.table, "table:style-name", rowStyleFor(ctx, rh));
    const domCells = Array.from(tr.cells);
    let di = 0;
    let col = 0;
    while (col < gridCols) {
      const occ = occupied[col];
      if (occ && occ.rows > 0) {
        for (let k = 0; k < occ.span; k++) wtr.appendChild(covered());
        occ.rows--;
        col += occ.span;
        continue;
      }
      const td = domCells[di++] as HTMLTableCellElement | undefined;
      if (!td) break;
      const cs = td.colSpan || 1;
      const rs = td.rowSpan || 1;
      const tc = ctx.doc.createElementNS(NS.table, "table:table-cell");
      const bstyle = cellBorderStyleFor(ctx, td);
      const cstyle = bstyle ?? td.getAttribute("data-odt-cellstyle");
      if (cstyle) tc.setAttributeNS(NS.table, "table:style-name", cstyle);
      if (cs > 1) tc.setAttributeNS(NS.table, "table:number-columns-spanned", String(cs));
      if (rs > 1) tc.setAttributeNS(NS.table, "table:number-rows-spanned", String(rs));
      const cell = (td.querySelector(".docx-cell") as HTMLElement) ?? td;
      for (const node of Array.from(cell.childNodes)) {
        const b = htmlBlockToOdf(node, ctx);
        if (b) tc.appendChild(b);
      }
      if (!tc.firstChild) tc.appendChild(ctx.doc.createElementNS(NS.text, "text:p"));
      wtr.appendChild(tc);
      for (let k = 1; k < cs; k++) wtr.appendChild(covered()); // same-row covered (colspan)
      if (rs > 1) occupied[col] = { rows: rs - 1, span: cs };
      col += cs;
    }
    tbl.appendChild(wtr);
  }
  return tbl;
}

function htmlBlockToOdf(node: Node, ctx: OdfCtx): Element | null {
  if (node.nodeType === 3) {
    if (!(node.textContent ?? "").trim()) return null;
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    p.appendChild(ctx.doc.createTextNode(node.textContent ?? ""));
    return p;
  }
  if (node.nodeType !== 1) return null;
  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  if (el.classList.contains("docx-table")) {
    const skel = el.getAttribute("data-odt-xml");
    return skel ? rebuildOdtTable(el, skel, ctx) : buildNewOdtTable(el, ctx);
  }
  if (el.classList.contains("docx-field-toc")) return buildOdtToc(el, ctx);
  // A section-break page boundary is display-only; the real break rides the paragraph's style.
  if (el.getAttribute("data-docx-pagebreak")) return null;
  const stash = el.getAttribute("data-odt-xml");
  if (stash) return importPassthrough(ctx.doc, stash);
  if (tag === "ul" || tag === "ol") return htmlListToOdf(el, ctx);
  const applyAlign = (block: Element): void => {
    const base = el.getAttribute("data-rdoc-style") || undefined; // a named paragraph style
    const name = paraStyleFor(ctx.doc, ctx.auto, ctx.created, {
      align: el.style.textAlign || undefined,
      indentPx: parseFloat(el.style.marginLeft) || undefined,
      lineHeight: parseFloat(el.style.lineHeight) || undefined,
      spaceBeforePx: el.style.marginTop !== "" ? parseFloat(el.style.marginTop) || 0 : undefined,
      spaceAfterPx: el.style.marginBottom !== "" ? parseFloat(el.style.marginBottom) || 0 : undefined,
      parent: base,
      breakBefore: el.getAttribute("data-odt-break-before") || undefined,
      breakAfter: el.getAttribute("data-odt-break-after") || undefined,
      masterPage: el.getAttribute("data-odt-masterpage") || undefined,
      tabStops: parseTabStops(el.getAttribute("data-rdoc-tabstops")),
      shading: toHex6(el.style.backgroundColor) || undefined,
      borders: blockBorders(el),
    });
    // direct formatting -> an automatic style (parented to the named one); otherwise the named style itself
    if (name) block.setAttributeNS(NS.text, "text:style-name", name);
    else if (base) block.setAttributeNS(NS.text, "text:style-name", base);
  };
  const m = /^h([1-6])$/.exec(tag);
  if (m) {
    const h = ctx.doc.createElementNS(NS.text, "text:h");
    h.setAttributeNS(NS.text, "text:outline-level", String(Math.min(3, Number(m[1]))));
    applyAlign(h);
    htmlInlineToOdf(el, h, FMT0, ctx);
    return h;
  }
  // p, div, and anything else become a paragraph
  const p = ctx.doc.createElementNS(NS.text, "text:p");
  applyAlign(p);
  htmlInlineToOdf(el, p, FMT0, ctx);
  return p;
}

/** Return the document's master page, creating a minimal "Standard" one (with a page
    layout to reference) if the file has none, so a header/footer can be added from
    scratch. Most ODF files already have one; this covers the rare file that does not. */
function ensureMasterPage(doc: Document): Element | null {
  const existing = doc.getElementsByTagName("style:master-page")[0];
  if (existing) return existing;
  const root = doc.documentElement;
  if (!root) return null;
  let pl = doc.getElementsByTagName("style:page-layout")[0];
  let plName = pl?.getAttribute("style:name") ?? null;
  if (!plName) {
    pl = doc.createElementNS(NS.style, "style:page-layout");
    plName = "pm-rdoc";
    pl.setAttributeNS(NS.style, "style:name", plName);
    pl.appendChild(doc.createElementNS(NS.style, "style:page-layout-properties"));
    ensureAutoStyles(doc, "office:master-styles").appendChild(pl);
  }
  let ms = doc.getElementsByTagName("office:master-styles")[0];
  if (!ms) {
    ms = doc.createElementNS(NS.office, "office:master-styles");
    root.appendChild(ms);
  }
  const master = doc.createElementNS(NS.style, "style:master-page");
  master.setAttributeNS(NS.style, "style:name", "Standard");
  master.setAttributeNS(NS.style, "style:page-layout-name", plName);
  ms.appendChild(master);
  return master;
}

/** Write edited header/footer HTML back into the master pages in styles.xml. A part path is
    "header"/"footer" for the default master, or "header@<master>"/"footer@<master>" for a
    distinct per-section master. */
function applyHeaderFooter(files: Record<string, Uint8Array>, parts: { path: string; html: string }[]): void {
  // "header"/"footer" (default), "header@M"/"footer@M" (per-section master), "header-left@M" /
  // "header-first@M" (even / first-page variants, ODF style:header-left / style:header-first), and
  // the "header:even" / "header:first" sentinels from a freshly-toggled variant.
  const hf = parts.filter((p) => /^(header|footer)(-left|-first)?(@|$)/.test(p.path) || /^(header|footer):(even|first)$/.test(p.path));
  if (!hf.length || !files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const defaultMaster = ensureMasterPage(doc);
  const byName = (name: string): Element | null =>
    Array.from(doc.getElementsByTagName("style:master-page")).find((m) => m.getAttribute("style:name") === name) ?? null;
  const ctx: OdfCtx = {
    doc,
    auto: ensureAutoStyles(doc, "office:master-styles"),
    created: new Map(),
    files,
    pics: [],
    objs: [],
    refMeta: new Map(),
    rangedIds: new Set(),
    done: new Map(),
    changes: [],
  };
  for (const p of hf) {
    const sentinel = /^(header|footer):(even|first)$/.exec(p.path);
    let role: string, slot: string, master: Element | null;
    if (sentinel) { role = sentinel[1]!; slot = sentinel[2] === "even" ? "-left" : "-first"; master = defaultMaster; }
    else {
      const m = /^(header|footer)(-left|-first)?(?:@(.+))?$/.exec(p.path);
      if (!m) continue;
      role = m[1]!; slot = m[2] ?? ""; master = m[3] ? byName(m[3]) : defaultMaster;
    }
    if (!master) continue;
    const tag = (role === "header" ? "style:header" : "style:footer") + slot;
    let el = master.getElementsByTagName(tag)[0];
    if (!el) {
      el = doc.createElementNS(NS.style, tag);
      master.appendChild(el);
    }
    while (el.firstChild) el.removeChild(el.firstChild);
    const htmlDoc = new DOMParser().parseFromString(p.html || "<p><br></p>", "text/html");
    for (const node of Array.from(htmlDoc.body.childNodes)) {
      const block = htmlBlockToOdf(node, ctx);
      if (block) el.appendChild(block);
    }
    if (!el.firstChild) el.appendChild(doc.createElementNS(NS.text, "text:p"));
  }
  addManifestEntries(files, ctx.pics);
  files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Build the <text:tracked-changes> region from the changes collected while serializing. */
function buildTrackedChanges(ctx: OdfCtx): Element | null {
  if (!ctx.changes.length) return null;
  const tc = ctx.doc.createElementNS(NS.text, "text:tracked-changes");
  for (const ch of ctx.changes) {
    const region = ctx.doc.createElementNS(NS.text, "text:changed-region");
    region.setAttributeNS(NS.text, "text:id", ch.id);
    const kind = ctx.doc.createElementNS(NS.text, ch.type === "insertion" ? "text:insertion" : "text:deletion");
    const info = ctx.doc.createElementNS(NS.office, "office:change-info");
    if (ch.author) {
      const cr = ctx.doc.createElementNS(NS.dc, "dc:creator");
      cr.textContent = ch.author;
      info.appendChild(cr);
    }
    if (ch.date) {
      const dt = ctx.doc.createElementNS(NS.dc, "dc:date");
      dt.textContent = ch.date;
      info.appendChild(dt);
    }
    kind.appendChild(info);
    if (ch.type === "deletion") {
      for (const line of (ch.deleted ?? "").split("\n")) {
        const p = ctx.doc.createElementNS(NS.text, "text:p");
        if (line) p.textContent = line;
        kind.appendChild(p);
      }
    }
    region.appendChild(kind);
    tc.appendChild(region);
  }
  return tc;
}

/** Set a page-layout-properties element's size, orientation, margins and columns (px -> cm), in
    place. Shared by the document Page setup and per-section master pages. */
const ODT_NUMFMT_WRITE: Record<string, string> = { decimal: "1", lowerRoman: "i", upperRoman: "I", lowerLetter: "a", upperLetter: "A" };
function setPageLayoutGeom(doc: Document, props: Element, g: { w: number; h: number; mt: number; mr: number; mb: number; ml: number; cols?: number; colGap?: number; vertical?: boolean; rtl?: boolean; pageBorder?: PageBorder; pageNumFormat?: string }): void {
  // Size + orientation (page-width/height are stored already swapped for landscape).
  props.setAttributeNS(NS.fo, "fo:page-width", pxToCm(g.w));
  props.setAttributeNS(NS.fo, "fo:page-height", pxToCm(g.h));
  props.setAttributeNS(NS.style, "style:print-orientation", g.w > g.h ? "landscape" : "portrait");
  // Writing direction: tategaki (tb-rl) / horizontal RTL (rl-tb) / default (lr-tb).
  props.setAttributeNS(NS.style, "style:writing-mode", g.vertical ? "tb-rl" : g.rtl ? "rl-tb" : "lr-tb");
  // Margins.
  props.setAttributeNS(NS.fo, "fo:margin-top", pxToCm(g.mt));
  props.setAttributeNS(NS.fo, "fo:margin-right", pxToCm(g.mr));
  props.setAttributeNS(NS.fo, "fo:margin-bottom", pxToCm(g.mb));
  props.setAttributeNS(NS.fo, "fo:margin-left", pxToCm(g.ml));
  // Columns: a style:columns child with count + gap. Keep a custom layout (unequal rel-widths or a
  // separator line) untouched when the count is unchanged, since the editor models only count + gap;
  // otherwise (re)build a plain equal-width one, and remove it entirely when down to one column.
  const n = g.cols && g.cols > 1 ? g.cols : 1;
  const old = props.getElementsByTagName("style:columns")[0];
  const colChildren = old ? Array.from(old.getElementsByTagName("style:column")) : [];
  const hasSep = !!old && old.getElementsByTagName("style:column-sep").length > 0;
  const relWidths = colChildren.map((c) => c.getAttribute("style:rel-width"));
  const unequal = relWidths.length > 1 && new Set(relWidths).size > 1;
  const existingNum = old ? Number(old.getAttribute("fo:column-count")) || colChildren.length : 0;
  const keepCustom = !!old && n > 1 && (hasSep || unequal) && existingNum === n;
  if (!keepCustom) {
    if (old) old.parentNode!.removeChild(old);
    if (n > 1) {
      const colsEl = doc.createElementNS(NS.style, "style:columns");
      colsEl.setAttributeNS(NS.fo, "fo:column-count", String(n));
      colsEl.setAttributeNS(NS.fo, "fo:column-gap", pxToCm(g.colGap ?? 36));
      props.appendChild(colsEl);
    }
  }
  // Page border: a uniform fo:border on the page layout; clear any per-side variants too.
  for (const a of ["fo:border", "fo:border-top", "fo:border-right", "fo:border-bottom", "fo:border-left"]) props.removeAttributeNS(NS.fo, a.slice(3));
  if (g.pageBorder) {
    const pb = g.pageBorder;
    props.setAttributeNS(NS.fo, "fo:border", `${pxToCm(pb.widthPx)} ${pb.style} #${pb.color.toLowerCase()}`);
  }
  // Page-number format: style:num-format on the page layout (odt has no page-layout start number).
  props.removeAttributeNS(NS.style, "num-format");
  if (g.pageNumFormat && ODT_NUMFMT_WRITE[g.pageNumFormat]) props.setAttributeNS(NS.style, "style:num-format", ODT_NUMFMT_WRITE[g.pageNumFormat]!);
}

/** Remove a section master's style:header/style:footer when its section has linked back to the
    default (no data-rdoc-sec*key on the boundary paragraph), so a relink sticks on reopen. */
function reconcileSectionBands(files: Record<string, Uint8Array>, htmlDoc: Document): void {
  const secs = Array.from(htmlDoc.querySelectorAll("[data-odt-masterpage]"));
  if (!secs.length || !files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const byName = (name: string | null): Element | null =>
    name ? (Array.from(doc.getElementsByTagName("style:master-page")).find((m) => m.getAttribute("style:name") === name) ?? null) : null;
  let touched = false;
  const drop = (master: Element, tag: string) => { const e = master.getElementsByTagName(tag)[0]; if (e) { e.parentNode!.removeChild(e); touched = true; } };
  for (const el of secs) {
    const master = byName(el.getAttribute("data-odt-masterpage"));
    if (!master) continue;
    if (!el.getAttribute("data-rdoc-secheaderkey")) drop(master, "style:header");
    if (!el.getAttribute("data-rdoc-secfooterkey")) drop(master, "style:footer");
  }
  if (touched) files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Update the first page-layout in styles.xml from the edited document geometry. */
function applyPageMargins(files: Record<string, Uint8Array>, geometry: PageGeometry): void {
  if (!files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const pl = doc.getElementsByTagName("style:page-layout")[0];
  if (!pl) return;
  let props = pl.getElementsByTagName("style:page-layout-properties")[0];
  if (!props) {
    props = doc.createElementNS(NS.style, "style:page-layout-properties");
    pl.insertBefore(props, pl.firstChild);
  }
  setPageLayoutGeom(doc, props, { w: geometry.widthPx, h: geometry.heightPx, mt: geometry.margin.top, mr: geometry.margin.right, mb: geometry.margin.bottom, ml: geometry.margin.left, cols: geometry.columns, colGap: geometry.columnGapPx, vertical: geometry.vertical, rtl: geometry.rtl, pageBorder: geometry.pageBorder, pageNumFormat: geometry.pageNumFormat });
  // Line numbering: a document-level text:linenumbering-configuration in office:styles (odt has no
  // per-section / start; only on/off, interval and restart-each-page round-trip).
  for (const e of Array.from(doc.getElementsByTagName("text:linenumbering-configuration"))) e.parentNode!.removeChild(e);
  if (geometry.lineNumbers) {
    let officeStyles = doc.getElementsByTagName("office:styles")[0];
    if (!officeStyles) { officeStyles = doc.createElementNS(NS.office, "office:styles"); doc.documentElement!.appendChild(officeStyles); }
    const lnc = doc.createElementNS(NS.text, "text:linenumbering-configuration");
    lnc.setAttributeNS(NS.text, "text:number-lines", "true");
    lnc.setAttributeNS(NS.text, "text:count-empty-lines", "true");
    lnc.setAttributeNS(NS.text, "text:increment", String(Math.max(1, Math.round(geometry.lineNumberInterval ?? 1))));
    lnc.setAttributeNS(NS.text, "text:restart-on-page-change", geometry.lineNumberRestart === "newPage" ? "true" : "false");
    officeStyles.appendChild(lnc);
  }
  files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Ensure each in-document section (a paragraph carrying data-rdoc-secstart + a master-page name)
    has a master-page + page-layout in styles.xml with that geometry, creating them for inserted
    sections and updating them for edited ones. The paragraph's text:style-name already references
    the master via style:master-page-name (set in htmlBlockToOdf). */
function applySectionMasters(files: Record<string, Uint8Array>, htmlDoc: Document): void {
  const secs = Array.from(htmlDoc.querySelectorAll("[data-rdoc-secstart][data-odt-masterpage]"));
  if (!secs.length || !files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const root = doc.documentElement;
  if (!root) return;
  const layouts = ensureAutoStyles(doc); // style:page-layout lives in office:automatic-styles
  let masters = doc.getElementsByTagName("office:master-styles")[0]; // style:master-page lives here
  if (!masters) { masters = doc.createElementNS(NS.office, "office:master-styles"); root.appendChild(masters); }
  let touched = false;
  for (const el of secs) {
    const name = el.getAttribute("data-odt-masterpage")!;
    let g: { w: number; h: number; mt: number; mr: number; mb: number; ml: number; cols?: number; colGap?: number; vertical?: boolean; rtl?: boolean; pageBorder?: PageBorder; pageNumFormat?: string };
    try { g = JSON.parse(el.getAttribute("data-rdoc-secstart")!); } catch { continue; }
    let master = Array.from(doc.getElementsByTagName("style:master-page")).find((m) => m.getAttribute("style:name") === name);
    // Leave an untouched section's existing master byte-for-byte; only act on edited / inserted ones.
    if (master && el.getAttribute("data-rdoc-secedited") !== "1") continue;
    touched = true;
    let plName = master?.getAttribute("style:page-layout-name") ?? `${name}-pl`;
    let pl = Array.from(doc.getElementsByTagName("style:page-layout")).find((p) => p.getAttribute("style:name") === plName);
    if (!pl) {
      pl = doc.createElementNS(NS.style, "style:page-layout");
      pl.setAttributeNS(NS.style, "style:name", plName);
      layouts.appendChild(pl);
    }
    let props = pl.getElementsByTagName("style:page-layout-properties")[0];
    if (!props) { props = doc.createElementNS(NS.style, "style:page-layout-properties"); pl.insertBefore(props, pl.firstChild); }
    setPageLayoutGeom(doc, props, g);
    if (!master) {
      master = doc.createElementNS(NS.style, "style:master-page");
      master.setAttributeNS(NS.style, "style:name", name);
      master.setAttributeNS(NS.style, "style:page-layout-name", plName);
      masters.appendChild(master);
    }
  }
  if (touched) files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Rebuild an .odt from edited HTML, preserving every other part of the archive. */
/** Add user-authored styles to styles.xml's office:styles, translating the CSS-like props to a
    style:style (paragraph or text family). */
function addOdtStyles(files: Record<string, Uint8Array>, styles: NewStyle[]): void {
  const key = "styles.xml";
  const doc = files[key]
    ? new DOMParser().parseFromString(strFromU8(files[key]!), "application/xml")
    : new DOMParser().parseFromString(`<office:document-styles xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:fo="${NS.fo}" xmlns:text="${NS.text}"><office:styles/></office:document-styles>`, "application/xml");
  let officeStyles = doc.getElementsByTagName("office:styles")[0];
  if (!officeStyles) {
    officeStyles = doc.createElementNS(NS.office, "office:styles");
    doc.documentElement!.appendChild(officeStyles);
  }
  // Existing named styles by id, so editing replaces a style's properties in place.
  const existingById = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    const id = st.getAttribute("style:name");
    if (id) existingById.set(id, st);
  }
  // ODF keeps style properties as attributes on these elements, so on edit only the dialog's
  // own attributes are re-derived; the long tail (fo:keep-with-next, text-indent, ...) is kept.
  const directChild = (parent: Element, tag: string): Element | undefined => Array.from(parent.children).find((c) => c.tagName === tag);
  const PARA_FO = ["text-align", "margin-left", "margin-top", "margin-bottom", "line-height", "background-color", "border-top", "border-right", "border-bottom", "border-left"];
  const RUN_FO = ["font-weight", "font-style", "color", "font-size", "background-color", "font-family"];
  const RUN_STYLE = ["text-underline-style", "text-underline-width", "text-underline-color", "text-line-through-style", "text-line-through-type", "font-name"];
  for (const s of styles) {
    const c = s.css;
    const prev = existingById.get(s.id);
    const st = prev ?? doc.createElementNS(NS.style, "style:style");
    if (!prev) {
      st.setAttributeNS(NS.style, "style:name", s.id);
      st.setAttributeNS(NS.style, "style:display-name", s.name);
      st.setAttributeNS(NS.style, "style:family", s.kind === "paragraph" ? "paragraph" : "text");
    }
    if (s.kind === "paragraph") {
      const a = ODF_ALIGN[c["text-align"] ?? ""];
      const align = a && a !== "left" ? (a === "right" ? "end" : a === "center" ? "center" : "justify") : undefined;
      const pBorders = (["top", "right", "bottom", "left"] as const).map((side) => ({ side, b: parseCssBorder(c[`border-${side}`]) })).filter((x) => x.b);
      const hasPara = !!(align || c["margin-left"] || c["margin-top"] || c["margin-bottom"] || c["line-height"] || c["background-color"] || pBorders.length || c["--rdoc-tabstops"]);
      let pp = directChild(st, "style:paragraph-properties");
      if (pp || hasPara) {
        if (!pp) { pp = doc.createElementNS(NS.style, "style:paragraph-properties"); st.insertBefore(pp, st.firstChild); }
        for (const attr of PARA_FO) pp.removeAttributeNS(NS.fo, attr); // re-derive only the modeled attrs
        if (align) pp.setAttributeNS(NS.fo, "fo:text-align", align);
        if (c["margin-left"]) pp.setAttributeNS(NS.fo, "fo:margin-left", pxToCm(parseFloat(c["margin-left"])));
        if (c["margin-top"]) pp.setAttributeNS(NS.fo, "fo:margin-top", pxToCm(parseFloat(c["margin-top"])));
        if (c["margin-bottom"]) pp.setAttributeNS(NS.fo, "fo:margin-bottom", pxToCm(parseFloat(c["margin-bottom"])));
        if (c["line-height"]) pp.setAttributeNS(NS.fo, "fo:line-height", `${Math.round(parseFloat(c["line-height"]) * 100)}%`);
        if (c["background-color"]) pp.setAttributeNS(NS.fo, "fo:background-color", c["background-color"]);
        for (const { side, b } of pBorders) pp.setAttributeNS(NS.fo, `fo:border-${side}`, `${pxToCm(b!.px)} ${b!.style} #${b!.hex.toLowerCase()}`);
        // Tab stops as part of the style: re-derive style:tab-stops from the css JSON.
        for (const ch of Array.from(pp.children)) if (ch.tagName === "style:tab-stops") pp.removeChild(ch);
        let styleStops: { pos: number; val?: string; leader?: string }[] | undefined;
        try { const j = JSON.parse(c["--rdoc-tabstops"] ?? "[]"); if (Array.isArray(j) && j.length) styleStops = j; } catch { /* skip */ }
        const tsEl = buildOdtTabStops(doc, styleStops);
        if (tsEl) pp.appendChild(tsEl);
        if (!pp.attributes.length && !pp.children.length) st.removeChild(pp);
      }
    }
    {
      let tp = directChild(st, "style:text-properties");
      if (!tp) { tp = doc.createElementNS(NS.style, "style:text-properties"); st.appendChild(tp); }
      for (const attr of RUN_FO) tp.removeAttributeNS(NS.fo, attr);
      for (const attr of RUN_STYLE) tp.removeAttributeNS(NS.style, attr);
      if (/bold|[6-9]00/.test(c["font-weight"] ?? "")) tp.setAttributeNS(NS.fo, "fo:font-weight", "bold");
      if (c["font-style"] === "italic") tp.setAttributeNS(NS.fo, "fo:font-style", "italic");
      if (/underline/.test(c["text-decoration"] ?? "")) {
        tp.setAttributeNS(NS.style, "style:text-underline-style", "solid");
        tp.setAttributeNS(NS.style, "style:text-underline-width", "auto");
        tp.setAttributeNS(NS.style, "style:text-underline-color", "font-color");
      }
      if (/line-through/.test(c["text-decoration"] ?? "")) {
        tp.setAttributeNS(NS.style, "style:text-line-through-style", "solid");
        tp.setAttributeNS(NS.style, "style:text-line-through-type", "single");
      }
      if (c["color"]) tp.setAttributeNS(NS.fo, "fo:color", c["color"]);
      if (c["font-size"]) tp.setAttributeNS(NS.fo, "fo:font-size", c["font-size"]);
      if (s.kind === "character" && c["background-color"]) tp.setAttributeNS(NS.fo, "fo:background-color", c["background-color"]);
      const font = (c["font-family"] ?? "").replace(/['"]/g, "").split(",")[0]?.trim();
      if (font) {
        tp.setAttributeNS(NS.fo, "fo:font-family", font);
        tp.setAttributeNS(NS.style, "style:font-name", font);
      }
      if (!tp.attributes.length) st.removeChild(tp);
    }
    if (!prev) officeStyles.appendChild(st);
  }
  files[key] = strToU8(new XMLSerializer().serializeToString(doc));
}

export function htmlToOdt(
  html: string,
  original: Uint8Array,
  opts?: { done?: Map<string, boolean>; parts?: { path: string; html: string }[]; page?: PageGeometry; newStyles?: NewStyle[]; notes?: Note[] },
): Uint8Array {
  const files = unzipSync(original);
  const content = files["content.xml"];
  if (!content) throw new Error("not an .odt: content.xml missing");
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) throw new Error("not an .odt: office:text missing");

  while (body.firstChild) body.removeChild(body.firstChild);
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  // Gather comment metadata + which ids wrap a range, before serializing.
  const refMeta = new Map<string, RefMeta>();
  for (const ref of Array.from(htmlDoc.querySelectorAll(".docx-comment-ref"))) {
    const id = ref.getAttribute("data-comment-id") ?? "";
    if (!id) continue;
    refMeta.set(id, {
      author: ref.getAttribute("data-comment-author") ?? "",
      date: ref.getAttribute("data-comment-date") ?? "",
      text: ref.getAttribute("data-comment-text") ?? "",
      resolved: ref.getAttribute("data-comment-resolved") === "1",
      paraId: ref.getAttribute("data-comment-paraid") ?? id,
    });
  }
  const rangedIds = new Set(
    Array.from(htmlDoc.querySelectorAll(".docx-comment[data-comment-id]")).map((s) => s.getAttribute("data-comment-id") ?? ""),
  );
  const ctx: OdfCtx = { doc, auto: ensureAutoStyles(doc), created: new Map(), files, pics: [], objs: [], refMeta, rangedIds, done: opts?.done ?? new Map(), changes: [], notesById: new Map((opts?.notes ?? []).map((n) => [n.id, n])) };
  for (const node of Array.from(htmlDoc.body.childNodes)) {
    const block = htmlBlockToOdf(node, ctx);
    if (block) body.appendChild(block);
  }
  if (!body.firstChild) body.appendChild(doc.createElementNS(NS.text, "text:p"));
  const tc = buildTrackedChanges(ctx); // tracked-changes region goes first in office:text
  if (tc) body.insertBefore(tc, body.firstChild);
  addManifestEntries(files, ctx.pics); // register any images embedded above
  addObjectManifestEntries(files, ctx.objs); // register any formula objects embedded above
  if (opts?.parts) applyHeaderFooter(files, opts.parts); // header/footer -> styles.xml
  reconcileSectionBands(files, htmlDoc); // drop a relinked section master's header/footer
  if (opts?.page) applyPageMargins(files, opts.page); // margins -> styles.xml page-layout
  applySectionMasters(files, htmlDoc); // per-section page setup -> styles.xml master pages
  if (opts?.newStyles?.length) addOdtStyles(files, opts.newStyles); // authored styles -> styles.xml

  const out = new XMLSerializer().serializeToString(doc);
  // Re-zip. ODF requires the "mimetype" entry first and stored (uncompressed).
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
  if (files["mimetype"]) repacked["mimetype"] = [files["mimetype"], { level: 0 }];
  for (const [name, data] of Object.entries(files)) {
    if (name === "mimetype") continue;
    repacked[name] = name === "content.xml" ? strToU8(out) : data;
  }
  return zipSync(repacked as Record<string, Uint8Array>);
}

