import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createRichEditor } from "../../core/editor";
import { bytesToBase64, firstFontFamily, fontSizeToHalfPt, toHex6 } from "../../core/util";
import type { Adapter, CommentEdits, CommentMarkers, CommentThread, EditorOptions, NewCommentMeta, RichDoc, RichEditor } from "../../core/types";

// odtedit: a standalone, framework-agnostic, client-side OpenDocument Text (.odt) editor.
//
// An .odt file is a zip of OpenDocument XML; the document body lives in content.xml. We
// convert that body to HTML, edit it in a contenteditable rich-text surface, and on export
// rebuild content.xml from the edited HTML and re-zip, preserving every other part of the
// archive (styles, images, metadata) byte-for-byte.
//
// Honest scope: this edits text and common inline/block formatting (bold, italic,
// underline, headings, lists, links). Advanced layout (tables, frames, columns, exact
// paragraph styles) is preserved in the file's other parts but is not re-rendered through
// the editor, so editing a paragraph normalizes it to the supported subset.

const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
  dc: "http://purl.org/dc/elements/1.1/",
  loext: "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
} as const;

interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
  color?: string; // 6-hex, no leading #
  bg?: string; // highlight, 6-hex
  font?: string; // family name
  sizePt?: number; // font size in points
}
const FMT0: Fmt = { b: false, i: false, u: false };
const fmtKey = (f: Fmt): string =>
  [f.b ? "b" : "", f.i ? "i" : "", f.u ? "u" : "", f.color ? "c" + f.color : "", f.bg ? "g" + f.bg : "", f.font ? "f" + f.font : "", f.sizePt ? "s" + f.sizePt : ""].join("");

/** Paragraph-level formatting (text alignment), kept separate from run formatting. */
interface PFmt {
  align?: string; // left | right | center | justify
}
const ODF_ALIGN: Record<string, string> = { start: "left", left: "left", end: "right", right: "right", center: "center", justify: "justify" };

const escapeHtml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (s: string): string => escapeHtml(s).replace(/"/g, "&quot;");

// Passthrough: anything we don't model (tables, frames/images, comments, tracked changes,
// fields, ...) is kept verbatim so editing the text never drops it. The original XML is
// stored, namespace-decorated, in a data-odt-xml attribute and re-emitted on save.
const XMLNS = "http://www.w3.org/2000/xmlns/";
const NS_DECLS: Record<string, string> = {
  "xmlns:office": NS.office,
  "xmlns:text": NS.text,
  "xmlns:style": NS.style,
  "xmlns:fo": NS.fo,
  "xmlns:xlink": NS.xlink,
  "xmlns:table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  "xmlns:draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  "xmlns:svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  "xmlns:dc": "http://purl.org/dc/elements/1.1/",
  "xmlns:meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  "xmlns:math": "http://www.w3.org/1998/Math/MathML",
  "xmlns:loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
  "xmlns:officeooo": "http://openoffice.org/2009/office",
};
function serializePassthrough(node: Element): string {
  const clone = node.cloneNode(true) as Element;
  for (const [k, v] of Object.entries(NS_DECLS)) {
    try {
      clone.setAttributeNS(XMLNS, k, v);
    } catch {
      /* declaration already present */
    }
  }
  return new XMLSerializer().serializeToString(clone);
}
const passthroughAttr = (node: Element): string => ` data-odt-xml="${escapeAttr(serializePassthrough(node))}"`;
// Comments/notes carry out-of-band text; preserve them but don't show their text inline.
const HIDDEN_PASS = new Set(["office:annotation", "office:annotation-end", "text:note"]);
const inlinePass = (el: Element): string => {
  const txt = HIDDEN_PASS.has(el.tagName) ? "" : escapeHtml(el.textContent ?? "");
  return `<span class="docx-pass" contenteditable="false"${passthroughAttr(el)}>${txt}</span>`;
};
const blockPass = (el: Element): string => `<div class="docx-pass-block" contenteditable="false"${passthroughAttr(el)}>${escapeHtml(el.textContent ?? "")}</div>`;
function importPassthrough(doc: Document, xml: string): Element | null {
  try {
    const frag = new DOMParser().parseFromString(xml, "application/xml");
    if (frag.getElementsByTagName("parsererror").length || !frag.documentElement) return null;
    return doc.importNode(frag.documentElement, true) as Element;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// .odt -> HTML
// ---------------------------------------------------------------------------

const hex6 = (v: string | null | undefined): string | undefined => {
  if (!v || v === "transparent") return undefined;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(v.trim());
  return m ? m[1]!.toUpperCase() : undefined;
};
const sizePt = (v: string | null | undefined): number | undefined => {
  if (!v) return undefined;
  const m = /^([\d.]+)pt$/.exec(v.trim());
  return m ? Math.round(parseFloat(m[1]!) * 10) / 10 : undefined;
};

const IMG_MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff",
};
const PX_PER = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, px: 1 };
/** ODF length (e.g. "5.2cm", "48pt") to CSS px. */
function lenToPx(v: string | null | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^([\d.]+)(cm|mm|in|pt|px)$/.exec(v.trim());
  return m ? parseFloat(m[1]!) * PX_PER[m[2] as keyof typeof PX_PER] : undefined;
}
const pxToCm = (px: number): string => `${Math.round((px / (96 / 2.54)) * 1000) / 1000}cm`;

/** Read context: the archive (for image data), the resolved style maps, and comment state. */
interface ChangeInfo {
  type: "insertion" | "deletion";
  author: string;
  date: string;
  deleted: string; // for deletions: the removed text
}
interface RCtx {
  files: Record<string, Uint8Array>;
  styles: Map<string, Fmt>;
  paras: Map<string, PFmt>;
  threads: CommentThread[]; // comments collected while rendering, for the panel
  rangedNames: Set<string>; // annotation names that have a matching annotation-end
  openComment: Set<string>; // comment ranges currently open (reopened per paragraph)
  changes: Map<string, ChangeInfo>; // tracked-change id -> metadata
  openIns: Set<string>; // insertion ranges currently open (reopened per paragraph)
}

/** Parse text:tracked-changes into a map of change id -> metadata. */
function readChanges(body: Element): Map<string, ChangeInfo> {
  const map = new Map<string, ChangeInfo>();
  const tc = body.getElementsByTagName("text:tracked-changes")[0];
  if (!tc) return map;
  for (const region of Array.from(tc.getElementsByTagName("text:changed-region"))) {
    const id = region.getAttribute("text:id") ?? region.getAttribute("xml:id") ?? "";
    if (!id) continue;
    const ins = region.getElementsByTagName("text:insertion")[0];
    const del = region.getElementsByTagName("text:deletion")[0];
    const info = (ins ?? del)?.getElementsByTagName("office:change-info")[0];
    const author = info?.getElementsByTagName("dc:creator")[0]?.textContent ?? "";
    const date = info?.getElementsByTagName("dc:date")[0]?.textContent ?? "";
    if (ins) map.set(id, { type: "insertion", author, date, deleted: "" });
    else if (del) {
      const deleted = Array.from(del.getElementsByTagName("text:p")).map((p) => p.textContent ?? "").join("\n");
      map.set(id, { type: "deletion", author, date, deleted });
    }
  }
  return map;
}
const revAttrs = (c: ChangeInfo | undefined): string =>
  ` data-author="${escapeAttr(c?.author ?? "")}" data-date="${escapeAttr(c?.date ?? "")}"`;

/** Read a comment's author/date/text from an office:annotation element. */
function readAnnotation(an: Element): { author: string; date: string; text: string; resolved: boolean } {
  const author = an.getElementsByTagName("dc:creator")[0]?.textContent ?? "";
  const date = an.getElementsByTagName("dc:date")[0]?.textContent ?? "";
  const text = Array.from(an.getElementsByTagName("text:p"))
    .map((p) => p.textContent ?? "")
    .join("\n")
    .trim();
  const resolved = an.getAttribute("loext:resolved") === "true" || an.getAttribute("officeooo:resolved") === "true";
  return { author, date, text, resolved };
}

/** The clickable reference marker for a comment, carrying its metadata for save + panel. */
function commentRef(name: string, m: { author: string; date: string; text: string; resolved: boolean }): string {
  const meta = m.author ? `${m.author}${m.date ? " – " + m.date.slice(0, 10) : ""}` : "";
  return (
    `<span class="docx-comment-ref" contenteditable="false" data-comment-id="${escapeAttr(name)}" data-comment-paraid="${escapeAttr(name)}"` +
    ` data-comment-author="${escapeAttr(m.author)}" data-comment-date="${escapeAttr(m.date)}" data-comment-text="${escapeAttr(m.text)}"` +
    `${m.resolved ? ' data-comment-resolved="1"' : ""} data-comment-meta="${escapeAttr(meta)}"` +
    ` title="${escapeAttr(meta ? meta + ": " + m.text : m.text)}">\u{1F4AC}</span>`
  );
}

/** A draw:frame holding a draw:image -> an <img> with a data URL; otherwise passthrough. */
function imageHtml(frame: Element, ctx: RCtx): string {
  const imgEl = frame.getElementsByTagName("draw:image")[0];
  const href = (imgEl?.getAttribute("xlink:href") ?? "").replace(/^\.\//, "");
  const bytes = href ? ctx.files[href] : undefined;
  if (!bytes) return inlinePass(frame); // can't resolve (linked/object/unknown): keep verbatim
  const ext = (href.split(".").pop() ?? "png").toLowerCase();
  const mime = IMG_MIME[ext] ?? "image/png";
  const w = lenToPx(frame.getAttribute("svg:width"));
  const h = lenToPx(frame.getAttribute("svg:height"));
  const dims = (w ? ` width="${Math.round(w)}"` : "") + (h ? ` height="${Math.round(h)}"` : "");
  return `<img src="data:${mime};base64,${bytesToBase64(bytes)}" alt=""${dims}>`;
}

/** Map text:style-name -> run formatting, read from the automatic/text styles. */
function collectTextStyles(doc: Document): Map<string, Fmt> {
  const map = new Map<string, Fmt>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "text") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const tp = st.getElementsByTagName("style:text-properties")[0];
    map.set(name, {
      b: tp?.getAttribute("fo:font-weight") === "bold",
      i: tp?.getAttribute("fo:font-style") === "italic",
      u: !!tp && tp.getAttribute("style:text-underline-style") != null && tp.getAttribute("style:text-underline-style") !== "none",
      color: hex6(tp?.getAttribute("fo:color")),
      bg: hex6(tp?.getAttribute("fo:background-color")),
      font: tp?.getAttribute("fo:font-family")?.replace(/^['"]|['"]$/g, "") || tp?.getAttribute("style:font-name") || undefined,
      sizePt: sizePt(tp?.getAttribute("fo:font-size")),
    });
  }
  return map;
}

/** Map paragraph style-name -> alignment, read from the paragraph styles. */
function collectParaStyles(doc: Document): Map<string, PFmt> {
  const map = new Map<string, PFmt>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "paragraph") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const pp = st.getElementsByTagName("style:paragraph-properties")[0];
    const align = ODF_ALIGN[pp?.getAttribute("fo:text-align") ?? ""];
    if (align) map.set(name, { align });
  }
  return map;
}

const wrapFmt = (inner: string, f: Fmt): string => {
  let s = inner;
  const css: string[] = [];
  if (f.color) css.push(`color:#${f.color}`);
  if (f.bg) css.push(`background-color:#${f.bg}`);
  if (f.font) css.push(`font-family:'${f.font.replace(/'/g, "")}'`);
  if (f.sizePt) css.push(`font-size:${f.sizePt}pt`);
  if (css.length) s = `<span style="${css.join(";")}">${s}</span>`;
  if (f.u) s = `<u>${s}</u>`;
  if (f.i) s = `<em>${s}</em>`;
  if (f.b) s = `<strong>${s}</strong>`;
  return s;
};

function inlineToHtml(el: Element, ctx: RCtx): string {
  let html = "";
  for (const id of ctx.openIns) html += `<ins class="docx-ins"${revAttrs(ctx.changes.get(id))}>`;
  for (const id of ctx.openComment) html += `<span class="docx-comment" data-comment-id="${escapeAttr(id)}">`;
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === 3) {
      html += escapeHtml(node.textContent ?? "");
      continue;
    }
    if (node.nodeType !== 1) continue;
    const child = node as Element;
    switch (child.tagName) {
      case "text:span": {
        const f = ctx.styles.get(child.getAttribute("text:style-name") ?? "") ?? FMT0;
        html += wrapFmt(inlineToHtml(child, ctx), f);
        break;
      }
      case "text:a": {
        const href = child.getAttribute("xlink:href") ?? "";
        html += `<a href="${escapeHtml(href)}">${inlineToHtml(child, ctx)}</a>`;
        break;
      }
      case "draw:frame":
        html += imageHtml(child, ctx);
        break;
      case "office:annotation": {
        const name = child.getAttribute("office:name") ?? `c${ctx.threads.length}`;
        const m = readAnnotation(child);
        ctx.threads.push({ id: name, author: m.author, date: m.date, text: m.text, reactions: [], paraId: name, replies: [], resolved: m.resolved });
        html += commentRef(name, m);
        if (ctx.rangedNames.has(name)) {
          ctx.openComment.add(name);
          html += `<span class="docx-comment" data-comment-id="${escapeAttr(name)}">`;
        }
        break;
      }
      case "office:annotation-end": {
        const name = child.getAttribute("office:name") ?? "";
        if (ctx.openComment.delete(name)) html += "</span>";
        break;
      }
      case "text:change-start": {
        const id = child.getAttribute("text:change-id") ?? "";
        if (ctx.changes.get(id)?.type === "insertion") {
          ctx.openIns.add(id);
          html += `<ins class="docx-ins"${revAttrs(ctx.changes.get(id))}>`;
        }
        break;
      }
      case "text:change-end": {
        const id = child.getAttribute("text:change-id") ?? "";
        if (ctx.openIns.delete(id)) html += "</ins>";
        break;
      }
      case "text:change": {
        // a deletion point: show the removed text struck through, from the changed-region
        const id = child.getAttribute("text:change-id") ?? "";
        const c = ctx.changes.get(id);
        if (c?.type === "deletion") html += `<del class="docx-del"${revAttrs(c)}>${escapeHtml(c.deleted)}</del>`;
        break;
      }
      case "text:line-break":
        html += "<br>";
        break;
      case "text:tab":
        html += "    ";
        break;
      case "text:s": {
        const n = parseInt(child.getAttribute("text:c") ?? "1", 10) || 1;
        html += n > 1 ? " ".repeat(n) : " ";
        break;
      }
      default:
        // Unmodelled inline content (bookmarks, notes, change marks, ...) preserved verbatim.
        html += inlinePass(child);
    }
  }
  for (let i = 0; i < ctx.openComment.size; i++) html += "</span>"; // reopened in the next paragraph
  for (let i = 0; i < ctx.openIns.size; i++) html += "</ins>"; // reopened in the next paragraph
  return html;
}

function listToHtml(el: Element, ctx: RCtx): string {
  let items = "";
  for (const li of Array.from(el.children)) {
    if (li.tagName !== "text:list-item") continue;
    let inner = "";
    for (const block of Array.from(li.children)) {
      if (block.tagName === "text:list") inner += listToHtml(block, ctx);
      else inner += inlineToHtml(block, ctx);
    }
    items += `<li>${inner || "<br>"}</li>`;
  }
  return `<ul>${items}</ul>`;
}

function blockToHtml(el: Element, ctx: RCtx): string {
  const alignAttr = (): string => {
    const a = ctx.paras.get(el.getAttribute("text:style-name") ?? "")?.align;
    return a && a !== "left" ? ` style="text-align:${a}"` : "";
  };
  switch (el.tagName) {
    case "text:h": {
      const lvl = Math.min(3, Math.max(1, parseInt(el.getAttribute("text:outline-level") ?? "1", 10) || 1));
      const inner = inlineToHtml(el, ctx);
      return `<h${lvl}${alignAttr()}>${inner || "<br>"}</h${lvl}>`;
    }
    case "text:list":
      return listToHtml(el, ctx);
    case "text:p": {
      const inner = inlineToHtml(el, ctx);
      return `<p${alignAttr()}>${inner || "<br>"}</p>`;
    }
    default:
      // Tables, tracked-changes, sequence-decls, sections, ... preserved verbatim.
      return blockPass(el);
  }
}

/** Header/footer HTML, read from the master page in styles.xml. */
function readHeaderFooter(files: Record<string, Uint8Array>): { header: string; footer: string } {
  const raw = files["styles.xml"];
  if (!raw) return { header: "", footer: "" };
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const master = doc.getElementsByTagName("style:master-page")[0];
  if (!master) return { header: "", footer: "" };
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), threads: [], rangedNames: new Set(), openComment: new Set(), changes: new Map(), openIns: new Set() };
  const render = (tag: string): string => {
    const el = master.getElementsByTagName(tag)[0];
    if (!el) return "";
    let html = "";
    for (const block of Array.from(el.children)) html += blockToHtml(block, ctx);
    return html;
  };
  return { header: render("style:header"), footer: render("style:footer") };
}

/** Parse an .odt into the editable body HTML, the comment threads, and header/footer. */
export function odtToParts(bytes: Uint8Array): { body: string; comments: CommentThread[]; header: string; footer: string } {
  const files = unzipSync(bytes);
  const { header, footer } = readHeaderFooter(files);
  const content = files["content.xml"];
  if (!content) return { body: "", comments: [], header, footer };
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) return { body: "", comments: [], header, footer };
  const rangedNames = new Set(
    Array.from(doc.getElementsByTagName("office:annotation-end"))
      .map((e) => e.getAttribute("office:name"))
      .filter((n): n is string => !!n),
  );
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), threads: [], rangedNames, openComment: new Set(), changes: readChanges(body), openIns: new Set() };
  let html = "";
  for (const block of Array.from(body.children)) {
    if (block.tagName === "text:tracked-changes") continue; // metadata, parsed into ctx.changes
    html += blockToHtml(block, ctx);
  }
  return { body: html || "<p><br></p>", comments: ctx.threads, header, footer };
}

/** Convert an .odt's body to HTML. Returns "" if there is no editable text body. */
export function odtToHtml(bytes: Uint8Array): string {
  return odtToParts(bytes).body;
}

// ---------------------------------------------------------------------------
// HTML -> .odt
// ---------------------------------------------------------------------------

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
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_t${created.size}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "text");
  const tp = doc.createElementNS(NS.style, "style:text-properties");
  if (f.b) tp.setAttributeNS(NS.fo, "fo:font-weight", "bold");
  if (f.i) tp.setAttributeNS(NS.fo, "fo:font-style", "italic");
  if (f.u) {
    tp.setAttributeNS(NS.style, "style:text-underline-style", "solid");
    tp.setAttributeNS(NS.style, "style:text-underline-width", "auto");
    tp.setAttributeNS(NS.style, "style:text-underline-color", "font-color");
  }
  if (f.color) tp.setAttributeNS(NS.fo, "fo:color", `#${f.color}`);
  if (f.bg) tp.setAttributeNS(NS.fo, "fo:background-color", `#${f.bg}`);
  if (f.font) {
    tp.setAttributeNS(NS.fo, "fo:font-family", f.font);
    tp.setAttributeNS(NS.style, "style:font-name", f.font);
  }
  if (f.sizePt) tp.setAttributeNS(NS.fo, "fo:font-size", `${f.sizePt}pt`);
  st.appendChild(tp);
  auto.appendChild(st);
  created.set(key, name);
  return name;
}

/** Create (once) a paragraph style for an alignment and return its name. */
function paraStyleFor(doc: Document, auto: Element, created: Map<string, string>, align: string): string | null {
  const a = ODF_ALIGN[align];
  if (!a || a === "left") return null;
  const key = `p_${a}`;
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_${key}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "paragraph");
  const pp = doc.createElementNS(NS.style, "style:paragraph-properties");
  pp.setAttributeNS(NS.fo, "fo:text-align", a === "right" ? "end" : a === "center" ? "center" : a === "justify" ? "justify" : "start");
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
  refMeta: Map<string, RefMeta>; // comment id -> metadata, gathered from the body refs
  rangedIds: Set<string>; // comment ids that wrap a text range (vs a point comment)
  done: Map<string, boolean>; // resolve state keyed by comment paraId
  changes: { id: string; type: "insertion" | "deletion"; author: string; date: string; deleted?: string }[]; // tracked changes to emit
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

/** Turn an <img> (data URL) into a draw:frame, embedding the bytes in the archive. */
function buildImageFrame(img: HTMLElement, ctx: OdfCtx): Element | null {
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

  const wPx = parseFloat(img.getAttribute("width") ?? "") || undefined;
  const hPx = parseFloat(img.getAttribute("height") ?? "") || undefined;
  const frame = ctx.doc.createElementNS(NS.draw, "draw:frame");
  frame.setAttributeNS(NS.draw, "draw:name", `Image${idx + 1}`);
  frame.setAttributeNS(NS.text, "text:anchor-type", "as-char");
  if (wPx) frame.setAttributeNS(NS.svg, "svg:width", pxToCm(wPx));
  if (hPx) frame.setAttributeNS(NS.svg, "svg:height", pxToCm(hPx));
  const image = ctx.doc.createElementNS(NS.draw, "draw:image");
  image.setAttributeNS(NS.xlink, "xlink:href", path);
  image.setAttributeNS(NS.xlink, "xlink:type", "simple");
  image.setAttributeNS(NS.xlink, "xlink:show", "embed");
  image.setAttributeNS(NS.xlink, "xlink:actuate", "onLoad");
  frame.appendChild(image);
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

/** Append the inline content of an HTML node to an ODF block element. */
function htmlInlineToOdf(node: Node, parent: Element, f: Fmt, ctx: OdfCtx): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      const txt = child.textContent ?? "";
      if (!txt) continue;
      if (!fmtKey(f)) {
        parent.appendChild(ctx.doc.createTextNode(txt));
      } else {
        const span = ctx.doc.createElementNS(NS.text, "text:span");
        const name = styleFor(ctx.doc, ctx.auto, ctx.created, f);
        if (name) span.setAttributeNS(NS.text, "text:style-name", name);
        span.appendChild(ctx.doc.createTextNode(txt));
        parent.appendChild(span);
      }
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    const stash = el.getAttribute("data-odt-xml");
    if (stash) {
      const node2 = importPassthrough(ctx.doc, stash);
      if (node2) parent.appendChild(node2);
      continue;
    }
    if (tag === "br") {
      parent.appendChild(ctx.doc.createElementNS(NS.text, "text:line-break"));
      continue;
    }
    if (tag === "img") {
      const frame = buildImageFrame(el, ctx);
      if (frame) parent.appendChild(frame);
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
      color: toHex6(el.style.color) ?? f.color,
      bg: toHex6(el.style.backgroundColor) ?? f.bg,
      font: firstFontFamily(el.style.fontFamily) ?? f.font,
      sizePt: hp ? hp / 2 : f.sizePt,
    };
    htmlInlineToOdf(el, parent, next, ctx);
  }
}

function htmlListToOdf(el: HTMLElement, ctx: OdfCtx): Element {
  const list = ctx.doc.createElementNS(NS.text, "text:list");
  for (const li of Array.from(el.children)) {
    if (li.tagName.toLowerCase() !== "li") continue;
    const item = ctx.doc.createElementNS(NS.text, "text:list-item");
    const nested = li.querySelector(":scope > ul, :scope > ol");
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    htmlInlineToOdf(li, p, FMT0, ctx);
    item.appendChild(p);
    if (nested) item.appendChild(htmlListToOdf(nested as HTMLElement, ctx));
    list.appendChild(item);
  }
  return list;
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
  const stash = el.getAttribute("data-odt-xml");
  if (stash) return importPassthrough(ctx.doc, stash);
  if (tag === "ul" || tag === "ol") return htmlListToOdf(el, ctx);
  const applyAlign = (block: Element): void => {
    const name = el.style.textAlign ? paraStyleFor(ctx.doc, ctx.auto, ctx.created, el.style.textAlign) : null;
    if (name) block.setAttributeNS(NS.text, "text:style-name", name);
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

/** Write edited header/footer HTML back into the master page in styles.xml. */
function applyHeaderFooter(files: Record<string, Uint8Array>, parts: { path: string; html: string }[]): void {
  const hf = parts.filter((p) => p.path === "header" || p.path === "footer");
  if (!hf.length || !files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const master = doc.getElementsByTagName("style:master-page")[0];
  if (!master) return;
  const ctx: OdfCtx = {
    doc,
    auto: ensureAutoStyles(doc, "office:master-styles"),
    created: new Map(),
    files,
    pics: [],
    refMeta: new Map(),
    rangedIds: new Set(),
    done: new Map(),
    changes: [],
  };
  for (const p of hf) {
    const tag = p.path === "header" ? "style:header" : "style:footer";
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

/** Rebuild an .odt from edited HTML, preserving every other part of the archive. */
export function htmlToOdt(
  html: string,
  original: Uint8Array,
  opts?: { done?: Map<string, boolean>; parts?: { path: string; html: string }[] },
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
  const ctx: OdfCtx = { doc, auto: ensureAutoStyles(doc), created: new Map(), files, pics: [], refMeta, rangedIds, done: opts?.done ?? new Map(), changes: [] };
  for (const node of Array.from(htmlDoc.body.childNodes)) {
    const block = htmlBlockToOdf(node, ctx);
    if (block) body.appendChild(block);
  }
  if (!body.firstChild) body.appendChild(doc.createElementNS(NS.text, "text:p"));
  const tc = buildTrackedChanges(ctx); // tracked-changes region goes first in office:text
  if (tc) body.insertBefore(tc, body.firstChild);
  addManifestEntries(files, ctx.pics); // register any images embedded above
  if (opts?.parts) applyHeaderFooter(files, opts.parts); // header/footer -> styles.xml

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

// ---------------------------------------------------------------------------
// odt adapter over the shared engine
// ---------------------------------------------------------------------------

export type OdtEditorOptions = EditorOptions;
export type OdtEditor = RichEditor;

/** Wrap a .odt byte array as an engine adapter: parse, serialize, capabilities. */
export function createOdtAdapter(bytes: Uint8Array): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      let parts = { body: "<p><br></p>", comments: [] as CommentThread[], header: "", footer: "" };
      try {
        const p = odtToParts(bytes);
        parts = { body: p.body || "<p><br></p>", comments: p.comments, header: p.header, footer: p.footer };
      } catch (e) {
        console.warn("odtedit: failed to parse document", e);
      }
      return {
        body: parts.body,
        header: parts.header,
        footer: parts.footer,
        headerPath: parts.header ? "header" : undefined,
        footerPath: parts.footer ? "footer" : undefined,
        comments: parts.comments,
      };
    },
    write(bodyHtml: string, parts: { path: string; html: string }[], edits: CommentEdits): Uint8Array {
      return htmlToOdt(bodyHtml, original, { done: edits.done, parts });
    },
    newCommentMarkers(meta: NewCommentMeta): CommentMarkers {
      const cmark = (): HTMLElement => {
        const s = document.createElement("span");
        s.className = "docx-cmark";
        s.contentEditable = "false";
        return s;
      };
      const ref = document.createElement("span");
      ref.className = "docx-comment-ref";
      ref.contentEditable = "false";
      ref.textContent = "\u{1F4AC}";
      ref.setAttribute("data-comment-id", meta.id);
      ref.setAttribute("data-comment-new", "1");
      ref.setAttribute("data-comment-paraid", meta.paraId);
      ref.setAttribute("data-comment-author", meta.author);
      if (meta.date) ref.setAttribute("data-comment-date", meta.date);
      ref.setAttribute("data-comment-text", meta.text);
      ref.setAttribute("data-comment-meta", meta.date ? `${meta.author} – ${meta.date.slice(0, 10)}` : meta.author);
      return { start: cmark(), end: cmark(), ref };
    },
    capabilities: {
      comments: true,
      commentReplies: false,
      commentReactions: false,
      trackChanges: true,
      images: true,
      headerFooter: true,
      pageBreak: false,
      textColor: true,
      fontControls: true,
      alignment: true,
    },
  };
}

/** Mount a .odt editor in `container`: the odt adapter driving the shared engine. */
export function createOdtEditor(container: HTMLElement, bytes: Uint8Array, options: OdtEditorOptions = {}): OdtEditor {
  return createRichEditor(container, createOdtAdapter(bytes), options);
}
