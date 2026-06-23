import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createRichEditor } from "../../core/editor";
import { firstFontFamily, fontSizeToHalfPt, toHex6 } from "../../core/util";
import type { Adapter, EditorOptions, RichDoc, RichEditor } from "../../core/types";

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

function inlineToHtml(el: Element, styles: Map<string, Fmt>): string {
  let html = "";
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === 3) {
      html += escapeHtml(node.textContent ?? "");
      continue;
    }
    if (node.nodeType !== 1) continue;
    const child = node as Element;
    switch (child.tagName) {
      case "text:span": {
        const f = styles.get(child.getAttribute("text:style-name") ?? "") ?? FMT0;
        html += wrapFmt(inlineToHtml(child, styles), f);
        break;
      }
      case "text:a": {
        const href = child.getAttribute("xlink:href") ?? "";
        html += `<a href="${escapeHtml(href)}">${inlineToHtml(child, styles)}</a>`;
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
        // Unmodelled inline content (annotations, bookmarks, notes, change marks,
        // inline frames, ...) is preserved verbatim.
        html += inlinePass(child);
    }
  }
  return html;
}

function listToHtml(el: Element, styles: Map<string, Fmt>): string {
  let items = "";
  for (const li of Array.from(el.children)) {
    if (li.tagName !== "text:list-item") continue;
    let inner = "";
    for (const block of Array.from(li.children)) {
      if (block.tagName === "text:list") inner += listToHtml(block, styles);
      else inner += inlineToHtml(block, styles);
    }
    items += `<li>${inner || "<br>"}</li>`;
  }
  return `<ul>${items}</ul>`;
}

function blockToHtml(el: Element, styles: Map<string, Fmt>, paras: Map<string, PFmt>): string {
  const alignAttr = (): string => {
    const a = paras.get(el.getAttribute("text:style-name") ?? "")?.align;
    return a && a !== "left" ? ` style="text-align:${a}"` : "";
  };
  switch (el.tagName) {
    case "text:h": {
      const lvl = Math.min(3, Math.max(1, parseInt(el.getAttribute("text:outline-level") ?? "1", 10) || 1));
      const inner = inlineToHtml(el, styles);
      return `<h${lvl}${alignAttr()}>${inner || "<br>"}</h${lvl}>`;
    }
    case "text:list":
      return listToHtml(el, styles);
    case "text:p": {
      const inner = inlineToHtml(el, styles);
      return `<p${alignAttr()}>${inner || "<br>"}</p>`;
    }
    default:
      // Tables, frames, tracked-changes, sequence-decls, sections, ... preserved verbatim.
      return blockPass(el);
  }
}

/** Convert an .odt's body to HTML. Returns "" if there is no editable text body. */
export function odtToHtml(bytes: Uint8Array): string {
  const files = unzipSync(bytes);
  const content = files["content.xml"];
  if (!content) return "";
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) return "";
  const styles = collectTextStyles(doc);
  const paras = collectParaStyles(doc);
  let html = "";
  for (const block of Array.from(body.children)) html += blockToHtml(block, styles, paras);
  return html || "<p><br></p>";
}

// ---------------------------------------------------------------------------
// HTML -> .odt
// ---------------------------------------------------------------------------

/** Ensure <office:automatic-styles> exists, returning it. */
function ensureAutoStyles(doc: Document): Element {
  let auto = doc.getElementsByTagName("office:automatic-styles")[0];
  if (auto) return auto;
  auto = doc.createElementNS(NS.office, "office:automatic-styles");
  const root = doc.documentElement;
  const body = doc.getElementsByTagName("office:body")[0];
  root.insertBefore(auto, body ?? null);
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

interface OdfCtx {
  doc: Document;
  auto: Element;
  created: Map<string, string>;
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

/** Rebuild an .odt from edited HTML, preserving every other part of the archive. */
export function htmlToOdt(html: string, original: Uint8Array): Uint8Array {
  const files = unzipSync(original);
  const content = files["content.xml"];
  if (!content) throw new Error("not an .odt: content.xml missing");
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) throw new Error("not an .odt: office:text missing");

  while (body.firstChild) body.removeChild(body.firstChild);
  const ctx: OdfCtx = { doc, auto: ensureAutoStyles(doc), created: new Map() };
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  for (const node of Array.from(htmlDoc.body.childNodes)) {
    const block = htmlBlockToOdf(node, ctx);
    if (block) body.appendChild(block);
  }
  if (!body.firstChild) body.appendChild(doc.createElementNS(NS.text, "text:p"));

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
      let body = "<p><br></p>";
      try {
        body = odtToHtml(bytes) || "<p><br></p>";
      } catch (e) {
        console.warn("odtedit: failed to parse document", e);
      }
      return { body, header: "", footer: "", comments: [] };
    },
    write(bodyHtml: string): Uint8Array {
      return htmlToOdt(bodyHtml, original);
    },
    newCommentMarkers(): never {
      // Comments are capability-gated off for odt, so the engine never calls this.
      throw new Error("odt: comments are not supported yet");
    },
    capabilities: {
      comments: false,
      trackChanges: false,
      images: false,
      headerFooter: false,
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
