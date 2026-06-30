// Shared odt (OpenDocument) primitives used by both halves of the adapter: namespaces,
// the run/paragraph formatting model, escaping, and passthrough (preserve-as-XML) helpers.

export const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
  dc: "http://purl.org/dc/elements/1.1/",
  loext: "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
  math: "http://www.w3.org/1998/Math/MathML",
} as const;

export interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
  strike?: boolean; // style:text-line-through-style
  vertAlign?: "super" | "sub"; // style:text-position
  color?: string; // 6-hex, no leading #
  bg?: string; // highlight, 6-hex
  font?: string; // family name
  sizePt?: number; // font size in points
  cStyle?: string; // named character style id (text:style-name)
}
export const FMT0: Fmt = { b: false, i: false, u: false };
export const fmtKey = (f: Fmt): string =>
  [f.b ? "b" : "", f.i ? "i" : "", f.u ? "u" : "", f.strike ? "k" : "", f.vertAlign ? "v" + f.vertAlign : "", f.color ? "c" + f.color : "", f.bg ? "g" + f.bg : "", f.font ? "f" + f.font : "", f.sizePt ? "s" + f.sizePt : ""].join("");

/** Paragraph-level formatting (text alignment), kept separate from run formatting. */
export interface PFmt {
  align?: string; // left | right | center | justify
  indentPx?: number; // fo:margin-left, in px
  lineHeight?: number; // fo:line-height (% form), as a multiple
  spaceBeforePx?: number; // fo:margin-top, in px
  spaceAfterPx?: number; // fo:margin-bottom, in px
  shading?: string; // fo:background-color (paragraph shading), as CSS "#rrggbb"
}
export const ODF_ALIGN: Record<string, string> = { start: "left", left: "left", end: "right", right: "right", center: "center", justify: "justify" };

export const IMG_MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff",
};

export const escapeHtml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const escapeAttr = (s: string): string => escapeHtml(s).replace(/"/g, "&quot;");

// Passthrough: anything we don't model (tables, frames/images, comments, tracked changes,
// fields, ...) is kept verbatim so editing the text never drops it. The original XML is
// stored, namespace-decorated, in a data-odt-xml attribute and re-emitted on save.
export const XMLNS = "http://www.w3.org/2000/xmlns/";
export const NS_DECLS: Record<string, string> = {
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
export function serializePassthrough(node: Element): string {
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
export const passthroughAttr = (node: Element): string => ` data-odt-xml="${escapeAttr(serializePassthrough(node))}"`;
// Comments/notes carry out-of-band text; preserve them but don't show their text inline.
export const HIDDEN_PASS = new Set(["office:annotation", "office:annotation-end", "text:note"]);
export const inlinePass = (el: Element): string => {
  const txt = HIDDEN_PASS.has(el.tagName) ? "" : escapeHtml(el.textContent ?? "");
  return `<span class="docx-pass" contenteditable="false"${passthroughAttr(el)}>${txt}</span>`;
};
export const blockPass = (el: Element): string => `<div class="docx-pass-block" contenteditable="false"${passthroughAttr(el)}>${escapeHtml(el.textContent ?? "")}</div>`;
export function importPassthrough(doc: Document, xml: string): Element | null {
  try {
    const frag = new DOMParser().parseFromString(xml, "application/xml");
    if (frag.getElementsByTagName("parsererror").length || !frag.documentElement) return null;
    return doc.importNode(frag.documentElement, true) as Element;
  } catch {
    return null;
  }
}

