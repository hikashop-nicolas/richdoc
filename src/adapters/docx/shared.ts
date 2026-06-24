// Shared docx (OOXML) primitives used by both the read and write halves of the adapter:
// namespaces, the run-formatting model, and colour/alignment lookup tables.

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
export const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const XMLNS = "http://www.w3.org/2000/xmlns/";

// Namespace declarations injected onto a passthrough fragment so it parses standalone.
export const NS_DECLS: Record<string, string> = {
  "xmlns:w": W,
  "xmlns:r": R,
  "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "xmlns:v": "urn:schemas-microsoft-com:vml",
  "xmlns:o": "urn:schemas-microsoft-com:office:office",
  "xmlns:w10": "urn:schemas-microsoft-com:office:word",
  "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
  "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
  "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
};

export const IMG_MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", tif: "image/tiff", tiff: "image/tiff", svg: "image/svg+xml", webp: "image/webp",
};

export interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  vertAlign?: "super" | "sub"; // w:vertAlign (superscript / subscript)
  color?: string; // 6-hex, no leading '#'
  highlight?: string; // OOXML named highlight (e.g. "yellow")
  shading?: string; // 6-hex arbitrary background (w:shd fill)
  sizeHalfPt?: number; // w:sz value, in half-points
  font?: string; // ascii font family
}
export const FMT0: Fmt = { b: false, i: false, u: false, strike: false };
export const escapeHtml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const escapeAttr = (s: string): string => escapeHtml(s).replace(/"/g, "&quot;");

// OOXML named highlight colours -> CSS, and the reverse (CSS hex -> name) for round-trip.
export const HL_CSS: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff",
  blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080",
  darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000",
  darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};
export const HL_BY_HEX = new Map<string, string>(Object.entries(HL_CSS).map(([n, c]) => [c.slice(1).toUpperCase(), n]));

/** Normalise a CSS colour (#rgb, #rrggbb, rgb()/rgba()) to 6 upper-hex, or undefined. */
// w:jc value -> CSS text-align, and the reverse (CSS text-align -> w:jc) for round-trip.
export const JC_TO_ALIGN: Record<string, string> = { both: "justify", distribute: "justify", center: "center", right: "right", end: "right", left: "left", start: "left" };
export const JC_BY_ALIGN: Record<string, string> = { justify: "both", center: "center", right: "right", left: "left" };
