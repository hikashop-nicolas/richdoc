import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { t } from "./i18n";
import "./docxedit.css";

// docxedit: a standalone, framework-agnostic, client-side Office Open XML (.docx) editor.
//
// A .docx is a zip of OOXML; the document body lives in word/document.xml. We convert that
// body to HTML, edit it in a contenteditable rich-text surface, and on export rebuild
// word/document.xml from the edited HTML and re-zip, preserving every other part of the
// archive (styles, headers/footers, images, numbering) byte-for-byte.
//
// Honest scope: edits text and common formatting (bold, italic, underline, headings,
// links, line breaks; lists best-effort). The paragraph body is regenerated from the
// supported subset, so per-paragraph direct formatting in edited content normalizes; the
// rest of the document is untouched.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const XMLNS = "http://www.w3.org/2000/xmlns/";

// Namespace declarations injected onto a passthrough fragment so it parses standalone.
const NS_DECLS: Record<string, string> = {
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

const IMG_MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", tif: "image/tiff", tiff: "image/tiff", svg: "image/svg+xml", webp: "image/webp",
};

interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  color?: string; // 6-hex, no leading '#'
  highlight?: string; // OOXML named highlight (e.g. "yellow")
  shading?: string; // 6-hex arbitrary background (w:shd fill)
  sizeHalfPt?: number; // w:sz value, in half-points
  font?: string; // ascii font family
}
const FMT0: Fmt = { b: false, i: false, u: false, strike: false };
const escapeHtml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (s: string): string => escapeHtml(s).replace(/"/g, "&quot;");

// OOXML named highlight colours -> CSS, and the reverse (CSS hex -> name) for round-trip.
const HL_CSS: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff",
  blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080",
  darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000",
  darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};
const HL_BY_HEX = new Map<string, string>(Object.entries(HL_CSS).map(([n, c]) => [c.slice(1).toUpperCase(), n]));

/** Normalise a CSS colour (#rgb, #rrggbb, rgb()/rgba()) to 6 upper-hex, or undefined. */
function toHex6(c: string | undefined): string | undefined {
  if (!c) return undefined;
  const s = c.trim();
  let m = /^#?([0-9a-fA-F]{6})$/.exec(s);
  if (m) return m[1]!.toUpperCase();
  m = /^#?([0-9a-fA-F]{3})$/.exec(s);
  if (m) return m[1]!.split("").map((x) => x + x).join("").toUpperCase();
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (m) return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase();
  return undefined;
}

// w:jc value -> CSS text-align, and the reverse (CSS text-align -> w:jc) for round-trip.
const JC_TO_ALIGN: Record<string, string> = { both: "justify", distribute: "justify", center: "center", right: "right", end: "right", left: "left", start: "left" };
const JC_BY_ALIGN: Record<string, string> = { justify: "both", center: "center", right: "right", left: "left" };

function fontSizeToHalfPt(v: string | undefined): number | undefined {
  if (!v) return undefined;
  let m = /^([\d.]+)pt$/.exec(v.trim());
  if (m) return Math.round(parseFloat(m[1]!) * 2);
  m = /^([\d.]+)px$/.exec(v.trim());
  if (m) return Math.round(parseFloat(m[1]!) * 0.75 * 2);
  return undefined;
}
const firstFontFamily = (v: string | undefined): string | undefined =>
  v ? (v.split(",")[0] ?? "").trim().replace(/^['"]|['"]$/g, "") || undefined : undefined;

interface Reaction {
  emoji: string;
  people: string[];
}
interface Comment {
  author: string;
  date: string;
  text: string;
  reactions: Reaction[];
  lastParaId: string; // last w14:paraId, the key used by commentsExtended threading
}

// Google Docs exports emoji reactions as plain-text paragraphs inside the comment; detect
// and strip them so the panel can render an icon + count instead of raw text.
const REACTION_HEADER = /^\s*(nombre total de r[ée]actions|total reactions|r[ée]actions)\s*:/i;
const REACTION_LINE = /^\s*(.+?)\s+(?:a r[ée]agi avec|reacted with)\s+(\S+)/i;

function parseReactions(paras: string[]): { text: string; reactions: Reaction[] } {
  const kept: string[] = [];
  const byEmoji = new Map<string, string[]>();
  for (const p of paras) {
    if (REACTION_HEADER.test(p)) continue;
    const m = REACTION_LINE.exec(p);
    if (m) {
      const emoji = m[2]!;
      const list = byEmoji.get(emoji) ?? [];
      list.push(m[1]!.trim());
      byEmoji.set(emoji, list);
      continue;
    }
    kept.push(p);
  }
  return { text: kept.join("\n").trim(), reactions: Array.from(byEmoji, ([emoji, people]) => ({ emoji, people })) };
}

// Context shared by the read helpers: archive files plus the relationships and numbering
// for the part being rendered (the body or a header/footer).
interface RenderCtx {
  files: Record<string, Uint8Array>;
  rels: Map<string, string>;
  numbering: Map<string, boolean>;
  comments: Map<string, Comment>;
  openComments: Set<string>; // comment ids whose range spans across paragraphs (render state)
  commentOrder: string[]; // ids in document order, for the side panel
}

/** Read word/comments.xml into id -> comment (reactions parsed out, paraIds captured). */
function readComments(file: Uint8Array | undefined): Map<string, Comment> {
  const map = new Map<string, Comment>();
  if (!file) return map;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  for (const c of Array.from(doc.getElementsByTagName("w:comment"))) {
    const id = c.getAttribute("w:id");
    if (!id) continue;
    const paras: string[] = [];
    const paraIds: string[] = [];
    for (const p of Array.from(c.getElementsByTagName("w:p"))) {
      const pid = p.getAttribute("w14:paraId");
      if (pid) paraIds.push(pid);
      let s = "";
      for (const t of Array.from(p.getElementsByTagName("w:t"))) s += t.textContent ?? "";
      paras.push(s);
    }
    const { text, reactions } = parseReactions(paras);
    map.set(id, {
      author: c.getAttribute("w:author") ?? "",
      date: c.getAttribute("w:date") ?? "",
      text,
      reactions,
      lastParaId: paraIds[paraIds.length - 1] ?? "",
    });
  }
  return map;
}

/** Read commentsExtended.xml threading: a comment's lastParaId -> its parent's lastParaId. */
function readCommentParents(file: Uint8Array | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!file) return map;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  for (const ex of Array.from(doc.getElementsByTagName("w15:commentEx"))) {
    const pid = ex.getAttribute("w15:paraId");
    const parent = ex.getAttribute("w15:paraIdParent");
    if (pid && parent) map.set(pid, parent);
  }
  return map;
}

/** Read which comment paraIds are marked resolved (w15:done="1"). */
function readCommentDone(file: Uint8Array | undefined): Set<string> {
  const set = new Set<string>();
  if (!file) return set;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  for (const ex of Array.from(doc.getElementsByTagName("w15:commentEx"))) {
    const pid = ex.getAttribute("w15:paraId");
    const d = ex.getAttribute("w15:done");
    if (pid && (d === "1" || d === "true")) set.add(pid);
  }
  return set;
}

// Serialize an element with namespace declarations so it can be re-parsed standalone and
// re-emitted verbatim on save (the "passthrough" of anything we don't model as editable).
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
const passthroughAttr = (node: Element): string => ` data-docx-xml="${escapeAttr(serializePassthrough(node))}"`;

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(bin);
}

/** Resolve a relationship target (relative to word/) to bytes in the archive. */
function resolveMedia(target: string, files: Record<string, Uint8Array>): Uint8Array | undefined {
  const clean = target.replace(/^\.\//, "");
  if (clean.startsWith("/")) return files[clean.slice(1)];
  return files["word/" + clean] ?? files[clean];
}

// De-obfuscate an embedded .odttf font (first 32 bytes XOR'd with the reversed GUID key).
export function deobfuscateFont(bytes: Uint8Array, target: string, key: string): Uint8Array {
  const hex = key.replace(/[^0-9a-fA-F]/g, "");
  const obfuscated = target.toLowerCase().endsWith(".odttf") || (hex.length === 32 && /[1-9a-f]/i.test(hex));
  if (!obfuscated || hex.length !== 32) return bytes;
  const k: number[] = [];
  for (let i = 0; i < 16; i++) k.push(parseInt(hex.substring(i * 2, i * 2 + 2), 16));
  k.reverse();
  const data = bytes.slice();
  for (let i = 0; i < 32 && i < data.length; i++) data[i] ^= k[i % 16]!;
  return data;
}

/** Build @font-face rules for fonts embedded in the .docx (fontTable.xml + word/fonts/). */
function loadEmbeddedFonts(files: Record<string, Uint8Array>): { css: string; urls: string[] } {
  const out = { css: "", urls: [] as string[] };
  const ft = files["word/fontTable.xml"];
  if (!ft) return out;
  const rels = readRels(files["word/_rels/fontTable.xml.rels"]);
  const doc = new DOMParser().parseFromString(strFromU8(ft), "application/xml");
  const variants: [string, string, string][] = [
    ["w:embedRegular", "normal", "normal"],
    ["w:embedBold", "bold", "normal"],
    ["w:embedItalic", "normal", "italic"],
    ["w:embedBoldItalic", "bold", "italic"],
  ];
  for (const font of Array.from(doc.getElementsByTagName("w:font"))) {
    const name = font.getAttribute("w:name");
    if (!name) continue;
    for (const [tag, weight, style] of variants) {
      const e = font.getElementsByTagName(tag)[0];
      if (!e) continue;
      const rid = e.getAttribute("r:id") ?? e.getAttributeNS(R, "id");
      const target = rid ? rels.get(rid) : undefined;
      const bytes = target ? resolveMedia(target, files) : undefined;
      if (!target || !bytes) continue;
      const data = deobfuscateFont(bytes, target, e.getAttribute("w:fontKey") ?? "");
      const url = URL.createObjectURL(new Blob([data as BlobPart], { type: "font/ttf" }));
      out.urls.push(url);
      out.css += `@font-face{font-family:"${name.replace(/"/g, "")}";font-weight:${weight};font-style:${style};font-display:swap;src:url(${url});}\n`;
    }
  }
  return out;
}

/** The document default font (styles.xml docDefaults), used as the page base font. */
function defaultFont(files: Record<string, Uint8Array>): string | undefined {
  const sx = files["word/styles.xml"];
  if (!sx) return undefined;
  const doc = new DOMParser().parseFromString(strFromU8(sx), "application/xml");
  const rf = doc.getElementsByTagName("w:docDefaults")[0]?.getElementsByTagName("w:rFonts")[0];
  return rf?.getAttribute("w:ascii") || rf?.getAttribute("w:hAnsi") || undefined;
}

const emuToPx = (v: string | null): number | undefined => {
  const n = v ? Number(v) : NaN;
  return Number.isFinite(n) ? Math.round(n / 9525) : undefined;
};

/** Render an image-bearing run (w:drawing / w:pict) to an <img>, preserving the run. */
function imageHtml(run: Element, ctx: RenderCtx): string {
  const blip = run.getElementsByTagName("a:blip")[0] ?? run.getElementsByTagName("pic:blipFill")[0]?.getElementsByTagName("a:blip")[0];
  let rid = blip?.getAttribute("r:embed") ?? blip?.getAttribute("r:link") ?? undefined;
  if (!rid) rid = run.getElementsByTagName("v:imagedata")[0]?.getAttribute("r:id") ?? undefined;
  const target = rid ? ctx.rels.get(rid) : undefined;
  const bytes = target ? resolveMedia(target, ctx.files) : undefined;
  const ext = (target?.split(".").pop() ?? "").toLowerCase();
  const mime = IMG_MIME[ext];
  const pass = passthroughAttr(run);
  if (!bytes || !mime || ext === "emf" || ext === "wmf") {
    return `<span class="docx-img-ph" contenteditable="false"${pass}>[image]</span>`;
  }
  const extent = run.getElementsByTagName("wp:extent")[0];
  const wpx = emuToPx(extent?.getAttribute("cx") ?? null);
  const cy = emuToPx(extent?.getAttribute("cy") ?? null);
  const dims = wpx ? ` width="${wpx}"${cy ? ` height="${cy}"` : ""}` : "";
  return `<img src="data:${mime};base64,${bytesToBase64(bytes)}" alt="" contenteditable="false"${pass}${dims}>`;
}

const hasDrawing = (run: Element): boolean =>
  run.getElementsByTagName("w:drawing").length > 0 || run.getElementsByTagName("w:pict").length > 0 || run.getElementsByTagName("w:object").length > 0;

// w:pBdr / table borders -> a CSS border declaration. OOXML sizes are in eighths of a point.
function borderCss(b: Element | undefined): string {
  if (!b) return "";
  const val = b.getAttribute("w:val") ?? "single";
  if (val === "nil" || val === "none") return "none";
  const sz = Number(b.getAttribute("w:sz") ?? "4");
  const px = Math.max(1, Math.round((sz / 8) * (96 / 72)));
  const color = b.getAttribute("w:color");
  const css = color && color !== "auto" && /^[0-9a-fA-F]{6}$/.test(color) ? `#${color}` : "#000";
  const style = val === "dashed" || val === "dotted" || val === "double" ? val : "solid";
  return `${px}px ${style} ${css}`;
}
function paragraphBorderStyle(pPr: Element | undefined): string {
  const pBdr = pPr?.getElementsByTagName("w:pBdr")[0];
  if (!pBdr) return "";
  const parts: string[] = [];
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const css = borderCss(pBdr.getElementsByTagName(`w:${side}`)[0]);
    if (css && css !== "none") parts.push(`border-${side}:${css}`);
  }
  if (parts.length) parts.push("padding:2px 6px");
  return parts.join(";");
}

/** Render a w:tbl to a read-only HTML table; the real table is preserved as passthrough. */
function tableHtml(tbl: Element, ctx: RenderCtx): string {
  const tblPr = tbl.getElementsByTagName("w:tblPr")[0];
  const tblBorders = tblPr?.getElementsByTagName("w:tblBorders")[0];
  const cellBorder = borderCss(tblBorders?.getElementsByTagName("w:top")[0]) || "1px solid #999";
  const showBorder = cellBorder !== "none";
  let rows = "";
  for (const tr of Array.from(tbl.children)) {
    if (tr.tagName !== "w:tr") continue;
    let cells = "";
    for (const tc of Array.from(tr.children)) {
      if (tc.tagName !== "w:tc") continue;
      const tcPr = tc.getElementsByTagName("w:tcPr")[0];
      const vMerge = tcPr?.getElementsByTagName("w:vMerge")[0];
      if (vMerge && (vMerge.getAttribute("w:val") ?? "continue") === "continue") continue; // continuation cell
      const span = tcPr?.getElementsByTagName("w:gridSpan")[0]?.getAttribute("w:val");
      const shd = tcPr?.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill");
      const bg = shd && shd !== "auto" && /^[0-9a-fA-F]{6}$/.test(shd) ? `background:#${shd};` : "";
      let inner = "";
      for (const p of Array.from(tc.children)) {
        if (p.tagName === "w:p") inner += `<div>${inlineToHtml(p, ctx) || "<br>"}</div>`;
        else if (p.tagName === "w:tbl") inner += tableHtml(p, ctx);
      }
      const cs = Number(span) > 1 ? ` colspan="${Number(span)}"` : "";
      const bdr = showBorder ? `border:${cellBorder};` : "";
      cells += `<td${cs} style="${bdr}${bg}padding:3px 6px;vertical-align:top">${inner || "<br>"}</td>`;
    }
    rows += `<tr>${cells}</tr>`;
  }
  return `<table class="docx-table" contenteditable="false"${passthroughAttr(tbl)} style="border-collapse:collapse;margin:0 0 .6em">${rows}</table>`;
}

// ---------------------------------------------------------------------------
// .docx -> HTML
// ---------------------------------------------------------------------------

function readRels(file: Uint8Array | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!file) return map;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
    const id = r.getAttribute("Id");
    const target = r.getAttribute("Target");
    if (id && target) map.set(id, target);
  }
  return map;
}

/** numId -> true if ordered (decimal-ish), best-effort from numbering.xml. */
function readNumbering(file: Uint8Array | undefined): Map<string, boolean> {
  const ordered = new Map<string, boolean>();
  if (!file) return ordered;
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  const absFmt = new Map<string, boolean>(); // abstractNumId -> ordered
  for (const an of Array.from(doc.getElementsByTagName("w:abstractNum"))) {
    const id = an.getAttributeNS(W, "abstractNumId") ?? an.getAttribute("w:abstractNumId");
    const fmt = an.getElementsByTagName("w:numFmt")[0]?.getAttribute("w:val") ?? "";
    if (id) absFmt.set(id, fmt !== "bullet" && fmt !== "none" && fmt !== "");
  }
  for (const num of Array.from(doc.getElementsByTagName("w:num"))) {
    const numId = num.getAttributeNS(W, "numId") ?? num.getAttribute("w:numId");
    const abs = num.getElementsByTagName("w:abstractNumId")[0]?.getAttribute("w:val") ?? "";
    if (numId) ordered.set(numId, absFmt.get(abs) ?? false);
  }
  return ordered;
}

const onFlag = (rPr: Element | undefined, tag: string): boolean => {
  const el = rPr?.getElementsByTagName(tag)[0];
  if (!el) return false;
  const v = el.getAttribute("w:val");
  return v !== "false" && v !== "0" && v !== "none";
};

function runStyle(f: Fmt): string {
  const parts: string[] = [];
  if (f.color) parts.push(`color:#${f.color}`);
  const bg = (f.highlight && HL_CSS[f.highlight]) || (f.shading ? `#${f.shading}` : "");
  if (bg) parts.push(`background-color:${bg}`);
  if (f.sizeHalfPt) parts.push(`font-size:${f.sizeHalfPt / 2}pt`);
  if (f.font) parts.push(`font-family:'${f.font.replace(/'/g, "")}', serif`);
  return parts.join(";");
}

const wrapFmt = (inner: string, f: Fmt): string => {
  let s = inner;
  if (f.u) s = `<u>${s}</u>`;
  if (f.strike) s = `<s>${s}</s>`;
  if (f.i) s = `<em>${s}</em>`;
  if (f.b) s = `<strong>${s}</strong>`;
  const style = runStyle(f);
  if (style) s = `<span style="${escapeAttr(style)}">${s}</span>`;
  return s;
};

const attrVal = (rPr: Element | undefined, tag: string): string | undefined =>
  rPr?.getElementsByTagName(tag)[0]?.getAttribute("w:val") ?? undefined;

// A page-break marker for the continuous ("pageless") view. "manual" breaks round-trip to
// w:br type=page; "auto" markers (Word's last-rendered page boundaries) are display-only.
const pageBreakHtml = (kind: "manual" | "auto"): string =>
  `<span class="docx-pagebreak${kind === "auto" ? " docx-pagebreak-auto" : ""}" contenteditable="false" data-docx-pagebreak="${kind}" data-label="${escapeAttr(t("pageBreak"))}"></span>`;

/** Read a w:rPr element into a Fmt (used for the run's props and rPrChange's old props). */
function readFmt(rPr: Element | undefined): Fmt {
  const color = attrVal(rPr, "w:color");
  const highlight = attrVal(rPr, "w:highlight");
  const sz = attrVal(rPr, "w:sz");
  const shd = rPr?.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill");
  return {
    b: onFlag(rPr, "w:b"),
    i: onFlag(rPr, "w:i"),
    u: onFlag(rPr, "w:u"),
    strike: onFlag(rPr, "w:strike"),
    color: color && color !== "auto" && /^[0-9a-fA-F]{6}$/.test(color) ? color.toUpperCase() : undefined,
    highlight: highlight && highlight !== "none" ? highlight : undefined,
    shading: shd && shd !== "auto" && /^[0-9a-fA-F]{6}$/.test(shd) ? shd.toUpperCase() : undefined,
    sizeHalfPt: sz && /^\d+$/.test(sz) ? Number(sz) : undefined,
    font: rPr?.getElementsByTagName("w:rFonts")[0]?.getAttribute("w:ascii") || rPr?.getElementsByTagName("w:rFonts")[0]?.getAttribute("w:hAnsi") || undefined,
  };
}

function runToHtml(run: Element): string {
  const rPr = run.getElementsByTagName("w:rPr")[0];
  const f = readFmt(rPr);
  // A formatting change (w:rPrChange) carries the previous run properties.
  const rPrChange = rPr ? Array.from(rPr.children).find((c) => c.tagName === "w:rPrChange") : undefined;
  const oldRPr = rPrChange ? Array.from(rPrChange.children).find((c) => c.tagName === "w:rPr") : undefined;
  let out = "";
  let text = "";
  const flush = () => {
    if (text) {
      out += wrapFmt(text, f);
      text = "";
    }
  };
  for (const node of Array.from(run.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.tagName === "w:t" || el.tagName === "w:delText") text += escapeHtml(el.textContent ?? "");
    else if (el.tagName === "w:tab") text += "    ";
    else if (el.tagName === "w:br") {
      if (el.getAttribute("w:type") === "page") {
        flush();
        out += pageBreakHtml("manual");
      } else text += "<br>";
    } else if (el.tagName === "w:lastRenderedPageBreak") {
      flush();
      out += pageBreakHtml("auto");
    }
  }
  flush();
  if (rPrChange && out) {
    const author = rPrChange.getAttribute("w:author") ?? "";
    const date = rPrChange.getAttribute("w:date") ?? "";
    out = `<span class="docx-rpr-change" data-old="${escapeAttr(JSON.stringify(readFmt(oldRPr)))}" data-rev-author="${escapeAttr(author)}" data-rev-date="${escapeAttr(date)}" title="${escapeAttr(author ? author + (date ? " – " + date.slice(0, 10) : "") : "")}">${out}</span>`;
  }
  return out;
}

// A zero-width passthrough for a comment range marker, so it round-trips on save.
const commentMark = (el: Element): string => `<span class="docx-cmark" contenteditable="false"${passthroughAttr(el)}></span>`;

/** A clickable comment reference marker carrying the comment's author/date/text. */
function commentRefHtml(run: Element, ctx: RenderCtx): string {
  const id = run.getElementsByTagName("w:commentReference")[0]?.getAttribute("w:id") ?? "";
  ctx.commentOrder.push(id);
  const c = ctx.comments.get(id);
  const meta = c ? `${c.author}${c.date ? " – " + c.date.slice(0, 10) : ""}` : "";
  return (
    `<span class="docx-comment-ref" contenteditable="false" data-comment-id="${escapeAttr(id)}"` +
    ` data-comment-meta="${escapeAttr(meta)}" data-comment-text="${escapeAttr(c?.text ?? "")}"` +
    ` title="${escapeAttr(meta ? meta + ": " + (c?.text ?? "") : c?.text ?? "")}"${passthroughAttr(run)}>\u{1F4AC}</span>`
  );
}

// A tracked change (w:ins / w:del) -> an <ins>/<del> carrying author/date, so it renders
// as a suggestion and round-trips (or can be accepted/rejected) later.
function revisionHtml(el: Element, ctx: RenderCtx): string {
  const tag = el.tagName === "w:del" ? "del" : "ins";
  const author = el.getAttribute("w:author") ?? "";
  const date = el.getAttribute("w:date") ?? "";
  let inner = "";
  for (const r of Array.from(el.getElementsByTagName("w:r"))) inner += hasDrawing(r) ? imageHtml(r, ctx) : runToHtml(r);
  if (!inner) return "";
  const title = author ? `${author}${date ? " – " + date.slice(0, 10) : ""}` : "";
  return `<${tag} class="docx-${tag}" data-author="${escapeAttr(author)}" data-date="${escapeAttr(date)}" title="${escapeAttr(title)}">${inner}</${tag}>`;
}

// Runs we can model fully; anything else in a run (w:sym, w:fldChar, w:footnoteReference,
// w:object, ...) means we must preserve the whole run verbatim rather than drop it.
const MODELED_RUN_CHILDREN = new Set(["w:rPr", "w:t", "w:delText", "w:tab", "w:br", "w:lastRenderedPageBreak"]);
const runIsModeled = (run: Element): boolean => Array.from(run.children).every((c) => MODELED_RUN_CHILDREN.has(c.tagName));
// Preserve an unmodelled element (bookmark, field, content control, math, ...) verbatim:
// show its text read-only and carry the original XML so it round-trips on save.
const inlinePassthrough = (el: Element): string => `<span class="docx-pass" contenteditable="false"${passthroughAttr(el)}>${escapeHtml(el.textContent ?? "")}</span>`;

function inlineToHtml(p: Element, ctx: RenderCtx): string {
  let html = "";
  for (const id of ctx.openComments) html += `<span class="docx-comment" data-comment-id="${escapeAttr(id)}">`;
  for (const node of Array.from(p.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.tagName === "w:commentRangeStart") {
      const id = el.getAttribute("w:id") ?? "";
      html += commentMark(el);
      html += `<span class="docx-comment" data-comment-id="${escapeAttr(id)}">`;
      ctx.openComments.add(id);
    } else if (el.tagName === "w:commentRangeEnd") {
      const id = el.getAttribute("w:id") ?? "";
      if (ctx.openComments.delete(id)) html += "</span>";
      html += commentMark(el);
    } else if (el.tagName === "w:ins" || el.tagName === "w:del") {
      html += revisionHtml(el, ctx);
    } else if (el.tagName === "w:r") {
      if (el.getElementsByTagName("w:commentReference").length) html += commentRefHtml(el, ctx);
      else if (hasDrawing(el)) html += imageHtml(el, ctx);
      else if (!runIsModeled(el)) html += inlinePassthrough(el); // run holds a field char, symbol, footnote ref, ...
      else html += runToHtml(el);
    } else if (el.tagName === "w:hyperlink") {
      const id = el.getAttributeNS(R, "id") ?? el.getAttribute("r:id") ?? "";
      const href = ctx.rels.get(id) ?? "";
      let inner = "";
      for (const r of Array.from(el.getElementsByTagName("w:r"))) inner += hasDrawing(r) ? imageHtml(r, ctx) : runToHtml(r);
      html += `<a href="${escapeHtml(href)}">${inner}</a>`;
    } else if (el.tagName !== "w:pPr") {
      // Anything else (bookmarks, fields, content controls, math, moves, ...) -> preserve.
      html += inlinePassthrough(el);
    }
  }
  for (let i = 0; i < ctx.openComments.size; i++) html += "</span>"; // closed here, reopened next paragraph
  return html;
}

interface PInfo {
  heading: number; // 0 = not a heading
  isList: boolean;
  ordered: boolean;
  align?: string; // CSS text-align
  border?: string; // CSS border declaration block
  pageBreakBefore: boolean;
  revPara?: "ins" | "del"; // paragraph-mark revision (split/merge)
  revAuthor?: string;
  revDate?: string;
}
function paragraphInfo(p: Element, numbering: Map<string, boolean>): PInfo {
  const pPr = p.getElementsByTagName("w:pPr")[0];
  let heading = 0;
  const styleVal = (pPr?.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") ?? "").replace(/\s+/g, "");
  const hm = /^heading([1-9])$/i.exec(styleVal);
  if (hm) heading = Number(hm[1]);
  const numPr = pPr?.getElementsByTagName("w:numPr")[0];
  const numId = numPr?.getElementsByTagName("w:numId")[0]?.getAttribute("w:val") ?? "";
  const jc = pPr?.getElementsByTagName("w:jc")[0]?.getAttribute("w:val") ?? "";
  // Paragraph-mark revision lives in pPr > rPr > w:ins / w:del.
  const markRPr = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:rPr") : undefined;
  const mark = markRPr && (Array.from(markRPr.children).find((c) => c.tagName === "w:ins" || c.tagName === "w:del") as Element | undefined);
  return {
    heading,
    isList: !!numPr,
    ordered: numbering.get(numId) ?? false,
    align: JC_TO_ALIGN[jc],
    border: paragraphBorderStyle(pPr),
    pageBreakBefore: onFlag(pPr, "w:pageBreakBefore"),
    revPara: mark ? (mark.tagName === "w:del" ? "del" : "ins") : undefined,
    revAuthor: mark?.getAttribute("w:author") ?? undefined,
    revDate: mark?.getAttribute("w:date") ?? undefined,
  };
}
function blockStyleAttr(info: PInfo): string {
  const parts: string[] = [];
  if (info.align && info.align !== "left") parts.push(`text-align:${info.align}`);
  if (info.border) parts.push(info.border);
  const style = parts.length ? ` style="${parts.join(";")}"` : "";
  if (!info.revPara) return style;
  return `${style} class="docx-para-${info.revPara}" data-rev-para="${info.revPara}" data-rev-author="${escapeAttr(info.revAuthor ?? "")}" data-rev-date="${escapeAttr(info.revDate ?? "")}"`;
}

/** Render a block container (w:body or a header/footer root) to HTML. */
function renderBlocks(container: Element, ctx: RenderCtx): string {
  let html = "";
  let listItems: string[] = [];
  let listOrdered = false;
  const flushList = () => {
    if (!listItems.length) return;
    html += `<${listOrdered ? "ol" : "ul"}>${listItems.join("")}</${listOrdered ? "ol" : "ul"}>`;
    listItems = [];
  };
  for (const node of Array.from(container.children)) {
    if (node.tagName === "w:tbl") {
      flushList();
      html += tableHtml(node, ctx);
      continue;
    }
    if (node.tagName !== "w:p") {
      flushList();
      // w:sectPr is re-appended separately on save; any other block (block-level content
      // control, etc.) is preserved verbatim so an edit doesn't drop it.
      if (node.tagName !== "w:sectPr") {
        html += `<div class="docx-pass-block" contenteditable="false"${passthroughAttr(node)}>${escapeHtml(node.textContent ?? "")}</div>`;
      }
      continue;
    }
    const info = paragraphInfo(node, ctx.numbering);
    const inner = inlineToHtml(node, ctx);
    if (info.isList) {
      listOrdered = info.ordered;
      listItems.push(`<li${blockStyleAttr(info)}>${inner || "<br>"}</li>`);
      continue;
    }
    flushList();
    if (info.pageBreakBefore) html += pageBreakHtml("manual");
    if (info.heading) {
      const lvl = Math.min(6, info.heading);
      html += `<h${lvl}${blockStyleAttr(info)}>${inner || "<br>"}</h${lvl}>`;
    } else {
      html += `<p${blockStyleAttr(info)}>${inner || "<br>"}</p>`;
    }
  }
  flushList();
  return html;
}

/** Render a referenced header/footer part (header1.xml etc.) to read-only HTML, or "". */
function renderRefPart(files: Record<string, Uint8Array>, target: string | undefined): string {
  if (!target) return "";
  const bytes = resolveMedia(target, files);
  if (!bytes) return "";
  const doc = new DOMParser().parseFromString(strFromU8(bytes), "application/xml");
  const root = doc.documentElement;
  if (!root) return "";
  const name = target.replace(/^.*\//, "").replace(/\.xml$/, "");
  const ctx: RenderCtx = {
    files,
    rels: readRels(files[`word/_rels/${name}.xml.rels`]),
    numbering: new Map(),
    comments: new Map(),
    openComments: new Set(),
    commentOrder: [],
  };
  return renderBlocks(root, ctx);
}

function refTarget(sectPr: Element | undefined, tag: string, rels: Map<string, string>): string | undefined {
  if (!sectPr) return undefined;
  const refs = Array.from(sectPr.getElementsByTagName(tag));
  const pick = refs.find((r) => (r.getAttribute("w:type") ?? "default") === "default") ?? refs[0];
  const id = pick?.getAttribute("r:id");
  return id ? rels.get(id) : undefined;
}

export interface CommentReaction {
  emoji: string;
  people: string[];
}
export interface CommentEntry {
  id: string;
  author: string;
  date: string;
  text: string;
  reactions: CommentReaction[];
  paraId: string; // last w14:paraId, used as the threading key
}
export interface CommentThread extends CommentEntry {
  replies: CommentEntry[];
  resolved: boolean;
}
export interface DocxParts {
  body: string;
  header: string;
  footer: string;
  headerPath?: string; // archive key of the header part, for write-back
  footerPath?: string;
  comments: CommentThread[]; // top-level threads in document order, replies nested
}

/** The archive key for a relationship target relative to word/ (e.g. "header1.xml"). */
function partKey(target: string | undefined, files: Record<string, Uint8Array>): string | undefined {
  if (!target) return undefined;
  const clean = target.replace(/^\.\//, "");
  if (clean.startsWith("/")) return files[clean.slice(1)] ? clean.slice(1) : undefined;
  if (files["word/" + clean]) return "word/" + clean;
  return files[clean] ? clean : undefined;
}

/** Parse a .docx into the editable body HTML plus the header/footer HTML and part keys. */
export function docxToParts(bytes: Uint8Array): DocxParts {
  const empty: DocxParts = { body: "", header: "", footer: "", comments: [] };
  const files = unzipSync(bytes);
  const docXml = files["word/document.xml"];
  if (!docXml) return empty;
  const doc = new DOMParser().parseFromString(strFromU8(docXml), "application/xml");
  const body = doc.getElementsByTagName("w:body")[0];
  if (!body) return empty;
  const rels = readRels(files["word/_rels/document.xml.rels"]);
  const comments = readComments(files["word/comments.xml"]);
  const ctx: RenderCtx = {
    files,
    rels,
    numbering: readNumbering(files["word/numbering.xml"]),
    comments,
    openComments: new Set(),
    commentOrder: [],
  };

  const html = renderBlocks(body, ctx);
  const sectPr = Array.from(body.getElementsByTagName("w:sectPr")).pop();
  const headerTarget = refTarget(sectPr, "w:headerReference", rels);
  const footerTarget = refTarget(sectPr, "w:footerReference", rels);

  // Group comments into threads via commentsExtended.xml (replies under their parent).
  const parents = readCommentParents(files["word/commentsExtended.xml"]);
  const done = readCommentDone(files["word/commentsExtended.xml"]);
  const byParaId = new Map<string, string>(); // lastParaId -> comment id
  for (const [id, c] of comments) if (c.lastParaId) byParaId.set(c.lastParaId, id);
  const entry = (id: string): CommentEntry => {
    const c = comments.get(id);
    return { id, author: c?.author ?? "", date: c?.date ?? "", text: c?.text ?? "", reactions: c?.reactions ?? [], paraId: c?.lastParaId ?? "" };
  };
  const parentOf = (id: string): string | undefined => {
    const c = comments.get(id);
    const pPara = c?.lastParaId ? parents.get(c.lastParaId) : undefined;
    return pPara ? byParaId.get(pPara) : undefined;
  };
  const threads: CommentThread[] = [];
  const threadById = new Map<string, CommentThread>();
  const seen = new Set<string>();
  for (const id of ctx.commentOrder) {
    if (seen.has(id)) continue;
    seen.add(id);
    const parent = parentOf(id);
    const parentThread = parent ? threadById.get(parent) : undefined;
    if (parentThread) {
      parentThread.replies.push(entry(id));
    } else {
      const th: CommentThread = Object.assign(entry(id), { replies: [] as CommentEntry[], resolved: done.has(comments.get(id)?.lastParaId ?? "") });
      threads.push(th);
      threadById.set(id, th);
    }
  }
  // Replies threaded in commentsExtended but without a body reference (e.g. ones we added).
  for (const [id] of comments) {
    if (seen.has(id)) continue;
    const parentThread = parentOf(id) ? threadById.get(parentOf(id)!) : undefined;
    if (parentThread) {
      seen.add(id);
      parentThread.replies.push(entry(id));
    }
  }
  return {
    body: html || "<p><br></p>",
    header: renderRefPart(files, headerTarget),
    footer: renderRefPart(files, footerTarget),
    headerPath: partKey(headerTarget, files),
    footerPath: partKey(footerTarget, files),
    comments: threads,
  };
}

/** Convert a .docx body to editable HTML. Returns "" if there is no document body. */
export function docxToHtml(bytes: Uint8Array): string {
  return docxToParts(bytes).body;
}

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

const fmtHasProps = (f: Fmt): boolean => !!(f.b || f.i || f.u || f.strike || f.color || f.highlight || f.shading || f.sizeHalfPt || f.font);

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
  const revPara = src.getAttribute("data-rev-para"); // "ins" | "del" paragraph-mark revision
  if (opts.heading || (opts.list && ctx.listNumId) || jc || revPara) {
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
  const stash = el.getAttribute("data-docx-xml");
  if (stash) {
    // A passthrough block (table, etc.): re-emit the original OOXML verbatim.
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
export function htmlToDocx(
  html: string,
  original: Uint8Array,
  parts?: { path: string; html: string }[],
  opts?: { reactions?: ReactionEdit[]; replies?: ReplyEdit[]; done?: Map<string, boolean>; deletedComments?: string[] },
): Uint8Array {
  const files = unzipSync(original);
  rebuildPart(files, "word/document.xml", html);
  for (const p of parts ?? []) rebuildPart(files, p.path, p.html);
  applyNewComments(files, html);
  applyReactions(files, opts?.reactions ?? []);
  applyReplies(files, opts?.replies ?? []);
  applyDone(files, opts?.done ?? new Map());
  applyDeletedComments(files, opts?.deletedComments ?? []);
  return zipSync(files);
}

// ---------------------------------------------------------------------------
// Editor
// ---------------------------------------------------------------------------

export interface DocxEditorOptions {
  onChange?: () => void;
  /** Author name stamped on comments added in the editor. */
  author?: string;
  /** ISO date string for added comments (injected so the build stays deterministic). */
  now?: string;
}
export interface DocxEditor {
  getBytes(): Promise<Uint8Array>;
  isDirty(): boolean;
  destroy(): void;
}

export function createDocxEditor(container: HTMLElement, bytes: Uint8Array, options: DocxEditorOptions = {}): DocxEditor {
  const original = bytes.slice();
  let dirty = false;

  const wrap = document.createElement("div");
  wrap.className = "docxedit-wrap";
  const toolbar = document.createElement("div");
  toolbar.className = "docxedit-toolbar";
  toolbar.setAttribute("role", "toolbar");
  toolbar.setAttribute("aria-label", t("toolbar"));
  const scroll = document.createElement("div");
  scroll.className = "docxedit-scroll";
  const page = document.createElement("div");
  page.className = "docxedit-page";
  const doc = document.createElement("div");
  doc.className = "docxedit-doc";
  doc.contentEditable = "true";
  doc.spellcheck = false;
  doc.setAttribute("role", "textbox");
  doc.setAttribute("aria-multiline", "true");
  doc.setAttribute("aria-label", t("documentText"));

  let parts: DocxParts = { body: "<p><br></p>", header: "", footer: "", comments: [] };
  try {
    parts = docxToParts(bytes);
  } catch (e) {
    console.warn("docxedit: failed to parse document", e);
  }
  doc.innerHTML = parts.body || "<p><br></p>";

  // Load fonts embedded in the .docx so the document renders in its own typefaces.
  const fontUrls: string[] = [];
  try {
    const fileMap = unzipSync(bytes);
    const ff = loadEmbeddedFonts(fileMap);
    if (ff.css) {
      const fs = document.createElement("style");
      fs.textContent = ff.css;
      wrap.appendChild(fs);
      fontUrls.push(...ff.urls);
    }
    const df = defaultFont(fileMap);
    if (df) page.style.setProperty("--docxedit-doc-font", `"${df.replace(/"/g, "")}"`);
  } catch {
    /* no embedded fonts */
  }

  const band = (cls: string, label: string, html: string): HTMLElement | null => {
    if (!html) return null;
    const el = document.createElement("div");
    el.className = cls;
    el.contentEditable = "true";
    el.spellcheck = false;
    el.setAttribute("role", "textbox");
    el.setAttribute("aria-multiline", "true");
    el.setAttribute("aria-label", label);
    el.innerHTML = html;
    return el;
  };
  const header = band("docxedit-header", t("header"), parts.header);
  const footer = band("docxedit-footer", t("footer"), parts.footer);
  if (header) page.appendChild(header);
  page.appendChild(doc);
  if (footer) page.appendChild(footer);

  // Keep the page centred, with the comments column in the right margin (Google-Docs
  // style): an empty left spacer balances the right comments area so the page stays centred.
  const canvas = document.createElement("div");
  canvas.className = "docxedit-canvas";
  const leftSpacer = document.createElement("div");
  leftSpacer.className = "docxedit-margin";
  const rightArea = document.createElement("div");
  rightArea.className = "docxedit-margin";
  const cmtPanel = document.createElement("div");
  cmtPanel.className = "docxedit-comments";
  rightArea.appendChild(cmtPanel);
  canvas.append(leftSpacer, page, rightArea);
  scroll.appendChild(canvas);
  wrap.append(toolbar, scroll);
  container.appendChild(wrap);

  // The editable regions (body + header/footer). Toolbar actions target whichever last
  // had focus, so formatting works inside the header and footer too.
  const regions = [doc, header, footer].filter(Boolean) as HTMLElement[];
  let activeEl: HTMLElement = doc;

  // Next comment id: one past the highest already present.
  let nextCommentId = 0;
  for (const m of Array.from(wrap.querySelectorAll("[data-comment-id]"))) {
    const n = Number(m.getAttribute("data-comment-id"));
    if (Number.isFinite(n)) nextCommentId = Math.max(nextCommentId, n + 1);
  }

  // Comments side panel: one card per thread (replies grouped), anchored vertically to the
  // commented range, with reactions and a "more" toggle for long text.
  const pendingReactions: { commentId: string; emoji: string; person: string }[] = [];
  const pendingReplies: { id: string; paraId: string; parentParaId: string; author: string; date: string; text: string }[] = [];
  const pendingDone = new Map<string, boolean>(); // thread paraId -> done
  const deletedComments: string[] = []; // comment ids removed from the document
  // thread membership, kept current as replies/comments are added.
  const threadOf = new Map<string, string>(); // any comment id -> its thread id
  const threadMembers = new Map<string, string[]>();
  const registerThread = (threadId: string, memberIds: string[]) => {
    threadMembers.set(threadId, memberIds);
    for (const m of memberIds) threadOf.set(m, threadId);
  };
  let paraSeed = 0x7f000000;
  const freshParaId = () => (paraSeed++).toString(16).toUpperCase().padStart(8, "0");
  const REACT_CHOICES = ["\u{1F44D}", "❤️", "\u{1F602}", "\u{1F389}", "\u{1F440}", "\u{1F64F}"];
  const metaLine = (c: { author: string; date: string }) => (c.date ? `${c.author} – ${c.date.slice(0, 10)}` : c.author);

  const renderReactions = (row: HTMLElement, entry: CommentEntry) => {
    row.querySelectorAll(".docxedit-react").forEach((n) => n.remove());
    const addBtn = row.querySelector(".docxedit-react-add");
    for (const r of entry.reactions) {
      if (!r.people.length) continue;
      const span = document.createElement("span");
      span.className = "docxedit-react";
      span.title = r.people.join(", ");
      span.textContent = r.emoji + (r.people.length > 1 ? " " + r.people.length : "");
      row.insertBefore(span, addBtn);
    }
  };

  const buildItem = (entry: CommentEntry, isReply: boolean): HTMLElement => {
    const item = document.createElement("div");
    item.className = "docxedit-cmt-item" + (isReply ? " docxedit-cmt-reply" : "");
    const meta = document.createElement("b");
    meta.textContent = metaLine(entry);
    const text = document.createElement("div");
    text.className = "docxedit-cmt-text";
    text.textContent = entry.text;
    const more = document.createElement("button");
    more.type = "button";
    more.className = "docxedit-cmt-more is-hidden";
    more.textContent = t("more");
    more.addEventListener("click", (e) => {
      e.stopPropagation();
      text.classList.add("expanded");
      more.classList.add("is-hidden");
      positionCards();
    });
    const row = document.createElement("div");
    row.className = "docxedit-cmt-react-row";
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "docxedit-react-add";
    addBtn.title = t("addReaction");
    addBtn.textContent = "\u{1F642}+";
    addBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      openReactionPicker(addBtn, entry, row);
    });
    row.appendChild(addBtn);
    renderReactions(row, entry);
    item.append(meta, text, more, row);
    return item;
  };

  const openReactionPicker = (anchor: HTMLElement, entry: CommentEntry, row: HTMLElement) => {
    document.querySelector(".docxedit-react-pop")?.remove();
    const pop = document.createElement("div");
    pop.className = "docxedit-react-pop";
    for (const emoji of REACT_CHOICES) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = emoji;
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        const person = options.author || "Author";
        const existing = entry.reactions.find((r) => r.emoji === emoji);
        if (existing) {
          if (!existing.people.includes(person)) existing.people.push(person);
        } else entry.reactions.push({ emoji, people: [person] });
        pendingReactions.push({ commentId: entry.id, emoji, person });
        renderReactions(row, entry);
        pop.remove();
        positionCards();
        mark();
      });
      pop.appendChild(b);
    }
    wrap.appendChild(pop);
    const ar = anchor.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    pop.style.left = `${Math.min(ar.left - wr.left, wrap.clientWidth - 200)}px`;
    pop.style.top = `${ar.bottom - wr.top + 4}px`;
  };

  const setActiveComment = (threadId: string | null) => {
    for (const c of Array.from(cmtPanel.children)) c.classList.toggle("active", (c as HTMLElement).dataset.commentId === threadId);
    const members = threadId ? threadMembers.get(threadId) ?? [threadId] : [];
    for (const r of Array.from(wrap.querySelectorAll(".docx-comment"))) {
      const rid = (r as HTMLElement).getAttribute("data-comment-id") ?? "";
      (r as HTMLElement).classList.toggle("active", members.includes(rid));
    }
  };

  const actionBtn = (label: string, title: string, fn: (e: Event) => void): HTMLElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-cmt-action";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("click", fn);
    return b;
  };

  const addReplyToThread = (card: HTMLElement, threadId: string, parentParaId: string, text: string) => {
    const id = String(nextCommentId++);
    const paraId = freshParaId();
    const author = options.author || "Author";
    const date = options.now || new Date().toISOString();
    pendingReplies.push({ id, paraId, parentParaId, author, date, text });
    const members = threadMembers.get(threadId) ?? [threadId];
    members.push(id);
    registerThread(threadId, members);
    const box = card.querySelector(".docxedit-cmt-replybox");
    card.insertBefore(buildItem({ id, author, date, text, reactions: [], paraId }, true), box);
    mark();
    positionCards();
  };

  const buildReplyBox = (card: HTMLElement, thread: { id: string }): HTMLElement => {
    const box = document.createElement("div");
    box.className = "docxedit-cmt-replybox";
    const showInput = () => {
      box.innerHTML = "";
      const ta = document.createElement("textarea");
      ta.className = "docxedit-cmt-replyinput";
      ta.rows = 2;
      ta.placeholder = t("reply");
      ta.addEventListener("click", (e) => e.stopPropagation());
      const send = document.createElement("button");
      send.type = "button";
      send.className = "docxedit-cmt-send";
      send.textContent = t("send");
      const commit = () => {
        const txt = ta.value.trim();
        box.replaceWith(buildReplyBox(card, thread));
        if (txt) addReplyToThread(card, thread.id, card.dataset.paraId ?? "", txt);
      };
      send.addEventListener("click", (e) => {
        e.stopPropagation();
        commit();
      });
      ta.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          commit();
        }
      });
      box.append(ta, send);
      ta.focus();
    };
    const btn2 = document.createElement("button");
    btn2.type = "button";
    btn2.className = "docxedit-cmt-replybtn";
    btn2.textContent = t("reply");
    btn2.addEventListener("click", (e) => {
      e.stopPropagation();
      showInput();
    });
    box.appendChild(btn2);
    return box;
  };

  const addThreadCard = (thread: CommentThread): HTMLElement => {
    const card = document.createElement("div");
    card.className = "docxedit-cmt-card" + (thread.resolved ? " resolved" : "");
    card.dataset.commentId = thread.id;
    card.dataset.paraId = thread.paraId || "";
    registerThread(thread.id, [thread.id, ...thread.replies.map((r) => r.id)]);

    const actions = document.createElement("div");
    actions.className = "docxedit-cmt-actions";
    actions.append(
      actionBtn("✓", t("resolve"), (e) => {
        e.stopPropagation();
        const resolved = !card.classList.contains("resolved");
        card.classList.toggle("resolved", resolved);
        pendingDone.set(card.dataset.paraId || "", resolved);
        if (resolved) setActiveComment(null);
        mark();
        positionCards();
      }),
      actionBtn("✕", t("deleteComment"), (e) => {
        e.stopPropagation();
        const members = threadMembers.get(thread.id) ?? [thread.id];
        for (const id of members) {
          deletedComments.push(id);
          // unwrap the highlight span (keep its text), then drop range/reference markers
          wrap.querySelectorAll(`.docx-comment[data-comment-id="${CSS.escape(id)}"]`).forEach((span) => {
            while (span.firstChild) span.parentNode?.insertBefore(span.firstChild, span);
            span.remove();
          });
          wrap.querySelectorAll(`.docx-comment-ref[data-comment-id="${CSS.escape(id)}"]`).forEach((n) => n.remove());
          for (const m of Array.from(wrap.querySelectorAll(".docx-cmark"))) {
            if ((m.getAttribute("data-docx-xml") ?? "").includes(`w:id="${id}"`)) m.remove();
          }
        }
        card.remove();
        mark();
        positionCards();
      }),
    );
    card.appendChild(actions);

    card.appendChild(buildItem(thread, false));
    for (const reply of thread.replies) card.appendChild(buildItem(reply, true));
    card.appendChild(buildReplyBox(card, thread));

    card.addEventListener("click", () => {
      setActiveComment(thread.id);
      wrap.querySelector(`.docx-comment[data-comment-id="${CSS.escape(thread.id)}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
    });
    cmtPanel.appendChild(card);
    return card;
  };

  // Anchor each card to its range's vertical position; stack to avoid overlap.
  const positionCards = () => {
    const cards = Array.from(cmtPanel.children) as HTMLElement[];
    if (!cards.length) return;
    const panelTop = cmtPanel.getBoundingClientRect().top;
    const measured = cards.map((card) => {
      const id = card.dataset.commentId ?? "";
      const marker =
        wrap.querySelector(`.docx-comment[data-comment-id="${CSS.escape(id)}"]`) ?? wrap.querySelector(`.docx-comment-ref[data-comment-id="${CSS.escape(id)}"]`);
      const y = marker ? (marker as HTMLElement).getBoundingClientRect().top - panelTop : 0;
      return { card, y };
    });
    measured.sort((a, b) => a.y - b.y);
    let prevBottom = 0;
    for (const { card, y } of measured) {
      // reveal "more" only when the text actually overflows
      for (const item of Array.from(card.querySelectorAll(".docxedit-cmt-item"))) {
        const txt = item.querySelector(".docxedit-cmt-text") as HTMLElement | null;
        const moreBtn = item.querySelector(".docxedit-cmt-more") as HTMLElement | null;
        if (txt && moreBtn && !txt.classList.contains("expanded")) {
          moreBtn.classList.toggle("is-hidden", txt.scrollHeight <= txt.clientHeight + 2);
        }
      }
      const top = Math.max(y, prevBottom);
      card.style.top = `${top}px`;
      prevBottom = top + card.offsetHeight + 10;
    }
    cmtPanel.style.height = `${Math.max(prevBottom, page.offsetHeight)}px`;
  };

  for (const thread of parts.comments) addThreadCard(thread);
  // Clicking commented (highlighted) text opens its thread; the inline icon is gone.
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement).closest?.(".docxedit-react-pop")) return;
    document.querySelector(".docxedit-react-pop")?.remove();
    const hit = (e.target as HTMLElement).closest?.(".docx-comment, .docx-comment-ref") as HTMLElement | null;
    if (hit) {
      const id = hit.getAttribute("data-comment-id") ?? "";
      const threadId = threadOf.get(id) ?? id;
      setActiveComment(threadId);
      cmtPanel.querySelector(`[data-comment-id="${CSS.escape(threadId)}"]`)?.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  });
  // Position now and once the layout settles (rAF is throttled in background tabs, so also
  // use a timeout), and on any reflow or edit.
  positionCards();
  requestAnimationFrame(positionCards);
  setTimeout(positionCards, 150);
  const repositionObserver = new ResizeObserver(() => positionCards());
  repositionObserver.observe(page);
  for (const r of regions) r.addEventListener("input", () => positionCards());

  // Image select: click an image to select it, drag the corner handle to resize, and use
  // the delete button (top-right) or the Delete/Backspace key to remove it.
  let selImg: HTMLImageElement | null = null;
  const imgHandle = document.createElement("div");
  imgHandle.className = "docxedit-img-handle is-hidden";
  const imgDel = document.createElement("button");
  imgDel.type = "button";
  imgDel.className = "docxedit-img-del is-hidden";
  imgDel.textContent = "✕";
  imgDel.title = t("deleteImage");
  wrap.append(imgHandle, imgDel);
  const placeHandle = () => {
    if (!selImg || !wrap.contains(selImg)) {
      imgHandle.classList.add("is-hidden");
      imgDel.classList.add("is-hidden");
      return;
    }
    const ir = selImg.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    // positions are runtime-computed; they must follow the selected image
    imgHandle.style.left = `${ir.right - wr.left - 6}px`;
    imgHandle.style.top = `${ir.bottom - wr.top - 6}px`;
    imgHandle.classList.remove("is-hidden");
    imgDel.style.left = `${ir.right - wr.left - 11}px`;
    imgDel.style.top = `${ir.top - wr.top - 11}px`;
    imgDel.classList.remove("is-hidden");
  };
  const selectImg = (img: HTMLImageElement | null) => {
    if (selImg) selImg.classList.remove("sel");
    selImg = img;
    if (selImg) selImg.classList.add("sel");
    placeHandle();
  };
  const deleteSelImg = () => {
    if (!selImg) return;
    selImg.remove();
    selImg = null;
    placeHandle();
    mark();
  };
  imgDel.addEventListener("mousedown", (e) => e.preventDefault());
  imgDel.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteSelImg();
  });
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement) === imgDel) return;
    const img = (e.target as HTMLElement).closest?.("img") as HTMLImageElement | null;
    selectImg(img && wrap.contains(img) ? img : null);
  });
  for (const r of regions)
    r.addEventListener("keydown", (e) => {
      if (selImg && (e.key === "Delete" || e.key === "Backspace")) {
        e.preventDefault();
        deleteSelImg();
      }
    });
  scroll.addEventListener("scroll", placeHandle);
  // Persist a resize: new images use the width/height attributes; existing ones carry the
  // original drawing in data-docx-xml, so update its extent (EMU) to the new size.
  const persistImgSize = (img: HTMLImageElement) => {
    const xml = img.getAttribute("data-docx-xml");
    const w = Number(img.getAttribute("width")) || 0;
    const h = Number(img.getAttribute("height")) || 0;
    if (!xml || !w) return;
    try {
      const frag = new DOMParser().parseFromString(xml, "application/xml");
      for (const tag of ["wp:extent", "a:ext"]) {
        const el = frag.getElementsByTagName(tag)[0];
        if (el) {
          el.setAttribute("cx", String(Math.round(w * 9525)));
          el.setAttribute("cy", String(Math.round(h * 9525)));
        }
      }
      img.setAttribute("data-docx-xml", new XMLSerializer().serializeToString(frag.documentElement!));
    } catch {
      /* leave as-is */
    }
  };
  imgHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    const img = selImg;
    if (!img) return;
    const startX = e.clientX;
    const rect = img.getBoundingClientRect();
    const startW = rect.width;
    const ratio = rect.height / startW || 1;
    const onMove = (ev: MouseEvent) => {
      const w = Math.max(16, Math.round(startW + (ev.clientX - startX)));
      img.setAttribute("width", String(w));
      img.setAttribute("height", String(Math.round(w * ratio)));
      placeHandle();
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      persistImgSize(img);
      mark();
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
  // Emit inline CSS (text-align, font-weight, ...) the serializer reads back, not legacy tags.
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch {
    /* not supported; legacy tags still round-trip */
  }

  const mark = () => {
    dirty = true;
    options.onChange?.();
  };
  for (const r of regions) {
    r.addEventListener("input", mark);
    r.addEventListener("focusin", () => {
      activeEl = r;
    });
  }
  const exec = (cmd: string, val?: string) => {
    activeEl.focus();
    document.execCommand(cmd, false, val);
    mark();
  };
  // Wrap the current selection in a span carrying one CSS property (for font size, which
  // has no execCommand equivalent in CSS mode). No-op on a collapsed selection.
  const styleSel = (prop: string, val: string) => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (range.collapsed) return;
    const span = document.createElement("span");
    (span.style as unknown as Record<string, string>)[prop] = val;
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
    } catch {
      return;
    }
    sel.removeAllRanges();
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.addRange(r2);
    mark();
  };
  const btn = (label: string, title: string, fn: () => void, cls = "") => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    if (cls) b.className = cls;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", fn);
    return b;
  };
  const sep = () => {
    const s = document.createElement("span");
    s.className = "sep";
    return s;
  };
  const alignIcon = (rows: [number, number][]): string =>
    `<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">${rows
      .map(([x, w], k) => `<rect x="${x}" y="${3 + k * 4}" width="${w}" height="1.6" rx=".6"/>`)
      .join("")}</svg>`;
  const iconBtn = (svg: string, title: string, fn: () => void) => {
    const b = btn("", title, fn);
    b.innerHTML = svg;
    return b;
  };

  const block = document.createElement("select");
  block.title = t("paragraphStyle");
  block.setAttribute("aria-label", t("paragraphStyle"));
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(t(key), v));
  }
  block.addEventListener("mousedown", () => activeEl.focus());
  block.addEventListener("change", () => exec("formatBlock", block.value));

  // A select whose first option is a non-selectable title; firing fn(value) on change.
  const pickerSelect = (title: string, opts: [string, string][], fn: (v: string) => void): HTMLSelectElement => {
    const s = document.createElement("select");
    s.title = title;
    s.setAttribute("aria-label", title);
    const head = new Option(title, "");
    head.disabled = true;
    head.selected = true;
    s.add(head);
    for (const [v, label] of opts) s.add(new Option(label, v));
    s.addEventListener("mousedown", () => activeEl.focus());
    s.addEventListener("change", () => {
      if (s.value) fn(s.value);
      s.selectedIndex = 0;
    });
    return s;
  };

  // Text colour: a native colour input that applies w:color via foreColor (CSS mode).
  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.value = "#000000";
  colorInput.title = t("textColor");
  colorInput.setAttribute("aria-label", t("textColor"));
  colorInput.className = "docxedit-color";
  colorInput.addEventListener("mousedown", () => activeEl.focus());
  colorInput.addEventListener("input", () => {
    beginFormatChange();
    exec("foreColor", colorInput.value);
  });

  // Background colour: a free colour picker (maps to w:highlight when it matches a named
  // highlight exactly, otherwise to arbitrary w:shd shading). A button clears it.
  const bgWrap = document.createElement("span");
  bgWrap.className = "docxedit-bg";
  const bgInput = document.createElement("input");
  bgInput.type = "color";
  bgInput.value = "#ffff00";
  bgInput.title = t("highlight");
  bgInput.setAttribute("aria-label", t("highlight"));
  bgInput.className = "docxedit-color";
  bgInput.addEventListener("mousedown", () => activeEl.focus());
  bgInput.addEventListener("input", () => {
    beginFormatChange();
    exec("hiliteColor", bgInput.value);
  });
  const bgClear = btn("⌫", t("none"), () => exec("hiliteColor", "transparent"), "docxedit-bg-clear");
  bgWrap.append(bgInput, bgClear);

  const FONTS = ["Arial", "Calibri", "Century", "Courier New", "Georgia", "Times New Roman", "Verdana"];
  const fontSel = pickerSelect(t("font"), FONTS.map((f) => [f, f] as [string, string]), (v) => {
    beginFormatChange();
    exec("fontName", v);
  });

  const SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48"];
  const sizeSel = pickerSelect(t("size"), SIZES.map((s) => [s, s] as [string, string]), (v) => {
    beginFormatChange();
    styleSel("fontSize", `${v}pt`);
  });

  const insertPageBreak = () => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const el = document.createElement("span");
    el.className = "docx-pagebreak";
    el.contentEditable = "false";
    el.setAttribute("data-docx-pagebreak", "manual");
    el.setAttribute("data-label", t("pageBreak"));
    range.insertNode(el);
    range.setStartAfter(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
  };
  const pbIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="3" y="1.5" width="10" height="4" rx=".5"/><rect x="3" y="10.5" width="10" height="4" rx=".5"/>' +
    '<line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="2 1.6"/></svg>';

  // Insert an image: read a file, show it via a data URL, and let the serializer embed it.
  const insertImage = () => {
    const sel = window.getSelection();
    const savedRange = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp";
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const buf = new Uint8Array(await file.arrayBuffer());
      const dataUrl = `data:${file.type};base64,${bytesToBase64(buf)}`;
      const probe = new Image();
      probe.onload = () => {
        const maxW = 600;
        let w = probe.naturalWidth || 200;
        let h = probe.naturalHeight || 200;
        if (w > maxW) {
          h = Math.round((h * maxW) / w);
          w = maxW;
        }
        const img = document.createElement("img");
        img.src = dataUrl;
        img.setAttribute("width", String(w));
        img.setAttribute("height", String(h));
        const range = savedRange ?? (() => {
          const r = document.createRange();
          r.selectNodeContents(activeEl);
          r.collapse(false);
          return r;
        })();
        range.collapse(false);
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
        const s2 = window.getSelection();
        if (s2) {
          s2.removeAllRanges();
          s2.addRange(range);
        }
        mark();
      };
      probe.src = dataUrl;
    });
    fileInput.click();
  };
  const imgIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
    '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

  // Add a comment over the current selection: wrap it in comment-range markers and a
  // reference marker that carries the text, so the serializer can build comments.xml.
  const addComment = () => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed) {
      const tip = document.createElement("div");
      tip.className = "docxedit-cmt-pop docxedit-cmt-tip";
      tip.textContent = t("commentSelect");
      wrap.appendChild(tip);
      setTimeout(() => tip.remove(), 1800);
      return;
    }
    const text = prompt(t("commentPrompt"));
    if (!text) return;
    const id = String(nextCommentId++);
    const author = options.author || "Author";
    const date = options.now || new Date().toISOString();
    const paraId = freshParaId();
    const ns = `xmlns:w="${W}"`;
    const markerSpan = (xml: string): HTMLElement => {
      const s = document.createElement("span");
      s.className = "docx-cmark";
      s.contentEditable = "false";
      s.setAttribute("data-docx-xml", xml);
      return s;
    };
    const start = markerSpan(`<w:commentRangeStart ${ns} w:id="${id}"/>`);
    const end = markerSpan(`<w:commentRangeEnd ${ns} w:id="${id}"/>`);
    const ref = document.createElement("span");
    ref.className = "docx-comment-ref";
    ref.contentEditable = "false";
    ref.textContent = "\u{1F4AC}";
    ref.setAttribute("data-comment-id", id);
    ref.setAttribute("data-comment-new", "1");
    ref.setAttribute("data-comment-paraid", paraId);
    ref.setAttribute("data-comment-author", author);
    if (date) ref.setAttribute("data-comment-date", date);
    ref.setAttribute("data-comment-text", text);
    ref.setAttribute("data-comment-meta", date ? `${author} – ${date.slice(0, 10)}` : author);
    ref.setAttribute(
      "data-docx-xml",
      `<w:r ${ns}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`,
    );
    const range = sel.getRangeAt(0);
    const visual = document.createElement("span");
    visual.className = "docx-comment";
    visual.setAttribute("data-comment-id", id);
    visual.appendChild(range.extractContents());
    range.insertNode(visual);
    const parent = visual.parentNode;
    if (parent) {
      parent.insertBefore(start, visual);
      parent.insertBefore(end, visual.nextSibling);
      parent.insertBefore(ref, end.nextSibling);
    }
    addThreadCard({ id, author, date, text, reactions: [], replies: [], paraId, resolved: false });
    setActiveComment(id);
    positionCards();
    mark();
  };
  const cmtIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<path d="M2 3.5h12v8H6l-3 2.5V11.5H2z"/><line x1="5" y1="6.2" x2="11" y2="6.2"/><line x1="5" y1="8.6" x2="9" y2="8.6"/></svg>';

  // --- Suggestion mode (track changes) ---------------------------------------
  let suggesting = false;
  const sugAuthor = () => options.author || "Author";
  const sugDate = () => options.now || new Date().toISOString();
  const blockOf = (n: Node | null): HTMLElement | null => {
    const start = n && (n.nodeType === 3 ? n.parentElement : (n as Element));
    const el = start?.closest?.("p,h1,h2,h3,h4,h5,h6,li,div") as HTMLElement | null;
    // never return a region container (the doc/header/footer itself)
    if (!el || el.classList.contains("docxedit-doc") || el.classList.contains("docxedit-header") || el.classList.contains("docxedit-footer")) return null;
    return el;
  };
  const markPara = (el: HTMLElement, kind: "ins" | "del") => {
    el.setAttribute("data-rev-para", kind);
    el.setAttribute("data-rev-author", sugAuthor());
    el.setAttribute("data-rev-date", sugDate());
    el.classList.add(`docx-para-${kind}`);
  };
  // Capture current run formatting from an element's computed style (for rPrChange "old").
  const captureFmt = (el: Element | null): Record<string, unknown> => {
    if (!el) return {};
    const cs = getComputedStyle(el);
    const deco = cs.textDecorationLine || "";
    const wght = Number(cs.fontWeight);
    return {
      b: cs.fontWeight === "bold" || wght >= 600,
      i: cs.fontStyle === "italic",
      u: /underline/.test(deco),
      strike: /line-through/.test(deco),
      color: toHex6(cs.color),
      sizeHalfPt: fontSizeToHalfPt(cs.fontSize),
      font: firstFontFamily(cs.fontFamily),
    };
  };
  const fmtToStyle = (f: Record<string, unknown>): string => {
    const p: string[] = [];
    if (f.b) p.push("font-weight:bold");
    if (f.i) p.push("font-style:italic");
    const dec = [f.u ? "underline" : "", f.strike ? "line-through" : ""].filter(Boolean).join(" ");
    if (dec) p.push(`text-decoration:${dec}`);
    if (f.color) p.push(`color:#${f.color}`);
    if (f.sizeHalfPt) p.push(`font-size:${(f.sizeHalfPt as number) / 2}pt`);
    if (f.font) p.push(`font-family:'${String(f.font).replace(/'/g, "")}', serif`);
    return p.join(";");
  };
  // Record a formatting change: wrap the selection so its old props become an rPrChange.
  const beginFormatChange = () => {
    if (!suggesting) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed) return;
    const range = sel.getRangeAt(0);
    const anc = range.commonAncestorContainer;
    const el = anc.nodeType === 3 ? anc.parentElement : (anc as Element);
    if (el?.closest?.(".docx-rpr-change")) return; // already inside a change
    const span = document.createElement("span");
    span.className = "docx-rpr-change";
    span.setAttribute("data-old", JSON.stringify(captureFmt(el)));
    span.setAttribute("data-rev-author", sugAuthor());
    span.setAttribute("data-rev-date", sugDate());
    span.appendChild(range.extractContents());
    range.insertNode(span);
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.removeAllRanges();
    sel.addRange(r2);
  };

  const insertSuggestText = (data: string) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    // Merge into an adjacent same-author insertion: the caret's own .docx-ins, or one
    // ending just before the caret.
    const sc = range.startContainer;
    let host = (sc.nodeType === 3 ? sc.parentElement : (sc as Element))?.closest?.(".docx-ins") as HTMLElement | null;
    if (!host && sc.nodeType === 1) {
      const prev = (sc as Element).childNodes[range.startOffset - 1];
      if (prev && prev.nodeType === 1) host = (prev as Element).closest?.(".docx-ins") as HTMLElement | null;
    }
    let after: Node;
    if (host && host.getAttribute("data-author") === sugAuthor()) {
      const tn = document.createTextNode(data);
      if (host.contains(sc)) range.insertNode(tn);
      else host.appendChild(tn);
      after = tn;
    } else {
      const ins = document.createElement("ins");
      ins.className = "docx-ins";
      ins.setAttribute("data-author", sugAuthor());
      ins.setAttribute("data-date", sugDate());
      ins.textContent = data;
      range.insertNode(ins);
      after = ins.firstChild ?? ins; // caret inside the ins, so the next keystroke merges
    }
    const r2 = document.createRange();
    r2.setStartAfter(after);
    r2.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r2);
  };

  const suggestDelete = (range: Range) => {
    if (range.collapsed) return;
    const anc = range.commonAncestorContainer;
    const ancEl = anc.nodeType === 3 ? anc.parentElement : (anc as Element);
    const ownIns = ancEl?.closest?.(".docx-ins");
    const frag = range.extractContents();
    if (ownIns && ownIns.getAttribute("data-author") === sugAuthor()) {
      range.collapse(true); // deleting your own pending insertion -> just drop it
    } else {
      const del = document.createElement("del");
      del.className = "docx-del";
      del.setAttribute("data-author", sugAuthor());
      del.setAttribute("data-date", sugDate());
      del.appendChild(frag);
      range.insertNode(del);
      range.setStartBefore(del);
      range.collapse(true);
    }
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  };

  const onBeforeInput = (e: Event) => {
    if (!suggesting) return;
    const ie = e as InputEvent;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const type = ie.inputType;
    if (type === "insertText" || type === "insertReplacementText" || type === "insertFromPaste") {
      ie.preventDefault();
      const data = ie.data ?? ie.dataTransfer?.getData("text/plain") ?? "";
      if (!data) return;
      const range = sel.getRangeAt(0);
      if (!range.collapsed) suggestDelete(range);
      insertSuggestText(data);
      mark();
      positionCards();
    } else if (type.startsWith("delete")) {
      ie.preventDefault();
      const range = sel.getRangeAt(0);
      let r = range;
      if (range.collapsed) {
        const tr = ie.getTargetRanges?.()[0];
        if (tr) {
          r = document.createRange();
          r.setStart(tr.startContainer, tr.startOffset);
          r.setEnd(tr.endContainer, tr.endOffset);
        }
      }
      if (r.collapsed) return;
      // A deletion that spans a paragraph boundary is a paragraph-mark deletion (merge):
      // mark the first block's mark as deleted instead of merging.
      const sb = blockOf(r.startContainer);
      const eb = blockOf(r.endContainer);
      if (sb && eb && sb !== eb) {
        if (!sb.hasAttribute("data-rev-para")) markPara(sb, "del");
      } else {
        suggestDelete(r);
      }
      mark();
      positionCards();
    }
  };
  for (const r of regions) r.addEventListener("beforeinput", onBeforeInput);
  // Paragraph split (Enter) in suggesting mode -> mark the first half's paragraph mark.
  for (const region of regions)
    region.addEventListener("input", (e) => {
      if (!suggesting || (e as InputEvent).inputType !== "insertParagraph") return;
      const sel = window.getSelection();
      const block = sel && sel.rangeCount ? blockOf(sel.getRangeAt(0).startContainer) : null;
      const first = block?.previousElementSibling as HTMLElement | null;
      if (first && !first.hasAttribute("data-rev-para")) markPara(first, "ins");
      positionCards();
    });

  const unwrap = (el: Element) => {
    while (el.firstChild) el.parentNode?.insertBefore(el.firstChild, el);
    el.remove();
  };
  const mergeWithNext = (el: HTMLElement) => {
    const next = el.nextElementSibling;
    if (next) {
      while (next.firstChild) el.appendChild(next.firstChild);
      next.remove();
    }
  };
  const clearPara = (el: HTMLElement) => {
    el.removeAttribute("data-rev-para");
    el.removeAttribute("data-rev-author");
    el.removeAttribute("data-rev-date");
    el.classList.remove("docx-para-ins", "docx-para-del");
  };
  const resolveChange = (el: Element, accept: boolean) => {
    if (el.classList.contains("docx-rpr-change")) {
      // formatting change: accept keeps the new look; reject restores the old props
      if (accept) unwrap(el);
      else {
        let old: Record<string, unknown> = {};
        try {
          old = JSON.parse(el.getAttribute("data-old") || "{}");
        } catch {
          /* keep default */
        }
        const span = document.createElement("span");
        const style = fmtToStyle(old);
        if (style) span.setAttribute("style", style);
        span.textContent = el.textContent;
        el.replaceWith(span);
      }
    } else if (el.hasAttribute("data-rev-para")) {
      const kind = el.getAttribute("data-rev-para");
      const merge = kind === "ins" ? !accept : accept; // ins-reject and del-accept both merge
      if (merge) mergeWithNext(el as HTMLElement);
      clearPara(el as HTMLElement);
    } else {
      const isDel = el.classList.contains("docx-del");
      if (accept ? isDel : !isDel) el.remove();
      else unwrap(el);
    }
    mark();
    positionCards();
  };
  const resolveAll = (accept: boolean) => {
    for (const el of Array.from(wrap.querySelectorAll(".docx-ins, .docx-del, .docx-rpr-change, [data-rev-para]"))) resolveChange(el, accept);
  };

  // Accept/reject popover when a tracked change is clicked.
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement).closest?.(".docxedit-change-pop")) return;
    document.querySelector(".docxedit-change-pop")?.remove();
    const ch = (e.target as HTMLElement).closest?.(".docx-ins, .docx-del, .docx-rpr-change") as HTMLElement | null;
    if (!ch) return;
    const pop = document.createElement("div");
    pop.className = "docxedit-change-pop";
    const mk = (label: string, title: string, accept: boolean) => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.title = title;
      b.addEventListener("click", (ev) => {
        ev.stopPropagation();
        resolveChange(ch, accept);
        pop.remove();
      });
      return b;
    };
    pop.append(mk("✓", t("accept"), true), mk("✕", t("reject"), false));
    wrap.appendChild(pop);
    const cr = ch.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    pop.style.left = `${Math.min(cr.left - wr.left, wrap.clientWidth - 80)}px`;
    pop.style.top = `${cr.bottom - wr.top + 3}px`;
  });

  const suggestIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" aria-hidden="true">' +
    '<path d="M2 11.5 10.5 3l2.5 2.5L4.5 14H2z"/><line x1="9" y1="4.5" x2="11.5" y2="7"/></svg>';
  const suggestBtn = iconBtn(suggestIcon, t("suggesting"), () => {
    suggesting = !suggesting;
    suggestBtn.classList.toggle("is-on", suggesting);
    activeEl.focus();
  });

  toolbar.append(
    btn("B", t("bold"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold"),
    btn("I", t("italic"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic"),
    btn("U", t("underline"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline"),
    colorInput,
    bgWrap,
    sep(),
    block,
    fontSel,
    sizeSel,
    sep(),
    btn(t("bulletedLabel"), t("bulleted"), () => exec("insertUnorderedList")),
    btn(t("numberedLabel"), t("numbered"), () => exec("insertOrderedList")),
    sep(),
    iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), t("alignLeft"), () => exec("justifyLeft")),
    iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), t("alignCenter"), () => exec("justifyCenter")),
    iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), t("alignRight"), () => exec("justifyRight")),
    iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), t("alignJustify"), () => exec("justifyFull")),
    sep(),
    iconBtn(imgIcon, t("insertImage"), insertImage),
    iconBtn(cmtIcon, t("addComment"), addComment),
    iconBtn(pbIcon, t("insertPageBreak"), insertPageBreak),
    btn(t("link"), t("linkAria"), () => {
      const url = prompt(t("linkPrompt"), "https://");
      if (url === null) return;
      if (url === "") exec("unlink");
      else exec("createLink", url);
    }),
    sep(),
    suggestBtn,
    btn("✓", t("acceptAll"), () => resolveAll(true)),
    btn("✕", t("rejectAll"), () => resolveAll(false)),
  );

  return {
    isDirty() {
      return dirty;
    },
    async getBytes() {
      if (!dirty) return original.slice();
      const editedParts: { path: string; html: string }[] = [];
      if (header && parts.headerPath) editedParts.push({ path: parts.headerPath, html: header.innerHTML });
      if (footer && parts.footerPath) editedParts.push({ path: parts.footerPath, html: footer.innerHTML });
      return htmlToDocx(doc.innerHTML, original, editedParts, {
        reactions: pendingReactions.map((r) => ({ ...r, date: options.now || new Date().toISOString() })),
        replies: pendingReplies,
        done: pendingDone,
        deletedComments,
      });
    },
    destroy() {
      for (const u of fontUrls) URL.revokeObjectURL(u);
      repositionObserver.disconnect();
      wrap.remove();
    },
  };
}
