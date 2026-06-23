import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { t } from "../../core/i18n";
import { bytesToBase64, toHex6, fontSizeToHalfPt, firstFontFamily } from "../../core/util";
import { createRichEditor } from "../../core/editor";
import type {
  Adapter,
  CommentEdits,
  CommentEntry,
  CommentThread,
  EditorOptions,
  NewCommentMeta,
  PageGeometry,
  RichDoc,
  RichEditor,
} from "../../core/types";
import "./docxedit.css";

export type { CommentReaction, CommentEntry, CommentThread } from "../../core/types";

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
// w:jc value -> CSS text-align, and the reverse (CSS text-align -> w:jc) for round-trip.
const JC_TO_ALIGN: Record<string, string> = { both: "justify", distribute: "justify", center: "center", right: "right", end: "right", left: "left", start: "left" };
const JC_BY_ALIGN: Record<string, string> = { justify: "both", center: "center", right: "right", left: "left" };

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

export interface DocxParts {
  body: string;
  header: string;
  footer: string;
  headerPath?: string; // archive key of the header part, for write-back
  footerPath?: string;
  comments: CommentThread[]; // top-level threads in document order, replies nested
  page?: PageGeometry; // page size/margins from w:sectPr, for the paginated view
}

// Convert OOXML twips (1/1440 inch) to CSS px at 96 dpi.
const twipToPx = (v: string | null | undefined): number | undefined => {
  if (!v) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n / 15 : undefined;
};

/** Read page size and margins from a w:sectPr into view geometry (px). */
function parsePageGeometry(sectPr: Element | undefined): PageGeometry | undefined {
  if (!sectPr) return undefined;
  const pgSz = sectPr.getElementsByTagName("w:pgSz")[0];
  const w = twipToPx(pgSz?.getAttributeNS(W, "w") ?? pgSz?.getAttribute("w:w"));
  const h = twipToPx(pgSz?.getAttributeNS(W, "h") ?? pgSz?.getAttribute("w:h"));
  if (!w || !h) return undefined; // no usable size; let the engine apply its default
  const pgMar = sectPr.getElementsByTagName("w:pgMar")[0];
  const m = (a: string) => Math.max(0, twipToPx(pgMar?.getAttributeNS(W, a) ?? pgMar?.getAttribute("w:" + a)) ?? 96);
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") } };
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
    page: parsePageGeometry(sectPr),
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
// docx adapter over the shared engine
// ---------------------------------------------------------------------------

export type DocxEditorOptions = EditorOptions;
export type DocxEditor = RichEditor;

/** Build the OOXML comment markers (range start/end + reference run) for a new comment. */
function docxCommentMarkers(meta: NewCommentMeta): { start: HTMLElement; end: HTMLElement; ref: HTMLElement } {
  const { id, author, date, text, paraId } = meta;
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
  return { start, end, ref };
}

/** Wrap a .docx byte array as an engine adapter: parse, serialize, comment markers, capabilities. */
export function createDocxAdapter(bytes: Uint8Array): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      let parts: DocxParts = { body: "<p><br></p>", header: "", footer: "", comments: [] };
      try {
        parts = docxToParts(bytes);
      } catch (e) {
        console.warn("docxedit: failed to parse document", e);
      }
      let fontCss = "";
      let fontUrls: string[] = [];
      let defaultFontName: string | undefined;
      try {
        const fileMap = unzipSync(bytes);
        const ff = loadEmbeddedFonts(fileMap);
        fontCss = ff.css;
        fontUrls = ff.urls;
        defaultFontName = defaultFont(fileMap);
      } catch {
        /* no embedded fonts */
      }
      return { ...parts, fontCss, fontUrls, defaultFont: defaultFontName };
    },
    write(bodyHtml: string, editedParts: { path: string; html: string }[], edits: CommentEdits): Uint8Array {
      return htmlToDocx(bodyHtml, original, editedParts, edits);
    },
    newCommentMarkers: docxCommentMarkers,
    capabilities: {
      comments: true,
      commentReplies: true,
      commentReactions: true,
      trackChanges: true,
      images: true,
      headerFooter: true,
      pageBreak: true,
      textColor: true,
      fontControls: true,
      alignment: true,
    },
  };
}

/** Mount a .docx editor in `container`: the docx adapter driving the shared engine. */
export function createDocxEditor(container: HTMLElement, bytes: Uint8Array, options: DocxEditorOptions = {}): DocxEditor {
  return createRichEditor(container, createDocxAdapter(bytes), options);
}
