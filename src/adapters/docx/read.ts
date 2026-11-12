// docx READ: parse a .docx archive into editable HTML, header/footer, comments, fonts and
// page geometry. Pure XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { t } from "../../core/i18n";
import { bytesToBase64, imageLayoutAttrs } from "../../core/util";
import type { CommentEntry, CommentThread, PageBorder, PageGeometry } from "../../core/types";
import { W, R, XMLNS, NS_DECLS, IMG_MIME, escapeHtml, escapeAttr, HL_CSS, JC_TO_ALIGN } from "./shared";
import { ommlToMathml } from "./omml";
import { ensureDingsFont } from "./materialdings";
import { SYMBOL, WEBDINGS, WINGDINGS2, WINGDINGS3 } from "./dingbats";
import type { Fmt } from "./shared";
import { readLayout } from "./image-layout";
import { metafileToSvgUrl } from "./metafile";
import { readDocDefaults, readThemeColors, readThemeFonts, rFontsLatin, themeColor, modColor } from "./fonts";
import { shapeHtml } from "./shapes";
import type { ThemeFonts } from "./fonts";

// The theme fonts of the document being read, so run and style fonts given as a theme
// reference (w:asciiTheme="minorHAnsi") resolve to a real typeface. Set by docxToParts.
let THEME: ThemeFonts = {};
let THEME_COLORS: Record<string, string> = {};


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
  numStart: Map<string, number>; // numId -> level-0 starting number (for <ol start>)
  numLevels?: Map<string, LvlFmt[]>; // numId -> per-level kind + marker style
  comments: Map<string, Comment>;
  openComments: Set<string>; // comment ids whose range spans across paragraphs (render state)
  commentOrder: string[]; // ids in document order, for the side panel
  bmNames?: Map<string, string>; // bookmark id -> name, so a w:bookmarkEnd (id only) can carry the name
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
export function loadEmbeddedFonts(files: Record<string, Uint8Array>): { css: string; urls: string[] } {
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

/** The document default font (styles.xml docDefaults + the Normal style, theme resolved). */
export function defaultFont(files: Record<string, Uint8Array>): string | undefined {
  return readDocDefaults(files).font;
}

const emuToPx = (v: string | null): number | undefined => {
  const n = v ? Number(v) : NaN;
  return Number.isFinite(n) ? Math.round(n / 9525) : undefined;
};

/** Render an image-bearing run (w:drawing / w:pict) to an <img>, preserving the run. */
function imageHtml(run: Element, ctx: RenderCtx): string {
  // A shape, text box or group: drawn read-only (its run is preserved verbatim).
  if (run.getElementsByTagName("wps:wsp").length || run.getElementsByTagName("wpg:wgp").length || run.getElementsByTagName("w:txbxContent").length) {
    const sh = shapeHtml(run, passthroughAttr(run), {
      colors: THEME_COLORS,
      blocks: (txbx) => renderBlocks(txbx, ctx),
      imageSrc: (rid) => {
        const tgt = ctx.rels.get(rid);
        const b = tgt ? resolveMedia(tgt, ctx.files) : undefined;
        const mime = IMG_MIME[(tgt?.split(".").pop() ?? "").toLowerCase()];
        return b && mime ? `data:${mime};base64,${bytesToBase64(b)}` : undefined;
      },
    });
    if (sh) return sh;
  }
  const blip = run.getElementsByTagName("a:blip")[0] ?? run.getElementsByTagName("pic:blipFill")[0]?.getElementsByTagName("a:blip")[0];
  let rid = blip?.getAttribute("r:embed") ?? blip?.getAttribute("r:link") ?? undefined;
  if (!rid) rid = run.getElementsByTagName("v:imagedata")[0]?.getAttribute("r:id") ?? undefined;
  const target = rid ? ctx.rels.get(rid) : undefined;
  const bytes = target ? resolveMedia(target, ctx.files) : undefined;
  const ext = (target?.split(".").pop() ?? "").toLowerCase();
  const mime = IMG_MIME[ext];
  const pass = passthroughAttr(run);
  const extent = run.getElementsByTagName("wp:extent")[0];
  let wpx = emuToPx(extent?.getAttribute("cx") ?? null);
  let cy = emuToPx(extent?.getAttribute("cy") ?? null);
  if (!wpx) {
    // A VML picture sizes itself in its shape's style ("width:120pt;height:80pt").
    const st = run.getElementsByTagName("v:shape")[0]?.getAttribute("style") ?? "";
    const dim = (k: string) => { const m = new RegExp(`(?:^|;)\\s*${k}:\\s*([\\d.]+)(pt|px|in)`).exec(st); return m ? Math.round(Number(m[1]) * (m[2] === "pt" ? 4 / 3 : m[2] === "in" ? 96 : 1)) : undefined; };
    wpx = dim("width");
    cy = dim("height");
  }
  if (bytes && (ext === "emf" || ext === "wmf")) {
    // A metafile: drawn as SVG inside the preserved placeholder (which the writer re-emits
    // verbatim, so the picture is saved as it was); else a box of the picture's size.
    const size = wpx && cy ? `width:${wpx}px;height:${cy}px` : "";
    const svg = metafileToSvgUrl(bytes);
    if (svg) return `<span class="docx-img-ph docx-img-meta" contenteditable="false"${pass}><img src="${escapeAttr(svg)}" alt=""${size ? ` style="${size}"` : ""}></span>`;
    return `<span class="docx-img-ph docx-img-vector" contenteditable="false"${pass}${size ? ` style="${size}"` : ""}>${escapeHtml(t("vectorPicture"))}</span>`;
  }
  if (!bytes || !mime) {
    return `<span class="docx-img-ph" contenteditable="false"${pass}>[image]</span>`;
  }
  const dims = wpx ? ` width="${wpx}"${cy ? ` height="${cy}"` : ""}` : "";
  // A floating (anchored) image: surface its wrap mode + position so it renders out of line
  // and the image toolbar can edit it. Inline images carry no layout attributes.
  const lay = imageLayoutAttrs(readLayout(run));
  const altText = run.getElementsByTagName("wp:docPr")[0]?.getAttribute("descr") ?? "";
  return `<img src="data:${mime};base64,${bytesToBase64(bytes)}" alt="${escapeAttr(altText)}" contenteditable="false"${pass}${dims}${lay}>`;
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
// Paragraph shading (a w:shd directly in w:pPr) -> a CSS background-color, or "" if none.
function paragraphShading(pPr: Element | undefined): string {
  const shd = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:shd") : undefined;
  const fill = shd?.getAttribute("w:fill");
  return fill && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill.toLowerCase()}` : "";
}

/** Render a w:tbl to a read-only HTML table; the real table is preserved as passthrough. */
// Preserve a property element (w:tblPr / w:tblGrid / w:tcPr) as a data-attribute, so a
// structurally-edited table can be rebuilt from the DOM without losing its styling.
const propAttr = (name: string, el: Element | undefined): string => (el ? ` data-docx-${name}="${escapeAttr(serializePassthrough(el))}"` : "");

const DOCX_BORDER_STYLE: Record<string, string> = {
  single: "solid", thick: "solid", double: "double", dashed: "dashed", dotted: "dotted",
  dashSmallGap: "dashed", dotDash: "dashed", dotDotDash: "dashed", wave: "solid", doubleWave: "double",
};
// A w:tblBorders/w:tcBorders side element -> "<w>px <style> <#color>", or null if absent/nil.
const docxBorderSpec = (el: Element | undefined): string | null => {
  if (!el) return null;
  const val = el.getAttribute("w:val");
  if (!val || val === "nil" || val === "none") return null;
  const style = DOCX_BORDER_STYLE[val] ?? "solid";
  const w = Math.max(1, Math.round((Number(el.getAttribute("w:sz")) || 4) / 6)); // eighths of pt -> px
  const c = el.getAttribute("w:color");
  const color = !c || c === "auto" ? "#000000" : `#${c.replace(/^#/, "")}`;
  return `${w}px ${style} ${color}`;
};
// Table styles (styles.xml w:type="table"), resolved along basedOn, so a table shows its
// style's borders, shading and header / banded-row looks. Set by docxToParts.
interface TblCond { rPr: Element[]; shd?: string; borders: Map<string, Element> }
interface TblStyle extends TblCond { rowBand: number; colBand: number; cond: Map<string, TblCond>; pCss: string }
let TBL_STYLES = new Map<string, TblStyle>();
// Paragraph styles that number their paragraphs (w:pPr/w:numPr in the style, along basedOn),
// such as Word's "List Number" / "List Bullet": styleId -> numId + level. Set by docxToParts.
let STYLE_NUM = new Map<string, { numId: string; level: number }>();
function readStyleNumbering(stylesXml: Uint8Array | undefined): Map<string, { numId: string; level: number }> {
  const out = new Map<string, { numId: string; level: number }>();
  if (!stylesXml) return out;
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const byId = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("w:style"))) {
    if ((st.getAttribute("w:type") ?? "paragraph") === "paragraph" && st.getAttribute("w:styleId")) byId.set(st.getAttribute("w:styleId")!, st);
  }
  for (const id of byId.keys()) {
    let numId: string | undefined, level: number | undefined;
    for (let st = byId.get(id), seen = new Set<Element>(); st && !seen.has(st); st = byId.get(kid(st, "w:basedOn")?.getAttribute("w:val") ?? "")) {
      seen.add(st);
      const numPr = kid(kid(st, "w:pPr"), "w:numPr");
      if (numId === undefined) numId = kid(numPr, "w:numId")?.getAttribute("w:val") ?? undefined;
      if (level === undefined) { const l = kid(numPr, "w:ilvl")?.getAttribute("w:val"); if (l) level = Number(l); }
    }
    if (numId && numId !== "0") out.set(id, { numId, level: Math.max(0, Math.min(8, level ?? 0)) });
  }
  return out;
}
const kids = (el: Element | undefined, tag: string): Element[] => (el ? Array.from(el.children).filter((c) => c.tagName === tag) : []);
const kid = (el: Element | undefined, tag: string): Element | undefined => kids(el, tag)[0];
const fillOf = (tcPr: Element | undefined): string | undefined => {
  const f = kid(tcPr, "w:shd")?.getAttribute("w:fill");
  return f && /^[0-9a-f]{6}$/i.test(f) ? f.toUpperCase() : undefined;
};
const bordersOf = (el: Element | undefined, into: Map<string, Element>): void => {
  for (const b of Array.from(el?.children ?? [])) into.set(b.tagName, b);
};
function readTableStyles(stylesXml: Uint8Array | undefined): Map<string, TblStyle> {
  const out = new Map<string, TblStyle>();
  if (!stylesXml) return out;
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const byId = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("w:style"))) {
    if (st.getAttribute("w:type") === "table" && st.getAttribute("w:styleId")) byId.set(st.getAttribute("w:styleId")!, st);
  }
  const resolve = (id: string, seen: Set<string>): TblStyle | undefined => {
    if (out.has(id)) return out.get(id);
    const st = byId.get(id);
    if (!st || seen.has(id)) return undefined;
    seen.add(id);
    const base = resolve(kid(st, "w:basedOn")?.getAttribute("w:val") ?? "", seen);
    const tblPr = kid(st, "w:tblPr");
    const num = (tag: string, d: number) => Number(kid(tblPr, tag)?.getAttribute("w:val")) || d;
    const ts: TblStyle = {
      rPr: [...(base?.rPr ?? []), ...kids(st, "w:rPr")],
      shd: fillOf(kid(st, "w:tcPr")) ?? base?.shd,
      borders: new Map(base?.borders),
      rowBand: num("w:tblStyleRowBandSize", base?.rowBand ?? 1),
      colBand: num("w:tblStyleColBandSize", base?.colBand ?? 1),
      cond: new Map(base?.cond),
      pCss: base?.pCss ?? "",
    };
    bordersOf(kid(tblPr, "w:tblBorders"), ts.borders);
    const sp = kid(kid(st, "w:pPr"), "w:spacing");
    if (sp) {
      const parts: string[] = [];
      const before = twipToPx(sp.getAttribute("w:before"));
      const after = twipToPx(sp.getAttribute("w:after"));
      const line = sp.getAttribute("w:line");
      if (before !== undefined) parts.push(`margin-top:${Math.round(before)}px`);
      if (after !== undefined) parts.push(`margin-bottom:${Math.round(after)}px`);
      if (line && (sp.getAttribute("w:lineRule") ?? "auto") === "auto") parts.push(`line-height:${Math.round((Number(line) / 240) * 100) / 100}`);
      if (parts.length) ts.pCss = parts.join(";");
    }
    for (const sp2 of kids(st, "w:tblStylePr")) {
      const type = sp2.getAttribute("w:type");
      if (!type) continue;
      const prev = ts.cond.get(type);
      const c: TblCond = { rPr: [...(prev?.rPr ?? []), ...kids(sp2, "w:rPr")], shd: fillOf(kid(sp2, "w:tcPr")) ?? prev?.shd, borders: new Map(prev?.borders) };
      bordersOf(kid(kid(sp2, "w:tcPr"), "w:tcBorders"), c.borders);
      bordersOf(kid(kid(sp2, "w:tblPr"), "w:tblBorders"), c.borders);
      ts.cond.set(type, c);
    }
    out.set(id, ts);
    return ts;
  };
  for (const id of byId.keys()) resolve(id, new Set());
  return out;
}
// Which of a table style's conditional formats a cell gets, lowest priority first (Word's order).
function cellConds(look: TblLook, ts: TblStyle, r: number, c: number, span: number, rows: number, cols: number): string[] {
  const out: string[] = [];
  const fr = look.firstRow && r === 0, lr = look.lastRow && r === rows - 1;
  const fc = look.firstCol && c === 0, lc = look.lastCol && c + span === cols;
  if (!look.noVBand && !fc && !lc) out.push(Math.floor((c - (look.firstCol ? 1 : 0)) / ts.colBand) % 2 ? "band2Vert" : "band1Vert");
  if (!look.noHBand && !fr && !lr) out.push(Math.floor((r - (look.firstRow ? 1 : 0)) / ts.rowBand) % 2 ? "band2Horz" : "band1Horz");
  if (fc) out.push("firstCol");
  if (lc) out.push("lastCol");
  if (fr) out.push("firstRow");
  if (lr) out.push("lastRow");
  if (fr && lc) out.push("neCell");
  if (fr && fc) out.push("nwCell");
  if (lr && lc) out.push("seCell");
  if (lr && fc) out.push("swCell");
  return out;
}
interface TblLook { firstRow: boolean; lastRow: boolean; firstCol: boolean; lastCol: boolean; noHBand: boolean; noVBand: boolean }
function readTblLook(el: Element | undefined): TblLook {
  // No w:tblLook: Word's default (header row and first column on, column bands off).
  if (!el) return { firstRow: true, lastRow: false, firstCol: true, lastCol: false, noHBand: false, noVBand: true };
  const hex = parseInt(el.getAttribute("w:val") ?? "", 16) || 0;
  const on = (name: string, bit: number) => {
    const v = el.getAttribute(`w:${name}`);
    return v !== null ? v === "1" || v === "true" || v === "on" : (hex & bit) !== 0;
  };
  return { firstRow: on("firstRow", 0x20), lastRow: on("lastRow", 0x40), firstCol: on("firstColumn", 0x80), lastCol: on("lastColumn", 0x100), noHBand: on("noHBand", 0x200), noVBand: on("noVBand", 0x400) };
}
// A table style's run properties as cell CSS (inherited by the cell's text).
function rPrCss(list: Element[], out: Record<string, string>): void {
  for (const rPr of list) {
    for (const c of Array.from(rPr.children)) {
      const v = c.getAttribute("w:val");
      const on = v === null || (v !== "0" && v !== "false");
      if (c.tagName === "w:b") out["font-weight"] = on ? "bold" : "normal";
      else if (c.tagName === "w:i") out["font-style"] = on ? "italic" : "normal";
      else if (c.tagName === "w:color" && v && /^[0-9a-f]{6}$/i.test(v)) out["color"] = `#${v}`;
      else if (c.tagName === "w:caps") out["text-transform"] = on ? "uppercase" : "none";
    }
  }
}

// A w:tblBorders-like element holding the style's sides overlaid with the table's own.
function mergedBorders(style: Map<string, Element> | undefined, own: Element | undefined): Element | undefined {
  if (!style?.size) return own;
  const m = new Map(style);
  bordersOf(own, m);
  const el = (own ?? style.values().next().value!).ownerDocument.createElementNS(W, "w:tblBorders");
  for (const b of m.values()) el.appendChild(b.cloneNode(true));
  return el;
}

// Resolve a cell side: an explicit tcBorders side wins (even nil = off); else the table's
// outer border on an edge cell, or its inside border for an interior edge.
const resolveCellBorder = (tcB: Element | undefined, tblB: Element | undefined, side: string, inside: string, edge: boolean): string | null => {
  const tcEl = tcB?.getElementsByTagName(side)[0];
  if (tcEl) return docxBorderSpec(tcEl);
  return docxBorderSpec(tblB?.getElementsByTagName(edge ? side : inside)[0]);
};

function tableHtml(tbl: Element, ctx: RenderCtx): string {
  const tblPr = tbl.getElementsByTagName("w:tblPr")[0];
  const tblGrid = tbl.getElementsByTagName("w:tblGrid")[0];
  // The table style's borders under the table's own (w:tblBorders sides override per side).
  const ts = TBL_STYLES.get(kid(tblPr, "w:tblStyle")?.getAttribute("w:val") ?? "");
  const look = readTblLook(kid(tblPr, "w:tblLook"));
  const tblB = mergedBorders(ts?.borders, tblPr?.getElementsByTagName("w:tblBorders")[0]);
  // Grid model: assign each cell its grid column (cumulative gridSpan); a vMerge=restart cell
  // spans down over the following rows' vMerge=continue cells in the same column.
  const trs = Array.from(tbl.children).filter((c) => c.tagName === "w:tr");
  const grid = trs.map((tr) => {
    const out: { tc: Element; tcPr: Element | undefined; gridCol: number; gridSpan: number; restart: boolean; cont: boolean }[] = [];
    let gc = 0;
    for (const tc of Array.from(tr.children)) {
      if (tc.tagName !== "w:tc") continue;
      const tcPr = tc.getElementsByTagName("w:tcPr")[0];
      const gridSpan = Number(tcPr?.getElementsByTagName("w:gridSpan")[0]?.getAttribute("w:val")) || 1;
      const vm = tcPr?.getElementsByTagName("w:vMerge")[0];
      const restart = !!vm && vm.getAttribute("w:val") === "restart";
      out.push({ tc, tcPr, gridCol: gc, gridSpan, restart, cont: !!vm && !restart });
      gc += gridSpan;
    }
    return out;
  });
  const rowspanOf = (r: number, gc: number): number => {
    let span = 1;
    for (let rr = r + 1; rr < grid.length; rr++) {
      if (grid[rr].some((c) => c.cont && c.gridCol === gc)) span++;
      else break;
    }
    return span;
  };
  const totalRows = grid.length;
  const totalCols = Math.max(0, ...grid.flat().map((c) => c.gridCol + c.gridSpan));
  // Column widths from w:tblGrid (twips -> px), only when the grid matches the cells.
  const gridW = kids(tblGrid, "w:gridCol").map((g) => twipToPx(g.getAttribute("w:w")) ?? 0);
  const colPx = gridW.length === totalCols && gridW.every((w) => w > 0) ? gridW.map((w) => Math.round(w)) : [];
  let rows = "";
  for (let r = 0; r < grid.length; r++) {
    let cells = "";
    for (const ci of grid[r]) {
      if (ci.cont) continue; // a vMerge continuation cell: covered by the restart above
      // Cell content goes through the same block pipeline as the body, so in-cell
      // alignment, spacing, named styles, lists and the pPr stash render and survive.
      const inner = renderBlocks(ci.tc, ctx);
      const cs = ci.gridSpan > 1 ? ` colspan="${ci.gridSpan}"` : "";
      const rsN = ci.restart ? rowspanOf(r, ci.gridCol) : 1;
      const rs = rsN > 1 ? ` rowspan="${rsN}"` : "";
      // Resolve the cell's real borders (Word's tcBorders-over-tblBorders rules) into the
      // editor's per-side model, so the document's borders show on load.
      // A conditional format's sides (header row rule, ...) sit between the cell's own and the table's.
      const conds = ts ? cellConds(look, ts, r, ci.gridCol, ci.gridSpan, totalRows, totalCols).map((k) => ts.cond.get(k)).filter((c): c is TblCond => !!c) : [];
      const condB = new Map<string, Element>();
      for (const c of conds) for (const [k, v] of c.borders) condB.set(k, v);
      bordersOf(ci.tcPr?.getElementsByTagName("w:tcBorders")[0], condB);
      const tcB = condB.size ? mergedBorders(condB, undefined) : undefined;
      const bAttr = (k: string, spec: string | null): string => (spec ? ` data-rdoc-b${k}="${spec}"` : "");
      const borders =
        bAttr("t", resolveCellBorder(tcB, tblB, "w:top", "w:insideH", r === 0)) +
        bAttr("b", resolveCellBorder(tcB, tblB, "w:bottom", "w:insideH", r + rsN === totalRows)) +
        bAttr("l", resolveCellBorder(tcB, tblB, "w:left", "w:insideV", ci.gridCol === 0)) +
        bAttr("r", resolveCellBorder(tcB, tblB, "w:right", "w:insideV", ci.gridCol + ci.gridSpan === totalCols));
      // Structure is locked (contenteditable=false on the table); each cell's content is its
      // own editable region. Cell shading round-trips via the preserved tcPr, and is shown here:
      // a shaded header with white text was otherwise white on white, an empty-looking row.
      // Display-only cell look (the writer keeps the cell's tcPr): its own shading, else the
      // table style's; the style's text colour / weight; and the grid column width.
      const cellCss: Record<string, string> = {};
      if (ts) rPrCss([...ts.rPr, ...conds.flatMap((c) => c.rPr)], cellCss);
      const fill = fillOf(ci.tcPr) ?? [...conds].reverse().find((c) => c.shd)?.shd ?? ts?.shd;
      if (fill) cellCss["background-color"] = `#${fill}`;
      const wPx = colPx.slice(ci.gridCol, ci.gridCol + ci.gridSpan).reduce((a, b) => a + b, 0);
      if (wPx > 0) cellCss["width"] = `${wPx}px`;
      const cssBody = Object.entries(cellCss).map(([k, v]) => `${k}:${v}`).join(";");
      const bg = cssBody ? ` style="${cssBody}"` : "";
      cells += `<td${cs}${rs}${bg}${propAttr("tcpr", ci.tcPr)}${borders}><div class="docx-cell" contenteditable="true">${inner || "<br>"}</div></td>`;
    }
    rows += `<tr>${cells}</tr>`;
  }
  // Table indent (w:tblInd) -> inline margin-left so it shows and round-trips.
  const tblInd = tblPr?.getElementsByTagName("w:tblInd")[0];
  const indPx = twipToPx(tblInd?.getAttributeNS(W, "w") ?? tblInd?.getAttribute("w:w"));
  const tStyle: string[] = [];
  if (indPx && indPx > 0) tStyle.push(`margin-left:${Math.round(indPx)}px`);
  // Table width (dxa twips or pct fiftieths of a percent) and fixed layout, for display.
  const tblW = kid(tblPr, "w:tblW");
  const tw = tblW?.getAttribute("w:w") ?? "";
  const twType = tblW?.getAttribute("w:type");
  if (twType === "dxa" && Number(tw) > 0) tStyle.push(`width:${Math.round(twipToPx(tw)!)}px`);
  else if (twType === "pct" && tw) tStyle.push(`width:${tw.endsWith("%") ? tw : `${Number(tw) / 50}%`}`);
  if (kid(tblPr, "w:tblLayout")?.getAttribute("w:type") === "fixed" && colPx.length) tStyle.push("table-layout:fixed");
  const indStyle = tStyle.length ? ` style="${tStyle.join(";")}"` : "";
  const tsAttr = ts?.pCss ? ` data-rdoc-tstyle="${escapeAttr(kid(tblPr, "w:tblStyle")!.getAttribute("w:val")!)}"` : "";
  return `<table class="docx-table" contenteditable="false"${passthroughAttr(tbl)}${indStyle}${tsAttr}${propAttr("tblpr", tblPr)}${propAttr("tblgrid", tblGrid)}>${rows}</table>`;
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

/** numId -> ordered flag + the level-0 starting number, best-effort from numbering.xml. The
    start is the numId's lvlOverride/startOverride for level 0, else the abstractNum's level-0
    w:start, else 1 (so a file that begins a list at N round-trips). */
function readNumbering(file: Uint8Array | undefined): { ordered: Map<string, boolean>; start: Map<string, number>; levels: Map<string, LvlFmt[]>; css: string } {
  const ordered = new Map<string, boolean>();
  const start = new Map<string, number>();
  const levels = new Map<string, LvlFmt[]>();
  let css = "";
  if (!file) return { ordered, start, levels, css };
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  const absFmt = new Map<string, boolean>(); // abstractNumId -> ordered
  const absStart = new Map<string, number>(); // abstractNumId -> level-0 start
  const absLevels = new Map<string, LvlFmt[]>(); // abstractNumId -> per-level format
  const levelOf = (l: Element, ruleName: string): LvlFmt => {
    const f = lvlFmt(l, ruleName);
    if (f.rule) css += f.rule;
    const kidVal = (tag: string) => Array.from(l.children).find((c) => c.tagName === tag)?.getAttribute("w:val") ?? undefined;
    return { ordered: f.ordered, listStyle: f.listStyle, cls: f.cls, fmt: kidVal("w:numFmt") ?? "decimal", text: kidVal("w:lvlText") ?? "", start: Number(kidVal("w:start") ?? "1") || 0, lgl: !!Array.from(l.children).find((c) => c.tagName === "w:isLgl"), suff: kidVal("w:suff") ?? "tab" };
  };
  for (const an of Array.from(doc.getElementsByTagName("w:abstractNum"))) {
    const id = an.getAttributeNS(W, "abstractNumId") ?? an.getAttribute("w:abstractNumId");
    if (id) {
      const lv: LvlFmt[] = [];
      for (const l of Array.from(an.getElementsByTagName("w:lvl"))) {
        const i = Number(l.getAttribute("w:ilvl") ?? "0");
        if (!(i >= 0 && i <= 8)) continue;
        lv[i] = levelOf(l, `rdoc-num-${id.replace(/[^\w-]/g, "_")}-${i}`);
      }
      absLevels.set(id, lv);
    }
    const fmt = an.getElementsByTagName("w:numFmt")[0]?.getAttribute("w:val") ?? "";
    if (!id) continue;
    absFmt.set(id, fmt !== "bullet" && fmt !== "none" && fmt !== "");
    const lvl0 = Array.from(an.getElementsByTagName("w:lvl")).find((l) => (l.getAttribute("w:ilvl") ?? "0") === "0");
    const s = Number(lvl0?.getElementsByTagName("w:start")[0]?.getAttribute("w:val"));
    absStart.set(id, Number.isFinite(s) ? s : 1);
  }
  for (const num of Array.from(doc.getElementsByTagName("w:num"))) {
    const numId = num.getAttributeNS(W, "numId") ?? num.getAttribute("w:numId");
    const abs = num.getElementsByTagName("w:abstractNumId")[0]?.getAttribute("w:val") ?? "";
    if (!numId) continue;
    ordered.set(numId, absFmt.get(abs) ?? false);
    const lv = absLevels.get(abs);
    // A w:lvlOverride with its own w:lvl redefines that level for this list only.
    const own = Array.from(num.getElementsByTagName("w:lvlOverride")).flatMap((o) => Array.from(o.children).filter((c) => c.tagName === "w:lvl").map((l) => [Number(o.getAttribute("w:ilvl") ?? "0"), l] as const));
    if (own.length) {
      const merged = [...(lv ?? [])];
      for (const [i, l] of own) if (i >= 0 && i <= 8) merged[i] = levelOf(l, `rdoc-num-o${numId.replace(/[^\w-]/g, "_")}-${i}`);
      levels.set(numId, merged);
    } else if (lv) levels.set(numId, lv);
    const ov = Array.from(num.getElementsByTagName("w:lvlOverride")).find((o) => (o.getAttribute("w:ilvl") ?? "0") === "0");
    const so = Number(ov?.getElementsByTagName("w:startOverride")[0]?.getAttribute("w:val") ?? ov?.getElementsByTagName("w:lvl")[0]?.getElementsByTagName("w:start")[0]?.getAttribute("w:val"));
    start.set(numId, Number.isFinite(so) ? so : absStart.get(abs) ?? 1);
  }
  return { ordered, start, levels, css };
}

/** One list level's kind and CSS marker (list-style-type value), from a w:lvl. */
interface LvlFmt {
  ordered: boolean;
  listStyle?: string;
  cls?: string; // a class on the list carrying a ::marker rule (multi-level labels)
  fmt?: string; // the raw w:numFmt, w:lvlText, w:start, w:isLgl and w:suff (heading numbering)
  text?: string;
  start?: number;
  lgl?: boolean;
  suff?: string;
}

/** CSS numbering the h1-h6 headings whose Heading style carries a w:numPr ("1 Introduction",
    "1.2 Scope"): CSS counters, so the numbers follow headings added or removed while editing,
    and a ::before label, so nothing enters the text that is saved. Deeper levels restart with
    counter-set on the one root counter (a sibling counter-reset does not scope in Chrome). */
function headingNumberCss(styleNum: Map<string, { numId: string; level: number }>, levels: Map<string, LvlFmt[]>): string {
  const heads: { h: number; numId: string; level: number }[] = [];
  for (const [id, sn] of styleNum) {
    const m = /^heading([1-6])$/i.exec(id.replace(/\s+/g, ""));
    if (m && levels.get(sn.numId)?.[sn.level]) heads.push({ h: Number(m[1]), ...sn });
  }
  if (!heads.length) return "";
  const name = (numId: string, lvl: number) => `rdoc-hn-${numId.replace(/[^\w-]/g, "_")}-${lvl}`;
  const startOf = (numId: string, lvl: number) => (levels.get(numId)?.[lvl]?.start ?? 1) - 1;
  const roots = new Set<string>();
  let css = "";
  for (const hd of heads) {
    const lv = levels.get(hd.numId)![hd.level]!;
    if (lv.fmt === "bullet" || lv.fmt === "none") continue;
    roots.add(`${name(hd.numId, hd.level)} ${startOf(hd.numId, hd.level)}`);
    // Deeper numbered headings of the same list restart under this one.
    const deeper = heads.filter((o) => o.numId === hd.numId && o.level > hd.level).map((o) => `${name(o.numId, o.level)} ${startOf(o.numId, o.level)}`);
    css += `.docxedit-doc h${hd.h}{counter-increment:${name(hd.numId, hd.level)}${deeper.length ? `;counter-set:${deeper.join(" ")}` : ""}}\n`;
    const parts: string[] = [];
    for (const tok of (lv.text ?? "").split(/(%[1-9])/)) {
      if (!tok) continue;
      const k = /^%([1-9])$/.exec(tok);
      if (!k) { parts.push(cssStr(tok)); continue; }
      const ref = Number(k[1]) - 1;
      const refLv = levels.get(hd.numId)?.[ref];
      const style = lv.lgl ? "decimal" : NUMFMT_CSS[refLv?.fmt ?? "decimal"] ?? "decimal";
      parts.push(heads.some((o) => o.numId === hd.numId && o.level === ref) ? `counter(${name(hd.numId, ref)},${style})` : cssStr(String(refLv?.start ?? 1)));
    }
    const suffix = lv.suff === "nothing" ? "" : lv.suff === "space" ? " " : "\u2003";
    if (parts.length) css += `.docxedit-doc h${hd.h}::before{content:${parts.join(" ")}${suffix ? " " + cssStr(suffix) : ""}}\n`;
  }
  return roots.size ? `.docxedit-doc{counter-reset:${[...roots].join(" ")}}\n${css}` : "";
}
// Word numFmt -> a CSS list-style-type with the same counter.
const NUMFMT_CSS: Record<string, string> = {
  decimal: "decimal", decimalZero: "decimal-leading-zero", lowerLetter: "lower-alpha", upperLetter: "upper-alpha",
  lowerRoman: "lower-roman", upperRoman: "upper-roman", aiueo: "katakana", aiueoFullWidth: "katakana", iroha: "katakana-iroha",
  irohaFullWidth: "katakana-iroha", ideographDigital: "cjk-decimal", japaneseCounting: "japanese-informal",
  chineseCounting: "simp-chinese-informal", chineseCountingThousand: "simp-chinese-informal", taiwaneseCounting: "trad-chinese-informal",
  koreanCounting: "korean-hangul-formal", hebrew1: "hebrew", arabicAlpha: "arabic-indic", russianLower: "lower-russian" , russianUpper: "upper-russian",
  lowerGreek: "lower-greek", decimalEnclosedCircle: "decimal", decimalFullWidth: "decimal",
};
// Symbol / Wingdings private-use bullet glyphs Word writes in w:lvlText, as Unicode.
const BULLET_PUA: Record<string, string> = { "\uF0B7": "•", "\uF0A7": "▪", "\uF0D8": "➢", "\uF076": "❖", "\uF0FC": "✓", "\uF0A8": "□", "\uF06E": "■", "\uF0E0": "➔" };
const cssStr = (v: string): string => `"${v.replace(/["\\]/g, "\\$&")}"`;
function lvlFmt(l: Element, ruleName: string): LvlFmt & { rule?: string } {
  const fmt = l.getElementsByTagName("w:numFmt")[0]?.getAttribute("w:val") ?? "decimal";
  const text = l.getElementsByTagName("w:lvlText")[0]?.getAttribute("w:val") ?? "";
  if (fmt === "none") return { ordered: false, listStyle: "none" };
  if (fmt === "bullet") {
    const ch = BULLET_PUA[text] ?? (/^[\uE000-\uF8FF]?$/.test(text) ? "•" : text);
    const kw = ch === "•" ? "disc" : ch === "o" || ch === "◦" ? "circle" : ch === "▪" || ch === "■" ? "square" : undefined;
    return { ordered: false, listStyle: kw ?? cssStr(ch + " ") };
  }
  const base = NUMFMT_CSS[fmt] ?? "decimal";
  // A single-counter label ("%1)", "(%2)", "%1.") keeps its punctuation via @counter-style;
  // any other multi-counter one falls back to its own counter with the default suffix.
  // "%1.%2." (every level decimal-like, in order): the nested counters joined by dots.
  const ml = /^%1((?:\.%[2-9])+)([^%]*)$/.exec(text);
  if (ml && base === "decimal" && ml[1]!.split(".%").slice(1).every((n, k) => Number(n) === k + 2)) {
    return { ordered: true, cls: ruleName, rule: `.docxedit-doc ol.${ruleName}>li::marker{content:counters(list-item,".") ${cssStr(ml[2]! === "." ? ". " : ml[2]! + " ")}}\n` };
  }
  const m = /^([^%]*)%[1-9]([^%]*)$/.exec(text);
  if (!m || (m[1] === "" && m[2] === ".")) return { ordered: true, listStyle: base === "decimal" ? undefined : base };
  const rule = `@counter-style ${ruleName}{system:extends ${base};prefix:${cssStr(m[1]!)};suffix:${cssStr(m[2]! + " ")};}\n`;
  return { ordered: true, listStyle: ruleName, rule };
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
  if (f.vertAlign === "super") s = `<sup>${s}</sup>`;
  else if (f.vertAlign === "sub") s = `<sub>${s}</sub>`;
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

// Which rPr/pPr content the editor models (rebuilt from the DOM on save). Anything
// else is stashed on the emitted HTML so the writer can merge it back, instead of
// silently dropping it document-wide on every save.
const MODELED_RPR = new Set(["w:rStyle", "w:b", "w:i", "w:strike", "w:vertAlign", "w:color", "w:sz", "w:szCs", "w:highlight", "w:shd", "w:rPrChange"]);
function rPrHasUnmodeled(rPr: Element): boolean {
  for (const c of Array.from(rPr.children)) {
    const tag = c.tagName;
    if (MODELED_RPR.has(tag)) continue;
    if (tag === "w:u") {
      // A plain single underline is fully modeled; any other flavour/colour must survive.
      if ((c.getAttribute("w:val") ?? "single") === "single" && c.attributes.length <= 1) continue;
      return true;
    }
    if (tag === "w:rFonts") {
      for (const a of Array.from(c.attributes)) if (a.name !== "w:ascii" && a.name !== "w:hAnsi") return true;
      continue;
    }
    return true; // smallCaps, caps, dstrike, spacing, kern, position, vanish, lang, bCs, ...
  }
  return false;
}
const MODELED_PPR = new Set(["w:pStyle", "w:numPr", "w:pBdr", "w:shd", "w:tabs", "w:jc", "w:rPr", "w:sectPr", "w:pageBreakBefore"]);
function pPrHasUnmodeled(pPr: Element): boolean {
  for (const c of Array.from(pPr.children)) {
    const tag = c.tagName;
    if (MODELED_PPR.has(tag)) continue;
    if (tag === "w:spacing") {
      const auto = (c.getAttribute("w:lineRule") ?? "auto") === "auto";
      for (const a of Array.from(c.attributes)) {
        if (a.name === "w:before" || a.name === "w:after") continue;
        if ((a.name === "w:line" || a.name === "w:lineRule") && auto) continue;
        return true; // exact/atLeast line rule, beforeLines, afterAutospacing, ...
      }
      continue;
    }
    if (tag === "w:ind") {
      for (const a of Array.from(c.attributes)) if (a.name !== "w:left" && a.name !== "w:start") return true;
      continue;
    }
    return true; // keepNext, keepLines, widowControl, outlineLvl, bidi, contextualSpacing, ...
  }
  return false;
}

/** Read a w:rPr element into a Fmt (used for the run's props and rPrChange's old props). */
function readFmt(rPr: Element | undefined): Fmt {
  const colorEl = rPr?.getElementsByTagName("w:color")[0];
  let color = attrVal(rPr, "w:color");
  if ((!color || color === "auto") && colorEl?.getAttribute("w:themeColor")) {
    const base = themeColor(THEME_COLORS, colorEl.getAttribute("w:themeColor"));
    const byte = (a: string) => { const v = colorEl.getAttribute(a); return v && /^[0-9a-f]{2}$/i.test(v) ? parseInt(v, 16) : undefined; };
    const shade = byte("w:themeShade"), tint = byte("w:themeTint");
    if (base) color = modColor(base, { shade: shade !== undefined ? (shade / 255) * 100000 : undefined, tint: tint !== undefined ? (tint / 255) * 100000 : undefined });
  }
  const highlight = attrVal(rPr, "w:highlight");
  const sz = attrVal(rPr, "w:sz");
  const shd = rPr?.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill");
  return {
    b: onFlag(rPr, "w:b"),
    i: onFlag(rPr, "w:i"),
    u: onFlag(rPr, "w:u"),
    strike: onFlag(rPr, "w:strike"),
    vertAlign: ((va) => (va === "superscript" ? "super" : va === "subscript" ? "sub" : undefined))(attrVal(rPr, "w:vertAlign")),
    color: color && color !== "auto" && /^[0-9a-fA-F]{6}$/.test(color) ? color.toUpperCase() : undefined,
    highlight: highlight && highlight !== "none" ? highlight : undefined,
    shading: shd && shd !== "auto" && /^[0-9a-fA-F]{6}$/.test(shd) ? shd.toUpperCase() : undefined,
    sizeHalfPt: sz && /^\d+$/.test(sz) ? Number(sz) : undefined,
    font: rFontsLatin(rPr?.getElementsByTagName("w:rFonts")[0], THEME),
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
    else if (el.tagName === "w:tab") text += '<span class="docx-tab" data-docx-tab="1" contenteditable="false">\t</span>';
    else if (el.tagName === "w:br") {
      if (el.getAttribute("w:type") === "page") {
        flush();
        out += pageBreakHtml("manual");
      } else text += "<br>";
    } else if (el.tagName === "w:lastRenderedPageBreak") {
      flush();
      out += pageBreakHtml("auto");
    } else if (el.tagName === "w:ruby") {
      // Where Word actually puts furigana: inside the run, not beside it.
      flush();
      out += rubyHtml(el);
    }
  }
  flush();
  // Unmodeled run properties (small caps, character spacing, eastAsia fonts, underline
  // flavours, lang, ...) ride along so the writer can merge them back on save.
  if (rPr && out && rPrHasUnmodeled(rPr)) out = `<span data-docx-rpr="${escapeAttr(serializePassthrough(rPr))}"${rPrDisplayAttrs(rPr)}>${out}</span>`;
  // A named character style on the run, so it round-trips and picks up the injected style CSS.
  const rStyle = attrVal(rPr, "w:rStyle");
  if (rStyle && out) out = `<span data-rdoc-cstyle="${escapeAttr(rStyle)}">${out}</span>`;
  if (rPrChange && out) {
    const author = rPrChange.getAttribute("w:author") ?? "";
    const date = rPrChange.getAttribute("w:date") ?? "";
    out = `<span class="docx-rpr-change" data-old="${escapeAttr(JSON.stringify(readFmt(oldRPr)))}" data-rev-author="${escapeAttr(author)}" data-rev-date="${escapeAttr(date)}" title="${escapeAttr(author ? author + (date ? " – " + date.slice(0, 10) : "") : "")}">${out}</span>`;
  }
  return out;
}

// How the unmodeled run properties look: classes and styles the writer never reads back
// (the stashed rPr is the source of truth), so showing them cannot change the file.
const U_CLASS: Record<string, string> = { double: "double", thick: "thick", dotted: "dotted", dottedHeavy: "dotted", dash: "dashed", dashedHeavy: "dashed", dashLong: "dashed", dashLongHeavy: "dashed", dotDash: "dashed", dotDotDash: "dashed", wave: "wavy", wavyHeavy: "wavy", wavyDouble: "wavy" };
function rPrDisplayAttrs(rPr: Element): string {
  const own = (tag: string) => Array.from(rPr.children).find((c) => c.tagName === tag);
  const flag = (tag: string) => { const e = own(tag); if (!e) return false; const v = e.getAttribute("w:val"); return v !== "false" && v !== "0" && v !== "off"; };
  const cls: string[] = [];
  if (flag("w:caps")) cls.push("docx-caps");
  if (flag("w:smallCaps")) cls.push("docx-smallcaps");
  if (flag("w:dstrike")) cls.push("docx-dstrike");
  if (flag("w:vanish")) cls.push("docx-hidden");
  const u = U_CLASS[own("w:u")?.getAttribute("w:val") ?? ""];
  if (u) cls.push(`docx-u-${u}`);
  const st: string[] = [];
  const spacing = Number(own("w:spacing")?.getAttribute("w:val"));
  if (spacing) st.push(`letter-spacing:${spacing / 20}pt`);
  const pos = Number(own("w:position")?.getAttribute("w:val"));
  if (pos) st.push(`vertical-align:${pos / 2}pt`);
  return (cls.length ? ` class="${cls.join(" ")}"` : "") + (st.length ? ` style="${st.join(";")}"` : "");
}

// Furigana: a w:ruby (base + reading) -> an HTML <ruby>base<rt>reading</rt></ruby>, which the
// browser renders natively above (or, in vertical writing, beside) the base. The w:rubyPr is
// stashed so alignment / sizing round-trips.
function rubyHtml(el: Element): string {
  const runsOf = (c: Element | undefined): string =>
    c ? Array.from(c.children).filter((x) => x.tagName === "w:r").map((r) => runToHtml(r)).join("") || escapeHtml(c.textContent ?? "") : "";
  const baseHtml = runsOf(el.getElementsByTagName("w:rubyBase")[0]);
  const rtHtml = runsOf(el.getElementsByTagName("w:rt")[0]);
  const rubyPr = el.getElementsByTagName("w:rubyPr")[0];
  const prAttr = rubyPr ? ` data-docx-rubypr="${escapeAttr(serializePassthrough(rubyPr))}"` : "";
  return `<ruby${prAttr}>${baseHtml || "&#8203;"}<rt>${rtHtml}</rt></ruby>`;
}

// A zero-width passthrough for a comment range marker, so it round-trips on save.
const commentMark = (el: Element): string =>
  `<span class="docx-cmark" contenteditable="false" data-comment-id="${escapeAttr(el.getAttribute("w:id") ?? "")}"${passthroughAttr(el)}></span>`;

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
const MODELED_RUN_CHILDREN = new Set(["w:rPr", "w:t", "w:delText", "w:tab", "w:br", "w:lastRenderedPageBreak", "w:ruby"]);
const runIsModeled = (run: Element): boolean => Array.from(run.children).every((c) => MODELED_RUN_CHILDREN.has(c.tagName));
// Preserve an unmodelled element (bookmark, field, content control, math, ...) verbatim:
// show its text read-only and carry the original XML so it round-trips on save.
const inlinePassthrough = (el: Element): string => `<span class="docx-pass" contenteditable="false"${passthroughAttr(el)}>${escapeHtml(el.textContent ?? "")}</span>`;
// An equation (m:oMath): rendered as MathML, with the original OMML kept in data-docx-xml so an
// un-edited equation rewrites verbatim. If conversion fails, fall back to opaque passthrough.
function eqHtml(oMath: Element): string {
  let mathml = "";
  try { mathml = ommlToMathml(oMath); } catch { /* fall through to passthrough */ }
  if (!mathml) return inlinePassthrough(oMath);
  return `<span class="docx-eq" data-rdoc-eq contenteditable="false"${passthroughAttr(oMath)}>${mathml}</span>`;
}

// Pick the symbol-font -> Unicode table for a font (Symbol, Webdings, Wingdings 2/3). Wingdings
// itself is not here: it renders via the bundled MaterialDings font instead.
function unicodeMap(font: string): Record<number, string> | undefined {
  if (/^symbol$/i.test(font)) return SYMBOL;
  if (/^webdings$/i.test(font)) return WEBDINGS;
  if (/^wingdings 2$/i.test(font)) return WINGDINGS2;
  if (/^wingdings 3$/i.test(font)) return WINGDINGS3;
  return undefined;
}
// A run carrying a w:sym (a symbol-font glyph): show the glyph while stashing the whole run so it
// rewrites verbatim. Symbol / Webdings / Wingdings 2-3 map to portable Unicode; Wingdings renders via
// the bundled MaterialDings open replacement; any other font is a best effort in the named font.
// Without this the run is an empty passthrough span, i.e. invisible.
function symHtml(run: Element, sym: Element): string {
  const font = sym.getAttributeNS(W, "font") ?? sym.getAttribute("w:font") ?? "";
  const raw = parseInt(sym.getAttributeNS(W, "char") ?? sym.getAttribute("w:char") ?? "", 16);
  if (!Number.isFinite(raw)) return inlinePassthrough(run);
  const stash = passthroughAttr(run);
  const mapped = unicodeMap(font)?.[raw & 0xff];
  if (mapped) return `<span class="docx-sym" contenteditable="false"${stash}>${escapeHtml(mapped)}</span>`;
  if (/^wingdings$/i.test(font)) {
    // MaterialDings maps at the classic low codepoints (0x21-0xFF), so use the low byte of the PUA char.
    ensureDingsFont();
    return `<span class="docx-sym docx-dings" contenteditable="false"${stash} style="font-family:'MaterialDings'">${escapeHtml(String.fromCharCode(raw & 0xff))}</span>`;
  }
  const style = ` style="font-family:'${escapeAttr(font.replace(/'/g, ""))}'"`;
  return `<span class="docx-sym" contenteditable="false"${stash}${style}>${escapeHtml(String.fromCharCode(raw))}</span>`;
}

function inlineToHtml(p: Element, ctx: RenderCtx): string {
  let html = "";
  for (const id of ctx.openComments) html += `<span class="docx-comment" data-comment-id="${escapeAttr(id)}">`;
  let field: FieldState | null = null;
  for (const node of Array.from(p.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    // A complex field in progress: consume every node up to its matching end, then emit the field.
    if (field) {
      const fc = el.tagName === "w:r" ? fldCharType(el) : null;
      field.raw += inlinePassthrough(el);
      if (fc === "begin") field.depth++;
      else if (fc === "end") { if (--field.depth === 0) { html += finishField(field); field = null; } }
      else if (fc === "separate" && field.depth === 1) field.phase = "result";
      else if (field.depth === 1) {
        if (field.phase === "instr") field.instr += instrTextOf(el);
        else field.result += el.textContent ?? "";
      }
      continue;
    }
    if (el.tagName === "w:r" && fldCharType(el) === "begin") {
      field = { depth: 1, phase: "instr", instr: "", result: "", raw: inlinePassthrough(el) };
      continue;
    }
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
      const noteRef = el.getElementsByTagName("w:footnoteReference")[0] ?? el.getElementsByTagName("w:endnoteReference")[0];
      const sym = el.getElementsByTagName("w:sym")[0];
      if (el.getElementsByTagName("w:commentReference").length) html += commentRefHtml(el, ctx);
      else if (noteRef) html += noteRefHtml(noteRef.getAttribute("w:id") ?? "", noteRef.tagName === "w:endnoteReference" ? "endnote" : "footnote");
      else if (hasDrawing(el)) html += imageHtml(el, ctx);
      else if (sym && !el.getElementsByTagName("w:t").length) html += symHtml(el, sym); // a symbol-font glyph
      else if (!runIsModeled(el)) html += inlinePassthrough(el); // run holds a field char, symbol, footnote ref, ...
      else html += runToHtml(el);
    } else if (el.tagName === "w:ruby") {
      html += rubyHtml(el);
    } else if (el.tagName === "w:bookmarkStart") {
      const name = el.getAttribute("w:name") ?? "";
      const id = el.getAttribute("w:id");
      if (name && name !== "_GoBack") { if (id) ctx.bmNames?.set(id, name); html += bookmarkStartHtml(name, id); } // skip Word's transient bookmark
    } else if (el.tagName === "w:bookmarkEnd") {
      const id = el.getAttribute("w:id");
      const name = id ? ctx.bmNames?.get(id) : undefined;
      if (name) html += bookmarkEndHtml(id, name); // skip ends of dropped bookmarks (e.g. _GoBack)
    } else if (el.tagName === "w:fldSimple") {
      const instr = el.getAttribute("w:instr") ?? "";
      const ref = parseRefInstr(instr);
      const seq = parseSeqInstr(instr);
      const info = infoFieldKind(instr);
      if (ref) html += xrefHtml(ref.name, ref.fmt, el.textContent ?? "");
      else if (seq) html += seqFieldHtml(seq, el.textContent ?? "");
      else if (info) html += infoFieldHtml(info, el.textContent ?? "");
      else html += inlinePassthrough(el); // an unmodeled simple field stays as-is
    } else if (el.tagName === "w:hyperlink") {
      const id = el.getAttributeNS(R, "id") ?? el.getAttribute("r:id") ?? "";
      const anchor = el.getAttribute("w:anchor"); // an internal link to a bookmark
      const href = id ? (ctx.rels.get(id) ?? "") : anchor ? "#" + anchor : "";
      let inner = "";
      for (const r of Array.from(el.getElementsByTagName("w:r"))) inner += hasDrawing(r) ? imageHtml(r, ctx) : runToHtml(r);
      html += `<a href="${escapeHtml(href)}">${inner}</a>`;
    } else if (el.localName === "oMath") {
      html += eqHtml(el);
    } else if (el.localName === "oMathPara") {
      for (const om of Array.from(el.children)) if (om.localName === "oMath") html += eqHtml(om);
    } else if (el.tagName !== "w:pPr") {
      // Anything else (content controls, moves, unknown fields, ...) -> preserve.
      html += inlinePassthrough(el);
    }
  }
  if (field) html += field.raw; // a malformed field with no end: keep its runs verbatim
  for (let i = 0; i < ctx.openComments.size; i++) html += "</span>"; // closed here, reopened next paragraph
  return html;
}

interface PInfo {
  heading: number; // 0 = not a heading
  isList: boolean;
  ordered: boolean;
  listStyle?: string; // the level's CSS list-style-type, when not the browser default
  listCls?: string; // the level's marker class (multi-level labels)
  numId?: string; // the list's numId (to track restart/continue across lists)
  level: number; // list nesting level (w:ilvl), 0 = top
  align?: string; // CSS text-align
  indentPx?: number; // w:ind left indent, in px
  lineHeight?: number; // w:spacing line (auto rule), as a multiple
  spaceBeforePx?: number; // w:spacing @w:before, in px
  spaceAfterPx?: number; // w:spacing @w:after, in px
  border?: string; // CSS border declaration block
  shading?: string; // CSS background-color from w:shd (paragraph shading)
  styleId?: string; // a named paragraph style (w:pStyle val) that is not a heading
  pageBreakBefore: boolean;
  keepNext?: boolean; // w:keepNext: paginate with the next paragraph
  styleNum?: string; // "numId:level" when the numbering comes from the paragraph style
  dropCap?: string; // w:framePr/@w:dropCap: "drop" (in the text) or "margin"
  frame?: { side: "left" | "right" | "center"; widthPx?: number }; // a w:framePr text frame (not a drop cap)
  revPara?: "ins" | "del"; // paragraph-mark revision (split/merge)
  revAuthor?: string;
  revDate?: string;
  sectPr?: string; // a mid-document section break (w:pPr/w:sectPr), preserved verbatim
  sectBreak?: boolean; // that section starts a new page (type != continuous) -> show a page break
  secGeom?: string; // JSON page geometry of the section ending here, for per-section rendering
  secHeaderRid?: string; // r:id of this section's header part (keys into sectionBands)
  secFooterRid?: string; // r:id of this section's footer part
  tabStops?: string; // JSON [{pos,val,leader}] of the paragraph's custom tab stops (w:tabs)
  pPrXml?: string; // the original w:pPr, stashed when it holds properties the editor does not model
}
/** The r:id of a section's default header/footer reference (keys per-section header/footer). */
function refRid(sectPr: Element | undefined, tag: string): string | undefined {
  if (!sectPr) return undefined;
  const refs = Array.from(sectPr.getElementsByTagName(tag));
  const pick = refs.find((r) => (r.getAttribute("w:type") ?? "default") === "default") ?? refs[0];
  return pick?.getAttributeNS(R, "id") ?? pick?.getAttribute("r:id") ?? undefined;
}
// A paragraph's / style's custom tab stops (w:pPr/w:tabs), in px, skipping "clear" entries.
function parseWTabs(pPr: Element | undefined): { pos: number; val: string; leader?: string }[] {
  const tabsEl = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:tabs") : undefined;
  if (!tabsEl) return [];
  return Array.from(tabsEl.getElementsByTagName("w:tab"))
    .filter((tb) => (tb.getAttribute("w:val") ?? "left") !== "clear")
    .map((tb) => ({ pos: Math.round(twipToPx(tb.getAttribute("w:pos")) ?? 0), val: tb.getAttribute("w:val") ?? "left", leader: tb.getAttribute("w:leader") ?? undefined }))
    .filter((s) => s.pos > 0);
}
function paragraphInfo(p: Element, numbering: Map<string, boolean>, numLevels?: Map<string, LvlFmt[]>): PInfo {
  const pPr = p.getElementsByTagName("w:pPr")[0];
  let heading = 0;
  const rawStyle = pPr?.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") ?? "";
  const styleVal = rawStyle.replace(/\s+/g, "");
  const hm = /^heading([1-9])$/i.exec(styleVal);
  if (hm) heading = Number(hm[1]);
  const directNumPr = pPr?.getElementsByTagName("w:numPr")[0];
  // Numbering from the paragraph style (not for headings, which stay headings).
  const styleNum = !directNumPr && !heading ? STYLE_NUM.get(rawStyle) : undefined;
  const numPr = directNumPr ?? (styleNum ? pPr : undefined);
  const numId = directNumPr?.getElementsByTagName("w:numId")[0]?.getAttribute("w:val") ?? styleNum?.numId ?? "";
  const level = styleNum ? styleNum.level : Math.max(0, Math.min(8, Number(numPr?.getElementsByTagName("w:ilvl")[0]?.getAttribute("w:val")) || 0));
  const jc = pPr?.getElementsByTagName("w:jc")[0]?.getAttribute("w:val") ?? "";
  const ind = pPr?.getElementsByTagName("w:ind")[0];
  const indentPx = twipToPx(ind?.getAttribute("w:left") ?? ind?.getAttribute("w:start"));
  const sp = pPr?.getElementsByTagName("w:spacing")[0];
  const line = sp?.getAttribute("w:line");
  const rule = sp?.getAttribute("w:lineRule") ?? "auto";
  const lineHeight = line && rule === "auto" ? Math.round((Number(line) / 240) * 100) / 100 : undefined;
  const spaceBeforePx = twipToPx(sp?.getAttribute("w:before"));
  const spaceAfterPx = twipToPx(sp?.getAttribute("w:after"));
  // Paragraph-mark revision lives in pPr > rPr > w:ins / w:del.
  const markRPr = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:rPr") : undefined;
  const mark = markRPr && (Array.from(markRPr.children).find((c) => c.tagName === "w:ins" || c.tagName === "w:del") as Element | undefined);
  // A mid-document section break: w:sectPr inside this paragraph's w:pPr (the final section's
  // sectPr is a child of w:body and handled separately). Preserve it so editing keeps sections.
  const sectEl = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:sectPr") : undefined;
  const sectType = sectEl?.getElementsByTagName("w:type")[0]?.getAttribute("w:val");
  // Custom tab stops (w:pPr/w:tabs); skip "clear" entries. Preserved so an edit keeps them.
  const stops = parseWTabs(pPr);
  const kn = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:keepNext") : undefined;
  const keepNext = !!kn && !/^(0|false|off)$/.test(kn.getAttribute("w:val") ?? "");
  const fp = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:framePr") : undefined;
  const dc = fp?.getAttribute("w:dropCap");
  const xAlign = fp?.getAttribute("w:xAlign");
  const fw = twipToPx(fp?.getAttribute("w:w"));
  const frame = fp && (!dc || dc === "none") ? { side: (xAlign === "right" || xAlign === "outside" ? "right" : xAlign === "center" ? "center" : "left") as "left" | "right" | "center", widthPx: fw && fw > 0 ? Math.round(fw) : undefined } : undefined;
  return {
    frame,
    dropCap: dc === "drop" || dc === "margin" ? dc : undefined,
    keepNext,
    styleNum: styleNum ? `${styleNum.numId}:${styleNum.level}` : undefined,
    heading,
    isList: !!numPr,
    ordered: numLevels?.get(numId)?.[level]?.ordered ?? numbering.get(numId) ?? false,
    listStyle: numLevels?.get(numId)?.[level]?.listStyle,
    listCls: numLevels?.get(numId)?.[level]?.cls,
    numId: numId || undefined,
    level,
    align: JC_TO_ALIGN[jc],
    indentPx: indentPx && indentPx > 0 ? Math.round(indentPx) : undefined,
    lineHeight,
    spaceBeforePx: spaceBeforePx === undefined ? undefined : Math.round(spaceBeforePx),
    spaceAfterPx: spaceAfterPx === undefined ? undefined : Math.round(spaceAfterPx),
    styleId: !hm && rawStyle ? rawStyle : undefined,
    border: paragraphBorderStyle(pPr),
    shading: paragraphShading(pPr),
    pageBreakBefore: onFlag(pPr, "w:pageBreakBefore"),
    revPara: mark ? (mark.tagName === "w:del" ? "del" : "ins") : undefined,
    revAuthor: mark?.getAttribute("w:author") ?? undefined,
    revDate: mark?.getAttribute("w:date") ?? undefined,
    sectPr: sectEl ? serializePassthrough(sectEl) : undefined,
    sectBreak: sectEl ? sectType !== "continuous" : undefined,
    secGeom: sectEl ? secGeomJson(sectEl) : undefined,
    secHeaderRid: refRid(sectEl, "w:headerReference"),
    secFooterRid: refRid(sectEl, "w:footerReference"),
    tabStops: stops.length ? JSON.stringify(stops) : undefined,
    pPrXml: pPr && pPrHasUnmodeled(pPr) ? serializePassthrough(pPr) : undefined,
  };
}
// Compact JSON page geometry of a section, read from its w:sectPr, for per-section rendering.
function secGeomJson(sectPr: Element): string | undefined {
  const g = parsePageGeometry(sectPr);
  if (!g) return undefined;
  return JSON.stringify({ w: g.widthPx, h: g.heightPx, mt: g.margin.top, mr: g.margin.right, mb: g.margin.bottom, ml: g.margin.left, cols: g.columns, colGap: g.columnGapPx, vertical: g.vertical, rtl: g.rtl });
}
function blockStyleAttr(info: PInfo): string {
  const parts: string[] = [];
  if (info.align && info.align !== "left") parts.push(`text-align:${info.align}`);
  if (info.indentPx) parts.push(`margin-left:${info.indentPx}px`);
  if (info.lineHeight) parts.push(`line-height:${info.lineHeight}`);
  if (info.spaceBeforePx !== undefined) parts.push(`margin-top:${info.spaceBeforePx}px`);
  if (info.spaceAfterPx !== undefined) parts.push(`margin-bottom:${info.spaceAfterPx}px`);
  if (info.shading) parts.push(`background-color:${info.shading}`);
  if (info.border) parts.push(info.border);
  if (info.frame?.widthPx) parts.push(`width:${info.frame.widthPx}px`); // display only: the writer keeps the framePr
  const style = parts.length ? ` style="${parts.join(";")}"` : "";
  const styleAttr = info.styleId ? ` data-rdoc-style="${escapeAttr(info.styleId)}"` : "";
  const sectAttr = info.sectPr ? ` data-docx-sectpr="${escapeAttr(info.sectPr)}"` : "";
  const secGeomAttr = info.secGeom ? ` data-rdoc-secbreak="${escapeAttr(info.secGeom)}"` : "";
  const secHKey = info.secHeaderRid ? ` data-rdoc-secheaderkey="${escapeAttr(info.secHeaderRid)}"` : "";
  const secFKey = info.secFooterRid ? ` data-rdoc-secfooterkey="${escapeAttr(info.secFooterRid)}"` : "";
  const tabAttr = (info.tabStops ? ` data-rdoc-tabstops="${escapeAttr(info.tabStops)}"` : "") + (info.keepNext ? " data-rdoc-keepnext" : "") + (info.dropCap ? ` data-rdoc-dropcap="${info.dropCap}"` : "") + (info.frame ? ` data-rdoc-frame="${info.frame.side}"` : "");
  // Unmodeled paragraph properties + the list's original numbering id, for merge-back on save.
  const pprAttr = info.pPrXml ? ` data-docx-ppr="${escapeAttr(info.pPrXml)}"` : "";
  const numAttr = (info.isList && info.numId ? ` data-docx-numid="${escapeAttr(info.numId)}"` : "") + (info.styleNum ? ` data-docx-stylenum="${info.styleNum}"` : "");
  const sec = sectAttr + secGeomAttr + secHKey + secFKey + pprAttr + numAttr;
  if (!info.revPara) return style + styleAttr + sec + tabAttr;
  return `${style}${styleAttr}${sec}${tabAttr} class="docx-para-${info.revPara}" data-rev-para="${info.revPara}" data-rev-author="${escapeAttr(info.revAuthor ?? "")}" data-rev-date="${escapeAttr(info.revDate ?? "")}"`;
}

/** Build nested <ul>/<ol> from a flat list of items carrying their nesting level (w:ilvl).
    A deeper item opens new lists inside the current <li>; a shallower one closes them. */
function buildNestedList(items: { level: number; ordered: boolean; li: string; listStyle?: string; listCls?: string }[], rootStart = 1): string {
  let html = "";
  const open: string[] = []; // stack of open list tags
  for (const it of items) {
    const depth = it.level + 1; // number of lists that should be open
    while (open.length > depth) html += `</li></${open.pop()}>`;
    if (open.length === depth) html += "</li>"; // sibling at the same level
    while (open.length < depth) {
      const tag = it.ordered ? "ol" : "ul";
      const startAttr = open.length === 0 && tag === "ol" && rootStart > 1 ? ` start="${rootStart}"` : ""; // restart/continue
      const lsAttr = it.listStyle && open.length === it.level && !(open.length === 0 && it.listStyle === "disc") ? ` style="list-style-type:${escapeAttr(it.listStyle)}"` : "";
      const clsAttr = it.listCls && open.length === it.level ? ` class="${it.listCls}"` : "";
      html += `<${tag}${startAttr}${lsAttr}${clsAttr}>`;
      open.push(tag);
    }
    html += `<li${it.li}`;
  }
  while (open.length) html += `</li></${open.pop()}>`;
  return html;
}

/** Render a block container (w:body or a header/footer root) to HTML. */
function renderBlocks(container: Element, ctx: RenderCtx): string {
  let html = "";
  let listBuf: { level: number; ordered: boolean; li: string; numId?: string; listStyle?: string; listCls?: string }[] = [];
  // Running count of level-0 items per ordered numId across the document, so a list that
  // continues an earlier one (same numId) gets the right <ol start>, not a fresh 1.
  const numRun = new Map<string, number>();
  const flushList = () => {
    if (!listBuf.length) return;
    let rootStart = 1;
    const first = listBuf[0]!;
    if (first.ordered && first.numId) {
      const explicit = ctx.numStart.get(first.numId) ?? 1;
      const prev = numRun.get(first.numId);
      rootStart = prev !== undefined ? prev + 1 : explicit;
      const count0 = listBuf.filter((i) => i.level === 0 && i.numId === first.numId).length;
      numRun.set(first.numId, rootStart - 1 + count0);
    }
    html += buildNestedList(listBuf, rootStart);
    listBuf = [];
  };
  for (const node of Array.from(container.children)) {
    if (node.tagName === "w:tbl") {
      flushList();
      html += tableHtml(node, ctx);
      continue;
    }
    if (node.tagName !== "w:p") {
      flushList();
      // w:sectPr is re-appended separately on save and w:tcPr is preserved by the table
      // rebuild; any other block (block-level content control, etc.) is preserved
      // verbatim so an edit doesn't drop it.
      if (node.tagName !== "w:sectPr" && node.tagName !== "w:tcPr") {
        html += `<div class="docx-pass-block" contenteditable="false"${passthroughAttr(node)}>${escapeHtml(node.textContent ?? "")}</div>`;
      }
      continue;
    }
    const info = paragraphInfo(node, ctx.numbering, ctx.numLevels);
    const inner = inlineToHtml(node, ctx);
    if (info.isList) {
      // A top-level item of another list (another numId) starts a new list rather than joining.
      const top = listBuf.find((i) => i.level === 0);
      if (info.level === 0 && top && info.numId && top.numId && info.numId !== top.numId) flushList();
      listBuf.push({ level: info.level, ordered: info.ordered, listStyle: info.listStyle, listCls: info.listCls, numId: info.numId, li: `${blockStyleAttr(info)}>${inner || "<br>"}` });
      continue;
    }
    flushList();
    if (info.pageBreakBefore) html += pageBreakHtml("manual");
    if (info.heading) {
      const lvl = Math.min(6, info.heading);
      html += `<h${lvl}${blockStyleAttr(info)}>${inner || "<br>"}</h${lvl}>`;
    } else {
      // A paragraph carrying a caption sequence field is a figure/table/equation caption (type from the seq id).
      const capKind = /data-seq="Table"/i.test(inner) ? "table" : /data-seq="Equation"/i.test(inner) ? "equation" : "figure";
      const capAttr = inner.includes('data-field="seq"') ? ` data-rdoc-caption="${capKind}"` : "";
      html += `<p${blockStyleAttr(info)}${capAttr}>${inner || "<br>"}</p>`;
    }
    // A next-page section break ends the page after its paragraph; show it (display only, the
    // real break rides the preserved w:sectPr re-injected on save).
    if (info.sectPr && info.sectBreak) html += pageBreakHtml("auto");
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
    numStart: new Map(),
    comments: new Map(),
    openComments: new Set(),
    commentOrder: [],
  };
  return renderBlocks(root, ctx);
}

// An inline footnote / endnote reference; the visible number is filled by the engine (document
// order), so this carries only the stable id + kind. contentEditable=false makes it atomic.
function noteRefHtml(id: string, kind: "footnote" | "endnote"): string {
  return `<sup class="docx-fnref" data-fn-id="${escapeAttr(id)}" data-fn-kind="${kind}" contenteditable="false"></sup>`;
}

// Bookmark anchors: a zero-width start marker (carries the name) and an end marker (carries the
// matching id), so both point and range bookmarks round-trip and become jump / reference targets.
function bookmarkStartHtml(name: string, id: string | null): string {
  return `<a class="docx-bookmark" data-rdoc-bm="${escapeAttr(name)}"${id ? ` data-rdoc-bm-id="${escapeAttr(id)}"` : ""} contenteditable="false"></a>`;
}
function bookmarkEndHtml(id: string | null, name: string): string {
  return `<a class="docx-bookmark-end"${id ? ` data-rdoc-bm-id="${escapeAttr(id)}"` : ""} data-rdoc-bm-end="${escapeAttr(name)}" contenteditable="false"></a>`;
}
// A REF / PAGEREF field instruction (" REF name \h ", " PAGEREF name ") -> the cross-ref target +
// format, or null for any other field.
function parseRefInstr(instr: string): { name: string; fmt: "text" | "page" | "direction" } | null {
  const m = /^\s*(REF|PAGEREF)\s+("[^"]+"|\S+)/i.exec(instr);
  if (!m) return null;
  const fmt = /\\p\b/i.test(instr) ? "direction" : m[1]!.toUpperCase() === "PAGEREF" ? "page" : "text";
  return { name: m[2]!.replace(/^"|"$/g, ""), fmt };
}
// A cross-reference: a clickable span whose text the engine recomputes; the cached text is the
// fallback shown before the first reflow.
function xrefHtml(name: string, fmt: "text" | "page" | "direction", cached: string): string {
  return `<a class="docx-xref" data-rdoc-xref="${escapeAttr(name)}" data-rdoc-xref-fmt="${fmt}" contenteditable="false">${escapeHtml(cached)}</a>`;
}
// A " SEQ Figure \* ARABIC " field instruction -> the sequence id, or null for any other field.
function parseSeqInstr(instr: string): string | null {
  const m = /^\s*SEQ\s+(\S+)/i.exec(instr);
  return m ? m[1]! : null;
}
// The w:fldChar type ("begin" | "separate" | "end") carried by a run, or null if it has none.
function fldCharType(run: Element): string | null {
  return run.getElementsByTagName("w:fldChar")[0]?.getAttribute("w:fldCharType") ?? null;
}
// The instruction text held by a run's w:instrText children (a complex field's instr is split across runs).
function instrTextOf(run: Element): string {
  let s = "";
  for (const it of Array.from(run.getElementsByTagName("w:instrText"))) s += it.textContent ?? "";
  return s;
}
// State for a complex field (fldChar begin ... instr ... separate ... result ... end). `raw` is the
// per-run passthrough used verbatim when the field is not one we model; `result` is the cached text.
interface FieldState { depth: number; phase: "instr" | "result"; instr: string; result: string; raw: string }
// Turn a finished complex field into a cross-ref / caption seq element, or its verbatim passthrough.
function finishField(field: FieldState): string {
  const ref = parseRefInstr(field.instr);
  if (ref) return xrefHtml(ref.name, ref.fmt, field.result);
  const seq = parseSeqInstr(field.instr);
  if (seq) return seqFieldHtml(seq, field.result);
  const info = infoFieldKind(field.instr);
  if (info) return infoFieldHtml(info, field.result);
  return field.raw; // TOC / unknown: keep the original runs
}
// A caption auto-number field; the engine renumbers each sequence on reflow.
function seqFieldHtml(id: string, cached: string): string {
  return `<span class="docx-field docx-field-seq" data-field="seq" data-seq="${escapeAttr(id)}" contenteditable="false">${escapeHtml(cached)}</span>`;
}
// A live document field the engine keeps current: PAGE / NUMPAGES are recomputed per page (so a
// footer page number is right on every page, not the one value Word cached), and DATE / TIME /
// AUTHOR / FILENAME round-trip as real fields showing their cached snapshot. REF / PAGEREF / SEQ
// are matched earlier, so the leading-PAGE test never captures PAGEREF.
function infoFieldKind(instr: string): "PAGE" | "NUMPAGES" | "DATE" | "TIME" | "AUTHOR" | "FILENAME" | null {
  const t = instr.trim().toUpperCase();
  for (const k of ["PAGE", "NUMPAGES", "DATE", "TIME", "AUTHOR", "FILENAME"] as const)
    if (new RegExp(`^${k}(\\s|\\\\|$)`).test(t)) return k;
  return null;
}
function infoFieldHtml(kind: string, cached: string): string {
  return `<span class="docx-field" data-field="${kind}" contenteditable="false">${escapeHtml(cached)}</span>`;
}

// Footnote / endnote bodies from word/footnotes.xml or endnotes.xml, keyed by id (skipping the
// separator / continuation-separator notes, which have id <= 0 or a w:type).
function readNotes(files: Record<string, Uint8Array>, part: "footnotes" | "endnotes", kind: "footnote" | "endnote"): { id: string; kind: "footnote" | "endnote"; html: string }[] {
  const raw = files[`word/${part}.xml`];
  if (!raw) return [];
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const ctx: RenderCtx = { files, rels: readRels(files[`word/_rels/${part}.xml.rels`]), numbering: new Map(), numStart: new Map(), comments: new Map(), openComments: new Set(), commentOrder: [] };
  const tag = part === "footnotes" ? "w:footnote" : "w:endnote";
  const out: { id: string; kind: "footnote" | "endnote"; html: string }[] = [];
  for (const note of Array.from(doc.getElementsByTagName(tag))) {
    const id = note.getAttribute("w:id");
    if (!id || Number(id) <= 0 || note.getAttribute("w:type")) continue;
    out.push({ id, kind, html: renderBlocks(note, ctx) || "<p><br></p>" });
  }
  return out;
}

function refTarget(sectPr: Element | undefined, tag: string, rels: Map<string, string>, type: "default" | "first" | "even" = "default"): string | undefined {
  if (!sectPr) return undefined;
  const refs = Array.from(sectPr.getElementsByTagName(tag));
  const pick = type === "default"
    ? (refs.find((r) => (r.getAttribute("w:type") ?? "default") === "default") ?? refs[0])
    : refs.find((r) => r.getAttribute("w:type") === type);
  const id = pick?.getAttributeNS(R, "id") ?? pick?.getAttribute("r:id");
  return id ? rels.get(id) : undefined;
}

export interface DocxParts {
  body: string;
  header: string;
  footer: string;
  headerPath?: string; // archive key of the header part, for write-back
  footerPath?: string;
  headerFirst?: { html: string; path?: string }; // first-page / even-page header & footer variants
  footerFirst?: { html: string; path?: string };
  headerEven?: { html: string; path?: string };
  footerEven?: { html: string; path?: string };
  sectionBands?: Record<string, { html: string; path: string }>; // per-section header/footer by r:id
  notes?: { id: string; kind: "footnote" | "endnote"; html: string }[]; // footnote/endnote bodies by id
  comments: CommentThread[]; // top-level threads in document order, replies nested
  page?: PageGeometry; // page size/margins from w:sectPr, for the paginated view
  paragraphStyles?: { id: string; name: string }[]; // named paragraph styles, for the picker
  characterStyles?: { id: string; name: string }[]; // named character styles, for the picker
  styleDefs?: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[]; // each style's resolved CSS, for editing
  styleCss?: string; // CSS giving each named style its appearance
  noteCss?: string; // the footnote/endnote body style's text props, for the note area
}

// Convert OOXML twips (1/1440 inch) to CSS px at 96 dpi.
const twipToPx = (v: string | null | undefined): number | undefined => {
  if (!v) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n / 15 : undefined;
};

/** Read page size and margins from a w:sectPr into view geometry (px). `evenOdd` comes from
    settings.xml (a document-level flag), so it is passed in rather than read from the sectPr. */
function parsePageGeometry(sectPr: Element | undefined, evenOdd = false): PageGeometry | undefined {
  if (!sectPr) return undefined;
  const pgSz = sectPr.getElementsByTagName("w:pgSz")[0];
  const w = twipToPx(pgSz?.getAttributeNS(W, "w") ?? pgSz?.getAttribute("w:w"));
  const h = twipToPx(pgSz?.getAttributeNS(W, "h") ?? pgSz?.getAttribute("w:h"));
  if (!w || !h) return undefined; // no usable size; let the engine apply its default
  const pgMar = sectPr.getElementsByTagName("w:pgMar")[0];
  const m = (a: string) => Math.max(0, twipToPx(pgMar?.getAttributeNS(W, a) ?? pgMar?.getAttribute("w:" + a)) ?? 96);
  // East-Asian vertical text: w:sectPr/w:textDirection @w:val starting "tbRl" (tbRl, tbRlV).
  const td = sectPr.getElementsByTagName("w:textDirection")[0];
  const dir = td?.getAttributeNS(W, "val") ?? td?.getAttribute("w:val") ?? "";
  const vertical = dir.startsWith("tbRl") || dir === "tbV";
  // Horizontal RTL section: w:sectPr/w:bidi (a present, non-false element).
  const bidi = sectPr.getElementsByTagName("w:bidi")[0];
  const rtl = !!bidi && (bidi.getAttributeNS(W, "val") ?? bidi.getAttribute("w:val") ?? "1") !== "0";
  // Columns: w:cols @w:num (count) + @w:space (gap, twips). Equal-width columns only.
  const cols = sectPr.getElementsByTagName("w:cols")[0];
  const numCols = Number(cols?.getAttributeNS(W, "num") ?? cols?.getAttribute("w:num"));
  const columns = Number.isFinite(numCols) && numCols > 1 ? numCols : undefined;
  const colGap = twipToPx(cols?.getAttributeNS(W, "space") ?? cols?.getAttribute("w:space"));
  // Different first page: w:sectPr/w:titlePg (a present, non-false element).
  const tp = sectPr.getElementsByTagName("w:titlePg")[0];
  const titlePage = !!tp && (tp.getAttributeNS(W, "val") ?? tp.getAttribute("w:val") ?? "1") !== "0";
  // Page border: w:pgBorders, read as a uniform box from the top side (w:sz is eighths of a point).
  const pgB = sectPr.getElementsByTagName("w:pgBorders")[0];
  const pbTop = pgB?.getElementsByTagName("w:top")[0];
  let pageBorder: PageBorder | undefined;
  if (pbTop) {
    const val = pbTop.getAttributeNS(W, "val") ?? pbTop.getAttribute("w:val") ?? "single";
    if (val !== "none" && val !== "nil") {
      const sz = Number(pbTop.getAttributeNS(W, "sz") ?? pbTop.getAttribute("w:sz")) || 4;
      const col = pbTop.getAttributeNS(W, "color") ?? pbTop.getAttribute("w:color") ?? "auto";
      const sp = Number(pbTop.getAttributeNS(W, "space") ?? pbTop.getAttribute("w:space"));
      const style = val === "double" ? "double" : val === "dashed" ? "dashed" : val === "dotted" ? "dotted" : "solid";
      pageBorder = { style, widthPx: Math.max(1, Math.round(sz / 6)), color: /^[0-9a-fA-F]{6}$/.test(col) ? col.toUpperCase() : "000000", spacePt: Number.isFinite(sp) ? sp : undefined };
    }
  }
  // Page numbering: w:pgNumType @w:start (restart) + @w:fmt (format; only the common tokens).
  const pn = sectPr.getElementsByTagName("w:pgNumType")[0];
  const pnStart = Number(pn?.getAttributeNS(W, "start") ?? pn?.getAttribute("w:start"));
  const pageNumStart = Number.isFinite(pnStart) && pnStart >= 1 ? pnStart : undefined;
  const pnFmt = pn?.getAttributeNS(W, "fmt") ?? pn?.getAttribute("w:fmt") ?? "";
  const pageNumFormat = /^(decimal|lowerRoman|upperRoman|lowerLetter|upperLetter)$/.test(pnFmt) ? pnFmt : undefined;
  // Line numbering: w:lnNumType (@w:countBy interval, @w:restart, @w:start). Absent restart = newPage.
  const ln = sectPr.getElementsByTagName("w:lnNumType")[0];
  const lineNumbers = ln ? true : undefined;
  const lnCountBy = Number(ln?.getAttributeNS(W, "countBy") ?? ln?.getAttribute("w:countBy"));
  const lineNumberInterval = Number.isFinite(lnCountBy) && lnCountBy > 1 ? lnCountBy : undefined;
  const lnRestart = ln?.getAttributeNS(W, "restart") ?? ln?.getAttribute("w:restart") ?? "newPage";
  const lineNumberRestart = ln ? (/^(continuous|newPage|newSection)$/.test(lnRestart) ? (lnRestart as "continuous" | "newPage" | "newSection") : "newPage") : undefined;
  const lnStart = Number(ln?.getAttributeNS(W, "start") ?? ln?.getAttribute("w:start"));
  const lineNumberStart = Number.isFinite(lnStart) ? lnStart : undefined;
  // Page vertical alignment: w:vAlign @w:val (center / both / bottom; top is the default).
  const va = sectPr.getElementsByTagName("w:vAlign")[0];
  const vaVal = va?.getAttributeNS(W, "val") ?? va?.getAttribute("w:val") ?? "";
  const pageVAlign = /^(center|both|bottom)$/.test(vaVal) ? (vaVal as "center" | "both" | "bottom") : undefined;
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") }, vertical, rtl, columns, columnGapPx: columns ? Math.round(colGap ?? 36) : undefined, titlePage: titlePage || undefined, evenOdd: evenOdd || undefined, pageBorder, pageNumStart, pageNumFormat, lineNumbers, lineNumberInterval, lineNumberRestart, lineNumberStart, pageVAlign };
}

/** The archive key for a relationship target relative to word/ (e.g. "header1.xml"). */
function partKey(target: string | undefined, files: Record<string, Uint8Array>): string | undefined {
  if (!target) return undefined;
  const clean = target.replace(/^\.\//, "");
  if (clean.startsWith("/")) return files[clean.slice(1)] ? clean.slice(1) : undefined;
  if (files["word/" + clean]) return "word/" + clean;
  return files[clean] ? clean : undefined;
}

/** A CSS-safe attribute selector value: backslash-escape quotes and backslashes. */
const cssAttrValue = (v: string): string => v.replace(/[\\"]/g, "\\$&");

/** Read named paragraph and character styles from styles.xml into picker lists plus the CSS
    that gives each its appearance (keyed on data-rdoc-style / data-rdoc-cstyle), resolving the
    w:basedOn chain. Character styles contribute run properties only. */
function readStyles(stylesXml: Uint8Array | undefined): {
  paragraphStyles: { id: string; name: string }[];
  characterStyles: { id: string; name: string }[];
  styleDefs: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[];
  css: string;
  noteCss: string;
} {
  if (!stylesXml) return { paragraphStyles: [], characterStyles: [], styleDefs: [], css: "", noteCss: "" };
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const wv = (el: Element | undefined, name: string): string | null => el?.getAttributeNS(W, name) ?? el?.getAttribute("w:" + name) ?? null;
  const paraById = new Map<string, Element>();
  const charById = new Map<string, Element>();
  for (const st of Array.from(doc.getElementsByTagName("w:style"))) {
    const type = wv(st, "type") ?? "paragraph";
    const id = wv(st, "styleId");
    if (!id) continue;
    if (type === "paragraph") paraById.set(id, st);
    else if (type === "character") charById.set(id, st);
  }
  const flagOn = (rPr: Element | undefined, tag: string): boolean | undefined => {
    const e = rPr ? Array.from(rPr.children).find((c) => c.tagName === tag) : undefined;
    if (!e) return undefined;
    const v = wv(e, "val");
    return v === null || (v !== "0" && v !== "false");
  };
  // A style's CSS. `own` = only this style's direct properties (for the edit dialog, so saving
  // does not flatten inherited props into the style); otherwise the effective CSS, walking the
  // w:basedOn chain root-to-leaf (leaf overrides), for the rendered appearance.
  const cssFor = (byId: Map<string, Element>, id: string, withPara: boolean, seen: Set<string>, own = false): Record<string, string> => {
    const st = byId.get(id);
    if (!st || seen.has(id)) return {};
    seen.add(id);
    const based = wv(st, "basedOn");
    const out: Record<string, string> = !own && based ? cssFor(byId, based, withPara, seen, false) : {};
    const rPr = Array.from(st.children).find((c) => c.tagName === "w:rPr");
    if (withPara) {
      const pPr = Array.from(st.children).find((c) => c.tagName === "w:pPr");
      const jc = wv(pPr?.getElementsByTagName("w:jc")[0] ?? undefined, "val");
      if (jc && JC_TO_ALIGN[jc]) out["text-align"] = JC_TO_ALIGN[jc];
      const ind = pPr?.getElementsByTagName("w:ind")[0];
      const left = twipToPx(wv(ind ?? undefined, "left") ?? wv(ind ?? undefined, "start"));
      if (left && left > 0) out["margin-left"] = `${Math.round(left)}px`;
      const sp = pPr?.getElementsByTagName("w:spacing")[0];
      const before = twipToPx(wv(sp ?? undefined, "before"));
      const after = twipToPx(wv(sp ?? undefined, "after"));
      const line = wv(sp ?? undefined, "line");
      if (before !== undefined) out["margin-top"] = `${Math.round(before)}px`;
      if (after !== undefined) out["margin-bottom"] = `${Math.round(after)}px`;
      if (line && (wv(sp ?? undefined, "lineRule") ?? "auto") === "auto") out["line-height"] = String(Math.round((Number(line) / 240) * 100) / 100);
      const pShd = pPr?.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill");
      if (pShd && pShd !== "auto" && /^[0-9a-f]{6}$/i.test(pShd)) out["background-color"] = `#${pShd}`;
      const pBdr = pPr?.getElementsByTagName("w:pBdr")[0];
      if (pBdr) {
        let any = false;
        for (const side of ["top", "right", "bottom", "left"] as const) {
          const bc = borderCss(pBdr.getElementsByTagName(`w:${side}`)[0]);
          if (bc && bc !== "none") { out[`border-${side}`] = bc; any = true; }
        }
        if (any) out["padding"] = "2px 6px";
      }
      const tabs = parseWTabs(pPr);
      if (tabs.length) out["--rdoc-tabstops"] = JSON.stringify(tabs);
    }
    const shd = rPr?.getElementsByTagName("w:shd")[0]?.getAttribute("w:fill");
    if (shd && shd !== "auto" && /^[0-9a-f]{6}$/i.test(shd)) out["background-color"] = `#${shd}`;
    const b = flagOn(rPr, "w:b");
    if (b !== undefined) out["font-weight"] = b ? "bold" : "normal";
    const it = flagOn(rPr, "w:i");
    if (it !== undefined) out["font-style"] = it ? "italic" : "normal";
    const deco: string[] = [];
    if (flagOn(rPr, "w:u")) deco.push("underline");
    if (flagOn(rPr, "w:strike")) deco.push("line-through");
    if (deco.length) out["text-decoration"] = deco.join(" ");
    const color = wv(rPr?.getElementsByTagName("w:color")[0] ?? undefined, "val");
    if (color && color !== "auto" && /^[0-9a-f]{6}$/i.test(color)) out["color"] = `#${color}`;
    const sz = wv(rPr?.getElementsByTagName("w:sz")[0] ?? undefined, "val");
    if (sz && Number(sz) > 0) out["font-size"] = `${Number(sz) / 2}pt`;
    const font = rFontsLatin(rPr?.getElementsByTagName("w:rFonts")[0], THEME);
    if (font) out["font-family"] = `'${font.replace(/'/g, "")}'`;
    return out;
  };
  let css = "";
  const styleDefs: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[] = [];
  const collect = (byId: Map<string, Element>, kind: "paragraph" | "character", attr: string, skipId: (id: string) => boolean): { id: string; name: string }[] => {
    const list: { id: string; name: string }[] = [];
    for (const [id, st] of byId) {
      if (skipId(id)) continue;
      if (wv(st, "default") === "1") continue; // the default style is the plain "Paragraph" / no-style entry
      if (Array.from(st.children).some((c) => c.tagName === "w:semiHidden")) continue; // Word's hidden built-ins
      const name = wv(st.getElementsByTagName("w:name")[0] ?? undefined, "val") || id;
      list.push({ id, name });
      styleDefs.push({ id, kind, css: cssFor(byId, id, kind === "paragraph", new Set(), true) }); // own props, for editing
      const decls = cssFor(byId, id, kind === "paragraph", new Set()); // resolved, for the rendered appearance
      const body = Object.entries(decls).map(([k, v]) => `${k}:${v}`).join(";");
      if (body) css += `.docxedit-doc [${attr}="${cssAttrValue(id)}"]{${body}}\n`;
    }
    list.sort((a, b) => a.name.localeCompare(b.name));
    return list;
  };
  // Plain paragraphs take the document's paragraph defaults (pPrDefault, then the default
  // "Normal" style): spacing, line height, alignment, colour. Font and size are the page base.
  {
    const base: Record<string, string> = {};
    const pPrDef = doc.getElementsByTagName("w:pPrDefault")[0]?.getElementsByTagName("w:pPr")[0];
    if (pPrDef) {
      const sp = pPrDef.getElementsByTagName("w:spacing")[0];
      const before = twipToPx(wv(sp, "before"));
      const after = twipToPx(wv(sp, "after"));
      const line = wv(sp, "line");
      if (before !== undefined) base["margin-top"] = `${Math.round(before)}px`;
      if (after !== undefined) base["margin-bottom"] = `${Math.round(after)}px`;
      if (line && (wv(sp, "lineRule") ?? "auto") === "auto") base["line-height"] = String(Math.round((Number(line) / 240) * 100) / 100);
    }
    const defId = [...paraById].find(([, st]) => /^(1|true)$/.test(wv(st, "default") ?? ""))?.[0];
    const normal = defId ? cssFor(paraById, defId, true, new Set()) : {};
    for (const k of ["margin-top", "margin-bottom", "line-height", "text-align", "color"]) if (normal[k]) base[k] = normal[k]!;
    const body = Object.entries(base).map(([k, v]) => `${k}:${v}`).join(";");
    if (body) css += `.docxedit-doc p{${body}}\n`;
  }
  const paragraphStyles = collect(paraById, "paragraph", "data-rdoc-style", (id) => /^heading[1-9]$/i.test(id.replace(/\s+/g, "")));
  // Headings render as <h1>..<h6>, not data-rdoc-style: give them their style's look too.
  for (const id of paraById.keys()) {
    const hm = /^heading([1-6])$/i.exec(id.replace(/\s+/g, ""));
    if (!hm) continue;
    const body = Object.entries(cssFor(paraById, id, true, new Set())).map(([k, v]) => `${k}:${v}`).join(";");
    if (body) css += `.docxedit-doc h${hm[1]}{${body}}\n`;
  }
  const characterStyles = collect(charById, "character", "data-rdoc-cstyle", () => false);
  // The footnote/endnote body style's inheritable text props, applied to the note area so notes
  // render at the document's footnote size/font. Match by style id or name (Word: "FootnoteText"
  // / "footnote text"), falling back to the endnote style.
  const noteStyleId = (() => {
    const want = ["footnotetext", "endnotetext"];
    for (const w of want) {
      for (const [id, st] of paraById) {
        const nm = (wv(st.getElementsByTagName("w:name")[0] ?? undefined, "val") || id).toLowerCase().replace(/\s+/g, "");
        if (nm === w || id.toLowerCase() === w) return id;
      }
    }
    return null;
  })();
  let noteCss = "";
  if (noteStyleId) {
    const d = cssFor(paraById, noteStyleId, true, new Set());
    noteCss = (["font-family", "font-size", "line-height", "color"] as const).filter((k) => d[k]).map((k) => `${k}:${d[k]}`).join(";");
  }
  return { paragraphStyles, characterStyles, styleDefs, css, noteCss };
}

/** Parse a .docx into the editable body HTML plus the header/footer HTML and part keys. */
export function docxToParts(bytes: Uint8Array, preunzipped?: Record<string, Uint8Array>): DocxParts {
  const empty: DocxParts = { body: "", header: "", footer: "", comments: [] };
  const files = preunzipped ?? unzipSync(bytes);
  THEME = readThemeFonts(files);
  THEME_COLORS = readThemeColors(files);
  TBL_STYLES = readTableStyles(files["word/styles.xml"]);
  STYLE_NUM = readStyleNumbering(files["word/styles.xml"]);
  const docXml = files["word/document.xml"];
  if (!docXml) return empty;
  const doc = new DOMParser().parseFromString(strFromU8(docXml), "application/xml");
  const body = doc.getElementsByTagName("w:body")[0];
  if (!body) return empty;
  const rels = readRels(files["word/_rels/document.xml.rels"]);
  const comments = readComments(files["word/comments.xml"]);
  const num = readNumbering(files["word/numbering.xml"]);
  const ctx: RenderCtx = {
    files,
    rels,
    numbering: num.ordered,
    numStart: num.start,
    numLevels: num.levels,
    comments,
    openComments: new Set(),
    commentOrder: [],
    bmNames: new Map(),
  };

  const html = renderBlocks(body, ctx);
  const sectPr = Array.from(body.getElementsByTagName("w:sectPr")).pop();
  const headerTarget = refTarget(sectPr, "w:headerReference", rels);
  const footerTarget = refTarget(sectPr, "w:footerReference", rels);
  // First-page / even-page header & footer variant parts (read whichever the document declares).
  const variantBand = (tag: string, type: "first" | "even") => {
    const tgt = refTarget(sectPr, tag, rels, type);
    return tgt ? { html: renderRefPart(files, tgt), path: partKey(tgt, files) } : undefined;
  };
  // Different odd & even pages is a document-level setting in word/settings.xml.
  const evenOdd = (() => {
    const s = files["word/settings.xml"];
    if (!s) return false;
    const e = new DOMParser().parseFromString(strFromU8(s), "application/xml").getElementsByTagName("w:evenAndOddHeaders")[0];
    return !!e && (e.getAttributeNS(W, "val") ?? e.getAttribute("w:val") ?? "1") !== "0";
  })();
  // Distinct per-section header/footer parts, keyed by r:id (matching data-rdoc-sec*key on the
  // boundary paragraphs). Every header/footer reference in any sectPr resolves here.
  const sectionBands: Record<string, { html: string; path: string }> = {};
  for (const sp of Array.from(body.getElementsByTagName("w:sectPr"))) {
    for (const tag of ["w:headerReference", "w:footerReference"] as const) {
      const rid = refRid(sp, tag);
      if (!rid || sectionBands[rid]) continue;
      const target = rels.get(rid);
      const path = partKey(target, files);
      if (target && path) sectionBands[rid] = { html: renderRefPart(files, target), path };
    }
  }

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
  const ps = readStyles(files["word/styles.xml"]);
  return {
    body: html || "<p><br></p>",
    header: renderRefPart(files, headerTarget),
    footer: renderRefPart(files, footerTarget),
    headerPath: partKey(headerTarget, files),
    footerPath: partKey(footerTarget, files),
    headerFirst: variantBand("w:headerReference", "first"),
    footerFirst: variantBand("w:footerReference", "first"),
    headerEven: variantBand("w:headerReference", "even"),
    footerEven: variantBand("w:footerReference", "even"),
    sectionBands,
    notes: [...readNotes(files, "footnotes", "footnote"), ...readNotes(files, "endnotes", "endnote")],
    comments: threads,
    page: parsePageGeometry(sectPr, evenOdd),
    paragraphStyles: ps.paragraphStyles,
    characterStyles: ps.characterStyles,
    styleDefs: ps.styleDefs,
    styleCss: ps.css + num.css + headingNumberCss(STYLE_NUM, num.levels) + [...TBL_STYLES].filter(([, t]) => t.pCss).map(([id, t]) => `.docxedit-doc table[data-rdoc-tstyle="${cssAttrValue(id)}"] p{${t.pCss}}\n`).join(""),
    noteCss: ps.noteCss,
  };
}

/** Convert a .docx body to editable HTML. Returns "" if there is no document body. */
export function docxToHtml(bytes: Uint8Array): string {
  return docxToParts(bytes).body;
}

