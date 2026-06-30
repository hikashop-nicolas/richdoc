// docx READ: parse a .docx archive into editable HTML, header/footer, comments, fonts and
// page geometry. Pure XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { t } from "../../core/i18n";
import { bytesToBase64, imageLayoutAttrs } from "../../core/util";
import type { CommentEntry, CommentThread, PageGeometry } from "../../core/types";
import { W, R, XMLNS, NS_DECLS, IMG_MIME, escapeHtml, escapeAttr, HL_CSS, JC_TO_ALIGN } from "./shared";
import { ommlToMathml } from "./omml";
import { ensureDingsFont } from "./materialdings";
import { SYMBOL, WEBDINGS, WINGDINGS2, WINGDINGS3 } from "./dingbats";
import type { Fmt } from "./shared";
import { readLayout } from "./image-layout";


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

/** The document default font (styles.xml docDefaults), used as the page base font. */
export function defaultFont(files: Record<string, Uint8Array>): string | undefined {
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
  const tblB = tblPr?.getElementsByTagName("w:tblBorders")[0];
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
  let rows = "";
  for (let r = 0; r < grid.length; r++) {
    let cells = "";
    for (const ci of grid[r]) {
      if (ci.cont) continue; // a vMerge continuation cell: covered by the restart above
      let inner = "";
      for (const p of Array.from(ci.tc.children)) {
        if (p.tagName === "w:p") inner += `<div>${inlineToHtml(p, ctx) || "<br>"}</div>`;
        else if (p.tagName === "w:tbl") inner += tableHtml(p, ctx);
      }
      const cs = ci.gridSpan > 1 ? ` colspan="${ci.gridSpan}"` : "";
      const rsN = ci.restart ? rowspanOf(r, ci.gridCol) : 1;
      const rs = rsN > 1 ? ` rowspan="${rsN}"` : "";
      // Resolve the cell's real borders (Word's tcBorders-over-tblBorders rules) into the
      // editor's per-side model, so the document's borders show on load.
      const tcB = ci.tcPr?.getElementsByTagName("w:tcBorders")[0];
      const bAttr = (k: string, spec: string | null): string => (spec ? ` data-rdoc-b${k}="${spec}"` : "");
      const borders =
        bAttr("t", resolveCellBorder(tcB, tblB, "w:top", "w:insideH", r === 0)) +
        bAttr("b", resolveCellBorder(tcB, tblB, "w:bottom", "w:insideH", r + rsN === totalRows)) +
        bAttr("l", resolveCellBorder(tcB, tblB, "w:left", "w:insideV", ci.gridCol === 0)) +
        bAttr("r", resolveCellBorder(tcB, tblB, "w:right", "w:insideV", ci.gridCol + ci.gridSpan === totalCols));
      // Structure is locked (contenteditable=false on the table); each cell's content is its
      // own editable region. Cell shading round-trips via the preserved tcPr.
      cells += `<td${cs}${rs}${propAttr("tcpr", ci.tcPr)}${borders}><div class="docx-cell" contenteditable="true">${inner || "<br>"}</div></td>`;
    }
    rows += `<tr>${cells}</tr>`;
  }
  // Table indent (w:tblInd) -> inline margin-left so it shows and round-trips.
  const tblInd = tblPr?.getElementsByTagName("w:tblInd")[0];
  const indPx = twipToPx(tblInd?.getAttributeNS(W, "w") ?? tblInd?.getAttribute("w:w"));
  const indStyle = indPx && indPx > 0 ? ` style="margin-left:${Math.round(indPx)}px"` : "";
  return `<table class="docx-table" contenteditable="false"${passthroughAttr(tbl)}${indStyle}${propAttr("tblpr", tblPr)}${propAttr("tblgrid", tblGrid)}>${rows}</table>`;
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
function readNumbering(file: Uint8Array | undefined): { ordered: Map<string, boolean>; start: Map<string, number> } {
  const ordered = new Map<string, boolean>();
  const start = new Map<string, number>();
  if (!file) return { ordered, start };
  const doc = new DOMParser().parseFromString(strFromU8(file), "application/xml");
  const absFmt = new Map<string, boolean>(); // abstractNumId -> ordered
  const absStart = new Map<string, number>(); // abstractNumId -> level-0 start
  for (const an of Array.from(doc.getElementsByTagName("w:abstractNum"))) {
    const id = an.getAttributeNS(W, "abstractNumId") ?? an.getAttribute("w:abstractNumId");
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
    const ov = Array.from(num.getElementsByTagName("w:lvlOverride")).find((o) => (o.getAttribute("w:ilvl") ?? "0") === "0");
    const so = Number(ov?.getElementsByTagName("w:startOverride")[0]?.getAttribute("w:val"));
    start.set(numId, Number.isFinite(so) ? so : absStart.get(abs) ?? 1);
  }
  return { ordered, start };
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
    vertAlign: ((va) => (va === "superscript" ? "super" : va === "subscript" ? "sub" : undefined))(attrVal(rPr, "w:vertAlign")),
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
    else if (el.tagName === "w:tab") text += '<span class="docx-tab" data-docx-tab="1" contenteditable="false">\t</span>';
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
const MODELED_RUN_CHILDREN = new Set(["w:rPr", "w:t", "w:delText", "w:tab", "w:br", "w:lastRenderedPageBreak"]);
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
      if (ref) html += xrefHtml(ref.name, ref.fmt, el.textContent ?? "");
      else if (seq) html += seqFieldHtml(seq, el.textContent ?? "");
      else html += inlinePassthrough(el); // PAGE / other simple fields stay as-is
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
  revPara?: "ins" | "del"; // paragraph-mark revision (split/merge)
  revAuthor?: string;
  revDate?: string;
  sectPr?: string; // a mid-document section break (w:pPr/w:sectPr), preserved verbatim
  sectBreak?: boolean; // that section starts a new page (type != continuous) -> show a page break
  secGeom?: string; // JSON page geometry of the section ending here, for per-section rendering
  secHeaderRid?: string; // r:id of this section's header part (keys into sectionBands)
  secFooterRid?: string; // r:id of this section's footer part
  tabStops?: string; // JSON [{pos,val,leader}] of the paragraph's custom tab stops (w:tabs)
}
/** The r:id of a section's default header/footer reference (keys per-section header/footer). */
function refRid(sectPr: Element | undefined, tag: string): string | undefined {
  if (!sectPr) return undefined;
  const refs = Array.from(sectPr.getElementsByTagName(tag));
  const pick = refs.find((r) => (r.getAttribute("w:type") ?? "default") === "default") ?? refs[0];
  return pick?.getAttributeNS(R, "id") ?? pick?.getAttribute("r:id") ?? undefined;
}
function paragraphInfo(p: Element, numbering: Map<string, boolean>): PInfo {
  const pPr = p.getElementsByTagName("w:pPr")[0];
  let heading = 0;
  const rawStyle = pPr?.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") ?? "";
  const styleVal = rawStyle.replace(/\s+/g, "");
  const hm = /^heading([1-9])$/i.exec(styleVal);
  if (hm) heading = Number(hm[1]);
  const numPr = pPr?.getElementsByTagName("w:numPr")[0];
  const numId = numPr?.getElementsByTagName("w:numId")[0]?.getAttribute("w:val") ?? "";
  const level = Math.max(0, Math.min(8, Number(numPr?.getElementsByTagName("w:ilvl")[0]?.getAttribute("w:val")) || 0));
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
  const tabsEl = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:tabs") : undefined;
  const stops = tabsEl
    ? Array.from(tabsEl.getElementsByTagName("w:tab"))
        .filter((tb) => (tb.getAttribute("w:val") ?? "left") !== "clear")
        .map((tb) => ({ pos: Math.round(twipToPx(tb.getAttribute("w:pos")) ?? 0), val: tb.getAttribute("w:val") ?? "left", leader: tb.getAttribute("w:leader") ?? undefined }))
        .filter((s) => s.pos > 0)
    : [];
  return {
    heading,
    isList: !!numPr,
    ordered: numbering.get(numId) ?? false,
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
  const style = parts.length ? ` style="${parts.join(";")}"` : "";
  const styleAttr = info.styleId ? ` data-rdoc-style="${escapeAttr(info.styleId)}"` : "";
  const sectAttr = info.sectPr ? ` data-docx-sectpr="${escapeAttr(info.sectPr)}"` : "";
  const secGeomAttr = info.secGeom ? ` data-rdoc-secbreak="${escapeAttr(info.secGeom)}"` : "";
  const secHKey = info.secHeaderRid ? ` data-rdoc-secheaderkey="${escapeAttr(info.secHeaderRid)}"` : "";
  const secFKey = info.secFooterRid ? ` data-rdoc-secfooterkey="${escapeAttr(info.secFooterRid)}"` : "";
  const tabAttr = info.tabStops ? ` data-rdoc-tabstops="${escapeAttr(info.tabStops)}"` : "";
  const sec = sectAttr + secGeomAttr + secHKey + secFKey;
  if (!info.revPara) return style + styleAttr + sec + tabAttr;
  return `${style}${styleAttr}${sec}${tabAttr} class="docx-para-${info.revPara}" data-rev-para="${info.revPara}" data-rev-author="${escapeAttr(info.revAuthor ?? "")}" data-rev-date="${escapeAttr(info.revDate ?? "")}"`;
}

/** Build nested <ul>/<ol> from a flat list of items carrying their nesting level (w:ilvl).
    A deeper item opens new lists inside the current <li>; a shallower one closes them. */
function buildNestedList(items: { level: number; ordered: boolean; li: string }[], rootStart = 1): string {
  let html = "";
  const open: string[] = []; // stack of open list tags
  for (const it of items) {
    const depth = it.level + 1; // number of lists that should be open
    while (open.length > depth) html += `</li></${open.pop()}>`;
    if (open.length === depth) html += "</li>"; // sibling at the same level
    while (open.length < depth) {
      const tag = it.ordered ? "ol" : "ul";
      const startAttr = open.length === 0 && tag === "ol" && rootStart > 1 ? ` start="${rootStart}"` : ""; // restart/continue
      html += `<${tag}${startAttr}>`;
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
  let listBuf: { level: number; ordered: boolean; li: string; numId?: string }[] = [];
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
      listBuf.push({ level: info.level, ordered: info.ordered, numId: info.numId, li: `${blockStyleAttr(info)}>${inner || "<br>"}` });
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
  return field.raw; // PAGE / TOC / DATE / unknown: keep the original runs
}
// A caption auto-number field; the engine renumbers each sequence on reflow.
function seqFieldHtml(id: string, cached: string): string {
  return `<span class="docx-field docx-field-seq" data-field="seq" data-seq="${escapeAttr(id)}" contenteditable="false">${escapeHtml(cached)}</span>`;
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
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") }, vertical, rtl, columns, columnGapPx: columns ? Math.round(colGap ?? 36) : undefined, titlePage: titlePage || undefined, evenOdd: evenOdd || undefined };
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
    const font = wv(rPr?.getElementsByTagName("w:rFonts")[0] ?? undefined, "ascii");
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
  const paragraphStyles = collect(paraById, "paragraph", "data-rdoc-style", (id) => /^heading[1-9]$/i.test(id.replace(/\s+/g, "")));
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
  const num = readNumbering(files["word/numbering.xml"]);
  const ctx: RenderCtx = {
    files,
    rels,
    numbering: num.ordered,
    numStart: num.start,
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
    styleCss: ps.css,
    noteCss: ps.noteCss,
  };
}

/** Convert a .docx body to editable HTML. Returns "" if there is no document body. */
export function docxToHtml(bytes: Uint8Array): string {
  return docxToParts(bytes).body;
}

