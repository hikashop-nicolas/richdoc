// docx READ: parse a .docx archive into editable HTML, header/footer, comments, fonts and
// page geometry. Pure XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { t } from "../../core/i18n";
import { bytesToBase64 } from "../../core/util";
import type { CommentEntry, CommentThread, PageGeometry } from "../../core/types";
import { W, R, XMLNS, NS_DECLS, IMG_MIME, escapeHtml, escapeAttr, HL_CSS, JC_TO_ALIGN } from "./shared";
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
  const layout = readLayout(run);
  let lay = "";
  if (layout) {
    lay = ` data-rdoc-wrap="${layout.wrap}" data-rdoc-align="${layout.align}"`;
    if (layout.wrap === "behind" || layout.wrap === "front") {
      lay += ` data-rdoc-x="${layout.x}" data-rdoc-y="${layout.y}" style="left:${layout.x}px;top:${layout.y}px"`;
    }
  }
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
  level: number; // list nesting level (w:ilvl), 0 = top
  align?: string; // CSS text-align
  indentPx?: number; // w:ind left indent, in px
  lineHeight?: number; // w:spacing line (auto rule), as a multiple
  spaceBeforePx?: number; // w:spacing @w:before, in px
  spaceAfterPx?: number; // w:spacing @w:after, in px
  border?: string; // CSS border declaration block
  styleId?: string; // a named paragraph style (w:pStyle val) that is not a heading
  pageBreakBefore: boolean;
  revPara?: "ins" | "del"; // paragraph-mark revision (split/merge)
  revAuthor?: string;
  revDate?: string;
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
  return {
    heading,
    isList: !!numPr,
    ordered: numbering.get(numId) ?? false,
    level,
    align: JC_TO_ALIGN[jc],
    indentPx: indentPx && indentPx > 0 ? Math.round(indentPx) : undefined,
    lineHeight,
    spaceBeforePx: spaceBeforePx === undefined ? undefined : Math.round(spaceBeforePx),
    spaceAfterPx: spaceAfterPx === undefined ? undefined : Math.round(spaceAfterPx),
    styleId: !hm && rawStyle ? rawStyle : undefined,
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
  if (info.indentPx) parts.push(`margin-left:${info.indentPx}px`);
  if (info.lineHeight) parts.push(`line-height:${info.lineHeight}`);
  if (info.spaceBeforePx !== undefined) parts.push(`margin-top:${info.spaceBeforePx}px`);
  if (info.spaceAfterPx !== undefined) parts.push(`margin-bottom:${info.spaceAfterPx}px`);
  if (info.border) parts.push(info.border);
  const style = parts.length ? ` style="${parts.join(";")}"` : "";
  const styleAttr = info.styleId ? ` data-rdoc-style="${escapeAttr(info.styleId)}"` : "";
  if (!info.revPara) return style + styleAttr;
  return `${style}${styleAttr} class="docx-para-${info.revPara}" data-rev-para="${info.revPara}" data-rev-author="${escapeAttr(info.revAuthor ?? "")}" data-rev-date="${escapeAttr(info.revDate ?? "")}"`;
}

/** Build nested <ul>/<ol> from a flat list of items carrying their nesting level (w:ilvl).
    A deeper item opens new lists inside the current <li>; a shallower one closes them. */
function buildNestedList(items: { level: number; ordered: boolean; li: string }[]): string {
  let html = "";
  const open: string[] = []; // stack of open list tags
  for (const it of items) {
    const depth = it.level + 1; // number of lists that should be open
    while (open.length > depth) html += `</li></${open.pop()}>`;
    if (open.length === depth) html += "</li>"; // sibling at the same level
    while (open.length < depth) {
      const tag = it.ordered ? "ol" : "ul";
      html += `<${tag}>`;
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
  let listBuf: { level: number; ordered: boolean; li: string }[] = [];
  const flushList = () => {
    if (!listBuf.length) return;
    html += buildNestedList(listBuf);
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
      listBuf.push({ level: info.level, ordered: info.ordered, li: `${blockStyleAttr(info)}>${inner || "<br>"}` });
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
  paragraphStyles?: { id: string; name: string }[]; // named paragraph styles, for the picker
  characterStyles?: { id: string; name: string }[]; // named character styles, for the picker
  styleDefs?: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[]; // each style's resolved CSS, for editing
  styleCss?: string; // CSS giving each named style its appearance
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
  // East-Asian vertical text: w:sectPr/w:textDirection @w:val starting "tbRl" (tbRl, tbRlV).
  const td = sectPr.getElementsByTagName("w:textDirection")[0];
  const dir = td?.getAttributeNS(W, "val") ?? td?.getAttribute("w:val") ?? "";
  const vertical = dir.startsWith("tbRl") || dir === "tbV";
  // Horizontal RTL section: w:sectPr/w:bidi (a present, non-false element).
  const bidi = sectPr.getElementsByTagName("w:bidi")[0];
  const rtl = !!bidi && (bidi.getAttributeNS(W, "val") ?? bidi.getAttribute("w:val") ?? "1") !== "0";
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") }, vertical, rtl };
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
} {
  if (!stylesXml) return { paragraphStyles: [], characterStyles: [], styleDefs: [], css: "" };
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
  // Resolve a style's effective CSS by walking basedOn from root to leaf (leaf overrides).
  const cssFor = (byId: Map<string, Element>, id: string, withPara: boolean, seen: Set<string>): Record<string, string> => {
    const st = byId.get(id);
    if (!st || seen.has(id)) return {};
    seen.add(id);
    const based = wv(st, "basedOn");
    const out: Record<string, string> = based ? cssFor(byId, based, withPara, seen) : {};
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
      const decls = cssFor(byId, id, kind === "paragraph", new Set());
      styleDefs.push({ id, kind, css: decls });
      const body = Object.entries(decls).map(([k, v]) => `${k}:${v}`).join(";");
      if (body) css += `.docxedit-doc [${attr}="${cssAttrValue(id)}"]{${body}}\n`;
    }
    list.sort((a, b) => a.name.localeCompare(b.name));
    return list;
  };
  const paragraphStyles = collect(paraById, "paragraph", "data-rdoc-style", (id) => /^heading[1-9]$/i.test(id.replace(/\s+/g, "")));
  const characterStyles = collect(charById, "character", "data-rdoc-cstyle", () => false);
  return { paragraphStyles, characterStyles, styleDefs, css };
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
  const ps = readStyles(files["word/styles.xml"]);
  return {
    body: html || "<p><br></p>",
    header: renderRefPart(files, headerTarget),
    footer: renderRefPart(files, footerTarget),
    headerPath: partKey(headerTarget, files),
    footerPath: partKey(footerTarget, files),
    comments: threads,
    page: parsePageGeometry(sectPr),
    paragraphStyles: ps.paragraphStyles,
    characterStyles: ps.characterStyles,
    styleDefs: ps.styleDefs,
    styleCss: ps.css,
  };
}

/** Convert a .docx body to editable HTML. Returns "" if there is no document body. */
export function docxToHtml(bytes: Uint8Array): string {
  return docxToParts(bytes).body;
}

