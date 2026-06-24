// docx READ: parse a .docx archive into editable HTML, header/footer, comments, fonts and
// page geometry. Pure XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { t } from "../../core/i18n";
import { bytesToBase64 } from "../../core/util";
import type { CommentEntry, CommentThread, PageGeometry } from "../../core/types";
import { W, R, XMLNS, NS_DECLS, IMG_MIME, escapeHtml, escapeAttr, HL_CSS, JC_TO_ALIGN } from "./shared";
import type { Fmt } from "./shared";


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
  align?: string; // CSS text-align
  indentPx?: number; // w:ind left indent, in px
  lineHeight?: number; // w:spacing line (auto rule), as a multiple
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
  const ind = pPr?.getElementsByTagName("w:ind")[0];
  const indentPx = twipToPx(ind?.getAttribute("w:left") ?? ind?.getAttribute("w:start"));
  const sp = pPr?.getElementsByTagName("w:spacing")[0];
  const line = sp?.getAttribute("w:line");
  const rule = sp?.getAttribute("w:lineRule") ?? "auto";
  const lineHeight = line && rule === "auto" ? Math.round((Number(line) / 240) * 100) / 100 : undefined;
  // Paragraph-mark revision lives in pPr > rPr > w:ins / w:del.
  const markRPr = pPr ? Array.from(pPr.children).find((c) => c.tagName === "w:rPr") : undefined;
  const mark = markRPr && (Array.from(markRPr.children).find((c) => c.tagName === "w:ins" || c.tagName === "w:del") as Element | undefined);
  return {
    heading,
    isList: !!numPr,
    ordered: numbering.get(numId) ?? false,
    align: JC_TO_ALIGN[jc],
    indentPx: indentPx && indentPx > 0 ? Math.round(indentPx) : undefined,
    lineHeight,
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

