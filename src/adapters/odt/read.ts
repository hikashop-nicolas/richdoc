// odt READ: parse an .odt archive into editable HTML, header/footer and comments. Pure
// XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { bytesToBase64 } from "../../core/util";
import type { CommentThread, PageGeometry } from "../../core/types";
import { ODF_ALIGN, escapeHtml, escapeAttr, inlinePass, blockPass, passthroughAttr, FMT0, IMG_MIME } from "./shared";
import type { Fmt, PFmt } from "./shared";


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

const PX_PER = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, px: 1 };
/** ODF length (e.g. "5.2cm", "48pt") to CSS px. */
function lenToPx(v: string | null | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^([\d.]+)(cm|mm|in|pt|px)$/.exec(v.trim());
  return m ? parseFloat(m[1]!) * PX_PER[m[2] as keyof typeof PX_PER] : undefined;
}

/** Read context: the archive (for image data), the resolved style maps, and comment state. */
interface ChangeInfo {
  type: "insertion" | "deletion";
  author: string;
  date: string;
  deleted: string; // for deletions: the removed text
}
interface CellBorders {
  t: string | null;
  r: string | null;
  b: string | null;
  l: string | null;
}

interface RCtx {
  files: Record<string, Uint8Array>;
  styles: Map<string, Fmt>;
  paras: Map<string, PFmt>;
  cellStyles: Map<string, CellBorders>;
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
      strike: !!tp && tp.getAttribute("style:text-line-through-style") != null && tp.getAttribute("style:text-line-through-style") !== "none",
      vertAlign: ((p) => {
        const v = (p ?? "").trim().split(/\s+/)[0] ?? "";
        if (v === "super") return "super";
        if (v === "sub") return "sub";
        const pct = parseFloat(v); // positive % = super, negative = sub
        return Number.isFinite(pct) && pct !== 0 ? (pct > 0 ? "super" : "sub") : undefined;
      })(tp?.getAttribute("style:text-position")),
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
    const indentRaw = lenToPx(pp?.getAttribute("fo:margin-left"));
    const indentPx = indentRaw && indentRaw > 0 ? Math.round(indentRaw) : undefined;
    const lh = pp?.getAttribute("fo:line-height");
    const lineHeight = lh && lh.endsWith("%") ? Math.round((parseFloat(lh) / 100) * 100) / 100 : undefined;
    if (align || indentPx || lineHeight) map.set(name, { align, indentPx, lineHeight });
  }
  return map;
}

const ODT_BORDER_STYLE: Record<string, string> = {
  solid: "solid", dashed: "dashed", dotted: "dotted", double: "double",
  groove: "solid", ridge: "solid", inset: "solid", outset: "solid", fine: "solid",
};
// An fo:border value ("0.5pt solid #000000") -> the editor's "<w>px <style> <#color>", or null.
function parseOdtBorder(v: string | null | undefined): string | null {
  if (!v) return null;
  const s = v.trim();
  if (!s || /^(none|hidden)$/i.test(s)) return null;
  let style = "solid";
  let color = "#000000";
  let w = 1;
  for (const p of s.split(/\s+/)) {
    const lp = p.toLowerCase();
    if (ODT_BORDER_STYLE[lp]) style = ODT_BORDER_STYLE[lp]!;
    else if (/^#[0-9a-f]{3,8}$/i.test(p)) color = p;
    else {
      const px = lenToPx(p);
      if (px) w = Math.max(1, Math.round(px));
    }
  }
  return `${w}px ${style} ${color}`;
}

/** Map table-cell style-name -> its four border specs, resolving fo:border (shorthand),
    per-side fo:border-* and the style:parent-style-name chain. */
function collectCellStyles(doc: Document): Map<string, CellBorders> {
  const raw = new Map<string, { own: CellBorders; parent?: string }>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "table-cell") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const cp = st.getElementsByTagName("style:table-cell-properties")[0];
    const all = cp?.getAttribute("fo:border") ?? undefined;
    const side = (a: string): string | null => parseOdtBorder(cp?.getAttribute(a) ?? all);
    raw.set(name, {
      own: { t: side("fo:border-top"), r: side("fo:border-right"), b: side("fo:border-bottom"), l: side("fo:border-left") },
      parent: st.getAttribute("style:parent-style-name") ?? undefined,
    });
  }
  const empty: CellBorders = { t: null, r: null, b: null, l: null };
  const resolve = (name: string, seen: Set<string>): CellBorders => {
    const e = raw.get(name);
    if (!e || seen.has(name)) return empty;
    seen.add(name);
    const base = e.parent ? resolve(e.parent, seen) : empty;
    return { t: e.own.t ?? base.t, r: e.own.r ?? base.r, b: e.own.b ?? base.b, l: e.own.l ?? base.l };
  };
  const map = new Map<string, CellBorders>();
  for (const name of raw.keys()) map.set(name, resolve(name, new Set()));
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
  if (f.vertAlign === "super") s = `<sup>${s}</sup>`;
  else if (f.vertAlign === "sub") s = `<sub>${s}</sub>`;
  if (f.strike) s = `<s>${s}</s>`;
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

/** Render a table:table to an editable HTML table (cells editable, structure locked and
    preserved as passthrough), mirroring the docx adapter so the engine treats both alike. */
function odtTableHtml(el: Element, ctx: RCtx): string {
  let rows = "";
  for (const tr of Array.from(el.children)) {
    if (tr.tagName !== "table:table-row") continue;
    let cells = "";
    for (const tc of Array.from(tr.children)) {
      if (tc.tagName !== "table:table-cell") continue; // skip table:covered-table-cell (covered by a span)
      const span = Number(tc.getAttribute("table:number-columns-spanned")) || 1;
      const rspan = Number(tc.getAttribute("table:number-rows-spanned")) || 1;
      let inner = "";
      for (const child of Array.from(tc.children)) inner += blockToHtml(child, ctx);
      const cs = span > 1 ? ` colspan="${span}"` : "";
      const rs = rspan > 1 ? ` rowspan="${rspan}"` : "";
      const cstyle = tc.getAttribute("table:style-name"); // preserve cell style (borders/bg/padding)
      const cellStyle = cstyle ? ` data-odt-cellstyle="${escapeAttr(cstyle)}"` : "";
      // Resolve the cell style's borders into the editor's per-side model so they show on load.
      const cb = cstyle ? ctx.cellStyles.get(cstyle) : undefined;
      const bAttr = (k: string, spec: string | null | undefined): string => (spec ? ` data-rdoc-b${k}="${spec}"` : "");
      const borders = cb ? bAttr("t", cb.t) + bAttr("r", cb.r) + bAttr("b", cb.b) + bAttr("l", cb.l) : "";
      cells += `<td${cs}${rs}${cellStyle}${borders}><div class="docx-cell" contenteditable="true">${inner || "<br>"}</div></td>`;
    }
    rows += `<tr>${cells}</tr>`;
  }
  // Preserve the table style-name and the column declarations so a structural edit keeps styling.
  const tstyle = el.getAttribute("table:style-name");
  const tableStyle = tstyle ? ` data-odt-tablestyle="${escapeAttr(tstyle)}"` : "";
  const cols = Array.from(el.children)
    .filter((c) => c.tagName === "table:table-column")
    .map((c) => ({ s: c.getAttribute("table:style-name") ?? "", r: c.getAttribute("table:number-columns-repeated") ?? "1" }));
  const colsAttr = cols.length ? ` data-odt-cols="${escapeAttr(JSON.stringify(cols))}"` : "";
  return `<table class="docx-table" contenteditable="false"${passthroughAttr(el)}${tableStyle}${colsAttr}>${rows}</table>`;
}

function blockToHtml(el: Element, ctx: RCtx): string {
  const alignAttr = (): string => {
    const pf = ctx.paras.get(el.getAttribute("text:style-name") ?? "");
    if (!pf) return "";
    const css: string[] = [];
    if (pf.align && pf.align !== "left") css.push(`text-align:${pf.align}`);
    if (pf.indentPx) css.push(`margin-left:${pf.indentPx}px`);
    if (pf.lineHeight) css.push(`line-height:${pf.lineHeight}`);
    return css.length ? ` style="${css.join(";")}"` : "";
  };
  switch (el.tagName) {
    case "text:h": {
      const lvl = Math.min(3, Math.max(1, parseInt(el.getAttribute("text:outline-level") ?? "1", 10) || 1));
      const inner = inlineToHtml(el, ctx);
      return `<h${lvl}${alignAttr()}>${inner || "<br>"}</h${lvl}>`;
    }
    case "text:list":
      return listToHtml(el, ctx);
    case "table:table":
      return odtTableHtml(el, ctx);
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
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), cellStyles: collectCellStyles(doc), threads: [], rangedNames: new Set(), openComment: new Set(), changes: new Map(), openIns: new Set() };
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
/** Page size + margins from the first page-layout in styles.xml, in px. Landscape is
    already reflected because fo:page-width/height are written swapped. */
function parsePageGeometry(files: Record<string, Uint8Array>): PageGeometry | undefined {
  const raw = files["styles.xml"];
  if (!raw) return undefined;
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const props = doc.getElementsByTagName("style:page-layout-properties")[0];
  if (!props) return undefined;
  const w = lenToPx(props.getAttribute("fo:page-width"));
  const h = lenToPx(props.getAttribute("fo:page-height"));
  if (!w || !h) return undefined; // no usable size; the engine applies its default
  const m = (a: string) => Math.round(Math.max(0, lenToPx(props.getAttribute(`fo:margin-${a}`)) ?? 96));
  const wm = props.getAttribute("style:writing-mode") ?? "";
  const vertical = wm.startsWith("tb"); // tb-rl / tb
  const rtl = wm.startsWith("rl"); // rl-tb (horizontal right-to-left)
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") }, vertical, rtl };
}

export function odtToParts(bytes: Uint8Array): { body: string; comments: CommentThread[]; header: string; footer: string; page?: PageGeometry } {
  const files = unzipSync(bytes);
  const { header, footer } = readHeaderFooter(files);
  const page = parsePageGeometry(files);
  const content = files["content.xml"];
  if (!content) return { body: "", comments: [], header, footer, page };
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) return { body: "", comments: [], header, footer, page };
  const rangedNames = new Set(
    Array.from(doc.getElementsByTagName("office:annotation-end"))
      .map((e) => e.getAttribute("office:name"))
      .filter((n): n is string => !!n),
  );
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), cellStyles: collectCellStyles(doc), threads: [], rangedNames, openComment: new Set(), changes: readChanges(body), openIns: new Set() };
  let html = "";
  for (const block of Array.from(body.children)) {
    if (block.tagName === "text:tracked-changes") continue; // metadata, parsed into ctx.changes
    html += blockToHtml(block, ctx);
  }
  return { body: html || "<p><br></p>", comments: ctx.threads, header, footer, page };
}

/** Convert an .odt's body to HTML. Returns "" if there is no editable text body. */
export function odtToHtml(bytes: Uint8Array): string {
  return odtToParts(bytes).body;
}

