// odt READ: parse an .odt archive into editable HTML, header/footer and comments. Pure
// XML -> HTML; the write half lives in ./write.
import { strFromU8, unzipSync } from "fflate";
import { bytesToBase64, imageLayoutAttrs } from "../../core/util";
import { t } from "../../core/i18n";
import type { CommentThread, PageGeometry } from "../../core/types";
import { ODF_ALIGN, escapeHtml, escapeAttr, inlinePass, blockPass, passthroughAttr, FMT0, IMG_MIME } from "./shared";
import type { Fmt, PFmt } from "./shared";
import { collectGraphicStyles, readOdtLayout } from "./image-layout";
import type { GraphicStyleInfo } from "./image-layout";


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
  tableMargins: Map<string, number>; // table style-name -> left indent (px)
  listStyles: Map<string, boolean[]>; // list style-name -> per-level ordered flag
  namedStyles: Set<string>; // ids of named paragraph styles (from styles.xml)
  autoParent: Map<string, string>; // automatic paragraph style-name -> its parent style-name
  namedCharStyles: Set<string>; // ids of named character styles (from styles.xml)
  textAutoParent: Map<string, string>; // automatic text style-name -> its parent style-name
  threads: CommentThread[]; // comments collected while rendering, for the panel
  rangedNames: Set<string>; // annotation names that have a matching annotation-end
  openComment: Set<string>; // comment ranges currently open (reopened per paragraph)
  changes: Map<string, ChangeInfo>; // tracked-change id -> metadata
  openIns: Set<string>; // insertion ranges currently open (reopened per paragraph)
  graphicStyles: Map<string, GraphicStyleInfo>; // graphic style-name -> wrap properties
  paraBreaks: Map<string, ParaBreak>; // paragraph style-name -> section break (the odt sectPr equivalent)
  listStarts: Map<string, number>; // list style-name -> level-1 start number
  listRun: { last: number }; // running end-count of the last level-1 ordered list (for continue-numbering)
  tabStops: Map<string, TabStop[]>; // paragraph style-name -> custom tab stops (px)
  masterGeoms: Map<string, string>; // master-page name -> JSON page geometry (for per-section rendering)
  masterBands?: Map<string, { header: string; footer: string; headerLeft?: string; footerLeft?: string; headerFirst?: string; footerFirst?: string }>; // master-page name -> its header/footer HTML (left = even, first = first page)
  notes?: { id: string; kind: "footnote" | "endnote"; html: string }[]; // footnote/endnote bodies, collected inline
}

/** Page geometry (compact JSON) from a style:page-layout-properties element, for per-section
    rendering. Mirrors the docx secGeomJson shape. */
function geomFromLayoutProps(props: Element, lenToPx: (v: string | null) => number | undefined): string | undefined {
  const w = lenToPx(props.getAttribute("fo:page-width"));
  const h = lenToPx(props.getAttribute("fo:page-height"));
  if (!w || !h) return undefined;
  const m = (a: string) => Math.round(Math.max(0, lenToPx(props.getAttribute(`fo:margin-${a}`)) ?? 96));
  const colsEl = props.getElementsByTagName("style:columns")[0];
  const numCols = Number(colsEl?.getAttribute("fo:column-count"));
  const cols = Number.isFinite(numCols) && numCols > 1 ? numCols : undefined;
  const gap = colsEl?.getElementsByTagName("style:column-sep")[0]?.getAttribute("style:width");
  const wm = props.getAttribute("style:writing-mode") ?? "";
  return JSON.stringify({ w: Math.round(w), h: Math.round(h), mt: m("top"), mr: m("right"), mb: m("bottom"), ml: m("left"), cols, colGap: cols ? Math.round(lenToPx(gap) ?? 36) : undefined, vertical: wm.startsWith("tb") || undefined, rtl: wm.startsWith("rl") || undefined });
}

/** Map master-page name -> its page-layout geometry (JSON), so a section that switches master
    page renders at that page's size/orientation/margins/columns. */
function collectMasterGeoms(stylesXml: Uint8Array | undefined, lenToPx: (v: string | null) => number | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!stylesXml) return map;
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const layouts = new Map<string, Element>();
  for (const pl of Array.from(doc.getElementsByTagName("style:page-layout"))) {
    const name = pl.getAttribute("style:name");
    const props = pl.getElementsByTagName("style:page-layout-properties")[0];
    if (name && props) layouts.set(name, props);
  }
  for (const mp of Array.from(doc.getElementsByTagName("style:master-page"))) {
    const mname = mp.getAttribute("style:name");
    const props = layouts.get(mp.getAttribute("style:page-layout-name") ?? "");
    const g = props ? geomFromLayoutProps(props, lenToPx) : undefined;
    if (mname && g) map.set(mname, g);
  }
  return map;
}

/** A paragraph's custom tab stops, shared (as JSON) with the editor and the docx adapter. */
interface TabStop {
  pos: number; // px
  val: string; // left | center | right | decimal
  leader?: string; // "dot" for a dotted leader
}
const ODT_TAB_TYPE: Record<string, string> = { left: "left", center: "center", right: "right", char: "decimal" };

/** Index paragraph-family styles' tab stops (style:tab-stops), in px. */
function collectTabStops(doc: Document, lenToPx: (v: string | null) => number | undefined): Map<string, TabStop[]> {
  const map = new Map<string, TabStop[]>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "paragraph") continue;
    const name = st.getAttribute("style:name");
    const tabs = st.getElementsByTagName("style:tab-stops")[0];
    if (!name || !tabs) continue;
    const stops = Array.from(tabs.getElementsByTagName("style:tab-stop"))
      .map((tb) => ({
        pos: Math.round(lenToPx(tb.getAttribute("style:position")) ?? 0),
        val: ODT_TAB_TYPE[tb.getAttribute("style:type") ?? "left"] ?? "left",
        leader: tb.getAttribute("style:leader-style") && tb.getAttribute("style:leader-style") !== "none" ? "dot" : undefined,
      }))
      .filter((s) => s.pos > 0);
    if (stops.length) map.set(name, stops);
  }
  return map;
}

/** A paragraph style's section break: a page break before/after and/or a new page master. */
interface ParaBreak {
  before?: string; // fo:break-before
  after?: string; // fo:break-after
  master?: string; // style:master-page-name (a new page sequence, i.e. a section)
}

/** Index paragraph-family styles whose break / master-page makes them a section boundary. */
function collectParaBreaks(doc: Document): Map<string, ParaBreak> {
  const map = new Map<string, ParaBreak>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "paragraph") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const pp = st.getElementsByTagName("style:paragraph-properties")[0];
    const before = pp?.getAttribute("fo:break-before") ?? undefined;
    const after = pp?.getAttribute("fo:break-after") ?? undefined;
    const master = st.getAttribute("style:master-page-name") ?? undefined;
    if (before || after || master) map.set(name, { before, after, master });
  }
  return map;
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

// Bookmark markers. ODF pairs start/end by name, so the name doubles as the engine's pairing id.
function bmStartHtml(name: string): string {
  return `<a class="docx-bookmark" data-rdoc-bm="${escapeAttr(name)}" data-rdoc-bm-id="${escapeAttr(name)}" contenteditable="false"></a>`;
}
function bmEndHtml(name: string): string {
  return `<a class="docx-bookmark-end" data-rdoc-bm-id="${escapeAttr(name)}" data-rdoc-bm-end="${escapeAttr(name)}" contenteditable="false"></a>`;
}
// A caption auto-number (text:sequence) -> a seq field the engine renumbers; data-seq keeps the
// sequence name verbatim so it groups and round-trips.
function seqFieldHtml(name: string, cached: string): string {
  return `<span class="docx-field docx-field-seq" data-field="seq" data-seq="${escapeAttr(name)}" contenteditable="false">${escapeHtml(cached)}</span>`;
}

/** A draw:frame holding a draw:image -> an <img> with a data URL; otherwise passthrough. The
 *  frame is stashed (data-odt-xml) so its layout + the draw:image href survive a save, and its
 *  wrap/anchor surfaces as data-rdoc-* so the image toolbar can edit it. */
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
  const lay = imageLayoutAttrs(readOdtLayout(frame, ctx.graphicStyles, lenToPx));
  const alt = frame.getElementsByTagName("svg:desc")[0]?.textContent ?? frame.getElementsByTagName("svg:title")[0]?.textContent ?? "";
  return `<img src="data:${mime};base64,${bytesToBase64(bytes)}" alt="${escapeAttr(alt)}" contenteditable="false"${passthroughAttr(frame)}${dims}${lay}>`;
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
    const mt = pp?.getAttribute("fo:margin-top");
    const mb = pp?.getAttribute("fo:margin-bottom");
    const spaceBeforePx = mt != null ? Math.round(lenToPx(mt) ?? 0) : undefined;
    const spaceAfterPx = mb != null ? Math.round(lenToPx(mb) ?? 0) : undefined;
    if (align || indentPx || lineHeight || spaceBeforePx !== undefined || spaceAfterPx !== undefined)
      map.set(name, { align, indentPx, lineHeight, spaceBeforePx, spaceAfterPx });
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

/** Map list-style-name -> per-level ordered flag (index = level-1): number style = ordered,
    bullet style = unordered. Lets the reader emit <ol> vs <ul> at each nesting depth. */
function collectListStyles(doc: Document): Map<string, boolean[]> {
  const map = new Map<string, boolean[]>();
  for (const ls of Array.from(doc.getElementsByTagName("text:list-style"))) {
    const name = ls.getAttribute("style:name");
    if (!name) continue;
    const levels: boolean[] = [];
    for (const lvl of Array.from(ls.children)) {
      const n = Number(lvl.getAttribute("text:level")) || levels.length + 1;
      levels[n - 1] = lvl.tagName === "text:list-level-style-number";
    }
    map.set(name, levels);
  }
  return map;
}

/** Map list-style-name -> its level-1 start-value (when > 1), so a list that begins at N
    round-trips as <ol start="N">. */
function collectListStarts(doc: Document): Map<string, number> {
  const map = new Map<string, number>();
  for (const ls of Array.from(doc.getElementsByTagName("text:list-style"))) {
    const name = ls.getAttribute("style:name");
    if (!name) continue;
    const lvl1 = Array.from(ls.children).find((l) => l.tagName === "text:list-level-style-number" && (l.getAttribute("text:level") ?? "1") === "1");
    const s = Number(lvl1?.getAttribute("text:start-value"));
    if (Number.isFinite(s) && s > 1) map.set(name, s);
  }
  return map;
}

const cssAttrValue = (v: string): string => v.replace(/[\\"]/g, "\\$&");

/** Read named paragraph and character styles from styles.xml into picker lists, the CSS giving
    each its appearance (data-rdoc-style / data-rdoc-cstyle), and the named-id sets, resolving
    style:parent-style-name. Character (family "text") styles contribute run properties only. */
function readOdtStyles(stylesXml: Uint8Array | undefined): {
  paragraphStyles: { id: string; name: string }[];
  characterStyles: { id: string; name: string }[];
  styleDefs: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[];
  css: string;
  noteCss: string;
  namedPara: Set<string>;
  namedChar: Set<string>;
} {
  const namedPara = new Set<string>();
  const namedChar = new Set<string>();
  if (!stylesXml) return { paragraphStyles: [], characterStyles: [], styleDefs: [], css: "", noteCss: "", namedPara, namedChar };
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const mapFor = (family: string): Map<string, Element> => {
    const m = new Map<string, Element>();
    for (const s of Array.from(doc.getElementsByTagName("style:style"))) {
      if (s.getAttribute("style:family") !== family) continue;
      const id = s.getAttribute("style:name");
      if (id) m.set(id, s);
    }
    return m;
  };
  // `own` = only this style's direct properties (for the edit dialog, so saving does not flatten
  // inherited props in); otherwise the effective CSS, walking style:parent-style-name.
  const cssFor = (byId: Map<string, Element>, id: string, withPara: boolean, seen: Set<string>, own = false): Record<string, string> => {
    const s = byId.get(id);
    if (!s || seen.has(id)) return {};
    seen.add(id);
    const parent = s.getAttribute("style:parent-style-name");
    const out: Record<string, string> = !own && parent ? cssFor(byId, parent, withPara, seen, false) : {};
    const tp = s.getElementsByTagName("style:text-properties")[0];
    if (withPara) {
      const pp = s.getElementsByTagName("style:paragraph-properties")[0];
      const align = ODF_ALIGN[pp?.getAttribute("fo:text-align") ?? ""];
      if (align && align !== "left") out["text-align"] = align;
      const ml = lenToPx(pp?.getAttribute("fo:margin-left"));
      if (ml && ml > 0) out["margin-left"] = `${Math.round(ml)}px`;
      const mt = pp?.getAttribute("fo:margin-top");
      const mb = pp?.getAttribute("fo:margin-bottom");
      if (mt != null) out["margin-top"] = `${Math.round(lenToPx(mt) ?? 0)}px`;
      if (mb != null) out["margin-bottom"] = `${Math.round(lenToPx(mb) ?? 0)}px`;
      const lh = pp?.getAttribute("fo:line-height");
      if (lh && lh.endsWith("%")) out["line-height"] = String(Math.round((parseFloat(lh) / 100) * 100) / 100);
      const pbg = pp?.getAttribute("fo:background-color");
      if (pbg && /^#[0-9a-f]{6}$/i.test(pbg)) out["background-color"] = pbg;
    }
    const bg = tp?.getAttribute("fo:background-color");
    if (bg && /^#[0-9a-f]{6}$/i.test(bg)) out["background-color"] = bg;
    const fw = tp?.getAttribute("fo:font-weight");
    if (fw) out["font-weight"] = /bold|[6-9]00/.test(fw) ? "bold" : "normal";
    const fs = tp?.getAttribute("fo:font-style");
    if (fs) out["font-style"] = fs === "italic" ? "italic" : "normal";
    const deco: string[] = [];
    if ((tp?.getAttribute("style:text-underline-style") ?? "none") !== "none") deco.push("underline");
    if ((tp?.getAttribute("style:text-line-through-style") ?? "none") !== "none") deco.push("line-through");
    if (deco.length) out["text-decoration"] = deco.join(" ");
    const color = tp?.getAttribute("fo:color");
    if (color && /^#[0-9a-f]{6}$/i.test(color)) out["color"] = color;
    const sz = tp?.getAttribute("fo:font-size");
    if (sz && sz.endsWith("pt")) out["font-size"] = sz;
    const font = tp?.getAttribute("fo:font-family") ?? tp?.getAttribute("style:font-name");
    if (font) out["font-family"] = `'${font.replace(/'/g, "")}'`;
    return out;
  };
  let css = "";
  const styleDefs: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[] = [];
  const collect = (family: string, kind: "paragraph" | "character", attr: string, named: Set<string>, skip: (id: string, s: Element) => boolean): { id: string; name: string }[] => {
    const byId = mapFor(family);
    const list: { id: string; name: string }[] = [];
    for (const [id, s] of byId) {
      if (skip(id, s)) continue;
      named.add(id);
      const name = (s.getAttribute("style:display-name") || id).replace(/_20_/g, " ");
      list.push({ id, name });
      styleDefs.push({ id, kind, css: cssFor(byId, id, kind === "paragraph", new Set(), true) }); // own props, for editing
      const decls = cssFor(byId, id, kind === "paragraph", new Set()); // resolved, for the rendered appearance
      const body = Object.entries(decls).map(([k, v]) => `${k}:${v}`).join(";");
      if (body) css += `.docxedit-doc [${attr}="${cssAttrValue(id)}"]{${body}}\n`;
    }
    list.sort((a, b) => a.name.localeCompare(b.name));
    return list;
  };
  const paragraphStyles = collect("paragraph", "paragraph", "data-rdoc-style", namedPara, (id, s) => /^heading(_20_)?[1-9]$/i.test(id) || id === "Standard" || s.getAttribute("style:class") === "extra");
  const characterStyles = collect("text", "character", "data-rdoc-cstyle", namedChar, () => false);
  // The footnote/endnote body style's inheritable text props, for the note area. ODF names the
  // paragraph style "Footnote" (class "extra", so it is skipped from the picker above); fall back
  // to "Endnote".
  let noteCss = "";
  {
    const paraMap = mapFor("paragraph");
    const noteId = ["Footnote", "Endnote"].find((id) => paraMap.has(id)) ?? null;
    if (noteId) {
      const d = cssFor(paraMap, noteId, true, new Set());
      noteCss = (["font-family", "font-size", "line-height", "color"] as const).filter((k) => d[k]).map((k) => `${k}:${d[k]}`).join(";");
    }
  }
  return { paragraphStyles, characterStyles, styleDefs, css, noteCss, namedPara, namedChar };
}

/** Map an automatic style-name -> its parent style-name (from content.xml), for the given
    family, so a block/run that uses an automatic style derived from a named one still shows it. */
function collectAutoParents(doc: Document, family: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const s of Array.from(doc.getElementsByTagName("style:style"))) {
    if (s.getAttribute("style:family") !== family) continue;
    const id = s.getAttribute("style:name");
    const parent = s.getAttribute("style:parent-style-name");
    if (id && parent) map.set(id, parent);
  }
  return map;
}

/** Map table style-name -> left indent (px) from fo:margin-left on the table style. */
function collectTableMargins(doc: Document): Map<string, number> {
  const map = new Map<string, number>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "table") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const ml = lenToPx(st.getElementsByTagName("style:table-properties")[0]?.getAttribute("fo:margin-left"));
    if (ml && ml > 0) map.set(name, Math.round(ml));
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
        const sn = child.getAttribute("text:style-name") ?? "";
        const f = ctx.styles.get(sn) ?? FMT0;
        let inner = wrapFmt(inlineToHtml(child, ctx), f);
        // A named character style behind the span (direct, or the parent of an automatic style).
        const eff = ctx.namedCharStyles.has(sn) ? sn : ctx.namedCharStyles.has(ctx.textAutoParent.get(sn) ?? "") ? ctx.textAutoParent.get(sn)! : "";
        if (eff) inner = `<span data-rdoc-cstyle="${escapeAttr(eff)}">${inner}</span>`;
        html += inner;
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
        html += `<span class="docx-tab" data-docx-tab="1" contenteditable="false">\t</span>`;
        break;
      case "text:s": {
        const n = parseInt(child.getAttribute("text:c") ?? "1", 10) || 1;
        html += n > 1 ? " ".repeat(n) : " ";
        break;
      }
      case "text:note": {
        // Inline footnote / endnote -> an inline reference + a collected note body (engine renumbers).
        const kind = child.getAttribute("text:note-class") === "endnote" ? "endnote" : "footnote";
        const id = child.getAttribute("text:id") || `rdoc-note-${(ctx.notes?.length ?? 0) + 1}`;
        const body = Array.from(child.children).find((c) => c.tagName === "text:note-body");
        let bodyHtml = "";
        if (body) for (const b of Array.from(body.children)) bodyHtml += blockToHtml(b, ctx);
        ctx.notes?.push({ id, kind, html: bodyHtml || "<p><br></p>" });
        html += `<sup class="docx-fnref" data-fn-id="${escapeAttr(id)}" data-fn-kind="${kind}" contenteditable="false"></sup>`;
        break;
      }
      case "text:ruby": {
        // Furigana: text:ruby-base + text:ruby-text -> an HTML <ruby>base<rt>reading</rt></ruby>.
        const baseEl = Array.from(child.children).find((c) => c.tagName === "text:ruby-base");
        const textEl = Array.from(child.children).find((c) => c.tagName === "text:ruby-text");
        const baseHtml = baseEl ? inlineToHtml(baseEl, ctx) : "";
        const rtHtml = textEl ? escapeHtml(textEl.textContent ?? "") : "";
        html += `<ruby>${baseHtml || "&#8203;"}<rt>${rtHtml}</rt></ruby>`;
        break;
      }
      case "text:bookmark": {
        // A point bookmark -> an empty start/end pair, so the model stays uniform with ranges.
        const name = child.getAttribute("text:name") ?? "";
        if (name) html += bmStartHtml(name) + bmEndHtml(name);
        break;
      }
      case "text:bookmark-start": {
        // ODF pairs start/end by name; reuse the name as the engine's pairing id.
        const name = child.getAttribute("text:name") ?? "";
        if (name) html += bmStartHtml(name);
        break;
      }
      case "text:bookmark-end": {
        const name = child.getAttribute("text:name") ?? "";
        if (name) html += bmEndHtml(name);
        break;
      }
      case "text:bookmark-ref": {
        // A cross-reference to a bookmark; the engine recomputes the text, this cached value shows first.
        const name = child.getAttribute("text:ref-name") ?? "";
        const rf = child.getAttribute("text:reference-format");
        const fmt = rf === "page" ? "page" : rf === "direction" ? "direction" : "text";
        html += `<a class="docx-xref" data-rdoc-xref="${escapeAttr(name)}" data-rdoc-xref-fmt="${fmt}" contenteditable="false">${escapeHtml(child.textContent ?? "")}</a>`;
        break;
      }
      case "text:sequence":
        html += seqFieldHtml(child.getAttribute("text:name") ?? "Figure", child.textContent ?? "");
        break;
      default:
        // Unmodelled inline content (bookmarks, notes, change marks, ...) preserved verbatim.
        html += inlinePass(child);
    }
  }
  for (let i = 0; i < ctx.openComment.size; i++) html += "</span>"; // reopened in the next paragraph
  for (let i = 0; i < ctx.openIns.size; i++) html += "</ins>"; // reopened in the next paragraph
  return html;
}

function listToHtml(el: Element, ctx: RCtx, level = 1, styleName = ""): string {
  // A nested text:list usually omits the style-name; inherit the outer one to resolve the level.
  const style = el.getAttribute("text:style-name") || styleName;
  const ordered = ctx.listStyles.get(style)?.[level - 1] ?? false;
  // A top-level ordered list starts at its style's start-value, or continues the previous list
  // when text:continue-numbering / continue-list is set (the odt equivalent of restart/continue).
  let startAttr = "";
  if (level === 1 && ordered) {
    const cont = el.getAttribute("text:continue-numbering") === "true" || el.hasAttribute("text:continue-list");
    const explicit = ctx.listStarts.get(style) ?? 1;
    const start = cont ? ctx.listRun.last + 1 : explicit;
    const count = Array.from(el.children).filter((c) => c.tagName === "text:list-item").length;
    ctx.listRun.last = start - 1 + count;
    if (start > 1) startAttr = ` start="${start}"`;
  }
  let items = "";
  for (const li of Array.from(el.children)) {
    if (li.tagName !== "text:list-item") continue;
    let inner = "";
    for (const block of Array.from(li.children)) {
      if (block.tagName === "text:list") inner += listToHtml(block, ctx, level + 1, style);
      else inner += inlineToHtml(block, ctx);
    }
    items += `<li>${inner || "<br>"}</li>`;
  }
  const tag = ordered ? "ol" : "ul";
  return `<${tag}${startAttr}>${items}</${tag}>`;
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
  const indPx = tstyle ? ctx.tableMargins.get(tstyle) : undefined; // table left indent -> margin-left
  const indStyle = indPx ? ` style="margin-left:${indPx}px"` : "";
  const cols = Array.from(el.children)
    .filter((c) => c.tagName === "table:table-column")
    .map((c) => ({ s: c.getAttribute("table:style-name") ?? "", r: c.getAttribute("table:number-columns-repeated") ?? "1" }));
  const colsAttr = cols.length ? ` data-odt-cols="${escapeAttr(JSON.stringify(cols))}"` : "";
  return `<table class="docx-table" contenteditable="false"${passthroughAttr(el)}${indStyle}${tableStyle}${colsAttr}>${rows}</table>`;
}

function blockToHtml(el: Element, ctx: RCtx): string {
  const alignAttr = (): string => {
    const pf = ctx.paras.get(el.getAttribute("text:style-name") ?? "");
    if (!pf) return "";
    const css: string[] = [];
    if (pf.align && pf.align !== "left") css.push(`text-align:${pf.align}`);
    if (pf.indentPx) css.push(`margin-left:${pf.indentPx}px`);
    if (pf.lineHeight) css.push(`line-height:${pf.lineHeight}`);
    if (pf.spaceBeforePx !== undefined) css.push(`margin-top:${pf.spaceBeforePx}px`);
    if (pf.spaceAfterPx !== undefined) css.push(`margin-bottom:${pf.spaceAfterPx}px`);
    return css.length ? ` style="${css.join(";")}"` : "";
  };
  // The named paragraph style behind this block: a direct reference, or the parent of an
  // automatic style. Recorded as data-rdoc-style so it round-trips and drives the style CSS.
  const namedAttr = (): string => {
    const sn = el.getAttribute("text:style-name") ?? "";
    const eff = ctx.namedStyles.has(sn) ? sn : ctx.namedStyles.has(ctx.autoParent.get(sn) ?? "") ? ctx.autoParent.get(sn)! : "";
    return eff ? ` data-rdoc-style="${escapeAttr(eff)}"` : "";
  };
  // A section break: fo:break-before/after or a new page master on this paragraph's style (the
  // odt equivalent of a docx w:sectPr). Stash the break carried by the paragraph's *own*
  // (automatic) style so an edit re-emits it; a break on a named style rides data-rdoc-style.
  const sn = el.getAttribute("text:style-name") ?? "";
  const own = ctx.paraBreaks.get(sn) ?? {};
  const par = ctx.paraBreaks.get(ctx.autoParent.get(sn) ?? "") ?? {};
  const beforePage = own.before === "page" || !!own.master || par.before === "page" || !!par.master;
  const afterPage = own.after === "page" || par.after === "page";
  let breakAttr = "";
  if (own.before === "page") breakAttr += ` data-odt-break-before="page"`;
  if (own.after === "page") breakAttr += ` data-odt-break-after="page"`;
  if (own.master) breakAttr += ` data-odt-masterpage="${escapeAttr(own.master)}"`;
  // A new page master begins a section: surface its geometry so the section renders at that size.
  const secGeom = own.master ? ctx.masterGeoms.get(own.master) : undefined;
  if (secGeom) breakAttr += ` data-rdoc-secstart="${escapeAttr(secGeom)}"`;
  // Distinct per-section header/footer: key into sectionBands when this section's master has one.
  const mBands = own.master ? ctx.masterBands?.get(own.master) : undefined;
  if (mBands?.header) breakAttr += ` data-rdoc-secheaderkey="oh:${escapeAttr(own.master!)}"`;
  if (mBands?.footer) breakAttr += ` data-rdoc-secfooterkey="of:${escapeAttr(own.master!)}"`;
  const odtPageBreak = `<span class="docx-pagebreak docx-pagebreak-auto" contenteditable="false" data-docx-pagebreak="auto" data-label="${escapeAttr(t("pageBreak"))}"></span>`;
  const before = beforePage ? odtPageBreak : "";
  const after = afterPage ? odtPageBreak : "";
  // Custom tab stops from the paragraph's own or parent style, preserved as data-rdoc-tabstops.
  const stops = ctx.tabStops.get(sn) ?? ctx.tabStops.get(ctx.autoParent.get(sn) ?? "");
  const tabAttr = stops && stops.length ? ` data-rdoc-tabstops="${escapeAttr(JSON.stringify(stops))}"` : "";
  switch (el.tagName) {
    case "text:h": {
      const lvl = Math.min(3, Math.max(1, parseInt(el.getAttribute("text:outline-level") ?? "1", 10) || 1));
      const inner = inlineToHtml(el, ctx);
      return `${before}<h${lvl}${alignAttr()}${breakAttr}${tabAttr}>${inner || "<br>"}</h${lvl}>${after}`;
    }
    case "text:list":
      return listToHtml(el, ctx);
    case "table:table":
      return odtTableHtml(el, ctx);
    case "text:p": {
      const inner = inlineToHtml(el, ctx);
      // A paragraph carrying a caption sequence field is a figure/table caption (type from the seq name).
      const capAttr = inner.includes('data-field="seq"') ? ` data-rdoc-caption="${/data-seq="[^"]*[Tt]able/.test(inner) ? "table" : "figure"}"` : "";
      return `${before}<p${alignAttr()}${namedAttr()}${breakAttr}${tabAttr}${capAttr}>${inner || "<br>"}</p>${after}`;
    }
    default:
      // Tables, tracked-changes, sequence-decls, sections, ... preserved verbatim.
      return blockPass(el);
  }
}

/** Header/footer HTML, read from the master page in styles.xml. */
/** Header/footer HTML per master page (the default master plus any a section switches to). */
function collectMasterBands(files: Record<string, Uint8Array>): Map<string, { header: string; footer: string; headerLeft?: string; footerLeft?: string; headerFirst?: string; footerFirst?: string }> {
  const out = new Map<string, { header: string; footer: string; headerLeft?: string; footerLeft?: string; headerFirst?: string; footerFirst?: string }>();
  const raw = files["styles.xml"];
  if (!raw) return out;
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), cellStyles: collectCellStyles(doc), tableMargins: collectTableMargins(doc), listStyles: collectListStyles(doc), namedStyles: new Set(), autoParent: new Map(), namedCharStyles: new Set(), textAutoParent: new Map(), threads: [], rangedNames: new Set(), openComment: new Set(), changes: new Map(), openIns: new Set(), graphicStyles: collectGraphicStyles(doc, lenToPx), paraBreaks: collectParaBreaks(doc), listStarts: collectListStarts(doc), listRun: { last: 0 }, tabStops: collectTabStops(doc, lenToPx), masterGeoms: collectMasterGeoms(files["styles.xml"], lenToPx) };
  const render = (master: Element, tag: string): string => {
    const el = master.getElementsByTagName(tag)[0];
    if (!el) return "";
    let html = "";
    for (const block of Array.from(el.children)) html += blockToHtml(block, ctx);
    return html;
  };
  for (const master of Array.from(doc.getElementsByTagName("style:master-page"))) {
    const name = master.getAttribute("style:name");
    // style:header-left / style:footer-left hold the even (left) page variant; style:header-first /
    // style:footer-first (ODF 1.3) hold the first-page variant; absent = same as the default. Each
    // only renders its variant when non-empty.
    if (name) out.set(name, { header: render(master, "style:header"), footer: render(master, "style:footer"), headerLeft: render(master, "style:header-left") || undefined, footerLeft: render(master, "style:footer-left") || undefined, headerFirst: render(master, "style:header-first") || undefined, footerFirst: render(master, "style:footer-first") || undefined });
  }
  return out;
}
function readHeaderFooter(files: Record<string, Uint8Array>): { header: string; footer: string } {
  const raw = files["styles.xml"];
  if (!raw) return { header: "", footer: "" };
  const doc = new DOMParser().parseFromString(strFromU8(raw), "application/xml");
  const first = doc.getElementsByTagName("style:master-page")[0]; // the body's master is the default
  const name = first?.getAttribute("style:name");
  return (name && collectMasterBands(files).get(name)) || { header: "", footer: "" };
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
  // Columns: style:columns @fo:column-count on the page layout, with the gap from a column-sep
  // or the first column's margins.
  const colsEl = props.getElementsByTagName("style:columns")[0];
  const numCols = Number(colsEl?.getAttribute("fo:column-count"));
  const columns = Number.isFinite(numCols) && numCols > 1 ? numCols : undefined;
  const gapAttr = colsEl?.getElementsByTagName("style:column-sep")[0]?.getAttribute("style:width");
  return { widthPx: Math.round(w), heightPx: Math.round(h), margin: { top: m("top"), right: m("right"), bottom: m("bottom"), left: m("left") }, vertical, rtl, columns, columnGapPx: columns ? Math.round(lenToPx(gapAttr) ?? 36) : undefined };
}

export function odtToParts(bytes: Uint8Array): { body: string; comments: CommentThread[]; header: string; footer: string; headerEven?: { html: string; path?: string }; footerEven?: { html: string; path?: string }; headerFirst?: { html: string; path?: string }; footerFirst?: { html: string; path?: string }; sectionBands?: Record<string, { html: string; path: string }>; notes?: { id: string; kind: "footnote" | "endnote"; html: string }[]; page?: PageGeometry; paragraphStyles?: { id: string; name: string }[]; characterStyles?: { id: string; name: string }[]; styleDefs?: { id: string; kind: "paragraph" | "character"; css: Record<string, string> }[]; styleCss?: string; noteCss?: string } {
  const files = unzipSync(bytes);
  const { header, footer } = readHeaderFooter(files);
  // Per-master header/footer HTML, so a section that switches master shows its own. Keyed for the
  // engine as oh:/of:<master>; the write-back path is header@/footer@<master>.
  const masterBands = collectMasterBands(files);
  const sectionBands: Record<string, { html: string; path: string }> = {};
  for (const [name, b] of masterBands) {
    if (b.header) sectionBands[`oh:${name}`] = { html: b.header, path: `header@${name}` };
    if (b.footer) sectionBands[`of:${name}`] = { html: b.footer, path: `footer@${name}` };
  }
  const page = parsePageGeometry(files);
  // Even-page (left) header/footer variant from the body's default master (its first entry).
  const defName = [...masterBands.keys()][0];
  const defBand = defName ? masterBands.get(defName) : undefined;
  const headerEven = defBand?.headerLeft ? { html: defBand.headerLeft, path: `header-left@${defName}` } : undefined;
  const footerEven = defBand?.footerLeft ? { html: defBand.footerLeft, path: `footer-left@${defName}` } : undefined;
  const headerFirst = defBand?.headerFirst ? { html: defBand.headerFirst, path: `header-first@${defName}` } : undefined;
  const footerFirst = defBand?.footerFirst ? { html: defBand.footerFirst, path: `footer-first@${defName}` } : undefined;
  if (page && (headerEven || footerEven)) page.evenOdd = true;
  if (page && (headerFirst || footerFirst)) page.titlePage = true;
  const ps = readOdtStyles(files["styles.xml"]);
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
  // Graphic styles (for image wrap) and paragraph breaks (for sections) can live in content.xml's
  // automatic styles or styles.xml; merge both, preferring content.xml.
  const graphicStyles = collectGraphicStyles(doc, lenToPx);
  const paraBreaks = collectParaBreaks(doc);
  const listStarts = collectListStarts(doc);
  const tabStops = collectTabStops(doc, lenToPx);
  const stylesRaw = files["styles.xml"];
  if (stylesRaw) {
    const sdoc = new DOMParser().parseFromString(strFromU8(stylesRaw), "application/xml");
    for (const [k, v] of collectGraphicStyles(sdoc, lenToPx)) if (!graphicStyles.has(k)) graphicStyles.set(k, v);
    for (const [k, v] of collectParaBreaks(sdoc)) if (!paraBreaks.has(k)) paraBreaks.set(k, v);
    for (const [k, v] of collectListStarts(sdoc)) if (!listStarts.has(k)) listStarts.set(k, v);
    for (const [k, v] of collectTabStops(sdoc, lenToPx)) if (!tabStops.has(k)) tabStops.set(k, v);
  }
  const ctx: RCtx = { files, styles: collectTextStyles(doc), paras: collectParaStyles(doc), cellStyles: collectCellStyles(doc), tableMargins: collectTableMargins(doc), listStyles: collectListStyles(doc), namedStyles: ps.namedPara, autoParent: collectAutoParents(doc, "paragraph"), namedCharStyles: ps.namedChar, textAutoParent: collectAutoParents(doc, "text"), threads: [], rangedNames, openComment: new Set(), changes: readChanges(body), openIns: new Set(), graphicStyles, paraBreaks, listStarts, listRun: { last: 0 }, tabStops, masterGeoms: collectMasterGeoms(files["styles.xml"], lenToPx), masterBands, notes: [] };
  let html = "";
  for (const block of Array.from(body.children)) {
    if (block.tagName === "text:tracked-changes") continue; // metadata, parsed into ctx.changes
    html += blockToHtml(block, ctx);
  }
  return { body: html || "<p><br></p>", comments: ctx.threads, header, footer, headerEven, footerEven, headerFirst, footerFirst, sectionBands, notes: ctx.notes, page, paragraphStyles: ps.paragraphStyles, characterStyles: ps.characterStyles, styleDefs: ps.styleDefs, styleCss: ps.css, noteCss: ps.noteCss };
}

/** Convert an .odt's body to HTML. Returns "" if there is no editable text body. */
export function odtToHtml(bytes: Uint8Array): string {
  return odtToParts(bytes).body;
}

