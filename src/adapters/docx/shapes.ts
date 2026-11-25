// DrawingML / VML shapes and text boxes: a box with the shape's size, fill, outline and simple
// geometry, its text box content editable inside, positioned like a floating image. The run is
// preserved verbatim; the writer edits only what the user changed (text, size, fill, outline).
import { drawingColor, themeColor } from "./fonts";
import { escapeAttr } from "./shared";
import { readLayout } from "./image-layout";
import { wordArtSvg } from "./wordart";
import { customGeometry, presetGeometry, type GeomPath } from "./drawingml-geom";
import { SHADE, shadeColor } from "../../core/feature/shape-geom";

const EMU_PER_PX = 9525;
const emu = (v: string | null | undefined): number | undefined => {
  const n = v ? Number(v) : NaN;
  return Number.isFinite(n) ? n / EMU_PER_PX : undefined;
};
const kid = (el: Element | undefined, tag: string): Element | undefined => (el ? Array.from(el.children).find((c) => c.tagName === tag) : undefined);
const px = (n: number) => `${Math.round(n * 100) / 100}px`;

export interface ShapeCtx {
  colors: Record<string, string>;
  /** Render a w:txbxContent's blocks to HTML. */
  blocks: (txbx: Element) => string;
  /** A picture's data URL by relationship id, for pictures inside a group. */
  imageSrc: (rid: string) => string | undefined;
}

// Text box content is block HTML, but the shape sits inside a paragraph: block tags become spans
// with a class that restores their display, so the browser keeps the paragraph whole.
const BLOCK_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6", "UL", "OL", "LI", "TABLE", "THEAD", "TBODY", "TR", "TD", "TH", "DIV", "FIGURE", "BLOCKQUOTE", "PRE"]);
function inlineBlocks(html: string): string {
  const body = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html").body;
  for (const el of Array.from(body.querySelectorAll("*")).reverse()) {
    if (!BLOCK_TAGS.has(el.tagName)) continue;
    const span = body.ownerDocument.createElement("span");
    // Every attribute rides along (the writer restores the block from them); class is kept aside.
    for (const a of Array.from(el.attributes)) span.setAttribute(a.name === "class" ? "data-tb-class" : a.name, a.value);
    span.className = `docx-tb-${el.tagName.toLowerCase()}`;
    while (el.firstChild) span.appendChild(el.firstChild);
    el.replaceWith(span);
  }
  return body.innerHTML;
}

export interface Look { fill?: string; line?: string; lineW?: number; radius?: string; text?: string; rot?: number }
/** A wps:wsp's look: spPr fill / outline / geometry, falling back to its wps:style references. */
function wspLook(sp: Element, ctx: ShapeCtx): Look {
  const spPr = kid(sp, "wps:spPr");
  const style = kid(sp, "wps:style");
  const look: Look = {};
  if (kid(spPr, "a:noFill")) look.fill = undefined;
  else look.fill = drawingColor(kid(spPr, "a:solidFill"), ctx.colors) ?? (Number(kid(style, "a:fillRef")?.getAttribute("idx") ?? "0") > 0 ? drawingColor(kid(style, "a:fillRef"), ctx.colors) : undefined);
  const ln = kid(spPr, "a:ln");
  if (!kid(ln, "a:noFill")) {
    look.line = drawingColor(kid(ln, "a:solidFill"), ctx.colors) ?? (Number(kid(style, "a:lnRef")?.getAttribute("idx") ?? "0") > 0 ? drawingColor(kid(style, "a:lnRef"), ctx.colors) : undefined);
    look.lineW = emu(ln?.getAttribute("w")) ?? (look.line ? 1 : undefined);
  }
  look.text = drawingColor(kid(style, "a:fontRef"), ctx.colors);
  const prst = kid(spPr, "a:prstGeom")?.getAttribute("prst");
  if (prst === "ellipse" || prst === "flowChartConnector") look.radius = "50%";
  else if (prst === "roundRect" || prst === "flowChartAlternateProcess") look.radius = "16.7%";
  const rot = Number(kid(spPr, "a:xfrm")?.getAttribute("rot") ?? "0");
  if (rot) look.rot = rot / 60000;
  return look;
}

function boxCss(look: Look): string[] {
  const css: string[] = [];
  if (look.fill) css.push(`background-color:#${look.fill}`);
  if (look.line) css.push(`border:${px(Math.max(0.5, look.lineW ?? 1))} solid #${look.line}`);
  if (look.radius) css.push(`border-radius:${look.radius}`);
  if (look.text) css.push(`color:#${look.text}`);
  if (look.rot) css.push(`transform:rotate(${look.rot}deg)`);
  return css;
}

/** The text body of a wps:wsp (or a v:textbox): insets, vertical anchor and direction. */
function bodyHtml(txbx: Element | undefined, bodyPr: Element | undefined, ctx: ShapeCtx): string {
  if (!txbx) return "";
  const ins = (a: string, d: number) => emu(bodyPr?.getAttribute(a)) ?? d / EMU_PER_PX;
  const css = [`padding:${px(ins("tIns", 45720))} ${px(ins("rIns", 91440))} ${px(ins("bIns", 45720))} ${px(ins("lIns", 91440))}`];
  const anchor = bodyPr?.getAttribute("anchor");
  if (anchor === "ctr") css.push("justify-content:center");
  else if (anchor === "b") css.push("justify-content:flex-end");
  const vert = bodyPr?.getAttribute("vert");
  if (vert === "vert" || vert === "eaVert" || vert === "wordArtVert") css.push("writing-mode:vertical-rl");
  else if (vert === "vert270") css.push("writing-mode:vertical-rl;transform:rotate(180deg)");
  // Editable in place; the writer rebuilds this text box's w:txbxContent from it when edited.
  return `<span class="docx-shape-body" contenteditable="true" style="${css.join(";")}">${inlineBlocks(ctx.blocks(txbx)) || '<span class="docx-tb-p"><br></span>'}</span>`;
}

// Line ends (a:headEnd / a:tailEnd): sizes are multiples of the line width.
const END_SIZE: Record<string, number> = { sm: 2, med: 3, lg: 5 };

/** An arrowhead at `tip`, pointing away from `from`. */
export function lineEnd(end: Element | undefined, from: [number, number], tip: [number, number], lw: number): string {
  const type = end?.getAttribute("type");
  if (!type || type === "none") return "";
  const k = Math.max(lw, 2.65); // thin lines get heads sized for a 0.7 mm line, as Word and LibreOffice do
  const hw = (END_SIZE[end!.getAttribute("w") ?? "med"] ?? 3) * k / 2, len = (END_SIZE[end!.getAttribute("len") ?? "med"] ?? 3) * k;
  let dx = tip[0] - from[0], dy = tip[1] - from[1];
  const n = Math.hypot(dx, dy) || 1;
  dx /= n; dy /= n;
  // Points in the head's frame: u along the line toward the tip, v across it.
  const at = (u: number, v: number) => `${r2(tip[0] + dx * u - dy * v)},${r2(tip[1] + dy * u + dx * v)}`;
  const p = (d: string, kind = "solid") => `<path data-head="${kind}" d="${d}"/>`;
  if (type === "triangle") return p(`M${at(0, 0)} L${at(-len, hw)} L${at(-len, -hw)} Z`);
  if (type === "stealth") return p(`M${at(0, 0)} L${at(-len, hw)} L${at(-len * 0.6, 0)} L${at(-len, -hw)} Z`);
  if (type === "diamond") return p(`M${at(len / 2, 0)} L${at(0, hw)} L${at(-len / 2, 0)} L${at(0, -hw)} Z`);
  const rot = r2((Math.atan2(dy, dx) * 180) / Math.PI);
  if (type === "oval") return p(`M${at(len / 2, 0)} A${r2(len / 2)},${r2(hw)} ${rot} 1 0 ${at(-len / 2, 0)} A${r2(len / 2)},${r2(hw)} ${rot} 1 0 ${at(len / 2, 0)} Z`);
  if (type === "arrow") return p(`M${at(-len, hw)} L${at(0, 0)} L${at(-len, -hw)}`, "open");
  return "";
}

const BOX_PRESETS = new Set(["rect", "ellipse", "flowChartConnector", "flowChartProcess"]);
const ROUND_PRESETS = new Set(["roundRect", "flowChartAlternateProcess"]);

/** The outline of a shape that is not a plain rectangle, as an SVG over the box, painted from the
 *  look (the box keeps the look as CSS for the shape bar and the writer). */
export function geomSvg(spPr: Element | undefined, w: number, h: number, look: Look): string | undefined {
  const prstEl = kid(spPr, "a:prstGeom"), cust = kid(spPr, "a:custGeom");
  const prst = prstEl?.getAttribute("prst");
  let paths: GeomPath[] | undefined;
  if (cust) paths = customGeometry(cust, w, h);
  else if (prst) {
    const adj: Record<string, string> = {};
    for (const g of Array.from(kid(prstEl, "a:avLst")?.children ?? [])) if (g.getAttribute("name")) adj[g.getAttribute("name")!] = g.getAttribute("fmla") ?? "";
    // Boxes CSS draws exactly stay CSS boxes.
    if (BOX_PRESETS.has(prst) || (ROUND_PRESETS.has(prst) && !Object.keys(adj).length)) return undefined;
    paths = presetGeometry(prst, w, h, adj);
  }
  if (!paths) return undefined;
  const bw = look.line ? Math.round(Math.max(0.5, look.lineW ?? 1) * 100) / 100 : 0;
  const line = look.line ? `#${look.line}` : "none";
  let body = "";
  for (const g of paths) {
    const fill = g.fill === "none" || !look.fill ? "none" : shadeColor(`#${look.fill}`, SHADE[g.fill] ?? 0);
    const stroke = g.stroke && look.line ? ` stroke="${line}" stroke-width="${bw}"` : ` stroke="none"`;
    body += `<path d="${g.d}" data-fill="${g.fill}" data-stroke="${g.stroke ? 1 : 0}" fill="${fill}"${stroke}/>`;
  }
  // Arrowheads go on the first open, stroked path (a line, a connector, an arc).
  const ln = kid(spPr, "a:ln");
  const open = paths.find((g) => g.stroke && !g.closed && g.start && g.end);
  if (open && look.line) {
    const heads = lineEnd(kid(ln, "a:headEnd"), open.start![1], open.start![0], bw) + lineEnd(kid(ln, "a:tailEnd"), open.end![0], open.end![1], bw);
    body += heads.replace(/data-head="(\w+)"/g, (_m, kind: string) => `data-head="${kind}" ${kind === "open" ? `fill="none" stroke="${line}" stroke-width="${bw}"` : `fill="${line}" stroke="none"`}`);
  }
  const xfrm = kid(spPr, "a:xfrm");
  const fh = xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true";
  const fv = xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true";
  const flip = fh || fv ? ` transform="translate(${fh ? r2(w) : 0},${fv ? r2(h) : 0}) scale(${fh ? -1 : 1},${fv ? -1 : 1})"` : "";
  return `<svg class="rdoc-shape-geom" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${r2(w)} ${r2(h)}" preserveAspectRatio="none" aria-hidden="true" style="left:${-bw}px;top:${-bw}px;width:calc(100% + ${2 * bw}px);height:calc(100% + ${2 * bw}px)" stroke-linejoin="round"><g${flip}>${body}</g></svg>`;
}
const r2 = (n: number) => String(Math.round(n * 100) / 100);

function wspHtml(sp: Element, w: number | undefined, h: number | undefined, ctx: ShapeCtx, extraCss: string[] = []): string {
  const txbx = kid(kid(sp, "wps:txbx"), "w:txbxContent");
  const look = wspLook(sp, ctx);
  const geom = w !== undefined && h !== undefined ? geomSvg(kid(sp, "wps:spPr"), w, h, look) : undefined;
  if (geom) look.radius = undefined;
  // A line has no height (or width): its box keeps that size, the outline overflows it.
  const size = geom ? [`width:${px(w!)}`, `height:${px(h!)}`] : [...(w ? [`width:${px(w)}`] : []), ...(h ? [`height:${px(h)}`] : [])];
  const css = [...extraCss, ...size, ...boxCss(look)];
  return `<span class="docx-shape-box"${geom ? " data-geom" : ""} style="${css.join(";")}">${geom ?? ""}${bodyHtml(txbx, kid(sp, "wps:bodyPr"), ctx)}</span>`;
}

/** A wpg:wgp group: each member placed in the group's box, scaled from its child coordinates. */
function groupHtml(g: Element, w: number, h: number, ctx: ShapeCtx): string {
  const xfrm = kid(kid(g, "wpg:grpSpPr"), "a:xfrm");
  const chOff = kid(xfrm, "a:chOff"), chExt = kid(xfrm, "a:chExt");
  const cx = Number(chOff?.getAttribute("x") ?? 0), cy = Number(chOff?.getAttribute("y") ?? 0);
  const sx = w / ((emu(chExt?.getAttribute("cx")) ?? w) || w), sy = h / ((emu(chExt?.getAttribute("cy")) ?? h) || h);
  let out = "";
  for (const m of Array.from(g.children)) {
    const spPr = kid(m, "wps:spPr") ?? kid(m, "pic:spPr") ?? kid(m, "wpg:grpSpPr");
    const x = kid(spPr, "a:xfrm");
    const off = kid(x, "a:off"), ext = kid(x, "a:ext");
    if (!off || !ext) continue;
    const pos = [`position:absolute`, `left:${px(((Number(off.getAttribute("x")) - cx) / EMU_PER_PX) * sx)}`, `top:${px(((Number(off.getAttribute("y")) - cy) / EMU_PER_PX) * sy)}`];
    const mw = (emu(ext.getAttribute("cx")) ?? 0) * sx, mh = (emu(ext.getAttribute("cy")) ?? 0) * sy;
    if (m.tagName === "wps:wsp") out += wspHtml(m, mw, mh, ctx, pos);
    else if (m.tagName === "pic:pic") {
      const rid = m.getElementsByTagName("a:blip")[0]?.getAttribute("r:embed");
      const src = rid ? ctx.imageSrc(rid) : undefined;
      if (src) out += `<img src="${escapeAttr(src)}" alt="" style="${[...pos, `width:${px(mw)}`, `height:${px(mh)}`].join(";")}">`;
    } else if (m.tagName === "wpg:grpSp") out += `<span class="docx-shape-box" style="${[...pos, `width:${px(mw)}`, `height:${px(mh)}`].join(";")}">${groupHtml(m, mw, mh, ctx)}</span>`;
  }
  return out;
}

// VML (w:pict) fallbacks: fill / stroke attributes and a v:textbox.
const vmlColor = (v: string | null, ctx: ShapeCtx): string | undefined => {
  if (!v) return undefined;
  const m = /^#?([0-9a-f]{6})\b/i.exec(v.trim());
  if (m) return m[1]!.toUpperCase();
  const named: Record<string, string> = { black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", gray: "808080", silver: "C0C0C0" };
  return named[v.trim().toLowerCase()] ?? themeColor(ctx.colors, v.trim());
};
function vmlHtml(pict: Element, ctx: ShapeCtx): string | undefined {
  const shape = Array.from(pict.getElementsByTagName("*")).find((e) => /^v:(shape|rect|roundrect|oval)$/.test(e.tagName) && e.getElementsByTagName("w:txbxContent").length);
  if (!shape) return undefined;
  const st = shape.getAttribute("style") ?? "";
  const dim = (k: string) => { const m = new RegExp(`(?:^|;)\\s*${k}:\\s*([\\d.]+)(pt|px|in|cm|mm)?`).exec(st); return m ? Number(m[1]) * ({ pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, px: 1 } as Record<string, number>)[m[2] ?? "px"]! : undefined; };
  const look: Look = {};
  if (shape.getAttribute("filled") !== "f" && shape.getAttribute("filled") !== "false") look.fill = vmlColor(shape.getAttribute("fillcolor"), ctx) ?? "FFFFFF";
  if (shape.getAttribute("stroked") !== "f" && shape.getAttribute("stroked") !== "false") {
    look.line = vmlColor(shape.getAttribute("strokecolor"), ctx) ?? "000000";
    const sw = /^([\d.]+)pt$/.exec(shape.getAttribute("strokeweight") ?? "");
    look.lineW = sw ? Number(sw[1]) * (4 / 3) : 1;
  }
  if (shape.tagName === "v:oval") look.radius = "50%";
  if (shape.tagName === "v:roundrect") look.radius = "16.7%";
  const w = dim("width"), h = dim("height");
  const css = [...(w ? [`width:${px(w)}`] : []), ...(h ? [`height:${px(h)}`] : []), ...boxCss(look)];
  return `<span class="docx-shape-box" style="${css.join(";")}">${bodyHtml(shape.getElementsByTagName("w:txbxContent")[0], undefined, ctx)}</span>`;
}

/** A run's shape / text box / group as read-only HTML, or undefined when it holds none. */
export function shapeHtml(run: Element, pass: string, ctx: ShapeCtx): string | undefined {
  const drawing = run.getElementsByTagName("w:drawing")[0];
  const wrapAttrs = (lay: ReturnType<typeof readLayout>): string => {
    if (!lay) return "";
    let a = ` data-shape-wrap="${lay.wrap}" data-shape-align="${lay.align}"`;
    if (lay.wrap === "behind" || lay.wrap === "front") a += ` style="left:${lay.x}px;top:${lay.y}px"`;
    return a;
  };
  if (drawing) {
    const ext = drawing.getElementsByTagName("wp:extent")[0];
    const w = emu(ext?.getAttribute("cx")) ?? 0, h = emu(ext?.getAttribute("cy")) ?? 0;
    const sp = drawing.getElementsByTagName("wps:wsp")[0];
    const grp = drawing.getElementsByTagName("wpg:wgp")[0];
    if (!sp && !grp) return undefined;
    // WordArt (warped or effect text): drawn as a picture, the run kept verbatim.
    const art = !grp && sp ? wordArtSvg(sp, ctx.colors, w, h) : undefined;
    if (art) return `<span class="docx-shape docx-shape-art" contenteditable="false"${pass}${wrapAttrs(readLayout(run))}><span class="docx-shape-box" style="width:${px(w)};height:${px(h)}"><img src="data:image/svg+xml;charset=utf-8,${escapeAttr(encodeURIComponent(art))}" alt="" style="width:100%;height:100%"></span></span>`;
    const inner = grp ? `<span class="docx-shape-box" style="width:${px(w)};height:${px(h)}">${groupHtml(grp, w, h, ctx)}</span>` : wspHtml(sp!, w, h, ctx);
    return `<span class="docx-shape" contenteditable="false"${pass}${wrapAttrs(readLayout(run))}>${inner}</span>`;
  }
  const pict = run.getElementsByTagName("w:pict")[0];
  const v = pict && vmlHtml(pict, ctx);
  return v ? `<span class="docx-shape" contenteditable="false"${pass}>${v}</span>` : undefined;
}

/** Turn text box spans back into block HTML (docx-tb-p -> p, ...) for the block writer. */
export function blocksFromBody(body: Element): DocumentFragment {
  const doc = body.ownerDocument;
  const frag = doc.createDocumentFragment();
  const clone = body.cloneNode(true) as Element;
  for (const el of Array.from(clone.querySelectorAll("span[class^='docx-tb-']")).reverse()) {
    const tag = /^docx-tb-([a-z0-9]+)/.exec(el.className)?.[1];
    if (!tag) continue;
    const blk = doc.createElement(tag);
    for (const a of Array.from(el.attributes)) {
      if (a.name === "class") continue;
      blk.setAttribute(a.name === "data-tb-class" ? "class" : a.name, a.value);
    }
    while (el.firstChild) blk.appendChild(el.firstChild);
    el.replaceWith(blk);
  }
  // Loose inline content (typed outside any paragraph) gets a paragraph of its own.
  let loose: HTMLElement | null = null;
  for (const n of Array.from(clone.childNodes)) {
    const isBlock = n.nodeType === 1 && BLOCK_TAGS.has((n as Element).tagName);
    if (isBlock) { loose = null; frag.appendChild(n); continue; }
    if (n.nodeType === 3 && !(n.textContent ?? "").trim() && !loose) continue;
    if (!loose) { loose = doc.createElement("p"); frag.appendChild(loose); }
    loose.appendChild(n);
  }
  return frag;
}
