// Drawing shapes in an odt (lines, freeform paths, polylines, polygons, rectangles, ellipses,
// custom shapes, groups): drawn read-only as an SVG picture laid over the text where the
// file puts it. The element itself is kept verbatim (data-odt-xml), so a save writes it back as
// it was. Coordinates are px throughout; a shape's position is relative to its anchor (the
// paragraph or character it is anchored to, or the page).

import { evaluateGeometry, type Geometry } from "./geometry";
import { groupBox } from "./shape-edit";
import { escapeAttr, escapeHtml, passthroughAttr } from "./shared";

type LenToPx = (v: string | null | undefined) => number | undefined;

/** How a shape is drawn: the graphic properties that matter, resolved through parent styles. */
export interface Look {
  stroke: "none" | "solid" | "dash";
  strokeColor: string;
  strokeWidth: number; // px; 0 = hairline
  strokeOpacity: number;
  fill: boolean;
  fillColor: string;
  opacity: number;
  markerStart?: string;
  markerEnd?: string;
  markerStartWidth: number; // px
  markerEndWidth: number;
  /** A marker centred on the line's end (draw:marker-*-center) rather than tipped at it. */
  markerStartCenter?: boolean;
  markerEndCenter?: boolean;
  wrap?: string; // style:wrap
  runThrough?: string; // style:run-through
  hpos?: string; // style:horizontal-pos
  vpos?: string; // style:vertical-pos
  hrel?: string; // style:horizontal-rel
  vrel?: string; // style:vertical-rel
}

interface Marker {
  vb: [number, number, number, number];
  d: string;
}

export interface ShapeCtx {
  looks: Map<string, Look>;
  markers: Map<string, Marker>;
  lenToPx: LenToPx;
}

const SHAPE_TAGS = new Set([
  "draw:line", "draw:path", "draw:polyline", "draw:polygon", "draw:rect", "draw:ellipse", "draw:circle",
  "draw:custom-shape", "draw:connector", "draw:g", "draw:measure",
]);
export const isShapeTag = (tag: string): boolean => SHAPE_TAGS.has(tag);

const DEFAULT_LOOK: Look = {
  stroke: "solid", strokeColor: "#3465a4", strokeWidth: 0, strokeOpacity: 1, fill: true, fillColor: "#729fcf", opacity: 1,
  markerStartWidth: 0, markerEndWidth: 0,
};

/** Every graphic style (content.xml automatic + styles.xml) resolved through its parents, plus
 *  the arrowhead markers the documents define. */
export function collectShapeStyles(docs: (Document | undefined)[], lenToPx: LenToPx): ShapeCtx {
  const raw = new Map<string, Element>();
  let defaults: Element | undefined;
  const markers = new Map<string, Marker>();
  for (const doc of docs) {
    if (!doc) continue;
    for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
      if (st.getAttribute("style:family") !== "graphic") continue;
      const name = st.getAttribute("style:name");
      if (name && !raw.has(name)) raw.set(name, st);
    }
    for (const d of Array.from(doc.getElementsByTagName("style:default-style"))) {
      if (d.getAttribute("style:family") === "graphic" && !defaults) defaults = d;
    }
    for (const m of Array.from(doc.getElementsByTagName("draw:marker"))) {
      const vb = (m.getAttribute("svg:viewBox") ?? "").trim().split(/[\s,]+/).map(Number);
      const d = m.getAttribute("svg:d") ?? "";
      if (vb.length !== 4 || vb.some((n) => !Number.isFinite(n)) || !d) continue;
      for (const key of [m.getAttribute("draw:name"), m.getAttribute("draw:display-name")]) {
        if (key && !markers.has(key)) markers.set(key, { vb: vb as Marker["vb"], d });
      }
    }
  }
  const apply = (look: Look, el: Element | undefined): void => {
    const gp = el?.getElementsByTagName("style:graphic-properties")[0];
    if (!gp) return;
    const a = (n: string) => gp.getAttribute(n);
    const stroke = a("draw:stroke");
    if (stroke === "none" || stroke === "solid" || stroke === "dash") look.stroke = stroke;
    if (a("svg:stroke-color")) look.strokeColor = a("svg:stroke-color")!;
    const sw = lenToPx(a("svg:stroke-width"));
    if (sw !== undefined) look.strokeWidth = sw;
    const so = percent(a("svg:stroke-opacity"));
    if (so !== undefined) look.strokeOpacity = so;
    const fill = a("draw:fill");
    if (fill) look.fill = fill !== "none";
    if (a("draw:fill-color")) look.fillColor = a("draw:fill-color")!;
    const op = percent(a("draw:opacity"));
    if (op !== undefined) look.opacity = op;
    if (gp.hasAttribute("draw:marker-start")) look.markerStart = a("draw:marker-start") || undefined;
    if (gp.hasAttribute("draw:marker-end")) look.markerEnd = a("draw:marker-end") || undefined;
    const msw = lenToPx(a("draw:marker-start-width"));
    if (msw !== undefined) look.markerStartWidth = msw;
    const mew = lenToPx(a("draw:marker-end-width"));
    if (mew !== undefined) look.markerEndWidth = mew;
    if (gp.hasAttribute("draw:marker-start-center")) look.markerStartCenter = a("draw:marker-start-center") === "true";
    if (gp.hasAttribute("draw:marker-end-center")) look.markerEndCenter = a("draw:marker-end-center") === "true";
    if (a("style:wrap")) look.wrap = a("style:wrap")!;
    if (a("style:run-through")) look.runThrough = a("style:run-through")!;
    if (a("style:horizontal-pos")) look.hpos = a("style:horizontal-pos")!;
    if (a("style:vertical-pos")) look.vpos = a("style:vertical-pos")!;
    if (a("style:horizontal-rel")) look.hrel = a("style:horizontal-rel")!;
    if (a("style:vertical-rel")) look.vrel = a("style:vertical-rel")!;
  };
  const looks = new Map<string, Look>();
  const resolve = (name: string, seen = new Set<string>()): Look => {
    const done = looks.get(name);
    if (done) return done;
    const st = raw.get(name);
    const parent = st?.getAttribute("style:parent-style-name");
    const look: Look = parent && !seen.has(parent) ? { ...resolve(parent, seen.add(name)) } : { ...DEFAULT_LOOK };
    if (!parent || seen.has(parent)) apply(look, defaults);
    apply(look, st);
    looks.set(name, look);
    return look;
  };
  for (const name of raw.keys()) resolve(name);
  const base = { ...DEFAULT_LOOK };
  apply(base, defaults);
  looks.set("", base);
  return { looks, markers, lenToPx };
}

function percent(v: string | null): number | undefined {
  if (!v) return undefined;
  const n = parseFloat(v);
  return Number.isFinite(n) ? (v.trim().endsWith("%") ? n / 100 : n) : undefined;
}

type Pt = [number, number];

/** A shape in px, ready to draw: SVG markup, the extent it covers, and its ends (for arrowheads). */
interface Drawn {
  svg: string;
  box: [number, number, number, number]; // minX, minY, maxX, maxY
}

const fmt = (n: number): string => String(Math.round(n * 100) / 100);
const grow = (box: Drawn["box"], p: Pt, pad = 0): void => {
  box[0] = Math.min(box[0], p[0] - pad);
  box[1] = Math.min(box[1], p[1] - pad);
  box[2] = Math.max(box[2], p[0] + pad);
  box[3] = Math.max(box[3], p[1] + pad);
};
const emptyBox = (): Drawn["box"] => [Infinity, Infinity, -Infinity, -Infinity];

/** A path's commands with every coordinate passed through `map`, as absolute commands, plus the
 *  first and last segments' directions for arrowheads. Arcs keep their radii scaled by `scale`. */
function mapPath(d: string, map: (p: Pt) => Pt, scale: Pt): { d: string; pts: Pt[]; start?: [Pt, Pt]; end?: [Pt, Pt] } {
  const tokens = d.match(/[a-zA-Z]|-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  let i = 0;
  let cmd = "";
  let cur: Pt = [0, 0];
  let startPt: Pt = [0, 0];
  let prevCtrl: Pt | null = null;
  const out: string[] = [];
  const pts: Pt[] = [];
  // Per drawn segment: its start (the point, and the way it leaves it) and its end (the way it
  // arrives, and the point).
  const segs: { s: [Pt, Pt]; e: [Pt, Pt] }[] = [];
  const num = () => Number(tokens[i++]);
  const isNum = () => i < tokens.length && !/^[a-zA-Z]$/.test(tokens[i]!);
  const put = (c: string, ...ps: Pt[]) => {
    const m = ps.map(map);
    out.push(c + m.map((p) => `${fmt(p[0])},${fmt(p[1])}`).join(" "));
    pts.push(...m);
  };
  while (i < tokens.length) {
    if (!isNum()) cmd = tokens[i++]!;
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const at = (x: number, y: number): Pt => (rel ? [cur[0] + x, cur[1] + y] : [x, y]);
    if (C === "M") {
      cur = at(num(), num());
      startPt = cur;
      put("M", cur);
      cmd = rel ? "l" : "L"; // further pairs are line-tos
      prevCtrl = null;
    } else if (C === "L" || C === "H" || C === "V") {
      const p: Pt = C === "L" ? at(num(), num()) : C === "H" ? [rel ? cur[0] + num() : num(), cur[1]] : [cur[0], rel ? cur[1] + num() : num()];
      segs.push({ s: [cur, p], e: [cur, p] });
      put("L", p);
      cur = p;
      prevCtrl = null;
    } else if (C === "C" || C === "S") {
      const c1: Pt = C === "C" ? at(num(), num()) : prevCtrl ? [2 * cur[0] - prevCtrl[0], 2 * cur[1] - prevCtrl[1]] : cur;
      const c2 = at(num(), num());
      const p = at(num(), num());
      segs.push({ s: [cur, c1], e: [c2, p] });
      put("C", c1, c2, p);
      prevCtrl = c2;
      cur = p;
    } else if (C === "Q" || C === "T") {
      const c: Pt = C === "Q" ? at(num(), num()) : prevCtrl ? [2 * cur[0] - prevCtrl[0], 2 * cur[1] - prevCtrl[1]] : cur;
      const p = at(num(), num());
      segs.push({ s: [cur, c], e: [c, p] });
      put("Q", c, p);
      prevCtrl = c;
      cur = p;
    } else if (C === "A") {
      const rx = num(), ry = num(), rot = num(), large = num(), sweep = num();
      const p = at(num(), num());
      segs.push({ s: [cur, p], e: [cur, p] });
      const m = map(p);
      out.push(`A${fmt(Math.abs(rx * scale[0]))},${fmt(Math.abs(ry * scale[1]))} ${fmt(rot)} ${large ? 1 : 0} ${sweep ? 1 : 0} ${fmt(m[0])},${fmt(m[1])}`);
      pts.push(m);
      cur = p;
      prevCtrl = null;
    } else if (C === "Z") {
      out.push("Z");
      cur = startPt;
      prevCtrl = null;
    } else {
      i++; // unknown command: skip a token rather than loop forever
    }
  }
  // Arrowheads point along the first and last segments, away from the line.
  const moves = ([a, b]: [Pt, Pt]) => a[0] !== b[0] || a[1] !== b[1];
  const first = segs.find((g) => moves(g.s))?.s;
  const last = [...segs].reverse().find((g) => moves(g.e))?.e;
  return {
    d: out.join(" "),
    pts,
    // [tip, the way it points from]: the start's tip is where the path begins.
    start: first ? [map(first[0]), map(first[1])] : undefined,
    end: last ? [map(last[0]), map(last[1])] : undefined,
  };
}

/** An arrowhead at `tip`, coming from `from`: the marker's shape (tip at the top centre of its
 *  view box, as odf draws markers) scaled to `width` px and turned along the line. */
function arrowhead(marker: Marker, from: Pt, tip: Pt, width: number, color: string, box: Drawn["box"], center = false): string {
  const dx = tip[0] - from[0], dy = tip[1] - from[1];
  const len = Math.hypot(dx, dy) || 1;
  const [vx, vy, vw, vh] = marker.vb;
  const k = width / (vw || 1);
  const theta = (Math.atan2(dx / len, -dy / len) * 180) / Math.PI;
  grow(box, tip, Math.max(width, vh * k));
  return `<path d="${escapeAttr(marker.d)}" fill="${escapeAttr(color)}" stroke="none" transform="translate(${fmt(tip[0])},${fmt(tip[1])}) rotate(${fmt(theta)}) scale(${fmt(k)}) translate(${fmt(-(vx + vw / 2))},${fmt(-(center ? vy + vh / 2 : vy))})"/>`;
}

function strokeAttrs(look: Look): string {
  if (look.stroke === "none") return ` stroke="none"`;
  const w = look.strokeWidth || 1;
  const dash = look.stroke === "dash" ? ` stroke-dasharray="${fmt(w * 4)},${fmt(w * 3)}"` : "";
  const op = look.strokeOpacity < 1 ? ` stroke-opacity="${fmt(look.strokeOpacity)}"` : "";
  return ` stroke="${escapeAttr(look.strokeColor)}" stroke-width="${fmt(w)}" stroke-linecap="round" stroke-linejoin="round"${dash}${op}`;
}

/** Arrowheads on an open line's ends, and the line pulled back a little so it doesn't poke out of a pointed head. */
function ends(look: Look, ctx: ShapeCtx, start: [Pt, Pt] | undefined, end: [Pt, Pt] | undefined, box: Drawn["box"]): string {
  if (look.stroke === "none") return "";
  let out = "";
  const ms = look.markerStart ? ctx.markers.get(look.markerStart) : undefined;
  const me = look.markerEnd ? ctx.markers.get(look.markerEnd) : undefined;
  const w = look.strokeWidth || 1;
  if (ms && start) out += arrowhead(ms, start[1], start[0], look.markerStartWidth || w * 3, look.strokeColor, box, look.markerStartCenter);
  if (me && end) out += arrowhead(me, end[0], end[1], look.markerEndWidth || w * 3, look.strokeColor, box, look.markerEndCenter);
  return out;
}

/** Centred text a shape carries (a label in a box or ellipse), drawn as simple SVG lines. */
function shapeText(el: Element, cx: number, cy: number): string {
  const lines = Array.from(el.children)
    .filter((c) => c.tagName === "text:p")
    .map((p) => (p.textContent ?? "").trim())
    .filter(Boolean);
  if (!lines.length) return "";
  const size = 13, lead = size * 1.25;
  const top = cy - ((lines.length - 1) * lead) / 2;
  return lines
    .map((t, n) => `<text x="${fmt(cx)}" y="${fmt(top + n * lead)}" font-size="${size}" font-family="sans-serif" text-anchor="middle" dominant-baseline="middle" fill="#000">${escapeHtml(t)}</text>`)
    .join("");
}

/** `rotate(a) translate(x y)` (odf draw:transform) as a point mapping, or undefined. */
function transformOf(el: Element, ctx: ShapeCtx): ((p: Pt) => Pt) | undefined {
  const t = el.getAttribute("draw:transform");
  if (!t) return undefined;
  const ops = [...t.matchAll(/(rotate|translate|skewX|scale)\s*\(([^)]*)\)/g)];
  if (!ops.length) return undefined;
  const fns = ops.map(([, op, args]) => {
    const v = args!.trim().split(/[\s,]+/);
    if (op === "rotate") {
      const a = -Number(v[0]); // odf rotates counter-clockwise in radians
      const c = Math.cos(a), s = Math.sin(a);
      return (p: Pt): Pt => [p[0] * c - p[1] * s, p[0] * s + p[1] * c];
    }
    if (op === "translate") {
      const x = ctx.lenToPx(v[0]) ?? 0, y = ctx.lenToPx(v[1] ?? "0") ?? 0;
      return (p: Pt): Pt => [p[0] + x, p[1] + y];
    }
    if (op === "scale") {
      const sx = Number(v[0]) || 1, sy = Number(v[1] ?? v[0]) || 1;
      return (p: Pt): Pt => [p[0] * sx, p[1] * sy];
    }
    return (p: Pt): Pt => p;
  });
  return (p) => fns.reduce((q, f) => f(q), p);
}

/** Draw one shape (or group) into `acc`, in px relative to its anchor. */
/** A #rrggbb colour darkened toward black (k < 0) or lightened toward white (k > 0). */
function shadeColor(color: string, k: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(color.trim());
  if (!m) return color;
  const n = parseInt(m[1]!, 16);
  const ch = (v: number) => Math.round(k < 0 ? v * (1 + k) : v + (255 - v) * k);
  return "#" + [n >> 16, (n >> 8) & 255, n & 255].map((v) => ch(v).toString(16).padStart(2, "0")).join("");
}

/** An invisible wide stroke along a line, so a thin line can be clicked. */
const hitPath = (d: string) => `<path class="odt-shape-hit" d="${d}" fill="none" stroke="transparent" stroke-width="10"/>`;

/** A shape's own size in px (its width and height; a line's extent), or undefined for a group. */
export function shapeExtent(el: Element, lenToPx: LenToPx): { w: number; h: number } | undefined {
  const L = (n: string) => lenToPx(el.getAttribute(n)) ?? 0;
  if (el.tagName === "draw:g") return undefined;
  if (el.hasAttribute("svg:x1")) return { w: Math.abs(L("svg:x2") - L("svg:x1")), h: Math.abs(L("svg:y2") - L("svg:y1")) };
  return { w: L("svg:width"), h: L("svg:height") };
}

/** The look the shape bar edits, as data attributes on the shape element. */
function modelAttrs(el: Element, look: Look, ctx: ShapeCtx): string {
  const z = el.getAttribute("draw:z-index");
  const zAttr = z && /^\d+$/.test(z) ? ` data-shape-z="${z}"` : "";
  const hex = (c: string) => (/^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : "#000000");
  const paint = (lk: Look) => ` data-shape-fill="${lk.fill ? hex(lk.fillColor) : "none"}" data-shape-line="${lk.stroke === "none" ? "none" : hex(lk.strokeColor)}" data-shape-linew="${fmt(lk.stroke === "none" ? 0 : lk.strokeWidth || 1)}"`;
  const ext = shapeExtent(el, ctx.lenToPx);
  if (!ext) {
    // A group shows its first filled member's look; an edit applies to every member.
    const kids = Array.from(el.getElementsByTagName("*")).filter((c) => isShapeTag(c.tagName) && c.tagName !== "draw:g");
    const looks = kids.map((c) => ctx.looks.get(c.getAttribute("draw:style-name") ?? "") ?? look);
    const first = looks.find((l) => l.fill) ?? looks[0];
    const b = groupBox(el);
    return ` data-shape-group=""${first ? paint(first) : ""}${b ? ` data-shape-w="${fmt(b.w)}" data-shape-h="${fmt(b.h)}"` : ""}${zAttr}`;
  }
  return `${paint(look)} data-shape-w="${fmt(ext.w)}" data-shape-h="${fmt(ext.h)}"${zAttr}`;
}

function drawShape(el: Element, ctx: ShapeCtx, look: Look, acc: { svg: string[]; box: Drawn["box"] }): void {
  const L = (n: string) => ctx.lenToPx(el.getAttribute(n)) ?? 0;
  const tag = el.tagName;
  if (tag === "draw:g") {
    for (const child of Array.from(el.children)) {
      if (!isShapeTag(child.tagName)) continue;
      drawShape(child, ctx, ctx.looks.get(child.getAttribute("draw:style-name") ?? "") ?? look, acc);
    }
    return;
  }
  const xf = transformOf(el, ctx);
  const x = xf ? 0 : L("svg:x"), y = xf ? 0 : L("svg:y");
  const w = L("svg:width"), h = L("svg:height");
  const place = (p: Pt): Pt => (xf ? xf(p) : [p[0] + x, p[1] + y]);
  const fill = look.fill ? ` fill="${escapeAttr(look.fillColor)}"` : ` fill="none"`;
  const pad = (look.strokeWidth || 1) / 2 + 1;
  const pathOut = (d: string, closed: boolean, fromBox?: { vb: number[] }, paint?: { fill: boolean; stroke: boolean; shade: number }) => {
    // A path in its own view box, scaled into the shape's box (or already in px for a connector).
    const vb = fromBox?.vb;
    const sx = vb && vb[2] ? w / vb[2] : 1, sy = vb && vb[3] ? h / vb[3] : 1;
    const map = (p: Pt): Pt => (vb ? place([(p[0] - vb[0]) * sx, (p[1] - vb[1]) * sy]) : p);
    const m = mapPath(d, map, [sx, sy]);
    for (const p of m.pts) grow(acc.box, p, pad);
    if (paint) {
      const shaded = paint.shade ? ` fill="${escapeAttr(shadeColor(look.fillColor, paint.shade))}"` : fill;
      acc.svg.push(`<path d="${m.d}"${paint.fill ? shaded : ` fill="none"`}${paint.stroke ? strokeAttrs(look) : ` stroke="none"`}/>`);
      // An open outline (a connector drawn as a custom shape) takes the line's ends too.
      if (!closed && paint.stroke) acc.svg.push(hitPath(m.d), ends(look, ctx, m.start, m.end, acc.box));
      return;
    }
    if (!closed) acc.svg.push(hitPath(m.d));
    acc.svg.push(`<path d="${m.d}"${closed ? fill : ` fill="none"`}${strokeAttrs(look)}/>`);
    if (!closed) acc.svg.push(ends(look, ctx, m.start, m.end, acc.box));
  };
  if (tag === "draw:line" || tag === "draw:connector" || tag === "draw:measure") {
    const a = place([L("svg:x1") - (xf ? 0 : x), L("svg:y1") - (xf ? 0 : y)]);
    const b = place([L("svg:x2") - (xf ? 0 : x), L("svg:y2") - (xf ? 0 : y)]);
    const connectorD = tag === "draw:connector" ? el.getAttribute("svg:d") : null;
    if (connectorD && el.getAttribute("svg:viewBox")) {
      pathOut(connectorD, false, { vb: el.getAttribute("svg:viewBox")!.trim().split(/[\s,]+/).map(Number) });
      return;
    }
    grow(acc.box, a, pad);
    grow(acc.box, b, pad);
    const d = `M${fmt(a[0])},${fmt(a[1])} L${fmt(b[0])},${fmt(b[1])}`;
    acc.svg.push(hitPath(d), `<path d="${d}" fill="none"${strokeAttrs(look)}/>`);
    acc.svg.push(ends(look, ctx, [a, b], [a, b], acc.box));
    return;
  }
  const vb = (el.getAttribute("svg:viewBox") ?? "").trim().split(/[\s,]+/).map(Number);
  if (tag === "draw:path") {
    if (vb.length === 4) pathOut(el.getAttribute("svg:d") ?? "", /z\s*$/i.test(el.getAttribute("svg:d") ?? ""), { vb });
    return;
  }
  if (tag === "draw:polyline" || tag === "draw:polygon") {
    const nums = (el.getAttribute("draw:points") ?? "").trim().split(/[\s,]+/).map(Number);
    if (vb.length !== 4 || nums.length < 4) return;
    let d = "";
    for (let k = 0; k + 1 < nums.length; k += 2) d += `${k ? "L" : "M"}${nums[k]},${nums[k + 1]} `;
    pathOut(d + (tag === "draw:polygon" ? "Z" : ""), tag === "draw:polygon", { vb });
    return;
  }
  // Boxes: rectangles, ellipses, circles and the custom shapes that are one of those.
  let kind: "rect" | "ellipse" | "path" | "" = tag === "draw:rect" ? "rect" : tag === "draw:ellipse" || tag === "draw:circle" ? "ellipse" : "";
  let r = ctx.lenToPx(el.getAttribute("draw:corner-radius")) ?? 0;
  let geom: Geometry | undefined;
  if (tag === "draw:custom-shape") {
    const g = el.getElementsByTagName("draw:enhanced-geometry")[0];
    const type = g?.getAttribute("draw:type") ?? "";
    if (/^(rectangle|rect|mso-spt1)$/.test(type)) kind = "rect";
    else if (/^(round-rectangle|mso-spt2)$/.test(type)) {
      kind = "rect";
      r = Math.min(w, h) / 6;
    } else if (/^(ellipse|mso-spt3)$/.test(type)) kind = "ellipse";
    else if (g && w > 0 && h > 0) {
      // Everything else is drawn from its geometry: equations, handles and the enhanced path.
      geom = evaluateGeometry(g, w, h);
      if (geom) kind = "path";
    }
  }
  if (!kind || !(w > 0 && h > 0)) return;
  if (kind === "path" && geom) {
    for (const sub of geom.paths) pathOut(sub.d, /z\s*$/i.test(sub.d), { vb: geom.vb }, { fill: sub.fill && look.fill, stroke: sub.stroke, shade: sub.shade });
  } else {
    const corners: Pt[] = [place([0, 0]), place([w, 0]), place([w, h]), place([0, h])];
    for (const p of corners) grow(acc.box, p, pad);
    const matrix = xf ? (() => {
      const o = place([0, 0]), ex = place([1, 0]), ey = place([0, 1]);
      return ` transform="matrix(${fmt(ex[0] - o[0])},${fmt(ex[1] - o[1])},${fmt(ey[0] - o[0])},${fmt(ey[1] - o[1])},${fmt(o[0])},${fmt(o[1])})"`;
    })() : ` transform="translate(${fmt(x)},${fmt(y)})"`;
    acc.svg.push(
      kind === "rect"
        ? `<rect width="${fmt(w)}" height="${fmt(h)}"${r ? ` rx="${fmt(r)}"` : ""}${fill}${strokeAttrs(look)}${matrix}/>`
        : `<ellipse cx="${fmt(w / 2)}" cy="${fmt(h / 2)}" rx="${fmt(w / 2)}" ry="${fmt(h / 2)}"${fill}${strokeAttrs(look)}${matrix}/>`,
    );
  }
  const c = place([w / 2, h / 2]);
  acc.svg.push(shapeText(el, c[0], c[1]));
}

/** A drawn shape as HTML: a read-only box at its place, holding an SVG picture of it, with the
 *  element kept verbatim for the writer. `undefined` when nothing in it can be drawn. */
export function shapeHtml(el: Element, ctx: ShapeCtx, level: "inline" | "block"): string | undefined {
  const look = ctx.looks.get(el.getAttribute("draw:style-name") ?? "") ?? ctx.looks.get("") ?? DEFAULT_LOOK;
  const acc = { svg: [] as string[], box: emptyBox() };
  drawShape(el, ctx, look, acc);
  const [x0, y0, x1, y1] = acc.box;
  if (!acc.svg.length || !Number.isFinite(x0) || x1 <= x0 || y1 <= y0) return undefined;
  const bw = x1 - x0, bh = y1 - y0;
  const op = look.opacity < 1 ? ` opacity="${fmt(look.opacity)}"` : "";
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(bw)}" height="${fmt(bh)}" viewBox="${fmt(x0)} ${fmt(y0)} ${fmt(bw)} ${fmt(bh)}"><g${op}>${acc.svg.join("")}</g></svg>`;
  // Inline, so only what is painted takes a click (the shape bar selects a shape by it).
  const name = el.getAttribute("draw:name");
  const img = svg.replace("<svg ", `<svg class="odt-shape-svg" role="img"${name ? ` aria-label="${escapeAttr(name)}"` : ` aria-hidden="true"`} `);
  const model = modelAttrs(el, look, ctx);
  const anchor = el.getAttribute("text:anchor-type") ?? "paragraph";
  if (anchor === "as-char") {
    return `<span class="odt-shape" data-shape-wrap="inline" contenteditable="false"${model}${passthroughAttr(el)} style="width:${fmt(bw)}px;height:${fmt(bh)}px">${img}</span>`;
  }
  // Wrapped beside the text (style:wrap other than run-through): floated at its paragraph's start,
  // aligned as its style says. Running through the text: behind or in front of it, where it is.
  const wrapped = level === "inline" && look.wrap && look.wrap !== "run-through" ? (look.wrap === "none" ? "topbottom" : look.wrap === "dynamic" ? "tight" : "square") : undefined;
  if (wrapped) {
    const align = look.hpos === "center" || look.hpos === "right" ? look.hpos : "left";
    const margin = align === "left" && x0 > 0 ? `;margin-left:${fmt(x0)}px` : "";
    return `<span class="odt-shape" data-shape-wrap="${wrapped}" data-shape-align="${align}" contenteditable="false"${model}${passthroughAttr(el)} style="width:${fmt(bw)}px;height:${fmt(bh)}px${margin}"><span class="odt-shape-box" style="position:relative;width:${fmt(bw)}px;height:${fmt(bh)}px">${img}</span></span>`;
  }
  const wrap = look.wrap === "run-through" && look.runThrough === "background" ? "behind" : "front";
  const left = x0, top = y0;
  const style = `left:${fmt(left)}px;top:${fmt(top)}px;width:${fmt(bw)}px;height:${fmt(bh)}px`;
  const box = `<span class="odt-shape-box" style="${style}">${img}</span>`;
  // Its offsets are from the page or its margins when the style says so (the editor places it
  // from there); otherwise from its paragraph.
  const rel = (v: string | undefined, content: string) => (!v || !/^page/.test(v) ? "para" : v === content ? "margin" : "page");
  // A drawing anchored to the page is placed from the page it names.
  const onPage = anchor === "page" ? ` data-shape-page="${Number(el.getAttribute("text:anchor-page-number")) || 1}"` : "";
  const placed = ` data-shape-relh="${anchor === "page" && !look.hrel ? "page" : rel(look.hrel, "page-content")}" data-shape-relv="${anchor === "page" && !look.vrel ? "page" : rel(look.vrel, "page-content")}" data-shape-bx="${fmt(left)}" data-shape-by="${fmt(top)}"${onPage}`;
  // A page-anchored shape sits in the body between paragraphs: a zero-height block placed from the
  // document's corner. One anchored to a paragraph or character is placed from that paragraph.
  return level === "block"
    ? `<div class="odt-shape odt-shape-page" data-shape-wrap="${wrap}"${placed} contenteditable="false"${model}${passthroughAttr(el)}>${box}</div>`
    : `<span class="odt-shape" data-shape-wrap="${wrap}"${placed} contenteditable="false"${model}${passthroughAttr(el)}>${box}</span>`;
}
