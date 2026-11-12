/*! richdoc SmartArt drawing: a port of WordCraft's diagram reader (https://github.com/storytold/wordcraft).
 * Copyright (c) 2026 ArtCraft Team and the WordCraft contributors. @license MIT
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
 * associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute,
 * sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions: The above copyright notice and this
 * permission notice shall be included in all copies or substantial portions of the Software.
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
 * NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
 * DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
 * OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. */
// SmartArt: the ready-made drawing (dsp:drawing) Word keeps beside the diagram data, drawn as SVG.
// The diagram's layout isn't run; the saved drawing is the picture. Units here are CSS pixels.
import { strFromU8 } from "fflate";
import { drawingColor } from "./fonts";

const MAX_GROUP_DEPTH = 8;
const MAX_SHAPES = 10_000;
export const MAX_ITEMS = 50_000;
const EMU_PX = 9525;
const MAX_EMU = 1.0e9;
const MAX_LEN_PX = 19_200;
const LINE = 1.2;
// Estimated character width in em (wordcraft's chart CHAR_W).
const CHAR_W = 0.55;
const PT_PX = 4 / 3;

export type Align = "left" | "center" | "right";
export type SmartArtItem =
  | { k: "geom"; prst: string; rect: [number, number, number, number]; rot: number; flipH: boolean; flipV: boolean; adj: Record<string, number>; fill?: string; stroke?: string; strokeWidth: number }
  | { k: "image"; rect: [number, number, number, number]; href: string }
  | { k: "text"; rect: [number, number, number, number]; text: string; size: number; color: string; bold: boolean; align: Align; font?: string };

// Affine map per axis: x' = kx*x + ox, y' = ky*y + oy.
interface Xf { kx: number; ox: number; ky: number; oy: number }
const ID: Xf = { kx: 1, ox: 0, ky: 1, oy: 0 };
const then = (o: Xf, i: Xf): Xf => ({ kx: o.kx * i.kx, ox: o.kx * i.ox + o.ox, ky: o.ky * i.ky, oy: o.ky * i.oy + o.oy });
function mapRect(f: Xf, b: number[]): [number, number, number, number] {
  const x0 = f.kx * b[0]! + f.ox, x1 = f.kx * (b[0]! + b[2]!) + f.ox;
  const y0 = f.ky * b[1]! + f.oy, y1 = f.ky * (b[1]! + b[3]!) + f.oy;
  return [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
}

const child = (e: Element | undefined | null, tag: string): Element | undefined => (e ? Array.from(e.children).find((c) => c.tagName === tag) : undefined);
const children = (e: Element | undefined | null, tag: string): Element[] => (e ? Array.from(e.children).filter((c) => c.tagName === tag) : []);

// wordcraft's units::int: an integer or a finite decimal, rounded.
export function int(s: string | null | undefined): number | undefined {
  const t = s?.trim();
  if (!t || t.length > 40 || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return undefined;
  const v = Number(t);
  return Number.isFinite(v) ? Math.round(v) : undefined;
}
const px = (v: number) => (Number.isNaN(v) ? 0 : Math.max(-MAX_LEN_PX, Math.min(MAX_LEN_PX, v)));
function emu(e: Element | undefined, attr: string): number | undefined {
  const v = int(e?.getAttribute(attr));
  return v === undefined ? undefined : Math.max(-MAX_EMU, Math.min(MAX_EMU, v));
}
const on = (v: string | null | undefined) => v === "1" || v === "true";

// a:off / a:ext of an a:xfrm or dsp:txXfrm as [x, y, w, h] EMU.
function xfrmBox(x: Element | undefined): number[] | undefined {
  const off = child(x, "a:off"), ext = child(x, "a:ext");
  if (!off || !ext) return undefined;
  const v = [emu(off, "x"), emu(off, "y"), emu(ext, "cx"), emu(ext, "cy")];
  if (v.some((n) => n === undefined)) return undefined;
  return [v[0]!, v[1]!, Math.max(0, v[2]!), Math.max(0, v[3]!)];
}

// Child-space to parent-space map of a group, from its a:xfrm (off/ext against chOff/chExt).
function groupXf(grp: Element): Xf {
  const x = child(child(grp, "dsp:grpSpPr"), "a:xfrm");
  if (!x) return ID;
  const v = (tag: string, a: string) => emu(child(x, tag), a) ?? 0;
  const scale = (e: number, c: number) => (c > 0 && e >= 0 ? e / c : 1);
  const kx = scale(v("a:ext", "cx"), v("a:chExt", "cx"));
  const ky = scale(v("a:ext", "cy"), v("a:chExt", "cy"));
  return { kx, ox: v("a:off", "x") - v("a:chOff", "x") * kx, ky, oy: v("a:off", "y") - v("a:chOff", "y") * ky };
}

function collect(tree: Element, xf: Xf, depth: number, out: [Xf, Element][]) {
  for (const e of Array.from(tree.children)) {
    if (e.tagName === "dsp:sp" && out.length < MAX_SHAPES) out.push([xf, e]);
    else if (e.tagName === "dsp:grpSp" && depth < MAX_GROUP_DEPTH) collect(e, then(xf, groupXf(e)), depth + 1, out);
  }
}

/** The items of a SmartArt drawing, scaled so its extent fits the w x h (px) frame, never above 1:1. */
export function smartArtItems(drawing: Element, colors: Record<string, string>, w: number, h: number, image?: (rId: string) => string | undefined): SmartArtItem[] {
  const tree = drawing.getElementsByTagName("dsp:spTree")[0];
  if (!tree) return [];
  const shapes: [Xf, Element][] = [];
  collect(tree, ID, 0, shapes);
  const placed: [Xf, Element, number[]][] = [];
  for (const [xf, sp] of shapes) {
    const b = xfrmBox(child(child(sp, "dsp:spPr"), "a:xfrm"));
    if (b) placed.push([xf, sp, b]);
  }
  let extX = 0, extY = 0;
  for (const [xf, , b] of placed) {
    const r = mapRect(xf, b);
    extX = Math.max(extX, r[2]);
    extY = Math.max(extY, r[3]);
  }
  const frame = (f: number) => (Number.isFinite(f) && f > 0 ? f * EMU_PX : 0);
  const fitOf = (fr: number, ext: number) => (fr > 0 && ext > 0 ? fr / ext : 1);
  const fit = Math.min(1, fitOf(frame(w), extX), fitOf(frame(h), extY));
  const root: Xf = { kx: fit / EMU_PX, ox: 0, ky: fit / EMU_PX, oy: 0 };
  const out: SmartArtItem[] = [];
  for (const [xf, sp, b] of placed) {
    if (out.length >= MAX_ITEMS) break;
    shape(sp, then(root, xf), b, colors, fit, image, out);
  }
  return out;
}

// A preset's adjust values (a:avLst a:gd fmla="val N"), raw.
export function presetAdjusts(geom: Element | undefined): Record<string, number> {
  const out: Record<string, number> = {};
  for (const gd of children(child(geom, "a:avLst"), "a:gd")) {
    const name = gd.getAttribute("name"), f = gd.getAttribute("fmla")?.trim();
    if (!name || !f?.startsWith("val ")) continue;
    const v = Number(f.slice(4).trim());
    if (Number.isFinite(v)) out[name] = Math.max(-1e9, Math.min(1e9, v));
  }
  return out;
}

// One dsp:sp: its picture or outline, then its text.
function shape(sp: Element, xf: Xf, b: number[], colors: Record<string, string>, fit: number, image: ((rId: string) => string | undefined) | undefined, out: SmartArtItem[]) {
  const [x0, y0, x1, y1] = mapRect(xf, b);
  if (![x0, y0, x1, y1].every(Number.isFinite)) return;
  const w = x1 - x0, h = y1 - y0;
  const sppr = child(sp, "dsp:spPr");
  const xfrm = child(sppr, "a:xfrm");
  const r = int(xfrm?.getAttribute("rot")) ?? 0;
  const rot = (((r % 21_600_000) + 21_600_000) % 21_600_000) / 60_000;
  const rect: [number, number, number, number] = [px(x0), px(y0), px(w), px(h)];
  let filled: boolean;
  const blip = child(sppr, "a:blipFill")?.getElementsByTagName("a:blip")[0];
  if (blip) {
    // A picture fill stretches into the shape's box; no clipping to the outline.
    const id = blip.getAttribute("r:embed");
    const href = id && image ? image(id) : undefined;
    if (href) out.push({ k: "image", rect, href });
    filled = false;
  } else {
    const fill = fillOf(sp, sppr, colors);
    const [stroke, strokeWidth] = lineOf(sp, sppr, colors);
    const geom = child(sppr, "a:prstGeom");
    const prst = geom?.getAttribute("prst") || "rect";
    out.push({ k: "geom", prst, rect, rot, flipH: on(xfrm?.getAttribute("flipH")), flipV: on(xfrm?.getAttribute("flipV")), adj: presetAdjusts(geom), fill, stroke, strokeWidth });
    filled = fill !== undefined;
  }
  text(sp, xf, b, colors, filled, fit, out);
}

// A fill element's colour: a gradient's first stop, a pattern's foreground.
function paint(e: Element, colors: Record<string, string>): string | undefined {
  switch (e.tagName) {
    case "a:gradFill": return drawingColor(child(child(e, "a:gsLst"), "a:gs"), colors);
    case "a:pattFill": return drawingColor(child(e, "a:fgClr"), colors);
    case "a:noFill": case "a:blipFill": return undefined;
    default: return drawingColor(e, colors);
  }
}

// The colour of a shape's style reference (a:fillRef, a:lnRef); index 0 means none.
function styleColour(sp: Element, which: string, colors: Record<string, string>): string | undefined {
  const r = child(child(sp, "dsp:style"), which);
  if (!r || r.getAttribute("idx") === "0") return undefined;
  return drawingColor(r, colors);
}

function fillOf(sp: Element, sppr: Element | undefined, colors: Record<string, string>): string | undefined {
  const own = Array.from(sppr?.children ?? []).find((e) => /^a:(solidFill|noFill|gradFill|pattFill)$/.test(e.tagName));
  return own ? paint(own, colors) : styleColour(sp, "a:fillRef", colors);
}

// Outline colour and width in px (a:ln, else the style's line at 0.75 pt).
function lineOf(sp: Element, sppr: Element | undefined, colors: Record<string, string>): [string | undefined, number] {
  const ln = child(sppr, "a:ln");
  const wv = int(ln?.getAttribute("w"));
  const width = wv === undefined ? 0.75 * PT_PX : Math.max(0, Math.min(100 * PT_PX, wv / EMU_PX));
  const own = Array.from(ln?.children ?? []).find((e) => /^a:(solidFill|noFill|gradFill)$/.test(e.tagName));
  if (own?.tagName === "a:noFill") return [undefined, 0];
  if (own) return [paint(own, colors), width];
  return [styleColour(sp, "a:lnRef", colors), width];
}

// v limited to [0, hi] (hi below 0 counts as 0).
const lim = (v: number, hi: number) => Math.min(Math.max(v, 0), Math.max(hi, 0));
const mod = (a: number, n: number) => ((a % n) + n) % n;

type Poly = [Array<[number, number]>, boolean];

// A pie slice or block arc from adj1 to adj2 (60000ths of a degree, clockwise from 3 o'clock).
function sector(block: boolean, w: number, h: number, adj: Record<string, number>): Poly {
  const deg = (n: string, d: number) => (adj[n] ?? d) / 60_000;
  const [st, en] = block ? [deg("adj1", 10_800_000), deg("adj2", 0)] : [deg("adj1", 0), deg("adj2", 16_200_000)];
  const sweep = mod(en - st, 360);
  const hw = w / 2, hh = h / 2;
  const arc = (rx: number, ry: number): Array<[number, number]> => {
    const pts: Array<[number, number]> = [];
    for (let k = 0; k <= 36; k++) {
      const t = ((st + (sweep * k) / 36) * Math.PI) / 180;
      pts.push([hw + rx * Math.cos(t), hh + ry * Math.sin(t)]);
    }
    return pts;
  };
  if (block) {
    const ss = Math.min(w, h);
    const band = lim(ss * ((adj.adj3 ?? 25_000) / 100_000), ss / 2);
    return [[...arc(hw, hh), ...arc(hw - band, hh - band).reverse()], true];
  }
  return [[[hw, hh], ...arc(hw, hh)], true];
}

// A block arrow along len (tip at len) across thick; a1 shaft thickness, a2 head length.
function arrow(len: number, thick: number, a1: number, a2: number, double: boolean): Array<[number, number]> {
  const head = lim(Math.min(len, thick) * a2, len / 2);
  const half = lim((thick * a1) / 2, thick / 2);
  const c = thick / 2;
  if (double) return [[0, c], [head, 0], [head, c - half], [len - head, c - half], [len - head, 0], [len, c], [len - head, thick], [len - head, c + half], [head, c + half], [head, thick]];
  return [[0, c - half], [len - head, c - half], [len - head, 0], [len, c], [len - head, thick], [len - head, c + half], [0, c + half]];
}

// A w x h rectangle with corners of radius r, each corner in six steps.
function rounded(w: number, h: number, r: number): Array<[number, number]> {
  const pts: Array<[number, number]> = [];
  for (const [cx, cy, start] of [[w - r, r, -90], [w - r, h - r, 0], [r, h - r, 90], [r, r, 180]] as const) {
    for (let k = 0; k <= 6; k++) {
      const t = ((start + 15 * k) * Math.PI) / 180;
      pts.push([cx + r * Math.cos(t), cy + r * Math.sin(t)]);
    }
  }
  return pts;
}

/** The outline of a preset in a w x h box, or undefined for curved presets (ellipse, heart, stars). Unknown presets are rectangles. */
export function presetOutline(prst: string, w: number, h: number, adjRaw: Record<string, number> = {}): Poly | undefined {
  const ss = Math.min(w, h);
  const a = (n: string, d: number) => (adjRaw[n] ?? d * 100_000) / 100_000;
  const hw = w / 2, hh = h / 2;
  let pts: Array<[number, number]>;
  switch (prst) {
    case "line": case "straightConnector1": return [[[0, 0], [w, h]], false];
    case "pie": case "blockArc": return sector(prst === "blockArc", w, h, adjRaw);
    case "ellipse": case "heart": case "star5": case "star4": case "star6": return undefined;
    case "roundRect": return [rounded(w, h, lim(ss * a("adj", 0.16667), ss / 2)), true];
    case "triangle": pts = [[lim(w * a("adj", 0.5), w), 0], [w, h], [0, h]]; break;
    case "rtTriangle": pts = [[0, 0], [0, h], [w, h]]; break;
    case "diamond": pts = [[hw, 0], [w, hh], [hw, h], [0, hh]]; break;
    case "rightArrow": case "leftArrow": case "upArrow": case "downArrow": {
      const vert = prst === "upArrow" || prst === "downArrow";
      pts = arrow(vert ? h : w, vert ? w : h, a("adj1", 0.5), a("adj2", 0.5), false).map(([u, v]) =>
        prst === "leftArrow" ? [w - u, v] : prst === "upArrow" ? [v, h - u] : prst === "downArrow" ? [v, u] : [u, v]);
      break;
    }
    case "leftRightArrow": case "upDownArrow": {
      const vert = prst === "upDownArrow";
      pts = arrow(vert ? h : w, vert ? w : h, a("adj1", 0.5), a("adj2", 0.5), true).map(([u, v]) => (vert ? [v, u] : [u, v]));
      break;
    }
    case "chevron": { const x1 = lim(ss * a("adj", 0.5), w); pts = [[0, 0], [w - x1, 0], [w, hh], [w - x1, h], [0, h], [x1, hh]]; break; }
    case "homePlate": { const x1 = w - lim(ss * a("adj", 0.5), w); pts = [[0, 0], [x1, 0], [w, hh], [x1, h], [0, h]]; break; }
    case "hexagon": { const x1 = lim(ss * a("adj", 0.25), w / 2); pts = [[0, hh], [x1, 0], [w - x1, 0], [w, hh], [w - x1, h], [x1, h]]; break; }
    case "pentagon":
      pts = [0, 1, 2, 3, 4].map((k) => { const t = ((-90 + 72 * k) * Math.PI) / 180; return [hw + hw * Math.cos(t), hh + hh * Math.sin(t)]; });
      break;
    case "trapezoid": { const x2 = lim(ss * a("adj", 0.25), w / 2); pts = [[0, h], [x2, 0], [w - x2, 0], [w, h]]; break; }
    case "parallelogram": { const x2 = lim(ss * a("adj", 0.25), w); pts = [[0, h], [x2, 0], [w, 0], [w - x2, h]]; break; }
    default: pts = [[0, 0], [w, 0], [w, h], [0, h]];
  }
  return [pts, true];
}

/** The SVG path data of a preset in its box at (x, y), and whether it is closed; curved kinds as wordcraft pdf/paint.rs draws them. */
export function presetPath(prst: string, x: number, y: number, w: number, h: number, adj: Record<string, number>): [string, boolean] {
  const f = (n: number) => String(Math.round(n * 100) / 100);
  const cx = x + w / 2, cy = y + h / 2;
  if (prst === "ellipse") {
    const rx = w / 2, ry = h / 2;
    return [`M${f(cx + rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 1 ${f(cx - rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 1 ${f(cx + rx)} ${f(cy)}Z`, true];
  }
  if (prst === "star5" || prst === "star4" || prst === "star6") {
    const d = Array.from({ length: 10 }, (_, i) => {
      const a = (Math.PI * i) / 5 - Math.PI / 2, rad = i % 2 === 0 ? 1 : 0.4;
      return `${i ? "L" : "M"}${f(cx + (Math.cos(a) * w * rad) / 2)} ${f(cy + (Math.sin(a) * h * rad) / 2)}`;
    }).join("");
    return [d + "Z", true];
  }
  if (prst === "heart") {
    const x1 = x + w, y1 = y + h;
    return [`M${f(cx)} ${f(y + h * 0.3)}C${f(cx)} ${f(y)} ${f(x)} ${f(y)} ${f(x)} ${f(y + h * 0.3)}C${f(x)} ${f(y + h * 0.6)} ${f(cx - w * 0.1)} ${f(y + h * 0.75)} ${f(cx)} ${f(y1)}C${f(cx + w * 0.1)} ${f(y + h * 0.75)} ${f(x1)} ${f(y + h * 0.6)} ${f(x1)} ${f(y + h * 0.3)}C${f(x1)} ${f(y)} ${f(cx)} ${f(y)} ${f(cx)} ${f(y + h * 0.3)}Z`, true];
  }
  const [pts, closed] = presetOutline(prst, w, h, adj) ?? [[], true];
  return [pts.map(([px0, py0], i) => `${i ? "L" : "M"}${f(x + px0)} ${f(y + py0)}`).join("") + (closed && pts.length ? "Z" : ""), closed];
}

// The text of a dsp:txBody: one item per wrapped line, stacked and centred in the text box.
function text(sp: Element, xf: Xf, b: number[], colors: Record<string, string>, filled: boolean, fit: number, out: SmartArtItem[]) {
  const tb = child(sp, "dsp:txBody");
  if (!tb) return;
  const body = child(tb, "a:bodyPr");
  const fs = int(child(body, "a:normAutofit")?.getAttribute("fontScale"));
  const scale = fs === undefined ? 1 : Math.max(0.1, Math.min(1, fs / 100_000));
  const inset = (n: string, d: number) => { const v = int(body?.getAttribute(n)); return v === undefined ? d : Math.max(0, Math.min(MAX_EMU, v)); };
  let box = xfrmBox(child(sp, "dsp:txXfrm"));
  if (!box) {
    const l = inset("lIns", 91_440), t = inset("tIns", 45_720), r = inset("rIns", 91_440), bo = inset("bIns", 45_720);
    box = [b[0]! + l, b[1]! + t, Math.max(0, b[2]! - l - r), Math.max(0, b[3]! - t - bo)];
  }
  const [bx0, by0, bx1, by1] = mapRect(xf, box);
  const x0 = px(bx0), y0 = px(by0), tw = px(bx1 - bx0), th = px(by1 - by0);
  const room = Math.max(0, MAX_ITEMS - out.length);
  const fallback = (filled ? colors.lt1 : colors.dk1) ?? (filled ? "FFFFFF" : "000000");
  const fontRef = drawingColor(child(child(sp, "dsp:style"), "a:fontRef"), colors);
  const lines: Array<{ text: string; size: number; color: string; bold: boolean; align: Align; font?: string }> = [];
  for (const p of children(tb, "a:p")) {
    if (lines.length >= room) break;
    const runs = children(p, "a:r");
    const raw = runs.map((r) => child(r, "a:t")?.textContent ?? "").join("");
    // eslint-disable-next-line no-control-regex
    const s = raw.replace(/[\u0000-\u001f\u007f-\u009f]/g, " ").trim();
    if (!s) continue;
    const rpr = child(runs.find((r) => child(r, "a:t")), "a:rPr");
    const sz = int(rpr?.getAttribute("sz"));
    const base = sz === undefined ? 18 : Math.max(1, Math.min(400, sz / 100));
    const size = base * PT_PX * scale * fit;
    const color = drawingColor(child(rpr, "a:solidFill"), colors) ?? fontRef ?? fallback;
    const algn = child(p, "a:pPr")?.getAttribute("algn");
    const align: Align = algn === "l" || algn === "just" || algn === "dist" ? "left" : algn === "r" ? "right" : "center";
    const tf = child(rpr, "a:latin")?.getAttribute("typeface");
    const font = tf && !tf.startsWith("+") ? tf : undefined;
    for (const line of wrap(s, size, tw, room - lines.length)) lines.push({ text: line, size, color, bold: on(rpr?.getAttribute("b")), align, font });
  }
  const total = lines.reduce((t, l) => t + l.size * LINE, 0);
  let y = y0 + (th - total) / 2;
  for (const l of lines) {
    const lh = l.size * LINE;
    out.push({ k: "text", rect: [x0, y, tw, lh], ...l });
    y += lh;
  }
}

/** Estimated width of s at size (wordcraft's chart estimate: CHAR_W em a character). */
export const textWidth = (s: string, size: number) => Array.from(s).length * size * CHAR_W;

/** Splits text into at most max lines of whole words that fit width; a wider word keeps its own line. */
export function wrap(text: string, size: number, width: number, max: number): string[] {
  const limit = width > 0 && size > 0 ? width : Infinity;
  const space = textWidth(" ", size);
  const lines: string[] = [];
  let cur = "", curW = 0;
  for (const word of text.split(/\s+/).filter(Boolean)) {
    const w = textWidth(word, size);
    if (cur && curW + space + w > limit) {
      if (lines.length + 1 >= max) break;
      lines.push(cur);
      cur = "";
      curW = 0;
    }
    if (cur) { cur += " "; curW += space; }
    cur += word;
    curW += w;
  }
  if (cur && lines.length < max) lines.push(cur);
  return lines;
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const num = (n: number) => String(Math.round(n * 100) / 100);

/** SVG markup for SmartArt items (no root element). */
export function smartArtItemsSvg(items: SmartArtItem[]): string {
  const out: string[] = [];
  for (const it of items) {
    const [x, y, w, h] = it.rect;
    if (it.k === "image") {
      out.push(`<image x="${num(x)}" y="${num(y)}" width="${num(w)}" height="${num(h)}" preserveAspectRatio="none" href="${esc(it.href)}"/>`);
    } else if (it.k === "geom") {
      const [d, closed] = presetPath(it.prst, x, y, w, h, it.adj);
      if (!d) continue;
      const fill = closed && it.fill ? `#${it.fill}` : "none";
      const stroke = it.stroke && it.strokeWidth > 0 ? ` stroke="#${it.stroke}" stroke-width="${num(it.strokeWidth)}" stroke-linejoin="round"` : "";
      // Flip in the box, then turn about its centre (clockwise, y down).
      const cx = x + w / 2, cy = y + h / 2;
      const tf: string[] = [];
      if (it.rot) tf.push(`rotate(${num(it.rot)} ${num(cx)} ${num(cy)})`);
      if (it.flipH || it.flipV) tf.push(`translate(${num(cx)} ${num(cy)}) scale(${it.flipH ? -1 : 1} ${it.flipV ? -1 : 1}) translate(${num(-cx)} ${num(-cy)})`);
      out.push(`<path d="${d}" fill="${fill}"${stroke}${tf.length ? ` transform="${tf.join(" ")}"` : ""}/>`);
    } else {
      const ax = it.align === "left" ? x : it.align === "right" ? x + w : x + w / 2;
      const anchor = it.align === "left" ? "start" : it.align === "right" ? "end" : "middle";
      const family = it.font ? `'${esc(it.font).replace(/'/g, "")}', Calibri, Arial, sans-serif` : "Calibri, Arial, sans-serif";
      out.push(`<text x="${num(ax)}" y="${num(y + h / 2)}" dominant-baseline="central" text-anchor="${anchor}" font-family="${family}" font-size="${num(it.size)}"${it.bold ? ` font-weight="bold"` : ""} fill="#${it.color}">${esc(it.text)}</text>`);
    }
  }
  return out.join("");
}

function parseXml(xml: string): Document | undefined {
  try {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    return doc.getElementsByTagName("parsererror").length ? undefined : doc;
  } catch {
    return undefined;
  }
}

/** A standalone SVG for a SmartArt drawing part (dsp:drawing), fitted to widthPx x heightPx; undefined when there is nothing to draw. `image` maps a drawing-part rId to an href. */
export function smartArtSvg(drawingXml: string, colors: Record<string, string>, widthPx: number, heightPx: number, image?: (rId: string) => string | undefined): string | undefined {
  const doc = parseXml(drawingXml);
  if (!doc) return undefined;
  const items = smartArtItems(doc.documentElement, colors, widthPx, heightPx, image);
  if (!items.length) return undefined;
  const w = Number.isFinite(widthPx) && widthPx > 0 ? widthPx : 1, h = Number.isFinite(heightPx) && heightPx > 0 ? heightPx : 1;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${num(w)}" height="${num(h)}" viewBox="0 0 ${num(w)} ${num(h)}">${smartArtItemsSvg(items)}</svg>`;
}

// A relationship target (relative to word/) as a zip path.
function partPath(target: string): string {
  const parts = (target.startsWith("/") ? target.slice(1) : "word/" + target).split("/");
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p && p !== ".") out.push(p);
  }
  return out.join("/");
}

const isDrawing = (doc: Document | undefined) => !!doc && doc.documentElement.localName === "drawing" && doc.documentElement.namespaceURI === "http://schemas.microsoft.com/office/drawing/2008/diagram";

/** The dsp:drawing XML of a SmartArt a:graphicData: the drawing its data part names (dsp:dataModelExt relId, a document relationship), else the drawing numbered like the data part. */
export function findSmartArtDrawing(graphicData: Element, rels: Map<string, string>, files: Record<string, Uint8Array>): string | undefined {
  const dm = child(graphicData, "dgm:relIds")?.getAttribute("r:dm");
  const dataTarget = dm ? rels.get(dm) : undefined;
  if (!dataTarget) return undefined;
  const dataPath = partPath(dataTarget);
  const read = (path: string): [string, Document | undefined] | undefined => {
    const bytes = files[path];
    if (!bytes) return undefined;
    try {
      const xml = strFromU8(bytes);
      return [xml, parseXml(xml)];
    } catch {
      return undefined;
    }
  };
  const data = read(dataPath);
  const ext = data?.[1]?.getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2008/diagram", "dataModelExt")[0];
  const extId = ext?.getAttribute("relId");
  const byId = extId ? rels.get(extId) : undefined;
  if (byId && !/^[a-z][a-z0-9+.-]*:/i.test(byId)) {
    const d = read(partPath(byId));
    if (d && isDrawing(d[1])) return d[0];
  }
  // No usable relationship id: the drawing whose number matches the data part's.
  const numOf = (t: string) => (t.split("/").pop() ?? t).replace(/\D/g, "");
  const n = numOf(dataPath);
  if (!n) return undefined;
  for (const target of rels.values()) {
    if (/^[a-z][a-z0-9+.-]*:/i.test(target) || numOf(target) !== n || !/drawing/i.test(target.split("/").pop() ?? "")) continue;
    const d = read(partPath(target));
    if (d && isDrawing(d[1])) return d[0];
  }
  return undefined;
}
