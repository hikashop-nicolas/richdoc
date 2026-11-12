/*! richdoc metafile player: a port of wordcraft-metafile (https://github.com/storytold/wordcraft).
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
// EMF / WMF pictures drawn in the browser: the records are played into paths, bitmaps and text
// runs, then written as an SVG. WMF text, SVG and PNG output are richdoc's; see metafile.LICENSE.md.
import { zlibSync } from "fflate";

// ---- limits (hostile input yields a shorter picture, never a hang) ----
const MAX_ITEMS = 200_000;
const MAX_OBJECTS = 4096;
const MAX_PIXELS = 16_000_000;
const MAX_STACK = 1024;
const MAX_PATH_OPS = 2_000_000;
const MAX_SIDE = 30_000;

type P = [number, number];
// Path ops: [0,x,y] move, [1,x,y] line, [2,x1,y1,x2,y2,x3,y3] cubic, [3] close.
type Op = [0, number, number] | [1, number, number] | [2, number, number, number, number, number, number] | [3];
type RGBA = [number, number, number, number];

interface Img { w: number; h: number; px: Uint8Array }
interface RawText { text: string; origin: P; size: number; cell: boolean; angle: number; dx: number[]; font: string; bold: boolean; italic: boolean; color: RGBA; align: number }
type Raw = { k: "path"; ops: Op[]; fill?: RGBA; stroke?: [RGBA, number]; evenOdd: boolean } | { k: "bmp"; rect: number[]; img: Img } | { k: "text"; t: RawText };
interface Frame { rect: number[]; size?: P }

// ---- bounds-checked little-endian reads ----
class Bytes {
  private v: DataView;
  constructor(readonly b: Uint8Array) {
    this.v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  }
  get length() { return this.b.length; }
  private ok(off: number, n: number) { return Number.isInteger(off) && off >= 0 && off + n <= this.b.length; }
  u8(o: number) { return this.ok(o, 1) ? this.b[o]! : undefined; }
  u16(o: number) { return this.ok(o, 2) ? this.v.getUint16(o, true) : undefined; }
  i16(o: number) { return this.ok(o, 2) ? this.v.getInt16(o, true) : undefined; }
  u32(o: number) { return this.ok(o, 4) ? this.v.getUint32(o, true) : undefined; }
  i32(o: number) { return this.ok(o, 4) ? this.v.getInt32(o, true) : undefined; }
  f32(o: number) { return this.ok(o, 4) ? this.v.getFloat32(o, true) : undefined; }
  slice(o: number, n: number): Bytes | undefined { return this.ok(o, n) ? new Bytes(this.b.subarray(o, o + n)) : undefined; }
  tail(o: number): Bytes | undefined { return o >= 0 && o <= this.b.length ? new Bytes(this.b.subarray(o)) : undefined; }
  points(o: number, n: number, wide: boolean): P[] | undefined {
    const sz = wide ? 8 : 4;
    if (!this.ok(o, n * sz)) return undefined;
    const out: P[] = new Array(n);
    for (let i = 0; i < n; i++) {
      const at = o + i * sz;
      out[i] = wide ? [this.v.getInt32(at, true), this.v.getInt32(at + 4, true)] : [this.v.getInt16(at, true), this.v.getInt16(at + 2, true)];
    }
    return out;
  }
}

// ---- shapes as cubic Béziers ----
const rectOps = (l: number, t: number, r: number, b: number): Op[] => [[0, l, t], [1, r, t], [1, r, b], [1, l, b], [3]];
function arcOps(cx: number, cy: number, rx: number, ry: number, a0: number, sweep: number): Op[] {
  const n = Math.min(8, Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2))));
  const step = sweep / n;
  const k = (4 / 3) * Math.tan(step / 4);
  const out: Op[] = [];
  for (let i = 0; i < n; i++) {
    const t0 = a0 + step * i, t1 = t0 + step;
    const p0: P = [cx + rx * Math.cos(t0), cy + ry * Math.sin(t0)];
    const p3: P = [cx + rx * Math.cos(t1), cy + ry * Math.sin(t1)];
    out.push([2, p0[0] - k * rx * Math.sin(t0), p0[1] + k * ry * Math.cos(t0), p3[0] + k * rx * Math.sin(t1), p3[1] - k * ry * Math.cos(t1), p3[0], p3[1]]);
  }
  return out;
}
const center = (l: number, t: number, r: number, b: number): [number, number, number, number] => [(l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2, Math.abs(b - t) / 2];
function roundRect(l0: number, t0: number, r0: number, b0: number, ew: number, eh: number): Op[] {
  const l = Math.min(l0, r0), r = Math.max(l0, r0), t = Math.min(t0, b0), b = Math.max(t0, b0);
  const rx = Math.min(Math.abs(ew / 2), (r - l) / 2), ry = Math.min(Math.abs(eh / 2), (b - t) / 2);
  const H = Math.PI / 2;
  return [[0, l + rx, t], [1, r - rx, t], ...arcOps(r - rx, t + ry, rx, ry, -H, H), [1, r, b - ry], ...arcOps(r - rx, b - ry, rx, ry, 0, H),
    [1, l + rx, b], ...arcOps(l + rx, b - ry, rx, ry, H, H), [1, l, t + ry], ...arcOps(l + rx, t + ry, rx, ry, Math.PI, H), [3]];
}
function ellipse(l: number, t: number, r: number, b: number): Op[] {
  const [cx, cy, rx, ry] = center(l, t, r, b);
  return [[0, cx + rx, cy], ...arcOps(cx, cy, rx, ry, 0, Math.PI * 2), [3]];
}
type ArcKind = "arc" | "chord" | "pie";
function arcShape(l: number, t: number, r: number, b: number, s: P, e: P, kind: ArcKind): Op[] {
  const [cx, cy, rx, ry] = center(l, t, r, b);
  const ang = (p: P) => Math.atan2((p[1] - cy) / Math.max(ry, Number.MIN_VALUE), (p[0] - cx) / Math.max(rx, Number.MIN_VALUE));
  const a0 = ang(s), a1 = ang(e);
  let d = (((a0 - a1) % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
  if (d < 1e-12) d = 2 * Math.PI;
  const start: P = [cx + rx * Math.cos(a0), cy + ry * Math.sin(a0)];
  const ops: Op[] = kind === "pie" ? [[0, cx, cy], [1, start[0], start[1]]] : [[0, start[0], start[1]]];
  ops.push(...arcOps(cx, cy, rx, ry, a0, -d));
  if (kind === "chord") ops.push([3]);
  if (kind === "pie") ops.push([1, cx, cy], [3]);
  return ops;
}
function subpaths(pts: P[], counts: number[], close: boolean): Op[] {
  const ops: Op[] = [];
  let at = 0;
  for (const c of counts) {
    const end = Math.min(pts.length, at + c);
    if (end > at) {
      ops.push([0, pts[at]![0], pts[at]![1]]);
      for (let i = at + 1; i < end; i++) ops.push([1, pts[i]![0], pts[i]![1]]);
      if (close) ops.push([3]);
    }
    at = end;
  }
  return ops;
}
function cubics(pts: P[]): Op[] {
  const ops: Op[] = [];
  for (let i = 0; i + 2 < pts.length; i += 3) ops.push([2, pts[i]![0], pts[i]![1], pts[i + 1]![0], pts[i + 1]![1], pts[i + 2]![0], pts[i + 2]![1]]);
  return ops;
}

// ---- drawing state ----
const MM_TEXT = 1, MM_ISOTROPIC = 7, MM_ANISOTROPIC = 8;
const STOCK = 0x80000000;
const PS_NULL = 5, PS_STYLE_MASK = 0x0f;
const BS_SOLID = 0, BS_NULL = 1;
const WINDING = 2, R2_NOP = 11;
const PATCOPY = 0x00f00021, BLACKNESS = 0x00000042, WHITENESS = 0x00ff0062;
const TA_UPDATECP = 1;
const DEFAULT_FONT_HEIGHT = 16;
const DEFAULT_PX_PER_MM = 96 / 25.4;
type Xform = [number, number, number, number, number, number];
const IDENTITY: Xform = [1, 0, 0, 1, 0, 0];
const mul = (a: Xform, b: Xform): Xform => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];

interface Pen { style: number; width: number; color: RGBA }
interface Brush { style: number; color: RGBA }
interface Font { height: number; escapement: number; weight: number; italic: boolean; face: string; charset: number }
type Obj = { k: "pen"; v: Pen } | { k: "brush"; v: Brush } | { k: "font"; v: Font } | { k: "other" };
const colorref = (c: number): RGBA => [c & 0xff, (c >>> 8) & 0xff, (c >>> 16) & 0xff, 255];
const metricMm = (mode: number): number | undefined => ({ 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 })[mode];

class Mapping {
  wo: P = [0, 0]; we: P = [1, 1]; vo: P = [0, 0]; ve: P = [1, 1];
  mode = MM_TEXT; world: Xform = IDENTITY; pxMm: P = [DEFAULT_PX_PER_MM, DEFAULT_PX_PER_MM];
  clone(): Mapping { return Object.assign(new Mapping(), this, { wo: [...this.wo], we: [...this.we], vo: [...this.vo], ve: [...this.ve], world: [...this.world], pxMm: [...this.pxMm] }); }
  ratios(): P {
    if (this.mode === MM_TEXT) return [1, 1];
    const mm = metricMm(this.mode);
    if (mm !== undefined) return [mm * this.pxMm[0], -mm * this.pxMm[1]];
    const ratio = (n: number, d: number) => (d === 0 ? 1 : n / d);
    const rx = ratio(this.ve[0], this.we[0]), ry = ratio(this.ve[1], this.we[1]);
    if (this.mode === MM_ISOTROPIC) {
      const s = Math.min(Math.abs(rx), Math.abs(ry));
      return [rx < 0 ? -s : s, ry < 0 ? -s : s];
    }
    return [rx, ry];
  }
  device(p: P): P {
    const w = this.world;
    const x = p[0] * w[0] + p[1] * w[2] + w[4], y = p[0] * w[1] + p[1] * w[3] + w[5];
    const [sx, sy] = this.ratios();
    return [(x - this.wo[0]) * sx + this.vo[0], (y - this.wo[1]) * sy + this.vo[1]];
  }
  scale(): number {
    const w = this.world;
    const [sx, sy] = this.ratios();
    return Math.sqrt(Math.abs(w[0] * w[3] - w[1] * w[2])) * (Math.abs(sx) + Math.abs(sy)) / 2;
  }
  usesExtents() { return this.mode === MM_ISOTROPIC || this.mode === MM_ANISOTROPIC; }
  windowRect(): number[] {
    const a = this.device(this.wo), b = this.device([this.wo[0] + this.we[0], this.wo[1] + this.we[1]]);
    return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.abs(b[0] - a[0]), Math.abs(b[1] - a[1])];
  }
}

interface Saved { map: Mapping; pen: Pen; brush: Brush; evenOdd: boolean; nop: boolean; font?: Font; textColor: RGBA; textAlign: number }

class Canvas {
  items: Raw[] = [];
  map = new Mapping();
  windowSet = false;
  viewportSet = false;
  full = false;
  private saved: Saved[] = [];
  private objects: (Obj | undefined)[] = [];
  private pen: Pen = { style: 0, width: 0, color: [0, 0, 0, 255] };
  private brush: Brush = { style: BS_SOLID, color: [255, 255, 255, 255] };
  private evenOdd = true;
  private nop = false;
  private font?: Font;
  private textColor: RGBA = [0, 0, 0, 255];
  private textAlign = 0;
  private pos: P = [0, 0];
  private inPath = false;
  private path: Op[] = [];

  private push(r: Raw) {
    if (this.items.length >= MAX_ITEMS) this.full = true;
    else this.items.push(r);
  }
  setWindowOrg(p: P) { this.map.wo = p; }
  setWindowExt(e: P) { this.map.we = e; this.windowSet = true; }
  setViewportOrg(p: P) { this.map.vo = p; }
  setViewportExt(e: P) { this.map.ve = e; this.viewportSet = true; }
  /** Set up the mapping a player uses to play a WMF into a target of `size`. */
  playInto(window: number[] | undefined, size: P) {
    if (window) {
      this.map.mode = MM_ANISOTROPIC;
      this.map.wo = [window[0]!, window[1]!];
      this.map.we = [window[2]!, window[3]!];
      this.map.vo = [0, 0];
    }
    this.map.ve = size;
  }
  setMapMode(m: number) { this.map.mode = m; }
  setDeviceMm(p: P) { if (p[0] > 0 && p[1] > 0 && isFinite(p[0]) && isFinite(p[1])) this.map.pxMm = p; }
  setPolyfill(m: number) { this.evenOdd = m !== WINDING; }
  setRop2(m: number) { this.nop = m === R2_NOP; }
  modifyWorld(x: Xform, mode: number) {
    if (mode === 1) this.map.world = IDENTITY;
    else if (mode === 2) this.map.world = mul(x, this.map.world);
    else if (mode === 3) this.map.world = mul(this.map.world, x);
    else if (mode === 4) this.map.world = x;
  }
  saveDc() {
    if (this.saved.length >= MAX_STACK) { this.full = true; return; }
    this.saved.push({ map: this.map.clone(), pen: this.pen, brush: this.brush, evenOdd: this.evenOdd, nop: this.nop, font: this.font, textColor: this.textColor, textAlign: this.textAlign });
  }
  restoreDc(n: number) {
    const idx = n < 0 ? this.saved.length + n : n - 1;
    const s = idx >= 0 ? this.saved[idx] : undefined;
    if (!s) return;
    this.saved.length = idx;
    Object.assign(this, { map: s.map, pen: s.pen, brush: s.brush, evenOdd: s.evenOdd, nop: s.nop, font: s.font, textColor: s.textColor, textAlign: s.textAlign });
  }
  /** Store an object in `slot` (EMF index) or the lowest free slot (WMF handle table). */
  addObject(slot: number | undefined, o: Obj) {
    if (slot !== undefined) {
      if (slot >= 0 && slot < MAX_OBJECTS) this.objects[slot] = o;
      return;
    }
    for (let i = 0; i < this.objects.length; i++) if (!this.objects[i]) { this.objects[i] = o; return; }
    if (this.objects.length < MAX_OBJECTS) this.objects.push(o);
  }
  deleteObject(i: number) { if (i < this.objects.length) this.objects[i] = undefined; }
  select(i: number) {
    if ((i & STOCK) !== 0) return this.selectStock((i & ~STOCK) >>> 0);
    const o = this.objects[i];
    if (o?.k === "pen") this.pen = o.v;
    else if (o?.k === "brush") this.brush = o.v;
    else if (o?.k === "font") this.font = o.v;
  }
  private selectStock(i: number) {
    const gray = (v: number): Brush => ({ style: BS_SOLID, color: [v, v, v, 255] });
    if (i === 0 || i === 18) this.brush = gray(255);
    else if (i === 1) this.brush = gray(192);
    else if (i === 2) this.brush = gray(128);
    else if (i === 3) this.brush = gray(64);
    else if (i === 4) this.brush = gray(0);
    else if (i === 5) this.brush = { style: BS_NULL, color: [0, 0, 0, 255] };
    else if (i === 6) this.pen = { style: 0, width: 0, color: [255, 255, 255, 255] };
    else if (i === 7 || i === 19) this.pen = { style: 0, width: 0, color: [0, 0, 0, 255] };
    else if (i === 8) this.pen = { style: PS_NULL, width: 0, color: [0, 0, 0, 255] };
    else if (i >= 10 && i <= 17) this.font = undefined;
  }
  private mapOp(o: Op): Op {
    if (o[0] === 3) return o;
    if (o[0] === 2) {
      const a = this.map.device([o[1], o[2]]), b = this.map.device([o[3], o[4]]), c = this.map.device([o[5], o[6]]);
      return [2, a[0], a[1], b[0], b[1], c[0], c[1]];
    }
    const p = this.map.device([o[1], o[2]]);
    return [o[0], p[0], p[1]];
  }
  private fillColor(): RGBA | undefined { return this.brush.style === BS_SOLID ? this.brush.color : undefined; }
  private strokePen(): [RGBA, number] | undefined {
    if ((this.pen.style & PS_STYLE_MASK) === PS_NULL) return undefined;
    return [this.pen.color, this.pen.width === 0 ? 0 : Math.abs(this.pen.width) * this.map.scale()];
  }
  private pushPath(ops: Op[]) {
    for (const o of ops) { if (this.path.length >= MAX_PATH_OPS) return; this.path.push(o); }
  }
  /** Draw logical-space ops with the pen / brush, or add them to the open path bracket. */
  shape(ops: Op[], fill: boolean, stroke: boolean) {
    const dev = ops.map((o) => this.mapOp(o));
    if (this.inPath) return this.pushPath(dev);
    if (this.nop) return;
    const f = fill ? this.fillColor() : undefined, s = stroke ? this.strokePen() : undefined;
    if ((!f && !s) || dev.length < 2) return;
    this.push({ k: "path", ops: dev, fill: f, stroke: s, evenOdd: this.evenOdd });
  }
  moveTo(p: P) { if (this.inPath) this.pushPath([[0, ...this.map.device(p)]]); this.pos = p; }
  lineTo(p: P) { this.lineChain([p], false, true); }
  lineChain(pts: P[], fill: boolean, stroke: boolean) {
    const ops: Op[] = this.inPath && this.path.length ? [] : [[0, this.pos[0], this.pos[1]]];
    for (const p of pts) ops.push([1, p[0], p[1]]);
    if (pts.length) this.pos = pts[pts.length - 1]!;
    this.shape(ops, fill, stroke);
  }
  bezier(from: P, pts: P[]) { this.shape([[0, from[0], from[1]], ...cubics(pts)], false, true); }
  bezierTo(pts: P[]) {
    const ops: Op[] = this.inPath && this.path.length ? [] : [[0, this.pos[0], this.pos[1]]];
    ops.push(...cubics(pts));
    if (pts.length) this.pos = pts[pts.length - 1]!;
    this.shape(ops, false, true);
  }
  poly(pts: P[], close: boolean, fill: boolean, stroke: boolean) {
    if (!pts.length) return;
    const ops: Op[] = [[0, pts[0]![0], pts[0]![1]]];
    for (let i = 1; i < pts.length; i++) ops.push([1, pts[i]![0], pts[i]![1]]);
    if (close) ops.push([3]);
    this.shape(ops, fill, stroke);
  }
  beginPath() { this.inPath = true; this.path = []; }
  endPath() { this.inPath = false; }
  closeFigure() { if (this.inPath) this.pushPath([[3]]); }
  paintPath(fill: boolean, stroke: boolean) {
    const ops = this.path;
    this.path = [];
    if (this.nop) return;
    const f = fill ? this.fillColor() : undefined, s = stroke ? this.strokePen() : undefined;
    if (ops.length < 2 || (!f && !s)) return;
    this.push({ k: "path", ops, fill: f, stroke: s, evenOdd: this.evenOdd });
  }
  fillRectWith(r: number[], color: RGBA) {
    this.push({ k: "path", ops: rectOps(r[0]!, r[1]!, r[2]!, r[3]!).map((o) => this.mapOp(o)), fill: color, evenOdd: true });
  }
  static sourceless(rop: number) { return rop === PATCOPY || rop === BLACKNESS || rop === WHITENESS; }
  /** Fill [x, y, w, h] for a raster operation that has no source bitmap. */
  fillRop([x, y, w, h]: number[], rop: number) {
    const r = [x!, y!, x! + w!, y! + h!];
    if (rop === PATCOPY) { const c = this.fillColor(); if (c) this.fillRectWith(r, c); }
    else if (rop === BLACKNESS) this.fillRectWith(r, [0, 0, 0, 255]);
    else if (rop === WHITENESS) this.fillRectWith(r, [255, 255, 255, 255]);
  }
  /** The selected font's charset, to decode WMF ANSI text. */
  charset() { return this.font?.charset ?? 0; }
  setTextColor(c: RGBA) { this.textColor = c; }
  setTextAlign(a: number) { this.textAlign = a; }
  text(at0: P, text: string, dx: number[]) {
    if (!text || this.inPath) return;
    const font = this.font ?? { height: 0, escapement: 0, weight: 400, italic: false, face: "", charset: 0 };
    const at: P = this.textAlign & TA_UPDATECP ? this.pos : at0;
    if (this.textAlign & TA_UPDATECP) this.pos = [at[0] + dx.reduce((a, b) => a + b, 0), at[1]];
    const height = font.height === 0 ? DEFAULT_FONT_HEIGHT : font.height;
    const origin = this.map.device(at);
    const len = (v: P) => { const d = this.map.device([at[0] + v[0], at[1] + v[1]]); return Math.hypot(d[0] - origin[0], d[1] - origin[1]); };
    const base = this.map.device([at[0] + 1, at[1]]);
    const angle = Math.atan2(base[1] - origin[1], base[0] - origin[0]) - ((font.escapement / 10) * Math.PI) / 180;
    const size = len([0, Math.abs(height)]);
    const t: RawText = {
      text, origin, size, cell: height > 0, angle: isFinite(angle) ? angle : 0,
      dx: dx.map((d) => Math.sign(d) * len([d, 0])),
      font: font.face.trim() || "Arial", bold: font.weight >= 600, italic: font.italic, color: this.textColor, align: this.textAlign,
    };
    if (isFinite(origin[0]) && isFinite(origin[1]) && isFinite(size) && size > 0) this.push({ k: "text", t });
  }
  /** A bitmap stretched into the logical [x, y, w, h]; a negative extent mirrors it. */
  bitmap(rect: number[], img: Img) {
    const [x, y, w, h] = rect as [number, number, number, number];
    const a = this.map.device([x, y]), b = this.map.device([x + w, y + h]);
    flip(img, a[0] > b[0], a[1] > b[1]);
    this.push({ k: "bmp", rect: [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.abs(b[0] - a[0]), Math.abs(b[1] - a[1])], img });
  }
}

// ---- device-independent bitmaps ----
const BI_RGB = 0, BI_BITFIELDS = 3;
const ROP_SRCAND = 0x008800c6, ROP_SRCPAINT = 0x00ee0086, ROP_SRCINVERT = 0x00660046;
type Blend = { rop: number } | { constant: number; perPixel: boolean };
interface DibHeader { w: number; h: number; topDown: boolean; bpp: number; masks: [number, number, number]; palOff: number; palN: number }
function dibHeader(info: Bytes): DibHeader | undefined {
  const size = info.u32(0), w = info.i32(4), h = info.i32(8), bpp = info.u16(14), comp = info.u32(16), clrUsed = info.u32(32);
  if (size === undefined || w === undefined || h === undefined || bpp === undefined || comp === undefined || clrUsed === undefined) return undefined;
  if (size < 40 || w <= 0 || h === 0 || ![1, 4, 8, 16, 24, 32].includes(bpp) || (comp !== BI_RGB && comp !== BI_BITFIELDS)) return undefined;
  let masks: [number, number, number];
  if (comp === BI_BITFIELDS) {
    const r = info.u32(40), g = info.u32(44), b = info.u32(48);
    if (r === undefined || g === undefined || b === undefined) return undefined;
    masks = [r, g, b];
  } else masks = bpp === 16 ? [0x7c00, 0x03e0, 0x001f] : [0xff0000, 0xff00, 0xff];
  const palOff = size + (comp === BI_BITFIELDS && size === 40 ? 12 : 0);
  const palN = bpp <= 8 ? (clrUsed === 0 ? 1 << bpp : Math.min(clrUsed, 1 << bpp)) : Math.min(clrUsed, 256);
  return { w, h: Math.abs(h), topDown: h < 0, bpp, masks, palOff, palN };
}
/** Length of a DIB's info part (header, masks, palette), so its bits follow. */
function dibInfoLen(dib: Bytes): number | undefined {
  const h = dibHeader(dib);
  return h && h.palOff + h.palN * 4;
}
const channel = (mask: number) => {
  const m = mask >>> 0;
  if (!m) return () => 0;
  let shift = 0;
  while (!((m >>> shift) & 1)) shift++;
  const max = m >>> shift;
  return (v: number) => Math.round((((v & m) >>> shift) * 255) / max);
};
/** Decode the part of a DIB that `src` selects, with the part of `dst` it fills. */
function decodeClip(info: Bytes, bits: Bytes, keepAlpha: boolean, src: number[], dst: number[]): [Img, number[]] | undefined {
  const hd = dibHeader(info);
  if (!hd || hd.w * hd.h === 0 || hd.w * hd.h > MAX_PIXELS) return undefined;
  const stride = Math.floor((hd.w * hd.bpp + 31) / 32) * 4;
  const data = bits.slice(0, stride * hd.h);
  const pal = info.slice(hd.palOff, hd.palN * 4);
  if (!data || !pal) return undefined;
  const clip = clipRect(hd.w, hd.h, !hd.topDown, src, dst);
  if (!clip) return undefined;
  const { x0, x1, r0, r1, out } = clip;
  if (x1 - x0 > MAX_SIDE || r1 - r0 > MAX_SIDE) return undefined;
  const ow = x1 - x0, oh = r1 - r0;
  const px = new Uint8Array(ow * oh * 4);
  const d = data.b, p = pal.b;
  const palette = (i: number): [number, number, number] => (i * 4 + 4 <= p.length ? [p[i * 4 + 2]!, p[i * 4 + 1]!, p[i * 4]!] : [0, 0, 0]);
  const [cr, cg, cb] = hd.masks.map(channel) as [(v: number) => number, (v: number) => number, (v: number) => number];
  const defMasks32 = hd.masks[0] === 0xff0000 && hd.masks[1] === 0xff00 && hd.masks[2] === 0xff;
  for (let y = r0; y < r1; y++) {
    const row = (hd.topDown ? y : hd.h - 1 - y) * stride;
    for (let x = x0; x < x1; x++) {
      let c: number[];
      if (hd.bpp <= 8) {
        const idx = hd.bpp === 8 ? d[row + x]! : hd.bpp === 4 ? (x % 2 ? d[row + (x >> 1)]! & 15 : d[row + (x >> 1)]! >> 4) : (d[row + (x >> 3)]! >> (7 - (x % 8))) & 1;
        c = [...palette(idx), 255];
      } else if (hd.bpp === 24) {
        const o = row + x * 3;
        c = [d[o + 2]!, d[o + 1]!, d[o]!, 255];
      } else if (hd.bpp === 16) {
        const v = d[row + x * 2]! | (d[row + x * 2 + 1]! << 8);
        c = [cr(v), cg(v), cb(v), 255];
      } else {
        const o = row + x * 4;
        if (defMasks32) c = [d[o + 2]!, d[o + 1]!, d[o]!, keepAlpha ? d[o + 3]! : 255];
        else {
          const v = (d[o]! | (d[o + 1]! << 8) | (d[o + 2]! << 16) | (d[o + 3]! << 24)) >>> 0;
          c = [cr(v), cg(v), cb(v), keepAlpha ? v >>> 24 : 255];
        }
      }
      px.set(c, ((y - r0) * ow + (x - x0)) * 4);
    }
  }
  return [{ w: ow, h: oh, px }, out];
}
/** Clip a source rect (DIB pixel coords; a non-positive size = the whole bitmap) to the bitmap,
    shrinking and shifting the destination by the same ratio. */
function clipRect(w: number, h: number, bottomUp: boolean, src: number[], dst: number[]) {
  const [sx, sy, sw, sh] = src[2]! > 0 && src[3]! > 0 ? src : [0, 0, w, h];
  const top = bottomUp ? h - sy! - sh! : sy!;
  const x0 = Math.max(0, sx!), x1 = Math.min(w, sx! + sw!), r0 = Math.max(0, top), r1 = Math.min(h, top + sh!);
  if (x1 <= x0 || r1 <= r0) return undefined;
  const kx = dst[2]! / sw!, ky = dst[3]! / sh!;
  return { x0, x1, r0, r1, out: [dst[0]! + (x0 - sx!) * kx, dst[1]! + (r0 - top) * ky, (x1 - x0) * kx, (r1 - r0) * ky] };
}
/** Raster operations drawn as they look over white paper; alpha blends made straight. */
function prepare(img: Img, blend: Blend) {
  const px = img.px;
  if ("rop" in blend) {
    const clearOf = blend.rop === ROP_SRCAND ? 255 : blend.rop === ROP_SRCPAINT || blend.rop === ROP_SRCINVERT ? 0 : -1;
    if (clearOf < 0) return;
    for (let i = 0; i < px.length; i += 4) if (px[i] === clearOf && px[i + 1] === clearOf && px[i + 2] === clearOf) px[i + 3] = 0;
    return;
  }
  for (let i = 0; i < px.length; i += 4) {
    const a = px[i + 3]!;
    if (blend.perPixel && a && a !== 255) for (let c = 0; c < 3; c++) px[i + c] = Math.min(255, Math.round((px[i + c]! * 255) / a));
    px[i + 3] = Math.floor((a * blend.constant) / 255);
  }
}
function flip(img: Img, horizontal: boolean, vertical: boolean) {
  const { w, h, px } = img;
  const row = w * 4;
  if (horizontal) for (let y = 0; y < h; y++) for (let x = 0; x < w >> 1; x++) for (let c = 0; c < 4; c++) {
    const a = y * row + x * 4 + c, b = y * row + (w - 1 - x) * 4 + c;
    const t = px[a]!; px[a] = px[b]!; px[b] = t;
  }
  if (vertical) for (let y = 0; y < h >> 1; y++) {
    const a = px.slice(y * row, y * row + row);
    px.copyWithin(y * row, (h - 1 - y) * row, (h - y) * row);
    px.set(a, (h - 1 - y) * row);
  }
}

// ---- EMF (MS-EMF) ----
const EMF_SIGNATURE = 0x464d4520;
const HMM_TO_PT = 72 / 2540;
const SRCCOPY = 0x00cc0020;
const ETO_GLYPH_INDEX = 0x10, ETO_PDY = 0x2000;
const isEmf = (b: Bytes) => b.u32(0) === 1 && b.u32(40) === EMF_SIGNATURE;
const xy32 = (r: Bytes, o: number): P | undefined => { const x = r.i32(o), y = r.i32(o + 4); return x === undefined || y === undefined ? undefined : [x, y]; };
const rect32 = (r: Bytes, o: number): number[] | undefined => { const v = [r.i32(o), r.i32(o + 4), r.i32(o + 8), r.i32(o + 12)]; return v.every((n) => n !== undefined) ? (v as number[]) : undefined; };
const i32s = (r: Bytes, o: number) => r.i32(o) ?? 0;
function devicePxPerMm(b: Bytes): P | undefined {
  const dw = b.i32(72), dh = b.i32(76), mw = b.i32(80), mh = b.i32(84);
  return dw! > 0 && dh! > 0 && mw! > 0 && mh! > 0 ? [dw! / mw!, dh! / mh!] : undefined;
}
function emfFrame(b: Bytes): Frame | undefined {
  const fr = rect32(b, 24), pm = devicePxPerMm(b);
  if (!fr || !pm || fr[2]! <= fr[0]! || fr[3]! <= fr[1]!) return undefined;
  const sx = pm[0] / 100, sy = pm[1] / 100;
  return { rect: [fr[0]! * sx, fr[1]! * sy, (fr[2]! - fr[0]!) * sx, (fr[3]! - fr[1]!) * sy], size: [(fr[2]! - fr[0]!) * HMM_TO_PT, (fr[3]! - fr[1]!) * HMM_TO_PT] };
}
function parseEmf(input: Bytes): { cv: Canvas; frame?: Frame } | undefined {
  const hsize = input.u32(4), bounds = rect32(input, 8);
  if (hsize === undefined || !bounds) return undefined;
  const [l, t, r, b] = bounds as [number, number, number, number];
  const frame = emfFrame(input) ?? (r > l && b > t ? { rect: [l, t, r - l, b - t] } : undefined);
  const cv = new Canvas();
  const pm = devicePxPerMm(input);
  if (pm) cv.setDeviceMm(pm);
  let parsed = 0;
  for (let pos = Math.max(hsize, 88); !cv.full; ) {
    const typ = input.u32(pos), size = input.u32(pos + 4);
    if (typ === undefined || size === undefined || size < 8 || size % 4) break;
    const rec = input.slice(pos, size);
    if (!rec || typ === 14) break;
    emfRecord(cv, typ, rec);
    parsed++;
    pos += size;
  }
  return parsed ? { cv, frame } : undefined;
}
function emfRecord(cv: Canvas, typ: number, r: Bytes) {
  const pt = (o: number) => xy32(r, o);
  const counted = (wide: boolean) => { const n = r.u32(24); return n !== undefined && n <= MAX_PATH_OPS ? r.tail(28)?.points(0, n, wide) : undefined; };
  const groups = (wide: boolean): [number[], P[]] | undefined => {
    const n = r.u32(24);
    if (n === undefined || n > MAX_PATH_OPS || !r.slice(32, n * 4)) return undefined;
    const counts: number[] = [];
    for (let i = 0; i < n; i++) counts.push(r.u32(32 + i * 4)!);
    const total = counts.reduce((a, c) => a + c, 0);
    if (total > MAX_PATH_OPS) return undefined;
    const pts = r.tail(32 + n * 4)?.points(0, total, wide);
    return pts && [counts, pts];
  };
  const xform = (o: number): Xform | undefined => { const x = [0, 1, 2, 3, 4, 5].map((i) => r.f32(o + i * 4)); return x.every((v) => v !== undefined) ? (x as Xform) : undefined; };
  switch (typ) {
    case 9: { const p = pt(8); if (p) cv.setWindowExt(p); break; }
    case 10: { const p = pt(8); if (p) cv.setWindowOrg(p); break; }
    case 11: { const p = pt(8); if (p) cv.setViewportExt(p); break; }
    case 12: { const p = pt(8); if (p) cv.setViewportOrg(p); break; }
    case 17: { const m = r.u32(8); if (m !== undefined) cv.setMapMode(m); break; }
    case 19: { const m = r.u32(8); if (m !== undefined) cv.setPolyfill(m); break; }
    case 20: { const m = r.u32(8); if (m !== undefined) cv.setRop2(m); break; }
    case 33: cv.saveDc(); break;
    case 34: { const n = r.i32(8); if (n !== undefined) cv.restoreDc(n); break; }
    case 35: { const x = xform(8); if (x) cv.modifyWorld(x, 4); break; }
    case 36: { const x = xform(8), m = r.u32(32); if (x && m !== undefined) cv.modifyWorld(x, m); break; }
    case 37: { const i = r.u32(8); if (i !== undefined) cv.select(i); break; }
    case 38: { // CREATEPEN: handle, LogPen
      const slot = r.u32(8), style = r.u32(12), width = r.i32(16), color = r.u32(24);
      if (slot !== undefined && style !== undefined && width !== undefined && color !== undefined) cv.addObject(slot, { k: "pen", v: { style, width, color: colorref(color) } });
      break;
    }
    case 95: { // EXTCREATEPEN: cosmetic pens are hairlines; a null brush draws no line
      const slot = r.u32(8), style = r.u32(28), width = r.u32(32), brush = r.u32(36), color = r.u32(40);
      if (slot === undefined || style === undefined || width === undefined || brush === undefined || color === undefined) break;
      cv.addObject(slot, { k: "pen", v: { style: brush === BS_NULL ? PS_NULL : style & PS_STYLE_MASK, width: (style & 0x000f0000) === 0 ? 0 : width, color: colorref(color) } });
      break;
    }
    case 39: { const slot = r.u32(8), style = r.u32(12), color = r.u32(16); if (slot !== undefined && style !== undefined && color !== undefined) cv.addObject(slot, { k: "brush", v: { style, color: colorref(color) } }); break; }
    case 82: { // EXTCREATEFONTINDIRECTW
      const slot = r.u32(8), height = r.i32(12), esc = r.i32(20), weight = r.i32(28), italic = r.u8(32);
      if (slot === undefined || height === undefined || esc === undefined || weight === undefined || italic === undefined) break;
      let face = "";
      for (let i = 0; i < 32; i++) { const u = r.u16(40 + i * 2) ?? 0; if (!u) break; face += String.fromCharCode(u); }
      cv.addObject(slot, { k: "font", v: { height, escapement: esc, weight, italic: italic !== 0, face, charset: r.u8(35) ?? 0 } });
      break;
    }
    case 22: { const a = r.u32(8); if (a !== undefined) cv.setTextAlign(a); break; }
    case 24: { const c = r.u32(8); if (c !== undefined) cv.setTextColor(colorref(c)); break; }
    case 84: emfText(cv, r); break;
    case 40: { const i = r.u32(8); if (i !== undefined) cv.deleteObject(i); break; }
    case 27: { const p = pt(8); if (p) cv.moveTo(p); break; }
    case 54: { const p = pt(8); if (p) cv.lineTo(p); break; }
    case 42: { const b = rect32(r, 8); if (b) cv.shape(ellipse(b[0]!, b[1]!, b[2]!, b[3]!), true, true); break; }
    case 43: { const b = rect32(r, 8); if (b) cv.shape(rectOps(b[0]!, b[1]!, b[2]!, b[3]!), true, true); break; }
    case 44: { const b = rect32(r, 8), cw = r.i32(24), ch = r.i32(28); if (b && cw !== undefined && ch !== undefined) cv.shape(roundRect(b[0]!, b[1]!, b[2]!, b[3]!, cw, ch), true, true); break; }
    case 45: case 46: case 47: {
      const b = rect32(r, 8), s = pt(24), e = pt(32);
      if (!b || !s || !e) break;
      const kind: ArcKind = typ === 45 ? "arc" : typ === 46 ? "chord" : "pie";
      cv.shape(arcShape(b[0]!, b[1]!, b[2]!, b[3]!, s, e, kind), kind !== "arc", true);
      break;
    }
    case 4: case 87: { const p = counted(typ === 4); if (p) cv.poly(p, false, false, true); break; }
    case 3: case 86: { const p = counted(typ === 3); if (p) cv.poly(p, true, true, true); break; }
    case 6: case 89: { const p = counted(typ === 6); if (p) cv.lineChain(p, false, true); break; }
    case 2: case 85: { const p = counted(typ === 2); if (p?.length) cv.bezier(p[0]!, p.slice(1)); break; }
    case 5: case 88: { const p = counted(typ === 5); if (p) cv.bezierTo(p); break; }
    case 7: case 90: { const g = groups(typ === 7); if (g) cv.shape(subpaths(g[1], g[0], false), false, true); break; }
    case 8: case 91: { const g = groups(typ === 8); if (g) cv.shape(subpaths(g[1], g[0], true), true, true); break; }
    case 59: cv.beginPath(); break;
    case 60: cv.endPath(); break;
    case 61: cv.closeFigure(); break;
    case 62: cv.paintPath(true, false); break;
    case 63: cv.paintPath(true, true); break;
    case 64: cv.paintPath(false, true); break;
    case 76: case 77: { // BITBLT / STRETCHBLT
      const dst = [i32s(r, 24), i32s(r, 28), i32s(r, 32), i32s(r, 36)], rop = r.u32(40);
      if (rop === undefined) break;
      if (Canvas.sourceless(rop)) { cv.fillRop(dst, rop); break; }
      const src = [i32s(r, 44), i32s(r, 48), typ === 77 ? i32s(r, 100) : dst[2]!, typ === 77 ? i32s(r, 104) : dst[3]!];
      emfBlit(cv, r, dst, src, 84, 92, { rop });
      break;
    }
    case 80: { // SETDIBITSTODEVICE
      const cx = i32s(r, 40), cy = i32s(r, 44), start = r.u32(68);
      if (cx > 0 && cy > 0 && start !== undefined) emfBlit(cv, r, [i32s(r, 24), i32s(r, 28), cx, cy], [i32s(r, 32), i32s(r, 36) - start, cx, cy], 48, 56, { rop: SRCCOPY });
      break;
    }
    case 81: { // STRETCHDIBITS
      const rop = r.u32(68);
      if (rop !== undefined) emfBlit(cv, r, [i32s(r, 24), i32s(r, 28), i32s(r, 72), i32s(r, 76)], [i32s(r, 32), i32s(r, 36), i32s(r, 40), i32s(r, 44)], 48, 56, { rop });
      break;
    }
    case 114: { // ALPHABLEND
      const constant = r.u8(42), fmt = r.u8(43);
      if (constant !== undefined && fmt !== undefined) emfBlit(cv, r, [i32s(r, 24), i32s(r, 28), i32s(r, 32), i32s(r, 36)], [i32s(r, 44), i32s(r, 48), i32s(r, 100), i32s(r, 104)], 84, 92, { constant, perPixel: (fmt & 1) !== 0 });
      break;
    }
  }
}
function emfBlit(cv: Canvas, r: Bytes, dst: number[], src: number[], atBmi: number, atBits: number, blend: Blend) {
  const bo = r.u32(atBmi), bl = r.u32(atBmi + 4), po = r.u32(atBits), pl = r.u32(atBits + 4);
  if (bo === undefined || bl === undefined || po === undefined || !pl) return;
  const info = r.slice(bo, bl), bits = r.slice(po, pl);
  if (!info || !bits) return;
  const d = decodeClip(info, bits, "perPixel" in blend && blend.perPixel, src, dst);
  if (!d) return;
  prepare(d[0], blend);
  cv.bitmap(d[1], d[0]);
}
/** EXTTEXTOUTW: the reference point at 36, count at 44, string offset 48, options 52, dx offset 72. */
function emfText(cv: Canvas, r: Bytes) {
  const at = xy32(r, 36), n = r.u32(44), off = r.u32(48), opts = r.u32(52), offDx = r.u32(72);
  if (!at || !n || off === undefined || opts === undefined || offDx === undefined || opts & ETO_GLYPH_INDEX) return;
  const s = r.slice(off, n * 2);
  if (!s) return;
  const units: number[] = [];
  for (let i = 0; i < n; i++) units.push(s.u16(i * 2)!);
  const step = opts & ETO_PDY ? 8 : 4;
  const raw: number[] = [];
  const dxb = offDx ? r.slice(offDx, n * step) : undefined;
  if (dxb) for (let i = 0; i < n; i++) raw.push(dxb.i32(i * step)!);
  // A character takes the advances of all its UTF-16 units.
  let text = "";
  const dx: number[] = [];
  for (let i = 0; i < units.length; ) {
    const hi = units[i]!;
    const w = hi >= 0xd800 && hi < 0xdc00 && (units[i + 1] ?? 0) >= 0xdc00 && (units[i + 1] ?? 0) < 0xe000 ? 2 : 1;
    text += String.fromCharCode(...units.slice(i, i + w));
    if (raw.length) dx.push(raw.slice(i, i + w).reduce((a, b) => a + b, 0));
    i += w;
  }
  const trimmed = text.replace(/\0+$/, "");
  cv.text(at, trimmed, dx.slice(0, [...trimmed].length));
}

// ---- WMF (MS-WMF) ----
const PLACEABLE_KEY = 0x9ac6cdd7;
const isWmf = (b: Bytes) => b.u32(0) === PLACEABLE_KEY || ((b.u16(0) === 1 || b.u16(0) === 2) && b.u16(2) === 9 && (b.u16(4) === 0x100 || b.u16(4) === 0x300));
const xy16 = (r: Bytes, o: number): P | undefined => { const y = r.i16(o), x = r.i16(o + 2); return x === undefined || y === undefined ? undefined : [x, y]; };
const rect16 = (r: Bytes, o: number): number[] | undefined => { const b = r.i16(o), rr = r.i16(o + 2), t = r.i16(o + 4), l = r.i16(o + 6); return [l, t, rr, b].every((v) => v !== undefined) ? [l!, t!, rr!, b!] : undefined; };
const i16s = (r: Bytes, o: number) => r.i16(o) ?? 0;
const wmfHandle = (h: number) => (h & 0x8000 ? ((h & 0x7fff) | STOCK) >>> 0 : h);
// A WMF font's charset -> the TextDecoder label of its ANSI code page.
const CHARSET_ENC: Record<number, string> = { 128: "shift_jis", 129: "euc-kr", 134: "gbk", 136: "big5", 161: "windows-1253", 162: "windows-1254", 163: "windows-1258", 177: "windows-1255", 178: "windows-1256", 186: "windows-1257", 204: "windows-1251", 222: "windows-874", 238: "windows-1250" };
function ansi(bytes: Uint8Array, charset: number): string {
  try { return new TextDecoder(CHARSET_ENC[charset] ?? "windows-1252").decode(bytes); } catch { return String.fromCharCode(...bytes); }
}
function parseWmf(input: Bytes): { cv: Canvas; frame?: Frame } | undefined {
  const placeable = input.u32(0) === PLACEABLE_KEY;
  let hdr = 0;
  let bounds: { c: number[]; size: P } | undefined;
  if (placeable) {
    const c = [6, 8, 10, 12].map((o) => input.i16(o));
    if (c.some((v) => v === undefined)) return undefined;
    const [l, t, r, b] = c as number[];
    const inch = input.u16(14) || 1440;
    if (r! > l! && b! > t!) bounds = { c: [l!, t!, r!, b!], size: [((r! - l!) * 72) / inch, ((b! - t!) * 72) / inch] };
    hdr = 22;
  }
  const hsize = input.u16(hdr + 2);
  if (hsize === undefined || hsize < 9) return undefined;
  const cv = new Canvas();
  if (bounds) cv.playInto([bounds.c[0]!, bounds.c[1]!, bounds.c[2]! - bounds.c[0]!, bounds.c[3]! - bounds.c[1]!], bounds.size);
  let playerViewport = !!bounds;
  let parsed = 0;
  for (let pos = hdr + hsize * 2; !cv.full; ) {
    const words = input.u32(pos), fn = input.u16(pos + 4);
    if (words === undefined || fn === undefined || words * 2 < 6) break;
    const rec = input.slice(pos, words * 2);
    if (!rec || fn === 0) break;
    wmfRecord(cv, fn, rec);
    if (fn === 0x020c && cv.windowSet && !playerViewport && !cv.viewportSet) {
      cv.playInto(undefined, [Math.abs(cv.map.we[0]), Math.abs(cv.map.we[1])]);
      playerViewport = true;
    }
    parsed++;
    pos += words * 2;
  }
  if (!parsed) return undefined;
  let frame: Frame | undefined;
  if (bounds && !cv.viewportSet && cv.map.usesExtents()) frame = { rect: [0, 0, bounds.size[0], bounds.size[1]], size: bounds.size };
  else if (bounds) {
    const a = cv.map.device([bounds.c[0]!, bounds.c[1]!]), b = cv.map.device([bounds.c[2]!, bounds.c[3]!]);
    const rect = [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.abs(b[0] - a[0]), Math.abs(b[1] - a[1])];
    if (rect[2]! > 0 && rect[3]! > 0) frame = { rect, size: bounds.size };
  } else if (playerViewport && !cv.viewportSet && cv.map.usesExtents()) frame = { rect: [0, 0, cv.map.ve[0], cv.map.ve[1]] };
  else if (cv.windowSet) frame = { rect: cv.map.windowRect() };
  return { cv, frame };
}
function wmfRecord(cv: Canvas, fn: number, r: Bytes) {
  switch (fn) {
    case 0x0103: { const m = r.u16(6); if (m !== undefined) cv.setMapMode(m); break; }
    case 0x0106: { const m = r.u16(6); if (m !== undefined) cv.setPolyfill(m); break; }
    case 0x0104: { const m = r.u16(6); if (m !== undefined) cv.setRop2(m); break; }
    case 0x020b: { const p = xy16(r, 6); if (p) cv.setWindowOrg(p); break; }
    case 0x020c: { const p = xy16(r, 6); if (p) cv.setWindowExt(p); break; }
    case 0x020d: { const p = xy16(r, 6); if (p) cv.setViewportOrg(p); break; }
    case 0x020e: { const p = xy16(r, 6); if (p) cv.setViewportExt(p); break; }
    case 0x001e: cv.saveDc(); break;
    case 0x0127: { const n = r.i16(6); if (n !== undefined) cv.restoreDc(n); break; }
    case 0x0214: { const p = xy16(r, 6); if (p) cv.moveTo(p); break; }
    case 0x0213: { const p = xy16(r, 6); if (p) cv.lineTo(p); break; }
    case 0x0325: case 0x0324: {
      const n = r.u16(6), pts = n === undefined ? undefined : r.tail(8)?.points(0, n, false);
      if (pts) cv.poly(pts, fn === 0x0324, fn === 0x0324, true);
      break;
    }
    case 0x0538: { // POLYPOLYGON
      const n = r.u16(6);
      if (n === undefined) break;
      const counts: number[] = [];
      for (let i = 0; i < n; i++) { const c = r.u16(8 + i * 2); if (c === undefined) return; counts.push(c); }
      const total = counts.reduce((a, c) => a + c, 0);
      const pts = total <= MAX_PATH_OPS ? r.tail(8 + n * 2)?.points(0, total, false) : undefined;
      if (pts) cv.shape(subpaths(pts, counts, true), true, true);
      break;
    }
    case 0x041b: { const b = rect16(r, 6); if (b) cv.shape(rectOps(b[0]!, b[1]!, b[2]!, b[3]!), true, true); break; }
    case 0x061c: { const eh = r.i16(6), ew = r.i16(8), b = rect16(r, 10); if (b && ew !== undefined && eh !== undefined) cv.shape(roundRect(b[0]!, b[1]!, b[2]!, b[3]!, ew, eh), true, true); break; }
    case 0x0418: { const b = rect16(r, 6); if (b) cv.shape(ellipse(b[0]!, b[1]!, b[2]!, b[3]!), true, true); break; }
    case 0x0817: case 0x081a: case 0x0830: {
      const e = xy16(r, 6), s = xy16(r, 10), b = rect16(r, 14);
      if (!e || !s || !b) break;
      const kind: ArcKind = fn === 0x0817 ? "arc" : fn === 0x081a ? "pie" : "chord";
      cv.shape(arcShape(b[0]!, b[1]!, b[2]!, b[3]!, s, e, kind), kind !== "arc", true);
      break;
    }
    case 0x02fa: { const style = r.u16(6), width = r.i16(8), color = r.u32(12); if (style !== undefined && width !== undefined && color !== undefined) cv.addObject(undefined, { k: "pen", v: { style, width, color: colorref(color) } }); break; }
    case 0x02fc: { const style = r.u16(6), color = r.u32(8); if (style !== undefined && color !== undefined) cv.addObject(undefined, { k: "brush", v: { style, color: colorref(color) } }); break; }
    case 0x02fb: { // CREATEFONTINDIRECT (richdoc): height, width, escapement, orientation, weight, italic ..., face at 24
      const height = r.i16(6), esc = r.i16(10), weight = r.i16(14), italic = r.u8(16), charset = r.u8(19);
      if (height === undefined || esc === undefined || weight === undefined || italic === undefined || charset === undefined) { cv.addObject(undefined, { k: "other" }); break; }
      const faceBytes = r.b.subarray(24, Math.min(r.length, 56));
      const z = faceBytes.indexOf(0);
      cv.addObject(undefined, { k: "font", v: { height, escapement: esc, weight, italic: italic !== 0, face: ansi(z >= 0 ? faceBytes.subarray(0, z) : faceBytes, charset), charset } });
      break;
    }
    case 0x00f7: case 0x01f9: case 0x0142: case 0x06ff: cv.addObject(undefined, { k: "other" }); break;
    case 0x012d: { const h = r.u16(6); if (h !== undefined) cv.select(wmfHandle(h)); break; }
    case 0x01f0: { const h = r.u16(6); if (h !== undefined) cv.deleteObject(h); break; }
    case 0x0209: { const c = r.u32(6); if (c !== undefined) cv.setTextColor(colorref(c)); break; } // SETTEXTCOLOR
    case 0x012e: { const a = r.u16(6); if (a !== undefined) cv.setTextAlign(a); break; } // SETTEXTALIGN
    case 0x0521: { // TEXTOUT: length, string (padded), y, x
      const n = r.u16(6);
      if (!n) break;
      const at = 8 + n + (n % 2), y = r.i16(at), x = r.i16(at + 2);
      const s = r.slice(8, n);
      if (s && x !== undefined && y !== undefined) cv.text([x, y], ansi(s.b, cv.charset()), []);
      break;
    }
    case 0x0a32: { // EXTTEXTOUT: y, x, length, options, [rect], string (padded), [dx]
      const y = r.i16(6), x = r.i16(8), n = r.u16(10), opts = r.u16(12);
      if (y === undefined || x === undefined || !n || opts === undefined) break;
      const sAt = 14 + (opts & 0x0006 ? 8 : 0);
      const s = r.slice(sAt, n);
      if (!s) break;
      const text = ansi(s.b, cv.charset());
      const dxAt = sAt + n + (n % 2);
      const dx: number[] = [];
      // Advances are per byte; they map onto characters only for a single-byte code page.
      if ([...text].length === n && r.slice(dxAt, n * 2)) for (let i = 0; i < n; i++) dx.push(r.i16(dxAt + i * 2)!);
      cv.text([x, y], text, dx);
      break;
    }
    case 0x0f43: { // STRETCHDIB
      const rop = r.u32(6) ?? 0;
      wmfBlit(cv, r, 28, rop, [i16s(r, 26), i16s(r, 24), i16s(r, 22), i16s(r, 20)], [i16s(r, 18), i16s(r, 16), i16s(r, 14), i16s(r, 12)]);
      break;
    }
    case 0x0b41: { // DIBSTRETCHBLT
      const rop = r.u32(6);
      if (rop === undefined) break;
      if (r.length / 2 <= (fn >> 8) + 3) { cv.fillRop([i16s(r, 26), i16s(r, 24), i16s(r, 22), i16s(r, 20)], rop); break; }
      const dst = [i16s(r, 24), i16s(r, 22), i16s(r, 20), i16s(r, 18)];
      if (Canvas.sourceless(rop)) { cv.fillRop(dst, rop); break; }
      wmfBlit(cv, r, 26, rop, dst, [i16s(r, 16), i16s(r, 14), i16s(r, 12), i16s(r, 10)]);
      break;
    }
    case 0x0940: { // DIBBITBLT
      const rop = r.u32(6);
      if (rop === undefined) break;
      if (r.length / 2 <= (fn >> 8) + 3) { cv.fillRop([i16s(r, 22), i16s(r, 20), i16s(r, 18), i16s(r, 16)], rop); break; }
      const dst = [i16s(r, 20), i16s(r, 18), i16s(r, 16), i16s(r, 14)];
      if (Canvas.sourceless(rop)) { cv.fillRop(dst, rop); break; }
      wmfBlit(cv, r, 22, rop, dst, [i16s(r, 12), i16s(r, 10), dst[2]!, dst[3]!]);
      break;
    }
  }
}
function wmfBlit(cv: Canvas, r: Bytes, at: number, rop: number, dst: number[], src: number[]) {
  const dib = r.tail(at);
  const n = dib && dibInfoLen(dib);
  if (!dib || n === undefined) return;
  const info = dib.slice(0, n), bits = dib.tail(n);
  if (!info || !bits) return;
  const d = decodeClip(info, bits, false, src, dst);
  if (!d) return;
  prepare(d[0], { rop });
  cv.bitmap(d[1], d[0]);
}

// ---- picture: device space mapped into the frame ----
type Item =
  | { k: "path"; ops: Op[]; fill?: RGBA; stroke?: [RGBA, number]; evenOdd: boolean }
  | { k: "bmp"; rect: number[]; img: Img }
  | { k: "text"; t: RawText };
export interface Picture { width: number; height: number; items: Item[] }

function build(raw: Raw[], frame0: Frame | undefined): Picture {
  const frame = frame0 && frame0.rect.every(isFinite) && frame0.rect[2]! > 0 && frame0.rect[3]! > 0 ? frame0 : undefined;
  let rect: number[], size: P;
  if (frame) { rect = frame.rect; size = frame.size ?? [frame.rect[2]!, frame.rect[3]!]; }
  else {
    const b = bbox(raw);
    if (!b) return { width: 0, height: 0, items: [] };
    rect = [b[0], b[1], b[2] - b[0], b[3] - b[1]];
    size = [rect[2]!, rect[3]!];
  }
  const sx = rect[2]! > 0 ? size[0] / rect[2]! : 1, sy = rect[3]! > 0 ? size[1] / rect[3]! : 1;
  const tf = (x: number, y: number): P => [(x - rect[0]!) * sx, (y - rect[1]!) * sy];
  const diag = Math.hypot(size[0], size[1]);
  const items: Item[] = [];
  for (const r of raw) {
    if (r.k === "path") {
      const ops = r.ops.map((o): Op => {
        if (o[0] === 3) return o;
        if (o[0] === 2) { const a = tf(o[1], o[2]), b = tf(o[3], o[4]), c = tf(o[5], o[6]); return [2, a[0], a[1], b[0], b[1], c[0], c[1]]; }
        const p = tf(o[1], o[2]);
        return [o[0], p[0], p[1]];
      });
      if (ops.length < 2 || !ops.every((o) => o.every((v) => isFinite(v)))) continue;
      items.push({ k: "path", ops, fill: r.fill, stroke: r.stroke && [r.stroke[0], Math.min(diag, (r.stroke[1] * (sx + sy)) / 2)], evenOdd: r.evenOdd });
    } else if (r.k === "bmp") {
      const p = tf(r.rect[0]!, r.rect[1]!);
      items.push({ k: "bmp", rect: [p[0], p[1], r.rect[2]! * sx, r.rect[3]! * sy], img: r.img });
    } else {
      const p = tf(r.t.origin[0], r.t.origin[1]);
      const size2 = r.t.size * Math.abs(sy);
      if (size2 > 0 && isFinite(size2)) items.push({ k: "text", t: { ...r.t, origin: p, size: size2, dx: r.t.dx.map((d) => d * Math.abs(sx)) } });
    }
  }
  return { width: size[0], height: size[1], items };
}
function bbox(raw: Raw[]): [number, number, number, number] | undefined {
  let b: [number, number, number, number] | undefined;
  const grow = (x: number, y: number) => {
    if (!isFinite(x) || !isFinite(y)) return;
    b = b ? [Math.min(b[0], x), Math.min(b[1], y), Math.max(b[2], x), Math.max(b[3], y)] : [x, y, x, y];
  };
  for (const r of raw) {
    if (r.k === "path") for (const o of r.ops) { if (o[0] === 2) { grow(o[1], o[2]); grow(o[3], o[4]); grow(o[5], o[6]); } else if (o[0] !== 3) grow(o[1], o[2]); }
    else if (r.k === "bmp") { grow(r.rect[0]!, r.rect[1]!); grow(r.rect[0]! + r.rect[2]!, r.rect[1]! + r.rect[3]!); }
    else grow(r.t.origin[0], r.t.origin[1]);
  }
  return b;
}

/** Parse a WMF (placeable or plain) or EMF file into a picture, or undefined when it is neither. */
export function parseMetafile(bytes: Uint8Array): Picture | undefined {
  const b = new Bytes(bytes);
  try {
    const r = isEmf(b) ? parseEmf(b) : isWmf(b) ? parseWmf(b) : undefined;
    return r && build(r.cv.items, r.frame);
  } catch {
    return undefined;
  }
}

// ---- output: SVG with PNG bitmaps (richdoc) ----
const CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(parts: Uint8Array[]): number {
  let c = 0xffffffff;
  for (const p of parts) for (let i = 0; i < p.length; i++) c = CRC[(c ^ p[i]!) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
/** An RGBA image as a PNG (8-bit RGBA, no filtering). */
export function encodePng(img: Img): Uint8Array {
  const raw = new Uint8Array((img.w * 4 + 1) * img.h);
  for (let y = 0; y < img.h; y++) raw.set(img.px.subarray(y * img.w * 4, (y + 1) * img.w * 4), y * (img.w * 4 + 1) + 1);
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    const v = new DataView(out.buffer);
    v.setUint32(0, data.length);
    const t = new Uint8Array([...type].map((c) => c.charCodeAt(0)));
    out.set(t, 4);
    out.set(data, 8);
    v.setUint32(8 + data.length, crc32([t, data]));
    return out;
  };
  const ihdr = new Uint8Array(13);
  const iv = new DataView(ihdr.buffer);
  iv.setUint32(0, img.w);
  iv.setUint32(4, img.h);
  ihdr.set([8, 6, 0, 0, 0], 8);
  const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", zlibSync(raw)), chunk("IEND", new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
const n4 = (v: number) => String(Math.round(v * 1000) / 1000);
const hex = (c: RGBA) => `#${[c[0], c[1], c[2]].map((v) => v.toString(16).padStart(2, "0")).join("")}`;
const xmlEsc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
function b64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
/** A picture as an SVG document stretched (like Word) over its own box. Text uses the named font,
    with Arial-like metrics standing in for the font's own (cell heights, alignment). */
export function pictureToSvg(pic: Picture): string {
  let body = "";
  for (const it of pic.items) {
    if (it.k === "path") {
      let d = "";
      for (const o of it.ops) d += o[0] === 0 ? `M${n4(o[1])} ${n4(o[2])}` : o[0] === 1 ? `L${n4(o[1])} ${n4(o[2])}` : o[0] === 2 ? `C${n4(o[1])} ${n4(o[2])} ${n4(o[3])} ${n4(o[4])} ${n4(o[5])} ${n4(o[6])}` : "Z";
      const fill = it.fill ? `fill="${hex(it.fill)}"${it.fill[3] < 255 ? ` fill-opacity="${n4(it.fill[3] / 255)}"` : ""}${it.evenOdd ? ' fill-rule="evenodd"' : ""}` : 'fill="none"';
      const stroke = it.stroke ? ` stroke="${hex(it.stroke[0])}" stroke-width="${it.stroke[1] > 0 ? n4(it.stroke[1]) : '1" vector-effect="non-scaling-stroke'}" stroke-linejoin="round"` : "";
      body += `<path d="${d}" ${fill}${stroke}/>`;
    } else if (it.k === "bmp") {
      const [x, y, w, h] = it.rect as [number, number, number, number];
      body += `<image x="${n4(x)}" y="${n4(y)}" width="${n4(w)}" height="${n4(h)}" preserveAspectRatio="none" href="data:image/png;base64,${b64(encodePng(it.img))}"/>`;
    } else {
      const t = it.t;
      const em = t.cell ? t.size / 1.15 : t.size;
      const baseline = (t.align & 24) === 24 ? 0 : (t.align & 24) === 8 ? -0.21 * em : 0.9 * em;
      const h = t.align & 6;
      const chars = [...t.text];
      let xs = "";
      let anchor = h === 6 ? "middle" : h === 2 ? "end" : "start";
      if (t.dx.length === chars.length && chars.length > 1) {
        const total = t.dx.reduce((a, b) => a + b, 0);
        let x = h === 6 ? -total / 2 : h === 2 ? -total : 0;
        const pos: string[] = [];
        for (const d of t.dx) { pos.push(n4(x)); x += d; }
        xs = ` x="${pos.join(" ")}"`;
        anchor = "start";
      }
      const rot = t.angle ? ` rotate(${n4((t.angle * 180) / Math.PI)})` : "";
      body += `<text transform="translate(${n4(t.origin[0])} ${n4(t.origin[1])})${rot}"${xs} y="${n4(baseline)}" font-family="${xmlEsc(t.font)}, Arial, sans-serif" font-size="${n4(em)}"${t.bold ? ' font-weight="bold"' : ""}${t.italic ? ' font-style="italic"' : ""} fill="${hex(t.color)}" text-anchor="${anchor}" xml:space="preserve">${xmlEsc(t.text)}</text>`;
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n4(pic.width)} ${n4(pic.height)}" preserveAspectRatio="none">${body}</svg>`;
}

/** An EMF / WMF picture as an SVG data URL, or undefined when it cannot be read or draws nothing. */
export function metafileToSvgUrl(bytes: Uint8Array): string | undefined {
  const pic = parseMetafile(bytes);
  if (!pic || !pic.items.length || !(pic.width > 0) || !(pic.height > 0)) return undefined;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(pictureToSvg(pic))}`;
}
