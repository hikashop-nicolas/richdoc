/*! richdoc WordArt drawing: a port of WordCraft's WordArt reader and warps (https://github.com/storytold/wordcraft).
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
// WordArt text boxes (wps:wsp) as SVG: w14 text effects drawn as layers (reflection, shadow, glow,
// fill, outline) and a:prstTxWarp warps. Without glyph outlines, each letter is bent as a whole.
import { drawingColor, modColor, themeColor } from "./fonts";
import { int, presetAdjusts, presetPath, textWidth } from "./smartart";

const MAX_STOPS = 10;
const MAX_OUTLINE = 20;
const MAX_ADJ_VALUE = 100_000_000;
const MAX_LETTERS = 5000;
const PT_PX = 4 / 3;
const EMU_PX = 9525;
const LINE = 1.2;
// A letter's estimated box around its baseline, in em.
const ASC = 0.75, DESC = 0.2, EM = ASC + DESC;

/** Every DrawingML preset text warp (ST_TextShapeType); a warp with another name is dropped. */
export const WARP_PRESETS = ["textNoShape", "textPlain", "textStop", "textTriangle", "textTriangleInverted", "textChevron", "textChevronInverted", "textRingInside", "textRingOutside", "textArchUp", "textArchDown", "textCircle", "textButton", "textArchUpPour", "textArchDownPour", "textCirclePour", "textButtonPour", "textCurveUp", "textCurveDown", "textCanUp", "textCanDown", "textWave1", "textWave2", "textDoubleWave1", "textWave4", "textInflate", "textDeflate", "textInflateBottom", "textDeflateBottom", "textInflateTop", "textDeflateTop", "textDeflateInflate", "textDeflateInflateDeflate", "textFadeRight", "textFadeLeft", "textFadeUp", "textFadeDown", "textSlantUp", "textSlantDown", "textCascadeUp", "textCascadeDown"];

export interface TextWarp { preset: string; adj: Record<string, number> }
interface Stop { pos: number; color: string; transparency: number }
type TextFill = { t: "none" } | { t: "solid"; color: string; transparency: number } | { t: "grad"; stops: Stop[]; angle: number };
interface TextEffects {
  fill?: TextFill;
  outline?: { color: string | null; width: number; transparency: number };
  shadow?: { color: string; transparency: number; blur: number; distance: number; angle: number };
  glow?: { color: string; size: number; transparency: number };
  reflection?: { transparency: number; size: number; distance: number; blur: number };
}

const child = (e: Element | undefined | null, tag: string): Element | undefined => (e ? Array.from(e.children).find((c) => c.tagName === tag) : undefined);
const clamp = (v: number, lo: number, hi: number, d: number) => (Number.isFinite(v) ? Math.min(hi, Math.max(lo, v)) : d);
const mod = (a: number, n: number) => ((a % n) + n) % n;

/** A preset text warp (a:prstTxWarp) with its adjust values; undefined for unknown presets. */
export function readTextWarp(e: Element | undefined): TextWarp | undefined {
  const preset = e?.getAttribute("prst");
  if (!preset || !WARP_PRESETS.includes(preset)) return undefined;
  const adj: Record<string, number> = {};
  for (const gd of Array.from(child(e, "a:avLst")?.children ?? [])) {
    const name = gd.getAttribute("name"), f = gd.getAttribute("fmla")?.trim();
    if (!name || name.length > 16 || !/^[A-Za-z0-9]+$/.test(name) || !f?.startsWith("val")) continue;
    const v = Number(f.slice(3).trim());
    if (!Number.isFinite(v) || !Number.isInteger(v)) continue;
    if (!(name in adj) && Object.keys(adj).length >= 8) continue;
    adj[name] = Math.max(-MAX_ADJ_VALUE, Math.min(MAX_ADJ_VALUE, v));
  }
  return { preset, adj };
}

// ---- w14 run effects ([MS-DOCX] 2.6.1); lengths EMU, angles 60000ths of a degree ----

// The w14 colour inside e with lumMod/lumOff applied, and its transparency (w14:alpha), percent.
function w14Color(e: Element, colors: Record<string, string>): [string, number] | undefined {
  const c = Array.from(e.children).find((x) => x.tagName === "w14:srgbClr" || x.tagName === "w14:schemeClr");
  const val = c?.getAttribute("w14:val");
  if (!c || !val) return undefined;
  let rgb = c.tagName === "w14:srgbClr" ? (/^[0-9a-f]{6}$/i.test(val) ? val.toUpperCase() : undefined) : themeColor(colors, val);
  if (!rgb) return undefined;
  let transparency = 0;
  for (const m of Array.from(c.children)) {
    const raw = int(m.getAttribute("w14:val"));
    const v = raw === undefined ? 0 : clamp(raw / 100_000, -10, 10, 0);
    if (m.tagName === "w14:lumMod") rgb = modColor(rgb, { lumMod: v * 100_000 });
    else if (m.tagName === "w14:lumOff") rgb = modColor(rgb, { lumOff: v * 100_000 });
    // Unlike DrawingML's a:alpha, w14:alpha is the transparency.
    else if (m.tagName === "w14:alpha") transparency = clamp(v, 0, 1, 0) * 100;
  }
  return [rgb, transparency];
}

const emuPt = (e: Element, n: string) => { const v = int(e.getAttribute(n)); return v === undefined ? 0 : clamp(v / 12_700, 0, 2000, 0); };
const angleOf = (e: Element, n: string) => { const v = int(e.getAttribute(n)); return v === undefined ? 0 : mod(v, 21_600_000) / 60_000; };

function textFill(e: Element, colors: Record<string, string>): TextFill | undefined {
  const names = ["w14:noFill", "w14:solidFill", "w14:gradFill"];
  const f = names.includes(e.tagName) ? e : Array.from(e.children).find((c) => names.includes(c.tagName));
  if (!f) return undefined;
  if (f.tagName === "w14:noFill") return { t: "none" };
  if (f.tagName === "w14:solidFill") {
    const c = w14Color(f, colors);
    return c ? { t: "solid", color: c[0], transparency: clamp(c[1], 0, 100, 0) } : undefined;
  }
  const stops: Stop[] = [];
  for (const gs of Array.from(child(f, "w14:gsLst")?.children ?? []).filter((g) => g.tagName === "w14:gs").slice(0, MAX_STOPS)) {
    const c = w14Color(gs, colors);
    if (!c) continue;
    const p = int(gs.getAttribute("w14:pos"));
    stops.push({ pos: p === undefined ? 0 : clamp(p, 0, 100_000, 0) / 1000, color: c[0], transparency: clamp(c[1], 0, 100, 0) });
  }
  stops.sort((a, b) => a.pos - b.pos);
  if (!stops.length) return undefined;
  const lin = child(f, "w14:lin");
  return { t: "grad", stops, angle: lin ? angleOf(lin, "w14:ang") : 90 };
}

// One colour standing for a fill (a gradient's middle stop).
const mainColor = (f: TextFill | undefined): string | undefined => (f?.t === "solid" ? f.color : f?.t === "grad" ? f.stops[Math.floor(f.stops.length / 2)]?.color : undefined);

/** The WordArt effects of a w:rPr (w14:textFill, textOutline, shadow, glow, reflection). */
export function readTextEffects(rpr: Element | undefined, colors: Record<string, string>): TextEffects {
  const fx: TextEffects = {};
  for (const e of Array.from(rpr?.children ?? [])) {
    switch (e.tagName) {
      case "w14:textFill": fx.fill = textFill(e, colors); break;
      case "w14:textOutline": {
        const f = Array.from(e.children).find((c) => c.tagName === "w14:solidFill" || c.tagName === "w14:gradFill" || c.tagName === "w14:noFill");
        let color: string | null = "000000", transparency = 0;
        if (f?.tagName === "w14:noFill") color = null;
        else if (f) {
          const t = textFill(f, colors);
          if (t?.t === "solid") [color, transparency] = [t.color, t.transparency];
          else color = mainColor(t) ?? null;
        }
        fx.outline = { color, width: clamp(emuPt(e, "w14:w"), 0, MAX_OUTLINE, 0.75), transparency };
        break;
      }
      case "w14:shadow": {
        const [color, transparency] = w14Color(e, colors) ?? ["000000", 60];
        fx.shadow = { color, transparency: clamp(transparency, 0, 100, 60), blur: Math.min(100, emuPt(e, "w14:blurRad")), distance: Math.min(200, emuPt(e, "w14:dist")), angle: angleOf(e, "w14:dir") };
        break;
      }
      case "w14:glow": {
        const [color, transparency] = w14Color(e, colors) ?? ["156082", 60];
        const size = Math.min(150, emuPt(e, "w14:rad"));
        fx.glow = size > 0 ? { color, size, transparency: clamp(transparency, 0, 100, 60) } : undefined;
        break;
      }
      case "w14:reflection": {
        // Start opacity and the end position of the fade (both 1000ths of a percent).
        const pct = (n: string, d: number) => { const v = int(e.getAttribute(n)); return v === undefined ? d : Math.max(0, Math.min(100_000, v)) / 1000; };
        fx.reflection = { transparency: 100 - pct("w14:stA", 50), size: pct("w14:endPos", 50), distance: Math.min(100, emuPt(e, "w14:dist")), blur: emuPt(e, "w14:blurRad") };
        break;
      }
    }
  }
  return fx;
}

// ---- warps (wordcraft layout/src/wordart.rs) ----

type Edge = (u: number, k: number) => [number, number];
const slantUp: Edge = (u, k) => [k * (1 - u), 1 - k * u];
const slantDown: Edge = (u, k) => [k * u, 1 - k * (1 - u)];
const inflate: Edge = (u, k) => { const s = Math.sin(Math.PI * u); return [k * (1 - s), 1 - k * (1 - s)]; };
const deflate: Edge = (u, k) => { const s = Math.sin(Math.PI * u); return [k * s, 1 - k * s]; };
const wave: Edge = (u, k) => { const d = k * Math.sin(2 * Math.PI * u); return [k - d, 1 - k - d]; };
const chevronUp: Edge = (u, k) => { const t = Math.abs(2 * u - 1); return [k * t, 1 - k * (1 - t)]; };
const chevronDown: Edge = (u, k) => { const t = Math.abs(2 * u - 1); return [k * (1 - t), 1 - k * t]; };
const triangleUp: Edge = (u, k) => [k * Math.abs(2 * u - 1), 1];
const triangleDown: Edge = (u, k) => [0, 1 - k * Math.abs(2 * u - 1)];

type Bend = { edges: Edge } | { start: number; sweep: number; outward: boolean };
export interface Frame { x: number; y: number; w: number; h: number }

/** A warp fitted to a frame: maps the text's box (u, v in 0..1, v = 0 at the top) onto it. */
export class Warp {
  private constructor(private frame: Frame, private bend: Bend, private k: number, private thick: number) {}

  /** The warp w fitted to frame for text aspect times as tall as wide; undefined for warps drawn flat. */
  static create(w: TextWarp, frame: Frame, aspect: number): Warp | undefined {
    const PI = Math.PI;
    if (!(frame.w > 0.5 && frame.h > 0.5 && Number.isFinite(frame.w) && Number.isFinite(frame.h))) return undefined;
    const a = w.adj.adj;
    const frac = (d: number, max: number) => (a === undefined ? d : Math.max(0, Math.min(max, a / 100_000)));
    const sweep = a === undefined ? PI : (Math.max(10, Math.min(350, a / 60_000)) * PI) / 180;
    let bend: Bend, k = 0;
    switch (w.preset) {
      case "textSlantUp": bend = { edges: slantUp }; k = frac(0.4, 0.9); break;
      case "textSlantDown": bend = { edges: slantDown }; k = frac(0.4, 0.9); break;
      case "textInflate": bend = { edges: inflate }; k = frac(0.3, 0.45); break;
      case "textDeflate": bend = { edges: deflate }; k = frac(0.3, 0.45); break;
      case "textWave1": bend = { edges: wave }; k = frac(0.125, 0.3); break;
      case "textChevron": bend = { edges: chevronUp }; k = frac(0.4, 0.9); break;
      case "textChevronInverted": bend = { edges: chevronDown }; k = frac(0.4, 0.9); break;
      case "textTriangle": bend = { edges: triangleUp }; k = frac(0.5, 0.9); break;
      case "textTriangleInverted": bend = { edges: triangleDown }; k = frac(0.5, 0.9); break;
      case "textArchUp": bend = { start: PI / 2 + sweep / 2, sweep: -sweep, outward: true }; break;
      case "textArchDown": bend = { start: 1.5 * PI - sweep / 2, sweep, outward: false }; break;
      case "textCircle": bend = { start: 1.5 * PI - 0.02, sweep: -(2 * PI - 0.04), outward: true }; break;
      default: return undefined;
    }
    // Along an arc the text keeps its proportions, within reason.
    let thick = 0;
    if ("sweep" in bend) {
      const as = Number.isFinite(aspect) ? Math.max(0, Math.min(10, aspect)) : 0.2;
      thick = Math.max(0.08, Math.min(0.6, as * Math.abs(bend.sweep)));
    }
    return new Warp(frame, bend, k, thick);
  }

  /** The point for text box point (u, v). */
  map(u: number, v: number): [number, number] {
    u = Number.isFinite(u) ? u : 0;
    v = Number.isFinite(v) ? v : 0;
    const f = this.frame, b = this.bend;
    if ("edges" in b) {
      const [top, bottom] = b.edges(Math.max(-1, Math.min(2, u)), this.k);
      return [f.x + u * f.w, f.y + (top + (bottom - top) * v) * f.h];
    }
    const rx = f.w / 2, ry = f.h / 2, cx = f.x + rx, cy = f.y + ry;
    const a = b.start + b.sweep * u;
    const rho = b.outward ? 1 - v * this.thick : 1 - this.thick + v * this.thick;
    return [cx + rx * rho * Math.cos(a), cy - ry * rho * Math.sin(a)];
  }
}

// ---- reading the text box ----

interface Run { text: string; size: number; bold: boolean; italic: boolean; font?: string; color: string; fx: TextEffects }
interface Para { runs: Run[]; align: "left" | "center" | "right" }

const onOff = (e: Element | undefined) => !!e && !/^(0|false|off)$/.test(e.getAttribute("w:val") ?? "");

function readRun(r: Element, colors: Record<string, string>): Run {
  const rpr = child(r, "w:rPr");
  let text = "";
  for (const c of Array.from(r.children)) {
    if (c.tagName === "w:t") text += c.textContent ?? "";
    else if (c.tagName === "w:tab") text += " ";
  }
  const sz = int(child(rpr, "w:sz")?.getAttribute("w:val"));
  const col = child(rpr, "w:color");
  const cv = col?.getAttribute("w:val");
  const color = themeColor(colors, col?.getAttribute("w:themeColor")) ?? (cv && /^[0-9a-f]{6}$/i.test(cv) ? cv.toUpperCase() : "000000");
  const rf = child(rpr, "w:rFonts");
  const font = rf?.getAttribute("w:ascii") ?? rf?.getAttribute("w:hAnsi") ?? undefined;
  return { text: text.replace(/[\u0000-\u001f]/g, " "), size: (sz === undefined ? 11 : Math.max(1, Math.min(1638, sz / 2))) * PT_PX, bold: onOff(child(rpr, "w:b")), italic: onOff(child(rpr, "w:i")), font, color, fx: readTextEffects(rpr, colors) };
}

function readParas(wsp: Element, colors: Record<string, string>): Para[] {
  const content = wsp.getElementsByTagName("w:txbxContent")[0];
  if (!content) return [];
  return Array.from(content.getElementsByTagName("w:p")).map((p) => {
    const runs = Array.from(p.getElementsByTagName("w:r")).filter((r) => nearestP(r) === p).map((r) => readRun(r, colors)).filter((r) => r.text);
    const jc = child(child(p, "w:pPr"), "w:jc")?.getAttribute("w:val");
    return { runs, align: jc === "center" ? "center" : jc === "right" || jc === "end" ? "right" : "left" };
  });
}
function nearestP(e: Element): Element | null {
  let c = e.parentElement;
  while (c && c.tagName !== "w:p") c = c.parentElement;
  return c;
}

// ---- drawing ----

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const n2 = (n: number) => String(Math.round(n * 100) / 100);
let uid = 0;

interface Box { x0: number; y0: number; x1: number; y1: number }
const grow = (b: Box | undefined, x: number, y: number): Box => (b ? { x0: Math.min(b.x0, x), y0: Math.min(b.y0, y), x1: Math.max(b.x1, x), y1: Math.max(b.y1, y) } : { x0: x, y0: y, x1: x, y1: y });

// A run's placed pieces: SVG text elements awaiting paint attributes, and their box.
interface Placed { run: Run; pieces: Array<(attrs: string) => string>; box?: Box }

const fontAttrs = (r: Run) => `font-family="${r.font ? `'${esc(r.font).replace(/'/g, "")}', ` : ""}Calibri, Arial, sans-serif" font-size="${n2(r.size)}"${r.bold ? ` font-weight="bold"` : ""}${r.italic ? ` font-style="italic"` : ""}`;

/** A standalone SVG of a WordArt text box (wps:wsp with a:prstTxWarp or w14 text effects); undefined for other shapes. */
export function wordArtSvg(wsp: Element, colors: Record<string, string>, widthPx: number, heightPx: number): string | undefined {
  if (!(Number.isFinite(widthPx) && widthPx > 0 && Number.isFinite(heightPx) && heightPx > 0)) return undefined;
  const bodyPr = child(wsp, "wps:bodyPr");
  const warp = readTextWarp(child(bodyPr, "a:prstTxWarp"));
  const paras = readParas(wsp, colors);
  const hasFx = paras.some((p) => p.runs.some((r) => r.fx.fill || r.fx.outline || r.fx.shadow || r.fx.glow || r.fx.reflection));
  if (!warp && !hasFx) return undefined;
  if (!paras.some((p) => p.runs.some((r) => r.text.trim()))) return undefined;
  const W = widthPx, H = heightPx;
  const ins = (n: string, d: number) => { const v = int(bodyPr?.getAttribute(n)); return v === undefined ? d : Math.max(0, Math.min(W, v / EMU_PX)); };
  const l = ins("lIns", 91_440 / EMU_PX), t = ins("tIns", 45_720 / EMU_PX), r = ins("rIns", 91_440 / EMU_PX), b = ins("bIns", 45_720 / EMU_PX);
  const frame: Frame = { x: l, y: t, w: Math.max(1, W - l - r), h: Math.max(1, H - t - b) };
  const id = `wa${++uid}`;

  // Source layout: lines stacked, letters at estimated widths.
  const lines = paras.map((p) => {
    const size = p.runs.reduce((m, x) => Math.max(m, x.size), 0) || 11 * PT_PX;
    return { p, size, width: p.runs.reduce((s, x) => s + textWidth(x.text, x.size), 0) };
  });
  const maxW = lines.reduce((m, x) => Math.max(m, x.width), 0);
  let src: Box | undefined;
  type Letter = { run: Run; ch: string; x: number; w: number; base: number };
  const letters: Letter[] = [];
  let y = 0;
  for (const ln of lines) {
    const base = y + (ASC + (LINE - EM) / 2) * ln.size;
    let x = ln.p.align === "center" ? (maxW - ln.width) / 2 : ln.p.align === "right" ? maxW - ln.width : 0;
    for (const run of ln.p.runs) {
      for (const ch of Array.from(run.text)) {
        const w = textWidth(ch, run.size);
        if (ch.trim() && letters.length < MAX_LETTERS) {
          letters.push({ run, ch, x, w, base });
          src = grow(grow(src, x, base - ASC * run.size), x + w, base + DESC * run.size);
        }
        x += w;
      }
    }
    y += ln.size * LINE;
  }
  const srcW = src ? Math.max(1e-3, src.x1 - src.x0) : 1, srcH = src ? Math.max(1e-3, src.y1 - src.y0) : 1;
  const bent = src && warp ? Warp.create(warp, frame, srcH / srcW) : undefined;

  const placed = new Map<Run, Placed>();
  const at = (run: Run) => { let p = placed.get(run); if (!p) placed.set(run, (p = { run, pieces: [] })); return p; };
  if (bent && src) {
    // Each letter's box mapped through the warp, as an affine map of the letter.
    for (const lt of letters) {
      const s = lt.run.size;
      const u0 = (lt.x - src.x0) / srcW, u1 = (lt.x + lt.w - src.x0) / srcW, uc = (u0 + u1) / 2;
      const v0 = (lt.base - ASC * s - src.y0) / srcH, v1 = (lt.base + DESC * s - src.y0) / srcH, vm = (v0 + v1) / 2;
      const p0 = bent.map(uc, v0), pv = bent.map(uc, v1), pl = bent.map(u0, vm), pr = bent.map(u1, vm);
      const Hx = pr[0] - pl[0], Hy = pr[1] - pl[1], Vx = pv[0] - p0[0], Vy = pv[1] - p0[1];
      const m = [Hx / lt.w, Hy / lt.w, Vx / (EM * s), Vy / (EM * s), p0[0] - Hx / 2 + (ASC / EM) * Vx, p0[1] - Hy / 2 + (ASC / EM) * Vy];
      if (!m.every(Number.isFinite)) continue;
      const pl0 = at(lt.run);
      for (const [cu, cv] of [[u0, v0], [u1, v0], [u0, v1], [u1, v1]] as const) { const q = bent.map(cu, cv); pl0.box = grow(pl0.box, q[0], q[1]); }
      const tr = `matrix(${m.map(n2).join(" ")})`;
      pl0.pieces.push((attrs) => `<text x="0" y="0" ${fontAttrs(lt.run)} textLength="${n2(lt.w)}" lengthAdjust="spacingAndGlyphs" transform="${tr}"${attrs}>${esc(lt.ch)}</text>`);
    }
  } else {
    // Flat: lines from the frame's top (or as bodyPr anchors them), each run where the estimate puts it.
    const total = lines.reduce((s, x) => s + x.size * LINE, 0);
    const anchor = bodyPr?.getAttribute("anchor");
    let ly = frame.y + (anchor === "ctr" ? (frame.h - total) / 2 : anchor === "b" ? frame.h - total : 0);
    for (const ln of lines) {
      const base = ly + (ASC + (LINE - EM) / 2) * ln.size;
      const single = ln.p.runs.length === 1;
      let x = frame.x + (ln.p.align === "center" ? (frame.w - ln.width) / 2 : ln.p.align === "right" ? frame.w - ln.width : 0);
      for (const run of ln.p.runs) {
        const w = textWidth(run.text, run.size);
        const p = at(run);
        p.box = grow(grow(p.box, x, base - ASC * run.size), x + w, base + DESC * run.size);
        // A line of one run keeps its true alignment; several runs follow the estimate.
        const ax = single ? (ln.p.align === "center" ? frame.x + frame.w / 2 : ln.p.align === "right" ? frame.x + frame.w : frame.x) : x;
        const ta = single && ln.p.align !== "left" ? ` text-anchor="${ln.p.align === "center" ? "middle" : "end"}"` : "";
        p.pieces.push((attrs) => `<text x="${n2(ax)}" y="${n2(base)}"${ta} ${fontAttrs(run)} xml:space="preserve"${attrs}>${esc(run.text)}</text>`);
        x += w;
      }
      ly += ln.size * LINE;
    }
  }

  const defs: string[] = [];
  const body: string[] = [];
  // The shape behind the text, when it has a fill or line of its own.
  const sppr = child(wsp, "wps:spPr");
  const bg = drawingColor(child(sppr, "a:solidFill"), colors);
  const ln = child(sppr, "a:ln");
  const lc = drawingColor(child(ln, "a:solidFill"), colors);
  if (bg || lc) {
    const geom = child(sppr, "a:prstGeom");
    const [d, closed] = presetPath(geom?.getAttribute("prst") || "rect", 0, 0, W, H, presetAdjusts(geom));
    const lw = int(ln?.getAttribute("w"));
    body.push(`<path d="${d}" fill="${bg && closed ? `#${bg}` : "none"}"${lc ? ` stroke="#${lc}" stroke-width="${n2(lw === undefined ? PT_PX * 0.75 : Math.min(100, lw / EMU_PX))}"` : ""}/>`);
  }
  let k = 0;
  const blur = (sd: number) => {
    if (!(sd > 0.05)) return "";
    const fid = `${id}f${k++}`;
    defs.push(`<filter id="${fid}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${n2(sd)}"/></filter>`);
    return ` filter="url(#${fid})"`;
  };
  for (const p of placed.values()) {
    const bx = p.box;
    if (!bx || !p.pieces.length) continue;
    const { run } = p, fx = run.fx;
    const all = (attrs: string) => p.pieces.map((f) => f(attrs)).join("");
    const fill: TextFill = fx.fill ?? { t: "solid", color: run.color, transparency: 0 };
    const outline = fx.outline && fx.outline.color ? { color: fx.outline.color, width: Math.max(0.25, fx.outline.width) * PT_PX, a: 1 - fx.outline.transparency / 100 } : undefined;
    const ink = mainColor(fill) ?? outline?.color ?? run.color;
    // The letters as a clip, made when a gradient needs it.
    let clip = "";
    const clipId = () => { if (!clip) { clip = `${id}c${k++}`; defs.push(`<clipPath id="${clip}">${all("")}</clipPath>`); } return clip; };
    const rect = (paint: string) => `<rect x="${n2(bx.x0)}" y="${n2(bx.y0)}" width="${n2(bx.x1 - bx.x0)}" height="${n2(bx.y1 - bx.y0)}" fill="${paint}" clip-path="url(#${clipId()})"/>`;
    if (fx.reflection) {
      // Mirrored about the letters' bottom edge, distance below it, fading out downwards.
      const r = fx.reflection, gap = r.distance * PT_PX, g = `${id}g${k++}`;
      const a = 1 - r.transparency / 100;
      const end = bx.y1 - (bx.y1 - bx.y0) * Math.max(0.01, r.size / 100);
      defs.push(`<linearGradient id="${g}" gradientUnits="userSpaceOnUse" x1="${n2(bx.x0)}" y1="${n2(bx.y1)}" x2="${n2(bx.x0)}" y2="${n2(end)}"><stop offset="0" stop-color="#${ink}" stop-opacity="${n2(a)}"/><stop offset="1" stop-color="#${ink}" stop-opacity="0"/></linearGradient>`);
      body.push(`<g transform="matrix(1 0 0 -1 0 ${n2(2 * bx.y1 + gap)})">${rect(`url(#${g})`)}</g>`);
    }
    if (fx.shadow) {
      const s = fx.shadow, rad = (s.angle * Math.PI) / 180, d = s.distance * PT_PX;
      body.push(`<g transform="translate(${n2(d * Math.cos(rad))} ${n2(d * Math.sin(rad))})" opacity="${n2(1 - s.transparency / 100)}"${blur((s.blur * PT_PX) / 2)}>${all(` fill="#${s.color}"`)}</g>`);
    }
    if (fx.glow) {
      const g = fx.glow, sz = g.size * PT_PX;
      body.push(`<g opacity="${n2(1 - g.transparency / 100)}"${blur(sz / 2)}>${all(` fill="#${g.color}" stroke="#${g.color}" stroke-width="${n2(2 * sz)}" stroke-linejoin="round" vector-effect="non-scaling-stroke"`)}</g>`);
    }
    if (fill.t === "solid") body.push(all(` fill="#${fill.color}"${fill.transparency ? ` fill-opacity="${n2(1 - fill.transparency / 100)}"` : ""}`));
    else if (fill.t === "grad") {
      const g = `${id}g${k++}`, rad = (fill.angle * Math.PI) / 180, c = Math.cos(rad), s = Math.sin(rad);
      const cx = (bx.x0 + bx.x1) / 2, cy = (bx.y0 + bx.y1) / 2;
      const d = ((bx.x1 - bx.x0) * Math.abs(c) + (bx.y1 - bx.y0) * Math.abs(s)) / 2;
      const stops = fill.stops.map((st) => `<stop offset="${n2(st.pos / 100)}" stop-color="#${st.color}"${st.transparency ? ` stop-opacity="${n2(1 - st.transparency / 100)}"` : ""}/>`).join("");
      defs.push(`<linearGradient id="${g}" gradientUnits="userSpaceOnUse" x1="${n2(cx - d * c)}" y1="${n2(cy - d * s)}" x2="${n2(cx + d * c)}" y2="${n2(cy + d * s)}">${stops}</linearGradient>`);
      body.push(rect(`url(#${g})`));
    }
    if (outline) body.push(all(` fill="none" stroke="#${outline.color}" stroke-width="${n2(outline.width)}" stroke-linejoin="round"${outline.a < 1 ? ` stroke-opacity="${n2(outline.a)}"` : ""} vector-effect="non-scaling-stroke"`));
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(W)}" height="${n2(H)}" viewBox="0 0 ${n2(W)} ${n2(H)}">${defs.length ? `<defs>${defs.join("")}</defs>` : ""}${body.join("")}</svg>`;
}
