/*! richdoc chart renderer: a port of the WordCraft chart drawing (https://github.com/storytold/wordcraft).
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
// Chart parts (c:chartSpace) drawn from their cached data as an SVG in points: bar, column, line,
// area, pie, doughnut, scatter and radar; other types draw a placeholder. Text widths are estimated.
import { drawingColor } from "./fonts";

// ---- limits (hostile input yields a smaller chart, never a hang) ----
const MAX_SERIES = 256;
const MAX_POINTS = 10_000;
const MAX_CELLS = 100_000;
const MAX_LABEL = 200;
const MAX_SPOKES = 100;
const MAX_RADAR_SERIES = 50;
const MAX_TICKS = 20;
const MAX_FORMAT_CODE = 255;
const MAX_DEPTH = 256;
const MAX_LEN_PT = 14_400;

// ---- layout constants, points ----
const FONT = 9;
/** Average character width estimate, em. */
export const CHAR_W = 0.55;
const LEGEND_SHARE = 0.4;
const MAX_LEGEND_ROWS = 3;
const TITLE_SIZE = 14;
const PAD = 6;
const LINE = 1.3;
const GAP = 4;
const LINE_W = 2.25;
const SWATCH = 7;
const SEP = 10;
const TEXT = "595959";
const GRID = "D9D9D9";
const AXIS = "BFBFBF";
const WHITE = "FFFFFF";
const BLACK = "000000";
// Word's default inline chart size when the caller gives none.
const DEFAULT_W = 432;
const DEFAULT_H = 252;

// Theme slots in wordcraft order, with the Office theme as fallback.
const SCHEME = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const DEFAULT_THEME = ["000000", "FFFFFF", "0E2841", "E8E8E8", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D"];

type Rgb = string;
// Path ops: [0,x,y] move, [1,x,y] line, [2,x1,y1,x2,y2,x,y] cubic, [3] close.
type Seg = [0, number, number] | [1, number, number] | [2, number, number, number, number, number, number] | [3];
type Align = "left" | "center" | "right";
type Item =
  | { k: "shape"; rect: number[]; ellipse: boolean; fill?: Rgb; stroke?: Rgb; sw: number }
  | { k: "path"; segs: Seg[]; fill?: Rgb; stroke?: Rgb; sw: number }
  | { k: "text"; rect: number[]; text: string; size: number; color: Rgb; bold: boolean; align: Align };

interface Theme { slots: Rgb[]; named: Record<string, string> }

// ---- XML helpers (prefixed tag names, as the file writes them) ----
const els = (e: Element | undefined): Element[] => (e ? Array.from(e.children) : []);
const child = (e: Element | undefined, name: string): Element | undefined => (e ? Array.from(e.children).find((c) => c.tagName === name) : undefined);
const children = (e: Element | undefined, name: string): Element[] => els(e).filter((c) => c.tagName === name);
const find = (e: Element | undefined, name: string): Element | undefined => e?.getElementsByTagName(name)[0] ?? undefined;
const attr = (e: Element | undefined, name: string): string | undefined => e?.getAttribute(name) ?? undefined;
/** The element's own text (direct text nodes). */
export function ownText(e: Element | undefined): string {
  let s = "";
  for (const n of Array.from(e?.childNodes ?? [])) if (n.nodeType === 3 || n.nodeType === 4) s += n.nodeValue ?? "";
  return s;
}
const valOf = (e: Element | undefined, name: string) => attr(child(e, name), "val");
const chars = (s: string): string[] => Array.from(s);
const takeChars = (s: string, n: number) => chars(s).slice(0, Math.max(0, n)).join("");
const utf8Len = (s: string) => new TextEncoder().encode(s).length;

// ---- lenient numbers, as Rust's parsers take them ----
const NUM_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
const SPECIAL_RE = /^[+-]?(?:inf|infinity|nan)$/i;
/** A float as Rust's `str::parse::<f64>` reads it (no trimming). */
export function parseF64(s: string | undefined): number | undefined {
  if (s === undefined) return undefined;
  if (NUM_RE.test(s)) return Number(s);
  if (SPECIAL_RE.test(s)) return /nan/i.test(s) ? NaN : s.startsWith("-") ? -Infinity : Infinity;
  return undefined;
}
/** Plain integer: decimals rounded (half away from zero), clamped to the i64 range. */
function int(s: string | undefined): number | undefined {
  if (s === undefined) return undefined;
  const t = s.trim();
  if (!t || t.length > 40) return undefined;
  const v = parseF64(t);
  if (v === undefined || !Number.isFinite(v)) return undefined;
  return Math.max(-9.223372036854776e18, Math.min(9.223372036854776e18, Math.sign(v) * Math.round(Math.abs(v))));
}
/** An unsigned decimal integer (Rust `parse::<u32>` / `parse::<usize>`), else undefined. */
export function parseUint(s: string | undefined, max: number): number | undefined {
  if (s === undefined || !/^\+?\d+$/.test(s)) return undefined;
  const d = s.replace(/^\+?0*/, "");
  if (d.length > 20) return undefined;
  const v = Number(d || "0");
  return v <= max ? v : undefined;
}
/** A length in points: a number in file units (`perPt` a point) or a universal measure. */
function measure(s: string | undefined, perPt: number): number | undefined {
  if (s === undefined) return undefined;
  const t = s.trim();
  if (!t || t.length > 40) return undefined;
  let split = t.length;
  let inExp = false;
  for (let i = 0; i < t.length; i++) {
    const b = t[i]!;
    if (/[0-9.]/.test(b)) { inExp = false; continue; }
    if ((b === "-" || b === "+") && (i === 0 || inExp)) { inExp = false; continue; }
    if (b === "e" || b === "E") {
      const rest = t.slice(i + 1);
      const after = rest[0] === "+" || rest[0] === "-" ? rest.slice(1) : rest;
      if (/^[0-9]/.test(after)) { inExp = true; continue; }
    }
    split = i;
    break;
  }
  const v = parseF64(t.slice(0, split));
  if (v === undefined || !Number.isFinite(v)) return undefined;
  const units: Record<string, number> = { "": 1 / perPt, pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, pi: 12, px: 0.75, emu: 1 / 12_700 };
  const k = units[t.slice(split).trim().toLowerCase()];
  if (k === undefined) return undefined;
  const pt = v * k;
  return Number.isFinite(pt) ? Math.max(-MAX_LEN_PT, Math.min(MAX_LEN_PT, pt)) : undefined;
}
const truthy = (v: string) => v.trim() === "1" || v.trim() === "true";
const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

// ---- Rust float formatting ----
/** `|a|` with `d` decimals, ties to even on the exact binary value (Rust's `{:.d}`). */
function fixedAbs(a: number, d: number): string {
  if (!Number.isFinite(a)) return a !== a ? "NaN" : "inf";
  if (a >= 1e21) return BigInt(a).toString() + (d > 0 ? "." + "0".repeat(d) : "");
  const up = a.toFixed(d);
  const ext = a.toFixed(Math.min(100, d + 60));
  const cut = ext.indexOf(".") + (d > 0 ? d + 1 : 0);
  const tail = d > 0 ? ext.slice(cut) : ext.slice(cut + 1);
  if (!/^50*$/.test(tail)) return up;
  const down = ext.slice(0, d > 0 ? cut : cut);
  const last = down.replace(".", "").slice(-1);
  return Number(last) % 2 === 0 ? down : up;
}
/** `v` with `d` decimals as Rust prints it (a negative zero keeps its sign). */
function fixed(v: number, d: number): string {
  return (v < 0 || Object.is(v, -0) ? "-" : "") + fixedAbs(Math.abs(v), d);
}
/** Rust's `Display` for f64: the shortest round-trip digits, never an exponent. */
export function rustFloat(v: number): string {
  const s = String(v);
  const m = /^(-?)(\d+)(?:\.(\d+))?e([+-]\d+)$/.exec(s);
  if (!m) return s === "Infinity" ? "inf" : s === "-Infinity" ? "-inf" : s;
  const digits = m[2]! + (m[3] ?? "");
  const point = m[2]!.length + Number(m[4]);
  if (point <= 0) return `${m[1]}0.${"0".repeat(-point)}${digits}`;
  if (point >= digits.length) return m[1] + digits + "0".repeat(point - digits.length);
  return `${m[1]}${digits.slice(0, point)}.${digits.slice(point)}`;
}

// ---- number formats ----
type Tok = { t: "lit"; s: string } | { t: "num"; c: string } | { t: "exp"; plus: boolean };

function sections(code: string): string[] {
  const out: string[] = [];
  let cur = "", quoted = false, escaped = false;
  for (const c of chars(code)) {
    if (escaped) escaped = false;
    else if (c === "\\" && !quoted) escaped = true;
    else if (c === '"') quoted = !quoted;
    else if (c === ";" && !quoted) { out.push(cur); cur = ""; continue; }
    cur += c;
  }
  out.push(cur);
  return out;
}
const firstSection = (code: string) => sections(code)[0] ?? "";
const currency = (inner: string): string | undefined => (inner.startsWith("$") ? inner.slice(1).split("-")[0] ?? "" : undefined);

function formatToks(section: string): Tok[] {
  const toks: Tok[] = [];
  const cs = chars(section);
  let i = 0;
  const takeUntil = (end: string) => { let s = ""; while (i < cs.length) { const q = cs[i++]!; if (q === end) break; s += q; } return s; };
  while (i < cs.length) {
    const c = cs[i++]!;
    if (c === '"') toks.push({ t: "lit", s: takeUntil('"') });
    else if (c === "[") { const sym = currency(takeUntil("]")); if (sym !== undefined) toks.push({ t: "lit", s: sym }); }
    else if (c === "\\") { if (i < cs.length) toks.push({ t: "lit", s: cs[i++]! }); }
    else if (c === "_") { i++; toks.push({ t: "lit", s: " " }); }
    else if (c === "*") i++;
    else if ((c === "E" || c === "e") && (cs[i] === "+" || cs[i] === "-")) toks.push({ t: "exp", plus: cs[i++] === "+" });
    else if (c === "#" || c === "0" || c === "," || c === ".") toks.push({ t: "num", c });
    else toks.push({ t: "lit", s: c });
  }
  return toks;
}

/** `v` in a number format code (`#,##0`, `0.0%`, `0.00E+00`, `"EUR"0.00`...); `General` shows as many decimals as `step` needs. */
export function formatValue(v: number, code: string, step: number): string {
  v = Math.abs(v) < step * 1e-9 ? 0 : v;
  const secs = sections(code);
  const neg = secs[1];
  return neg !== undefined && v < 0 ? formatSection(-v, neg, step) : formatSection(v, secs[0] ?? "", step);
}

function formatSection(v: number, section: string, step: number): string {
  const toks = formatToks(section);
  const isNum = (t: Tok) => t.t !== "lit";
  const lit = (r: Tok[]) => r.map((t) => (t.t === "lit" ? t.s : "")).join("");
  const s = toks.findIndex(isNum);
  const e = toks.length - 1 - [...toks].reverse().findIndex(isNum);
  if (s < 0) return general(v, step);
  const body = toks.slice(s, e + 1);
  const pre = lit(toks.slice(0, s)), suf = lit(toks.slice(e + 1));
  const digitsOf = (r: Tok[]) => r.map((t) => (t.t === "num" ? t.c : "")).join("");
  const k = body.findIndex((t) => t.t === "exp");
  if (k >= 0) {
    const ex = body[k] as { t: "exp"; plus: boolean };
    return pre + scientific(v, decimals(digitsOf(body.slice(0, k))), digitsOf(body.slice(k + 1)).length, ex.plus) + suf;
  }
  const b = digitsOf(body);
  const pct = pre.includes("%") || suf.includes("%");
  if (pct) v *= 100;
  const digits = fixedAbs(Math.abs(v), decimals(b));
  const dot = digits.indexOf(".");
  const ip = dot < 0 ? digits : digits.slice(0, dot);
  const fp = dot < 0 ? "" : digits.slice(dot);
  let grouped = "";
  for (let i = 0; i < ip.length; i++) {
    if (b.includes(",") && i > 0 && (ip.length - i) % 3 === 0) grouped += ",";
    grouped += ip[i];
  }
  const sign = v < 0 && /[^0.]/.test(digits) ? "-" : "";
  return sign + pre + grouped + fp + suf;
}

function decimals(body: string): number {
  const dot = body.indexOf(".");
  return dot < 0 ? 0 : Math.min(12, chars(body.slice(dot + 1)).filter((c) => c === "0" || c === "#").length);
}

function general(v: number, step: number): string {
  let d = 0;
  while (d < 6) {
    const x = step * Math.pow(10, d);
    if (!(Math.abs(x - Math.trunc(x)) > 1e-9)) break;
    d++;
  }
  const s = fixed(v, d);
  return s.replace(/^[-0.]+/, "") === "" ? s.replace(/^-+/, "") : s;
}

function scientific(v: number, dec: number, expDigits: number, plus: boolean): string {
  const s = Math.abs(v).toExponential(dec);
  const [mantissa, e] = s.split("e");
  const exp = v === 0 ? 0 : Number(e ?? 0);
  const sign = exp < 0 ? "-" : plus ? "+" : "";
  return `${v < 0 ? "-" : ""}${mantissa}E${sign}${String(Math.abs(exp)).padStart(clamp(expDigits, 1, 4), "0")}`;
}

// ---- dates ----
interface Stamp { y: number; m: number; d: number; weekday: number; secs: number }

/** The moment of Excel serial day `serial` (1904 system when `date1904`), or undefined out of range. */
function excelStamp(serial: number, date1904: boolean): Stamp | undefined {
  if (date1904) serial += 1462;
  if (!(serial >= 0 && serial < 2_958_466)) return undefined;
  const whole = Math.floor(serial);
  const secs = Math.min(Math.round((serial - whole) * 86_400), 86_399);
  const day = whole;
  const [y, m, d] = day === 0 ? [1900, 1, 0] : day === 60 ? [1900, 2, 29] : day < 60 ? civil(day + 1) : civil(day);
  return { y, m, d, weekday: (((day - 1) % 7) + 7) % 7, secs };
}

/** Calendar date of serial day `n` counted from 30 December 1899. */
function civil(n: number): [number, number, number] {
  const t = new Date(Date.UTC(1899, 11, 30) + n * 86_400_000);
  return [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()];
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** Whether a format code (outside quotes and brackets) has date or time fields. */
function isDateCode(code: string): boolean {
  let quoted = false, bracket = false;
  for (const c of chars(firstSection(code))) {
    if (c === '"') { quoted = !quoted; continue; }
    if (c === "[" && !quoted) { bracket = true; continue; }
    if (c === "]" && bracket) { bracket = false; continue; }
    if (!quoted && !bracket && /[ydmhs]/i.test(c)) return true;
  }
  return false;
}

type DTok = { t: "lit"; s: string } | { t: "field"; c: string; n: number } | { t: "ampm"; long: boolean };

function dateToks(code: string): DTok[] {
  const sec = chars(firstSection(code));
  const toks: DTok[] = [];
  const startsCi = (at: number, pat: string) => chars(pat).every((p, k) => sec[at + k]?.toLowerCase() === p.toLowerCase());
  let i = 0;
  while (i < sec.length) {
    const at = i;
    const c = sec[i++]!;
    if (c === '"') {
      let lit = "";
      while (i < sec.length && sec[i] !== '"') lit += sec[i++];
      i++;
      toks.push({ t: "lit", s: lit });
    } else if (c === "[") {
      let inner = "";
      while (i < sec.length && sec[i] !== "]") inner += sec[i++];
      i++;
      const sym = currency(inner);
      if (sym !== undefined) toks.push({ t: "lit", s: sym });
    } else if (c === "\\" || c === "_" || c === "*") {
      const next = sec[i++];
      if (c === "\\" && next !== undefined) toks.push({ t: "lit", s: next });
      else if (c === "_") toks.push({ t: "lit", s: " " });
    } else if ((c === "a" || c === "A") && startsCi(at, "AM/PM")) {
      i += 4;
      toks.push({ t: "ampm", long: true });
    } else if ((c === "a" || c === "A") && startsCi(at, "A/P")) {
      i += 2;
      toks.push({ t: "ampm", long: false });
    } else if (/[ymdhs]/i.test(c)) {
      let n = 1;
      while (sec[i] !== undefined && sec[i]!.toLowerCase() === c.toLowerCase()) { i++; n++; }
      toks.push({ t: "field", c: c.toLowerCase(), n });
    } else toks.push({ t: "lit", s: c });
  }
  // `m` / `mm` after an hour or before seconds is minutes.
  const fields = toks.map((t, k) => (t.t === "field" ? k : -1)).filter((k) => k >= 0);
  const kind = (k: number | undefined) => { const t = k === undefined ? undefined : toks[k]; return t?.t === "field" ? t.c : undefined; };
  fields.forEach((k, j) => {
    const t = toks[k] as { t: "field"; c: string; n: number };
    const afterHour = j > 0 && kind(fields[j - 1]) === "h";
    const beforeSecs = kind(fields[j + 1]) === "s";
    if (t.c === "m" && t.n <= 2 && (afterHour || beforeSecs)) t.c = "n";
  });
  return toks;
}

function dateText(t: Stamp, code: string): string {
  const toks = dateToks(code);
  const twelve = toks.some((k) => k.t === "ampm");
  const hour = Math.floor(t.secs / 3600);
  let out = "";
  for (const tok of toks) {
    if (tok.t === "lit") out += tok.s;
    else if (tok.t === "ampm") out += hour < 12 ? (tok.long ? "AM" : "A") : tok.long ? "PM" : "P";
    else out += fieldText(t, tok.c, tok.n, twelve);
  }
  return out;
}

function fieldText(t: Stamp, c: string, n: number, twelve: boolean): string {
  const month = MONTHS[clamp(t.m, 1, 12) - 1] ?? "";
  const weekday = WEEKDAYS[t.weekday] ?? "";
  const pad = (v: number) => (n >= 2 ? String(v).padStart(2, "0") : String(v));
  const hour = Math.floor(t.secs / 3600);
  switch (c) {
    case "y": return n <= 2 ? String(((t.y % 100) + 100) % 100).padStart(2, "0") : String(t.y).padStart(4, "0");
    case "d": return n <= 2 ? pad(t.d) : n === 3 ? weekday.slice(0, 3) : weekday;
    case "m": return n <= 2 ? pad(t.m) : n === 3 ? month.slice(0, 3) : n === 4 ? month : month.slice(0, 1);
    case "h": return pad(twelve ? (hour % 12 === 0 ? 12 : hour % 12) : hour);
    case "n": return pad(Math.floor(t.secs / 60) % 60);
    default: return pad(t.secs % 60);
  }
}

/** A number in its format code: a date for date codes (serial days), else as `formatValue`. */
export function categoryText(v: number, code: string, step: number, date1904: boolean): string {
  const t = excelStamp(v, date1904);
  return t && isDateCode(code) ? dateText(t, code) : formatValue(v, code, step);
}

// ---- model ----
type Kind = "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "radar";
type Grouping = "clustered" | "stacked" | "percent";
type Pos = "r" | "l" | "t" | "b";
const KINDS: Record<string, Kind> = { "c:barChart": "bar", "c:lineChart": "line", "c:areaChart": "area", "c:pieChart": "pie", "c:doughnutChart": "doughnut", "c:scatterChart": "scatter", "c:radarChart": "radar" };
const kindOf = (name: string): Kind | undefined => (Object.prototype.hasOwnProperty.call(KINDS, name) ? KINDS[name] : undefined);

interface DataLabel { x: number; y: number; text: string; inside: boolean }
interface Axis { deleted: boolean; grid: boolean; line: boolean; mid: boolean; min?: number; max?: number; fmt?: string }
interface Series {
  name: string;
  idx?: number;
  color?: Rgb;
  lineW: number;
  line: boolean;
  values: (number | null)[];
  xs: (number | null)[];
  points: (Rgb | undefined)[];
  fmt?: string;
  xFmt?: string;
  labels?: boolean;
  raw: (number | null)[];
}
interface Chart {
  kind: Kind;
  horizontal: boolean;
  grouping: Grouping;
  gap: number;
  overlap: number;
  hole: number;
  firstAngle: number;
  radarFill: boolean;
  radarMarks: boolean;
  vary: boolean;
  n: number;
  cats: string[];
  series: Series[];
  accents: Rgb[];
  title?: string;
  legend?: Pos;
  cat: Axis;
  val: Axis;
  textSize: number;
  valFmt: string;
  date1904: boolean;
  plotBorder?: Rgb;
}

const resize = <T>(a: T[], n: number, v: T): T[] => { const out = a.slice(0, n); while (out.length < n) out.push(v); return out; };
const stacked = (c: Chart) => c.grouping !== "clustered";
const accent = (c: Chart, i: number) => c.accents[i % 6] ?? BLACK;
/** Series colour: its own, else the accent its `c:idx` picks, else its place in the file. */
const seriesColour = (c: Chart, j: number) => { const s = c.series[j]; return s?.color ?? accent(c, s?.idx ?? j); };
function pointColour(c: Chart, j: number, i: number): Rgb {
  const p = c.series[j]?.points[i];
  if (p) return p;
  return c.vary ? accent(c, i) : seriesColour(c, j);
}
function legendEntries(c: Chart): [string, Rgb][] {
  if (c.kind === "pie" || c.kind === "doughnut") return c.cats.map((s, i) => [s || String(i + 1), pointColour(c, 0, i)]);
  return c.series.map((s, j) => [s.name || `Series ${j + 1}`, seriesColour(c, j)]);
}

// ---- colours ----
function themeOf(colors: Record<string, string>): Theme {
  const named: Record<string, string> = {};
  const slots = SCHEME.map((k, i) => {
    const v = colors?.[k];
    const c = typeof v === "string" && /^[0-9a-f]{6}$/i.test(v) ? v.toUpperCase() : DEFAULT_THEME[i]!;
    named[k] = c;
    return c;
  });
  return { slots, named };
}
/** The first solid, gradient (first stop) or pattern (foreground) fill of an spPr or a:ln. */
function fillColour(e: Element | undefined, theme: Theme): Rgb | undefined {
  const f = els(e).find((c) => c.tagName === "a:solidFill" || c.tagName === "a:gradFill" || c.tagName === "a:pattFill");
  if (!f) return undefined;
  if (f.tagName === "a:gradFill") return drawingColor(child(child(f, "a:gsLst"), "a:gs"), theme.named);
  if (f.tagName === "a:pattFill") return drawingColor(child(f, "a:fgClr"), theme.named);
  return drawingColor(f, theme.named);
}

// ---- parsing ----
const label = (s: string) => takeChars(s, MAX_LABEL);

function parseChart(space: Element, theme: Theme): Chart | undefined {
  const chart = child(space, "c:chart");
  const plot = child(chart, "c:plotArea");
  const ty = els(plot).find((e) => kindOf(e.tagName));
  if (!chart || !plot || !ty) return undefined;
  const kind = kindOf(ty.tagName)!;
  const scatter = kind === "scatter";
  const d1904 = child(space, "c:date1904");
  const date1904 = !!d1904 && (attr(d1904, "val") === undefined || truthy(attr(d1904, "val")!));
  let series = readSeries(ty, scatter, theme);
  if (!series.length) return undefined;
  const catEl = children(ty, "c:ser").map((s) => child(s, "c:cat")).find((c) => c);
  let cats = catEl ? labels(catEl, MAX_POINTS, date1904) : [];
  const longest = Math.max(0, ...series.map((s) => s.values.length));
  const n = kind === "radar" ? Math.min(Math.max(longest, cats.length), MAX_SPOKES) : kind === "pie" || kind === "doughnut" ? Math.max(series[0]!.values.length, cats.length) : Math.max(longest, cats.length);
  // Every series is padded to n points, so keep only as many as the cell budget holds.
  if (n > 0) series = series.slice(0, Math.max(1, Math.floor(MAX_CELLS / n)));
  if (kind === "radar") series = series.slice(0, MAX_RADAR_SERIES);
  const g = valOf(ty, "c:grouping");
  const groupable = kind === "bar" || kind === "area" || kind === "line";
  const grouping: Grouping = g === "stacked" && groupable ? "stacked" : g === "percentStacked" && groupable ? "percent" : "clustered";
  cats = resize(cats, n, "");
  const radarStyle = valOf(ty, "c:radarStyle");
  const typeLabels = showValues(child(ty, "c:dLbls"));
  for (const s of series) {
    s.values = resize(s.values, n, null);
    s.xs = resize(s.xs, n, null);
    s.labels = s.labels ?? typeLabels ?? false;
    if (s.labels) s.raw = s.values.slice();
  }
  if (grouping === "percent") toPercent(series, n);
  const [cat, val] = chartAxes(plot, scatter);
  const valFmt = grouping === "percent" ? "0%" : val.fmt ?? series.find((s) => s.fmt !== undefined)?.fmt ?? "General";
  const num = (name: string) => int(valOf(ty, name));
  const vary = valOf(ty, "c:varyColors");
  return {
    kind,
    horizontal: kind === "bar" && valOf(ty, "c:barDir") === "bar",
    grouping,
    gap: clamp(num("c:gapWidth") ?? 150, 0, 500) / 100,
    overlap: clamp(num("c:overlap") ?? 0, -100, 100) / 100,
    hole: clamp(num("c:holeSize") ?? 50, 10, 90) / 100,
    firstAngle: clamp(num("c:firstSliceAng") ?? 0, 0, 360),
    radarFill: radarStyle === "filled",
    radarMarks: radarStyle === "marker",
    vary: (kind === "pie" || kind === "doughnut") && (vary === undefined || truthy(vary)),
    n,
    cats,
    title: titleOf(chart, series),
    series,
    accents: theme.slots.slice(4, 10),
    legend: legendPos(chart),
    cat,
    val,
    textSize: textSize(space),
    valFmt,
    date1904,
    plotBorder: fillColour(child(child(plot, "c:spPr"), "a:ln"), theme),
  };
}

function readSeries(ty: Element, scatter: boolean, theme: Theme): Series[] {
  let budget = MAX_CELLS;
  const out: Series[] = [];
  for (const ser of children(ty, "c:ser").slice(0, MAX_SERIES)) {
    const s = parseSeries(ser, scatter, theme, Math.min(budget, MAX_POINTS));
    budget = Math.max(0, budget - Math.max(s.values.length, s.xs.length));
    out.push(s);
  }
  return out;
}

function toPercent(series: Series[], n: number) {
  for (let i = 0; i < n; i++) {
    const total = series.reduce((t, s) => t + Math.abs(s.values[i] ?? 0), 0);
    for (const s of series) {
      const v = s.values[i];
      if (v !== null && v !== undefined) s.values[i] = total > 0 ? v / total : 0;
    }
  }
}

function chartAxes(plot: Element, scatter: boolean): [Axis, Axis] {
  const vals = children(plot, "c:valAx");
  const horiz = (a: Element) => { const p = valOf(a, "c:axPos"); return p === "b" || p === "t"; };
  if (scatter) {
    const x = vals.find(horiz) ?? vals[0];
    const y = vals.find((a) => !horiz(a)) ?? vals.find((a) => a !== x);
    return [axis(x), axis(y)];
  }
  return [axis(els(plot).find((e) => e.tagName === "c:catAx" || e.tagName === "c:dateAx")), axis(vals[0])];
}

/** The title; an empty c:title shows a lone series' name unless the automatic title is deleted. */
function titleOf(chart: Element, series: Series[]): string | undefined {
  const del = valOf(chart, "c:autoTitleDeleted");
  const auto = !(del !== undefined && truthy(del)) && series.length === 1 && series[0]!.name ? series[0]!.name : undefined;
  const t = child(chart, "c:title");
  return t ? titleText(t) ?? auto : undefined;
}

function showValues(d: Element | undefined): boolean | undefined {
  if (!d) return undefined;
  const del = valOf(d, "c:delete");
  if (del !== undefined && truthy(del)) return false;
  const v = child(d, "c:showVal");
  if (!v) return undefined;
  const a = attr(v, "val");
  return a === undefined || truthy(a);
}

function legendPos(chart: Element): Pos | undefined {
  const l = child(chart, "c:legend");
  if (!l) return undefined;
  const p = valOf(l, "c:legendPos") ?? "r";
  return p === "l" ? "l" : p === "t" || p === "tr" ? "t" : p === "b" ? "b" : "r";
}

function textSize(space: Element): number {
  const sz = int(attr(find(child(space, "c:txPr"), "a:defRPr"), "sz"));
  return sz === undefined ? FONT : clamp(sz / 100, 6, 24);
}

function parseSeries(ser: Element, scatter: boolean, theme: Theme, cap: number): Series {
  const v = find(child(ser, "c:tx"), "c:v");
  const name = v ? label(ownText(v).trim()) : "";
  const idx = parseUint(valOf(ser, "c:idx")?.trim(), 4_294_967_295);
  const sp = child(ser, "c:spPr");
  const ln = child(sp, "a:ln");
  const marker = fillColour(child(child(ser, "c:marker"), "c:spPr"), theme);
  const color = fillColour(sp, theme) ?? fillColour(ln, theme) ?? marker;
  const w = measure(attr(ln, "w"), 12_700);
  const lineW = w === undefined ? LINE_W : clamp(w, 0, 20);
  const line = !ln || !child(ln, "a:noFill");
  const [values, fmt] = numbers(child(ser, scatter ? "c:yVal" : "c:val"), cap);
  let [xs, xFmt] = scatter ? numbers(child(ser, "c:xVal"), cap) : [[] as (number | null)[], undefined];
  if (scatter && !xs.length) xs = values.map((_, i) => i + 1);
  const points: (Rgb | undefined)[] = new Array(values.length).fill(undefined);
  for (const dpt of children(ser, "c:dPt")) {
    const i = int(valOf(dpt, "c:idx"));
    if (i !== undefined && i >= 0 && i < points.length) points[i] = fillColour(child(dpt, "c:spPr"), theme);
  }
  return { name, idx, color, lineW, line, values, xs, points, fmt, xFmt, labels: showValues(child(ser, "c:dLbls")), raw: [] };
}

/** Category labels of a c:cat (string or number cache), capped; numbers in their format code. */
function labels(el: Element, cap: number, date1904: boolean): string[] {
  const cache = find(el, "c:strCache") ?? find(el, "c:numCache") ?? find(el, "c:strLit") ?? find(el, "c:numLit");
  if (!cache) return [];
  const f = child(cache, "c:formatCode");
  const raw = f ? ownText(f) : undefined;
  const code = raw && utf8Len(raw) <= MAX_FORMAT_CODE && raw.toLowerCase() !== "general" ? raw : undefined;
  return pts(cache, cap).map((p) => {
    const text = p ?? "";
    const v = parseF64(text.trim());
    return code !== undefined && v !== undefined && Number.isFinite(v) ? label(categoryText(v, code, 1, date1904)) : label(text);
  });
}

/** Values of a c:val-style element (number cache or literal), capped, with its format code. */
function numbers(el: Element | undefined, cap: number): [(number | null)[], string | undefined] {
  const cache = el ? find(el, "c:numCache") ?? find(el, "c:numLit") : undefined;
  if (!cache) return [[], undefined];
  const f = child(cache, "c:formatCode");
  const raw = f ? ownText(f) : undefined;
  const fmt = raw && utf8Len(raw) <= MAX_FORMAT_CODE ? raw : undefined;
  const vals = pts(cache, cap).map((p) => {
    const v = p === undefined ? undefined : parseF64(p.trim());
    return v !== undefined && Number.isFinite(v) ? v : null;
  });
  return [vals, fmt];
}

/** The c:pt texts of a cache by index; indexes and counts from the file are capped. */
function pts(cache: Element, cap: number): (string | undefined)[] {
  const declared = clamp(int(valOf(cache, "c:ptCount")) ?? 0, 0, cap);
  const out: (string | undefined)[] = new Array(declared).fill(undefined);
  for (const p of children(cache, "c:pt")) {
    const i = int(attr(p, "idx"));
    if (i === undefined || i < 0 || i >= cap) continue;
    while (out.length <= i) out.push(undefined);
    const v = child(p, "c:v");
    out[i] = v ? ownText(v) : undefined;
  }
  return out;
}

function axis(a: Element | undefined): Axis {
  if (!a) return { deleted: true, grid: false, line: false, mid: false };
  const scaling = child(a, "c:scaling");
  const num = (n: string) => { const v = parseF64(valOf(scaling, n)?.trim()); return v !== undefined && Number.isFinite(v) ? v : undefined; };
  const del = valOf(a, "c:delete");
  const nf = child(a, "c:numFmt");
  const fc = nf && attr(nf, "sourceLinked") !== "1" ? attr(nf, "formatCode") : undefined;
  return {
    deleted: del !== undefined && truthy(del),
    grid: !!child(a, "c:majorGridlines"),
    line: !child(child(child(a, "c:spPr"), "a:ln"), "a:noFill"),
    mid: valOf(a, "c:crossBetween") === "midCat",
    min: num("c:min"),
    max: num("c:max"),
    fmt: fc !== undefined && utf8Len(fc) <= MAX_FORMAT_CODE ? fc : undefined,
  };
}

function titleText(t: Element): string | undefined {
  const tx = child(t, "c:tx");
  if (!tx) return undefined;
  const rich = child(tx, "c:rich");
  let s = "";
  if (rich) s = richText(rich, 0);
  else { const v = find(tx, "c:v"); s = v ? ownText(v) : ""; }
  s = label(s.trim());
  return s || undefined;
}

/** The a:t runs under an element, in order. */
function richText(e: Element, depth: number): string {
  if (depth > MAX_DEPTH) return "";
  return els(e).map((c) => (c.tagName === "a:t" ? ownText(c) : richText(c, depth + 1))).join("");
}

// ---- canvas ----
interface Rect { x: number; y: number; w: number; h: number }
const right = (r: Rect) => r.x + r.w;
const bottom = (r: Rect) => r.y + r.h;
const arr = (r: Rect) => [r.x, r.y, r.w, r.h];

/** Estimated text width, points: CHAR_W em a character. */
export const textW = (s: string, size: number) => chars(s).length * size * CHAR_W;

/** `s` cut short with an ellipsis to fit `width` points by the estimate. */
function clipText(s: string, width: number, size: number): string {
  const f = Math.floor(width / (size * CHAR_W));
  const fit = Number.isFinite(f) ? Math.max(0, f) : 0;
  const cs = chars(s);
  if (cs.length <= fit) return s;
  return cs.slice(0, Math.max(0, fit - 1)).join("") + "…";
}

const widest = (items: string[], size: number) => items.reduce((m, s) => Math.max(m, textW(s, size)), 0);

function text(out: Item[], r: Rect, s: string, size: number, align: Align, bold: boolean) {
  out.push({ k: "text", rect: arr(r), text: s, size, color: TEXT, bold, align });
}
/** Text centred on `cx` with its top at `y`. */
function textAt(out: Item[], cx: number, y: number, s: string, size: number) {
  const w = textW(s, size) + 4;
  text(out, { x: cx - w / 2, y, w, h: size * LINE }, s, size, "center", false);
}
function fill(out: Item[], segs: Seg[], color: Rgb) {
  if (segs.length) out.push({ k: "path", segs, fill: color, sw: 0 });
}
function stroke(out: Item[], segs: Seg[], color: Rgb, width: number) {
  if (segs.length) out.push({ k: "path", segs, stroke: color, sw: width });
}
function swatch(out: Item[], x: number, y: number, color: Rgb) {
  out.push({ k: "shape", rect: [x, y, SWATCH, SWATCH], ellipse: false, fill: color, sw: 0 });
}
const seg = (x0: number, y0: number, x1: number, y1: number): Seg[] => [[0, x0, y0], [1, x1, y1]];
const rectSegs = (x0: number, y0: number, x1: number, y1: number): Seg[] => [[0, x0, y0], [1, x1, y0], [1, x1, y1], [1, x0, y1], [3]];

/** Polyline pieces: a gap (undefined) starts a new sub-path. */
function runs(points: ([number, number] | undefined)[]): Seg[] {
  const segs: Seg[] = [];
  let open = false;
  for (const p of points) {
    if (p) { segs.push(open ? [1, p[0], p[1]] : [0, p[0], p[1]]); open = true; } else open = false;
  }
  return segs;
}

/** Words of `s` in lines of at most `width` points (at most three lines). */
function wrap(s: string, width: number, size: number): string[] {
  const lines: string[] = [];
  let cur = "";
  for (const word of s.split(/\s+/).filter(Boolean)) {
    const joined = cur ? `${cur} ${word}` : word;
    if (!cur || textW(joined, size) <= width) cur = joined;
    else { lines.push(cur); cur = word; }
  }
  if (cur) lines.push(cur);
  if (!lines.length) lines.push("");
  return lines.slice(0, 3);
}

/** Legend on the given side of `area`; returns the area left for the plot. */
function legend(out: Item[], entries: [string, Rgb][], pos: Pos, area: Rect, fs: number): Rect {
  if (!entries.length) return area;
  // Entry texts are cut to a share of the chart, so a long name can't push the legend out.
  const cut = entries.map(([s, c]): [string, Rgb] => [clipText(s, area.w * LEGEND_SHARE, fs), c]);
  return pos === "r" || pos === "l" ? sideLegend(out, cut, pos === "l", area, fs) : rowLegend(out, cut, pos === "t", area, fs);
}

function sideLegend(out: Item[], all: [string, Rgb][], left: boolean, area: Rect, fs: number): Rect {
  const lh = fs * LINE;
  const fitN = Math.floor(area.h / lh);
  const entries = all.slice(0, Math.min(Number.isFinite(fitN) ? Math.max(0, fitN) : 0, all.length));
  if (!entries.length) return area;
  const tw = entries.reduce((m, [s]) => Math.max(m, textW(s, fs)), 0);
  const blockW = SWATCH + GAP + tw;
  const blockH = lh * entries.length;
  const x = left ? area.x : right(area) - blockW;
  const y0 = area.y + Math.max(0, (area.h - blockH) / 2);
  entries.forEach(([s, c], k) => {
    const y = y0 + lh * k;
    swatch(out, x, y + (lh - SWATCH) / 2, c);
    text(out, { x: x + SWATCH + GAP, y, w: tw + 2, h: lh }, s, fs, "left", false);
  });
  const used = blockW + PAD;
  return { ...area, x: left ? area.x + used : area.x, w: Math.max(1, area.w - used) };
}

function rowLegend(out: Item[], entries: [string, Rgb][], top: boolean, area: Rect, fs: number): Rect {
  const lh = fs * LINE;
  const r = Math.floor(area.h / 2 / lh);
  const maxRows = Math.min(MAX_LEGEND_ROWS, Number.isFinite(r) ? Math.max(1, r) : 1);
  const rows = legendRows(entries, area.w, fs, maxRows);
  const y0 = top ? area.y : bottom(area) - lh * rows.length;
  rows.forEach((row, ri) => {
    const total = row.reduce((t, [, , w]) => t + SWATCH + GAP + w, 0) + SEP * Math.max(0, row.length - 1);
    let x = area.x + Math.max(0, (area.w - total) / 2);
    const y = y0 + lh * ri;
    for (const [s, c, w] of row) {
      swatch(out, x, y + (lh - SWATCH) / 2, c);
      text(out, { x: x + SWATCH + GAP, y, w: w + 2, h: lh }, s, fs, "left", false);
      x += SWATCH + GAP + w + SEP;
    }
  });
  const used = lh * rows.length + PAD;
  return { ...area, y: top ? area.y + used : area.y, h: Math.max(1, area.h - used) };
}

/** Entries packed into at most `maxRows` rows of `width` points; the rest are left out. */
function legendRows(entries: [string, Rgb][], width: number, fs: number, maxRows: number): [string, Rgb, number][][] {
  const rows: [string, Rgb, number][][] = [];
  let rowW = 0;
  for (const [s, c] of entries) {
    const w = textW(s, fs);
    const entryW = SWATCH + GAP + w;
    if (rows.length && rowW + SEP + entryW <= width) rowW += SEP + entryW;
    else if (rows.length < maxRows) { rows.push([]); rowW = entryW; } else break;
    rows[rows.length - 1]!.push([s, c, w]);
  }
  return rows;
}

// ---- cartesian plots ----
/** Chart area: the chart space's fill and border (white and borderless by default). */
function backdrop(out: Item[], space: Element, theme: Theme, w: number, h: number) {
  const sp = child(space, "c:spPr");
  const f = sp ? (child(sp, "a:noFill") ? undefined : fillColour(sp, theme) ?? WHITE) : WHITE;
  const ln = child(sp, "a:ln");
  const s = fillColour(ln, theme);
  const sw = clamp(measure(attr(ln, "w"), 12_700) ?? 0.75, 0, 20);
  if (f || s) out.push({ k: "shape", rect: [0, 0, w, h], ellipse: false, fill: f, stroke: s, sw });
}

interface Scale { lo: number; hi: number; step: number }

function ticksOf(s: Scale): number[] {
  const raw = Math.floor((s.hi - s.lo) / s.step + 1e-9);
  const n = Math.min(Number.isFinite(raw) ? Math.max(0, raw) : 0, MAX_TICKS);
  const out: number[] = [];
  for (let k = 0; k <= n; k++) { const t = s.lo + k * s.step; if (t <= s.hi + 1e-9) out.push(t); }
  return out;
}

/** A scale for data dlo..dhi as Excel picks it: 5% headroom away from zero, steps of 1, 2, 2.5, 5 or 10 times a power of ten. */
export function niceScale(dlo: number, dhi: number, zero: boolean, min?: number, max?: number): Scale {
  if (!(dhi > dlo)) dhi = dlo + 1;
  const span = dhi - dlo;
  const [endsLo, endsHi] = zero ? [Math.min(dlo, 0), Math.max(dhi, 0)] : [dlo, dhi];
  const lo = min ?? (endsLo < 0 ? endsLo - 0.05 * span : endsLo);
  const hi = max ?? (endsHi > 0 ? endsHi + 0.05 * span : endsHi);
  const fallback = { lo: 0, hi: 1, step: 0.2 };
  if (!(hi > lo && Number.isFinite(hi) && Number.isFinite(lo))) return fallback;
  const raw = (hi - lo) / 7.5;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  const loR = (min !== undefined ? lo : Math.floor(lo / step + 1e-9) * step) + 0;
  const hiR = (max !== undefined ? hi : Math.ceil(hi / step - 1e-9) * step) + 0;
  if (!(Number.isFinite(step) && step > 0 && Number.isFinite(loR) && Number.isFinite(hiR) && hiR > loR)) return fallback;
  return { lo: loR, hi: hiR, step };
}

/** The plot rectangle and the scales that map values and categories onto it. */
class Frame {
  constructor(readonly plot: Rect, readonly horiz: boolean, readonly n: number, readonly vs: Scale, readonly xs: Scale | undefined, readonly mid: boolean) {}
  band() { return (this.horiz ? this.plot.h : this.plot.w) / Math.max(1, this.n); }
  bandStart(i: number) { return this.horiz ? bottom(this.plot) - (i + 1) * this.band() : this.plot.x + i * this.band(); }
  val(v: number) {
    const t = clamp((v - this.vs.lo) / (this.vs.hi - this.vs.lo), 0, 1);
    return this.horiz ? this.plot.x + t * this.plot.w : bottom(this.plot) - t * this.plot.h;
  }
  catX(i: number) {
    if (!this.mid) return this.plot.x + (i + 0.5) * this.band();
    return this.n <= 1 ? this.plot.x + this.plot.w / 2 : this.plot.x + (this.plot.w * i) / (this.n - 1);
  }
  scatX(x: number) {
    return this.xs ? this.plot.x + clamp((x - this.xs.lo) / (this.xs.hi - this.xs.lo), 0, 1) * this.plot.w : this.plot.x;
  }
}

interface Axes { vs: Scale; vticks: [number, string][]; xs?: Scale; xticks: [number, string][] }

function axesOf(c: Chart): Axes {
  const [dlo, dhi] = extent(c);
  const [vmin, vmax] = c.grouping === "percent" ? [0, 1] : [c.val.min, c.val.max];
  // Lines and scatters don't need zero for a narrow range.
  const zero = !(c.kind === "line" || c.kind === "scatter") || !narrow(dlo, dhi);
  const vs = niceScale(dlo, dhi, zero, vmin, vmax);
  const vticks = ticksOf(vs).map((t): [number, string] => [t, formatValue(t, c.valFmt, vs.step)]);
  let xs: Scale | undefined;
  let xticks: [number, string][] = [];
  if (c.kind === "scatter") {
    const [xl, xh] = xExtent(c);
    xs = niceScale(xl, xh, false, c.cat.min, c.cat.max);
    const xfmt = c.cat.fmt ?? c.series.find((s) => s.xFmt !== undefined)?.xFmt ?? "General";
    const s = xs;
    xticks = ticksOf(s).map((t): [number, string] => [t, categoryText(t, xfmt, s.step, c.date1904)]);
  }
  return { vs, vticks, xs, xticks };
}

interface Margins { left: number; top: number; bottom: number; wrapped: string[][]; skip: number }

function axisMargins(c: Chart, ax: Axes, area: Rect): Margins {
  const fs = c.textSize;
  const lh = fs * LINE;
  const scatter = c.kind === "scatter";
  const catOn = !c.cat.deleted, valOn = !c.val.deleted;
  const left = c.horizontal ? (catOn ? Math.min(widest(c.cats, fs), area.w * 0.4) + GAP : 0) : valOn ? ax.vticks.reduce((m, [, s]) => Math.max(m, textW(s, fs)), 0) + GAP : 0;
  const bandW = Math.max(1, area.w - left) / c.n;
  const wrapped = catOn && !scatter ? c.cats.map((s) => wrap(s, c.horizontal ? left - GAP : bandW, fs)) : [];
  const widestLine = wrapped.flat().reduce((m, s) => Math.max(m, textW(s, fs)), 0);
  const sk = Math.ceil(widestLine / bandW);
  const skip = Math.max(1, Number.isFinite(sk) ? sk : 0);
  const labelLines = wrapped.length ? Math.max(...wrapped.map((l) => l.length)) : 1;
  const bot = catOn && !c.horizontal ? (scatter ? lh + GAP : labelLines * lh + GAP) : c.horizontal && valOn ? lh + GAP : 0;
  const top = !c.horizontal && valOn ? lh / 2 : 0;
  return { left, top, bottom: bot, wrapped, skip };
}

/** A bar, column, line, area or scatter chart in `area`: gridlines, series, axes, then labels. */
function cartesian(out: Item[], c: Chart, area: Rect) {
  if (c.n === 0) return;
  const ax = axesOf(c);
  const m = axisMargins(c, ax, area);
  const plot = { x: area.x + m.left, y: area.y + m.top, w: area.w - m.left, h: area.h - m.top - m.bottom };
  if (!(plot.w >= 1 && plot.h >= 1)) return;
  const g = new Frame(plot, c.horizontal, c.n, ax.vs, ax.xs, c.val.mid);
  if (c.val.grid) drawGrid(out, g, ax.vticks);
  const lbls: DataLabel[] = [];
  if (c.kind === "bar") bars(out, c, g, lbls);
  else if (c.kind === "line") lines(out, c, g, lbls);
  else if (c.kind === "area") areas(out, c, g, lbls);
  else if (c.kind === "scatter") drawScatter(out, c, g, lbls);
  drawAxes(out, c, g);
  drawDataLabels(out, lbls, c.textSize, g.horiz && c.kind === "bar");
  if (!c.val.deleted) drawValueLabels(out, c, g, ax.vticks, { ...area, w: m.left });
  if (!c.cat.deleted) drawCategoryLabels(out, c, g, ax.xticks, m, area);
}

function drawGrid(out: Item[], g: Frame, ticks: [number, string][]) {
  const p = g.plot;
  for (const [t] of ticks) {
    const v = g.val(t);
    stroke(out, g.horiz ? seg(v, p.y, v, bottom(p)) : seg(p.x, v, right(p), v), GRID, 0.75);
  }
}

function drawAxes(out: Item[], c: Chart, g: Frame) {
  const p = g.plot;
  if (!c.cat.deleted && c.cat.line) stroke(out, g.horiz ? seg(p.x, p.y, p.x, bottom(p)) : seg(p.x, bottom(p), right(p), bottom(p)), AXIS, 0.75);
  if (c.plotBorder) stroke(out, rectSegs(p.x, p.y, right(p), bottom(p)), c.plotBorder, 0.75);
  if (!c.val.deleted && c.val.line) stroke(out, g.horiz ? seg(p.x, bottom(p), right(p), bottom(p)) : seg(p.x, p.y, p.x, bottom(p)), AXIS, 0.75);
}

function drawValueLabels(out: Item[], c: Chart, g: Frame, ticks: [number, string][], margin: Rect) {
  const fs = c.textSize, lh = fs * LINE;
  for (const [t, s] of ticks) {
    const p = g.val(t);
    if (g.horiz) textAt(out, p, bottom(g.plot) + GAP / 2, s, fs);
    else text(out, { x: margin.x, y: p - lh / 2, w: margin.w - GAP, h: lh }, s, fs, "right", false);
  }
}

function drawCategoryLabels(out: Item[], c: Chart, g: Frame, xticks: [number, string][], m: Margins, area: Rect) {
  const fs = c.textSize, lh = fs * LINE;
  const plot = g.plot;
  if (c.kind === "scatter") {
    for (const [t, s] of xticks) textAt(out, g.scatX(t), bottom(plot) + GAP / 2, s, fs);
  } else if (g.horiz) {
    m.wrapped.forEach((ls, i) => {
      const top = g.bandStart(i) + g.band() / 2 - (ls.length * lh) / 2;
      ls.forEach((s, k) => text(out, { x: area.x, y: top + k * lh, w: m.left - GAP, h: lh }, s, fs, "right", false));
    });
  } else {
    const band = g.band();
    m.wrapped.forEach((ls, i) => {
      if (i % m.skip !== 0) return;
      ls.forEach((s, k) => text(out, { x: g.catX(i) - band / 2, y: bottom(plot) + GAP / 2 + k * lh, w: band, h: lh }, s, fs, "center", false));
    });
  }
}

/** Whether data lo..hi is a narrow band away from zero (span under a sixth of the far end). */
export function narrow(lo: number, hi: number): boolean {
  if (lo > 0) return (hi - lo) / hi < 1 / 6;
  if (hi < 0) return (hi - lo) / Math.abs(lo) < 1 / 6;
  return false;
}

function extent(c: Chart): [number, number] {
  const ps: number[] = [];
  if (stacked(c)) {
    for (let i = 0; i < c.n; i++) {
      let p = 0, q = 0;
      for (const s of c.series) { const v = s.values[i]; if (v === null || v === undefined) continue; if (v >= 0) p += v; else q += v; }
      ps.push(p, q);
    }
  } else for (const s of c.series) for (const v of s.values) if (v !== null) ps.push(v);
  let lo = Infinity, hi = -Infinity;
  for (const v of ps) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
  return Number.isFinite(lo) && Number.isFinite(hi) ? [lo, hi] : [0, 0];
}

function xExtent(c: Chart): [number, number] {
  let lo = Infinity, hi = -Infinity;
  for (const s of c.series) for (const v of s.xs) if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
  return Number.isFinite(lo) && Number.isFinite(hi) ? [lo, hi] : [0, 1];
}

const hasValues = (s: Series) => s.values.some((v) => v !== null);

function bars(out: Item[], c: Chart, g: Frame, lbls: DataLabel[]) {
  const st = stacked(c);
  const slots = st ? 1 : Math.max(1, c.series.length);
  const ov = st ? 0 : c.overlap;
  const band = g.band();
  const thick = band / (slots - (slots - 1) * ov + c.gap);
  const group = thick * (slots - (slots - 1) * ov);
  const pos = new Array<number>(g.n).fill(0), neg = new Array<number>(g.n).fill(0);
  c.series.forEach((s, j) => {
    if (!hasValues(s)) return;
    const sl0 = st ? 0 : j;
    const slot = g.horiz ? Math.max(0, slots - 1 - sl0) : sl0;
    for (let i = 0; i < g.n; i++) {
      const v = s.values[i];
      if (v === null || v === undefined) continue;
      let v0 = 0, v1 = v;
      if (st && v >= 0) { v0 = pos[i]!; pos[i]! += v; v1 = pos[i]!; } else if (st) { v0 = neg[i]!; neg[i]! += v; v1 = neg[i]!; }
      const a = g.bandStart(i) + (band - group) / 2 + slot * thick * (1 - ov);
      const q0 = g.val(v0), q1 = g.val(v1);
      const p0 = Math.min(q0, q1), p1 = Math.max(q0, q1);
      fill(out, g.horiz ? rectSegs(p0, a, p1, a + thick) : rectSegs(a, p0, a + thick, p1), pointColour(c, j, i));
      const t = labelOf(s, i);
      if (t !== undefined) {
        const mid = a + thick / 2;
        // Beyond the bar's end when clustered: right of a positive bar, above a positive column.
        const [x, y] = g.horiz ? (st ? [(p0 + p1) / 2, mid] : [v >= 0 ? p1 : p0, mid]) : st ? [mid, (p0 + p1) / 2] : [mid, v >= 0 ? p0 : p1];
        lbls.push({ x, y, text: t, inside: st });
      }
    }
  });
}

/** Line series; stacked ones are drawn at the running total of the series up to them. */
function lines(out: Item[], c: Chart, g: Frame, lbls: DataLabel[]) {
  const total = new Array<number>(g.n).fill(0);
  c.series.forEach((s, j) => {
    if (!hasValues(s)) return;
    const p: ([number, number] | undefined)[] = [];
    for (let i = 0; i < g.n; i++) {
      const v = s.values[i];
      if (v === null || v === undefined) { p.push(undefined); continue; }
      let y = v;
      if (stacked(c)) { total[i]! += v; y = total[i]!; }
      p.push([g.catX(i), g.val(y)]);
    }
    p.forEach((q, i) => { const t = labelOf(s, i); if (q && t !== undefined) lbls.push({ x: q[0], y: q[1], text: t, inside: false }); });
    stroke(out, runs(p), seriesColour(c, j), s.lineW);
  });
}

function areas(out: Item[], c: Chart, g: Frame, lbls: DataLabel[]) {
  let base = new Array<number>(g.n).fill(0);
  c.series.forEach((s, j) => {
    if (!hasValues(s)) return;
    const top = base.map((b, i) => b + (s.values[i] ?? 0));
    const segs: Seg[] = top.map((t, i): Seg => (i === 0 ? [0, g.catX(i), g.val(t)] : [1, g.catX(i), g.val(t)]));
    for (let i = base.length - 1; i >= 0; i--) segs.push([1, g.catX(i), g.val(base[i]!)]);
    segs.push([3]);
    fill(out, segs, seriesColour(c, j));
    top.forEach((t, i) => { const l = labelOf(s, i); if (l !== undefined) lbls.push({ x: g.catX(i), y: g.val((t + base[i]!) / 2), text: l, inside: true }); });
    if (stacked(c)) base = top;
  });
}

/** Scatter series as lines, or as small markers when the series has no line. */
function drawScatter(out: Item[], c: Chart, g: Frame, lbls: DataLabel[]) {
  const MARKER = 6;
  c.series.forEach((s, j) => {
    const n = Math.min(s.xs.length, s.values.length);
    const p: ([number, number] | undefined)[] = [];
    for (let i = 0; i < n; i++) {
      const x = s.xs[i], y = s.values[i];
      p.push(x !== null && x !== undefined && y !== null && y !== undefined ? [g.scatX(x), g.val(y)] : undefined);
    }
    p.forEach((q, i) => { const t = labelOf(s, i); if (q && t !== undefined) lbls.push({ x: q[0], y: q[1] - MARKER / 2, text: t, inside: false }); });
    if (s.line) { stroke(out, runs(p), seriesColour(c, j), s.lineW); return; }
    for (const q of p) if (q) out.push({ k: "shape", rect: [q[0] - MARKER / 2, q[1] - MARKER / 2, MARKER, MARKER], ellipse: true, fill: seriesColour(c, j), sw: 0 });
  });
}

// ---- polar plots ----
function pie(out: Item[], c: Chart, area: Rect) {
  const s = c.series[0];
  if (!s) return;
  const vals = s.values.map((v) => Math.max(v ?? 0, 0));
  const total = vals.reduce((t, v) => t + v, 0);
  const r = Math.min(area.w, area.h) / 2 - 2;
  if (!(total > 0 && Number.isFinite(total) && r >= 1)) return;
  const cx = area.x + area.w / 2, cy = area.y + area.h / 2;
  const inner = c.kind === "doughnut" ? r * c.hole : 0;
  let a = (c.firstAngle * Math.PI) / 180;
  const lbls: DataLabel[] = [];
  vals.forEach((v, i) => {
    if (v <= 0) return;
    const b = a + (v / total) * Math.PI * 2;
    const segs: Seg[] = [];
    if (inner > 0) {
      segs.push([0, ...polar(cx, cy, r, a)]);
      arc(segs, cx, cy, r, a, b);
      segs.push([1, ...polar(cx, cy, inner, b)]);
      arc(segs, cx, cy, inner, b, a);
    } else {
      segs.push([0, cx, cy], [1, ...polar(cx, cy, r, a)]);
      arc(segs, cx, cy, r, a, b);
    }
    segs.push([3]);
    fill(out, segs, pointColour(c, 0, i));
    const t = labelOf(s, i);
    if (t !== undefined) {
      // Halfway across the ring (pie: two thirds out from the centre).
      const [x, y] = polar(cx, cy, inner > 0 ? (r + inner) / 2 : r * 0.65, (a + b) / 2);
      lbls.push({ x, y, text: t, inside: true });
    }
    a = b;
  });
  drawDataLabels(out, lbls, c.textSize, false);
}

/** Radar: spokes, grid polygons at the value ticks, a polygon per series and the category labels. */
function radar(out: Item[], c: Chart, area: Rect) {
  const MARKER = 5;
  if (c.n === 0) return;
  const fs = c.textSize, lh = fs * LINE;
  const labelled = c.cats.some((s) => s !== "");
  const side = labelled ? Math.min(widest(c.cats, fs), area.w * 0.25) + GAP + 2 : 0;
  const vert = labelled ? lh + GAP : 0;
  const r = Math.min(area.w / 2 - side, area.h / 2 - vert) - 2;
  if (Number.isNaN(r) || r < 8) return;
  const cx = area.x + area.w / 2, cy = area.y + area.h / 2;
  const vs = axesOf(c).vs;
  const span = vs.hi - vs.lo;
  const angle = (i: number) => (i / c.n) * Math.PI * 2;
  const at = (v: number, i: number) => polar(cx, cy, r * (span > 0 ? clamp((v - vs.lo) / span, 0, 1) : 0), angle(i));
  const ticks = ticksOf(vs);
  if (c.val.grid) {
    for (const t of ticks.slice(1)) {
      const segs: Seg[] = [];
      for (let i = 0; i < c.n; i++) segs.push([i === 0 ? 0 : 1, ...at(t, i)]);
      segs.push([3]);
      stroke(out, segs, GRID, 0.75);
    }
  }
  for (let i = 0; i < c.n; i++) stroke(out, [[0, cx, cy], [1, ...polar(cx, cy, r, angle(i))]], GRID, 0.75);
  const lbls: DataLabel[] = [];
  c.series.forEach((s, j) => {
    if (!hasValues(s)) return;
    const col = seriesColour(c, j);
    // A missing value counts as the centre, as Word draws it.
    const ring: [number, number][] = [];
    for (let i = 0; i < c.n; i++) ring.push(at(s.values[i] ?? vs.lo, i));
    const segs: Seg[] = ring.map((p, i): Seg => [i === 0 ? 0 : 1, p[0], p[1]]);
    segs.push([3]);
    if (c.radarFill) fill(out, segs.slice(), col);
    stroke(out, segs, col, s.lineW);
    ring.forEach((p, i) => {
      const v = s.values[i];
      if (c.radarMarks && v !== null && v !== undefined) out.push({ k: "shape", rect: [p[0] - MARKER / 2, p[1] - MARKER / 2, MARKER, MARKER], ellipse: true, fill: col, sw: 0 });
      const t = labelOf(s, i);
      if (t !== undefined) lbls.push({ x: p[0], y: p[1], text: t, inside: false });
    });
  });
  drawDataLabels(out, lbls, fs, false);
  if (!c.val.deleted) {
    for (const t of ticks) {
      const y = at(t, 0)[1];
      const s = formatValue(t, c.valFmt, vs.step);
      const w = textW(s, fs) + 4;
      text(out, { x: cx - GAP / 2 - w, y: y - lh / 2, w, h: lh }, s, fs, "right", false);
    }
  }
  if (labelled) {
    c.cats.forEach((s, i) => {
      if (s === "") return;
      const a = angle(i);
      const [x, y] = polar(cx, cy, r + GAP, a);
      const w = side - GAP;
      const txt = clipText(s, w, fs);
      const sx = Math.sin(a), sy = Math.cos(a);
      if (sx > 0.2) text(out, { x, y: y - lh / 2, w, h: lh }, txt, fs, "left", false);
      else if (sx < -0.2) text(out, { x: x - w, y: y - lh / 2, w, h: lh }, txt, fs, "right", false);
      else {
        const cw = Math.max(1, textW(txt, fs) + 4);
        text(out, { x: x - cw / 2, y: sy > 0 ? y - lh : y, w: cw, h: lh }, txt, fs, "center", false);
      }
    });
  }
}

/** The data label of point `i`, when the series shows them: the file's value in the series' format. */
function labelOf(s: Series, i: number): string | undefined {
  if (s.labels !== true) return undefined;
  const v = s.raw[i];
  if (v === null || v === undefined) return undefined;
  const f = s.fmt !== undefined && s.fmt.toLowerCase() !== "general" ? s.fmt : undefined;
  return f !== undefined ? formatValue(v, f, 1) : generalValue(v);
}

/** A value in General: whole numbers without decimals, others with up to six. */
function generalValue(v: number): string {
  if (v === Math.trunc(v) && Math.abs(v) < 1e15) return fixed(v, 0);
  const s = fixed(v, 6).replace(/0+$/, "").replace(/\.$/, "");
  return s === "-0" ? "0" : s;
}

/** Data labels over the drawn chart: centred inside (white), else above, or right of horizontal bars. */
function drawDataLabels(out: Item[], lbls: DataLabel[], fs: number, besideRight: boolean) {
  const lh = fs * LINE;
  for (const l of lbls) {
    const w = textW(l.text, fs) + 4;
    const [rect, align]: [Rect, Align] = l.inside
      ? [{ x: l.x - w / 2, y: l.y - lh / 2, w, h: lh }, "center"]
      : besideRight
        ? [{ x: l.x + GAP / 2, y: l.y - lh / 2, w, h: lh }, "left"]
        : [{ x: l.x - w / 2, y: l.y - lh, w, h: lh }, "center"];
    out.push({ k: "text", rect: arr(rect), text: l.text, size: fs, color: l.inside ? WHITE : TEXT, bold: false, align });
  }
}

/** A point at clockwise angle `t` (radians from 12 o'clock) on a circle of radius `r`. */
const polar = (cx: number, cy: number, r: number, t: number): [number, number] => [cx + r * Math.sin(t), cy - r * Math.cos(t)];

/** Cubic segments along a circle from angle `a` to `b`, each at most a quarter turn. */
function arc(segs: Seg[], cx: number, cy: number, r: number, a: number, b: number) {
  const q = Math.ceil(Math.abs(b - a) / (Math.PI / 2));
  const n = clamp(Number.isFinite(q) ? q : 0, 1, 16);
  const d = (b - a) / n;
  const k = (4 / 3) * Math.tan(d / 4);
  for (let j = 0; j < n; j++) {
    const t0 = a + d * j, t1 = t0 + d;
    const p0 = polar(cx, cy, r, t0), p1 = polar(cx, cy, r, t1);
    segs.push([2, p0[0] + k * r * Math.cos(t0), p0[1] + k * r * Math.sin(t0), p1[0] - k * r * Math.cos(t1), p1[1] - k * r * Math.sin(t1), p1[0], p1[1]]);
  }
}

// ---- the chart ----
function titleBlock(out: Item[], t: string, area: Rect): Rect {
  const lh = TITLE_SIZE * LINE;
  text(out, { x: area.x, y: area.y, w: area.w, h: lh }, t, TITLE_SIZE, "center", true);
  return { ...area, y: area.y + lh + PAD, h: Math.max(1, area.h - lh - PAD) };
}

const UNSUPPORTED: Record<string, string> = {
  "c:stockChart": "Stock chart",
  "c:bubbleChart": "Bubble chart",
  "c:surfaceChart": "Surface chart",
  "c:surface3DChart": "Surface chart",
  "c:ofPieChart": "Pie of pie chart",
  "c:bar3DChart": "3-D bar chart",
  "c:line3DChart": "3-D line chart",
  "c:area3DChart": "3-D area chart",
  "c:pie3DChart": "3-D pie chart",
};

/** A chart of only unsupported types: its frame, title and a line naming the type. */
function placeholder(out: Item[], space: Element, theme: Theme, w: number, h: number) {
  const chart = child(space, "c:chart");
  const plot = child(chart, "c:plotArea");
  if (!chart || !plot || els(plot).some((e) => kindOf(e.tagName))) return;
  const ty = els(plot).find((e) => e.tagName.endsWith("Chart"));
  if (!ty) return;
  backdrop(out, space, theme, w, h);
  out.push({ k: "shape", rect: [0, 0, w, h], ellipse: false, stroke: AXIS, sw: 0.75 });
  let area = { x: PAD, y: PAD, w: Math.max(1, w - 2 * PAD), h: Math.max(1, h - 2 * PAD) };
  const t = titleOf(chart, readSeries(ty, false, theme));
  if (t !== undefined) area = titleBlock(out, t, area);
  const size = textSize(space);
  const lh = size * LINE;
  const msg = `${Object.prototype.hasOwnProperty.call(UNSUPPORTED, ty.tagName) ? UNSUPPORTED[ty.tagName] : "Chart"} (not supported yet)`;
  text(out, { x: area.x, y: area.y + Math.max(0, (area.h - lh) / 2), w: area.w, h: lh }, clipText(msg, area.w, size), size, "center", false);
}

/** Items for a chart space sized `w` x `h` points: a placeholder for unsupported types, empty when nothing draws. */
function chartItems(space: Element, theme: Theme, w: number, h: number): Item[] {
  const out: Item[] = [];
  if (!(Number.isFinite(w) && Number.isFinite(h) && w > 0 && h > 0)) return out;
  const c = parseChart(space, theme);
  if (!c) { placeholder(out, space, theme, w, h); return out; }
  backdrop(out, space, theme, w, h);
  let area: Rect = { x: PAD, y: PAD, w: Math.max(1, w - 2 * PAD), h: Math.max(1, h - 2 * PAD) };
  if (c.title !== undefined) area = titleBlock(out, c.title, area);
  if (c.legend) {
    // Word lists bar series in drawing order: reversed for clustered bars and vertically listed stacked columns.
    const vertical = c.legend === "l" || c.legend === "r";
    const reverse = c.kind === "bar" && (c.horizontal ? !stacked(c) : stacked(c) && vertical);
    const entries = legendEntries(c);
    if (reverse) entries.reverse();
    area = legend(out, entries, c.legend, area, c.textSize);
  }
  if (c.kind === "pie" || c.kind === "doughnut") pie(out, c, area);
  else if (c.kind === "radar") radar(out, c, area);
  else cartesian(out, c, area);
  return out;
}

// ---- SVG ----
const num = (v: number) => String(Math.round(v * 1000) / 1000);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const finite = (...v: number[]) => v.every(Number.isFinite);

/** Path data with non-finite segments dropped, and anything after a broken move. */
function pathData(segs: Seg[]): string {
  const d: string[] = [];
  let open = false;
  for (const s of segs) {
    if (s[0] === 0) { open = finite(s[1], s[2]); if (open) d.push(`M${num(s[1])} ${num(s[2])}`); }
    else if (s[0] === 1) { if (open && finite(s[1], s[2])) d.push(`L${num(s[1])} ${num(s[2])}`); }
    else if (s[0] === 2) { if (open && finite(s[1], s[2], s[3], s[4], s[5], s[6])) d.push(`C${[s[1], s[2], s[3], s[4], s[5], s[6]].map(num).join(" ")}`); }
    else if (open) d.push("Z");
  }
  return d.join("");
}

function itemSvg(it: Item): string {
  const paint = (f?: Rgb, s?: Rgb, sw = 0) => `fill="${f ? "#" + f.toLowerCase() : "none"}"` + (s ? ` stroke="#${s.toLowerCase()}" stroke-width="${num(sw)}"` : "");
  if (it.k === "shape") {
    const [x, y, w, h] = it.rect as [number, number, number, number];
    if (!finite(x, y, w, h)) return "";
    return it.ellipse
      ? `<ellipse cx="${num(x + w / 2)}" cy="${num(y + h / 2)}" rx="${num(w / 2)}" ry="${num(h / 2)}" ${paint(it.fill, it.stroke, it.sw)}/>`
      : `<rect x="${num(x)}" y="${num(y)}" width="${num(w)}" height="${num(h)}" ${paint(it.fill, it.stroke, it.sw)}/>`;
  }
  if (it.k === "path") {
    const d = pathData(it.segs);
    return d ? `<path d="${d}" ${paint(it.fill, it.stroke, it.sw)}${it.stroke ? ' stroke-linejoin="round"' : ""}/>` : "";
  }
  const [x, y, w, h] = it.rect as [number, number, number, number];
  if (!it.text || !finite(x, y, w, h, it.size)) return "";
  const tx = it.align === "left" ? x : it.align === "center" ? x + w / 2 : x + w;
  const anchor = it.align === "left" ? "start" : it.align === "center" ? "middle" : "end";
  return `<text x="${num(tx)}" y="${num(y + h / 2)}" font-size="${num(clamp(it.size, 1, 400))}" text-anchor="${anchor}" dominant-baseline="central" fill="#${it.color.toLowerCase()}"${it.bold ? ' font-weight="bold"' : ""}>${esc(it.text)}</text>`;
}

/** Parse chart part XML; undefined when it isn't well-formed. */
export function parseChartXml(xml: string): Element | undefined {
  if (typeof xml !== "string" || !xml.trim()) return undefined;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) return undefined;
  return doc.documentElement ?? undefined;
}

/**
 * A chart part (c:chartSpace XML) drawn as a standalone SVG document `widthPt` x `heightPt` points,
 * in the theme `colors` (dk1, lt1, ..., accent1..6 as RRGGBB). Undefined when there is nothing to draw.
 */
export function chartToSvg(chartXml: string, colors: Record<string, string>, widthPt = DEFAULT_W, heightPt = DEFAULT_H): string | undefined {
  try {
    const space = parseChartXml(chartXml);
    if (!space) return undefined;
    const items = chartItems(space, themeOf(colors ?? {}), widthPt, heightPt);
    if (!items.length) return undefined;
    const body = items.map(itemSvg).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${num(widthPt)} ${num(heightPt)}" width="${num(widthPt)}pt" height="${num(heightPt)}pt" preserveAspectRatio="none" font-family="Calibri, Carlito, Arial, sans-serif" xml:space="preserve">${body}</svg>`;
  } catch {
    // Hostile input must never break the document; the caller shows the fallback picture.
    return undefined;
  }
}
