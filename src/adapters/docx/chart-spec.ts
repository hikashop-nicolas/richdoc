/*! richdoc chart model: a port of the WordCraft chart spec (https://github.com/storytold/wordcraft).
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
// Editable charts: the model, written as a chart part with literal data (c:strLit / c:numLit, no
// embedded workbook), and read back only when re-writing it gives the same markup.
import { ownText, parseChartXml, parseF64, parseUint, rustFloat } from "./chart";

/** Most series in a chart. */
export const MAX_SERIES = 255;
/** Most categories (points per series). */
export const MAX_POINTS = 4000;
/** Most cells (series x categories). */
export const MAX_CELLS = 100_000;
/** Longest title, series name or category label, characters. */
export const MAX_TEXT = 255;
const MAX_DEPTH = 256;
const USIZE_MAX = 18_446_744_073_709_551_615;

export type ChartType = "column" | "columnStacked" | "bar" | "barStacked" | "line" | "area" | "pie" | "doughnut" | "scatter";
export type LegendPos = "right" | "left" | "top" | "bottom";

/** One series: its name and a value per category (null is a gap). */
export interface ChartSeries {
  name: string;
  values: (number | null)[];
}

/** A chart richdoc can edit. */
export interface ChartSpec {
  type: ChartType;
  /** Category labels (scatter: the x values). */
  categories: string[];
  series: ChartSeries[];
  title?: string;
  /** No legend when absent. */
  legend?: LegendPos;
  /** Show each point's value next to it. */
  dataLabels: boolean;
}

export const CHART_TYPES: readonly ChartType[] = ["column", "columnStacked", "bar", "barStacked", "line", "area", "pie", "doughnut", "scatter"];
export const LEGEND_POSITIONS: readonly LegendPos[] = ["right", "top", "left", "bottom"];
const LABELS: Record<ChartType, string> = { column: "Clustered Column", columnStacked: "Stacked Column", bar: "Clustered Bar", barStacked: "Stacked Bar", line: "Line", area: "Area", pie: "Pie", doughnut: "Doughnut", scatter: "Scatter" };

/** Content type of a chart part. */
export const CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
/** a:graphicData URI of a chart. */
export const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// Axis ids (any distinct pair; Word keeps them).
const CAT_AX = "500000001";
const VAL_AX = "500000002";

export const chartTypeFromId = (id: string): ChartType | undefined => CHART_TYPES.find((t) => t.toLowerCase() === String(id).trim().toLowerCase());
export const legendFromId = (id: string): LegendPos | undefined => LEGEND_POSITIONS.find((p) => p === String(id).trim().toLowerCase());
export const chartTypeLabel = (t: ChartType): string => LABELS[t];
/** Pie and doughnut charts show one series, a slice per category. */
export const isRound = (t: ChartType): boolean => t === "pie" || t === "doughnut";

/** A new chart of type `type` with sample data: four quarters, three series (one for pie, x values for scatter), legend at the bottom. */
export function sampleChart(type: ChartType): ChartSpec {
  const [categories, series]: [string[], [string, number[]][]] = isRound(type)
    ? [["North", "East", "South", "West"], [["Share", [38, 24, 21, 17]]]]
    : type === "scatter"
      ? [["1", "2", "3", "4"], [["Series 1", [2.1, 3.4, 3.0, 4.6]], ["Series 2", [1.2, 1.9, 2.8, 3.1]]]]
      : [["Q1", "Q2", "Q3", "Q4"], [["Series 1", [3.2, 4.1, 3.6, 5.0]], ["Series 2", [2.0, 2.7, 3.9, 3.3]], ["Series 3", [1.5, 2.2, 2.4, 3.8]]]];
  return { type, categories, series: series.map(([name, v]) => ({ name, values: v.slice() })), legend: "bottom", dataLabels: false };
}

// Unicode White_Space, as Rust's str::trim sees it.
const WS = "[\\t\\n\\v\\f\\r \\u0085\\u00a0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]";
const trimStart = (s: string) => s.replace(new RegExp(`^${WS}+`), "");
const trimEnd = (s: string) => s.replace(new RegExp(`${WS}+$`), "");

/** `s` as chart text: control characters as spaces, ends trimmed, at most MAX_TEXT characters. */
export function cleanText(s: unknown): string {
  const cs = Array.from(typeof s === "string" ? s : "").slice(0, MAX_TEXT * 4).map((c) => (/^[\u0000-\u001f\u007f-\u009f]$/.test(c) ? " " : c));
  return trimEnd(Array.from(trimEnd(trimStart(cs.join("")))).slice(0, MAX_TEXT).join(""));
}

const finiteOrNull = (v: unknown): number | null => (typeof v === "number" && Number.isFinite(v) ? v : null);

/** The chart made safe to store and draw (a new object): text cleaned, sizes capped, non-finite values as gaps, every series as long as the categories. */
export function sanitize(spec: ChartSpec): ChartSpec {
  const type = CHART_TYPES.includes(spec?.type) ? spec.type : "column";
  let series = (Array.isArray(spec?.series) ? spec.series : []).slice(0, MAX_SERIES).map((s) => ({ name: s?.name, values: Array.isArray(s?.values) ? s.values : [] }));
  const cats = Array.isArray(spec?.categories) ? spec.categories : [];
  const longest = Math.max(0, ...series.map((s) => s.values.length));
  const n = Math.min(Math.max(cats.length, longest), MAX_POINTS);
  const keep = n > 0 ? Math.min(Math.max(Math.floor(MAX_CELLS / n), 1), MAX_SERIES) : MAX_SERIES;
  series = series.slice(0, keep);
  const categories = cats.slice(0, n).map(cleanText);
  while (categories.length < n) categories.push("");
  const out: ChartSpec = {
    type,
    categories,
    series: series.map((s) => {
      const values = s.values.slice(0, n).map(finiteOrNull);
      while (values.length < n) values.push(null);
      return { name: cleanText(s.name), values };
    }),
    dataLabels: spec?.dataLabels === true,
  };
  const title = spec?.title == null ? "" : cleanText(spec.title);
  if (title) out.title = title;
  if (spec?.legend && LEGEND_POSITIONS.includes(spec.legend)) out.legend = spec.legend;
  return out;
}

/** The spec held to the limits before sanitizing, so an oversized model never multiplies. */
function bounded(spec: ChartSpec): ChartSpec {
  return sanitize({
    ...spec,
    categories: (Array.isArray(spec?.categories) ? spec.categories : []).slice(0, MAX_POINTS),
    series: (Array.isArray(spec?.series) ? spec.series : []).slice(0, MAX_SERIES).map((s) => ({ name: s?.name, values: (Array.isArray(s?.values) ? s.values : []).slice(0, MAX_POINTS) })),
  });
}

// ---- writer ----
function xmlChar(cp: number): boolean {
  return cp === 0x9 || cp === 0xa || cp === 0xd || (cp >= 0x20 && cp <= 0xd7ff) || (cp >= 0xe000 && cp <= 0xfffd) || cp >= 0x10000;
}
/** Text escaped as WordCraft writes it (tabs and newlines as character references, invalid characters dropped). */
function esc(s: string): string {
  let o = "";
  for (const c of s) {
    switch (c) {
      case "&": o += "&amp;"; break;
      case "<": o += "&lt;"; break;
      case ">": o += "&gt;"; break;
      case '"': o += "&quot;"; break;
      case "'": o += "&apos;"; break;
      case "\t": o += "&#9;"; break;
      case "\n": o += "&#10;"; break;
      case "\r": o += "&#13;"; break;
      default: if (xmlChar(c.codePointAt(0)!)) o += c;
    }
  }
  return o;
}

class W {
  s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  private attrs(a: [string, string][]) { for (const [k, v] of a) this.s += ` ${k}="${esc(v)}"`; }
  open(name: string, a: [string, string][] = []) { this.s += `<${name}`; this.attrs(a); this.s += ">"; }
  empty(name: string, a: [string, string][] = []) { this.s += `<${name}`; this.attrs(a); this.s += "/>"; }
  close(name: string) { this.s += `</${name}>`; }
  leaf(name: string, t: string) { this.open(name); this.s += esc(t); this.close(name); }
  val(name: string, v: string) { this.empty(name, [["val", v]]); }
}

/** The chart part for `spec`: a complete XML document, written as `sanitize` leaves it. */
export function chartXml(spec: ChartSpec): string {
  return chartPart(spec).xml;
}

/** `chartXml`, and whether the limits left series or points of `spec` out. */
export function chartPart(spec: ChartSpec): { xml: string; cut: boolean } {
  const kept = bounded(spec);
  const ser = Array.isArray(spec?.series) ? spec.series : [];
  const points = Math.max(Array.isArray(spec?.categories) ? spec.categories.length : 0, ...ser.map((s) => (Array.isArray(s?.values) ? s.values.length : 0)));
  return { xml: writeChart(kept), cut: kept.series.length < ser.length || kept.categories.length < points };
}

function writeChart(spec: ChartSpec): string {
  const w = new W();
  w.open("c:chartSpace", [["xmlns:c", NS_C], ["xmlns:a", NS_A], ["xmlns:r", NS_R]]);
  w.val("c:roundedCorners", "0");
  w.open("c:chart");
  if (spec.title !== undefined) {
    w.open("c:title");
    w.open("c:tx");
    w.open("c:rich");
    w.empty("a:bodyPr");
    w.empty("a:lstStyle");
    w.open("a:p");
    w.open("a:r");
    w.leaf("a:t", spec.title);
    w.close("a:r");
    w.close("a:p");
    w.close("c:rich");
    w.close("c:tx");
    w.val("c:overlay", "0");
    w.close("c:title");
  }
  w.val("c:autoTitleDeleted", spec.title !== undefined ? "0" : "1");
  w.open("c:plotArea");
  w.empty("c:layout");
  plot(w, spec);
  w.close("c:plotArea");
  if (spec.legend) {
    w.open("c:legend");
    w.val("c:legendPos", { right: "r", left: "l", top: "t", bottom: "b" }[spec.legend]);
    w.val("c:overlay", "0");
    w.close("c:legend");
  }
  w.val("c:plotVisOnly", "1");
  w.val("c:dispBlanksAs", "gap");
  w.close("c:chart");
  // A white chart area without a border, in the document's own background colour.
  w.open("c:spPr");
  w.open("a:solidFill");
  w.val("a:schemeClr", "bg1");
  w.close("a:solidFill");
  w.open("a:ln");
  w.empty("a:noFill");
  w.close("a:ln");
  w.close("c:spPr");
  w.close("c:chartSpace");
  return w.s;
}

const isBar = (k: ChartType) => k === "column" || k === "columnStacked" || k === "bar" || k === "barStacked";

/** The chart type element and its series, then the axes. */
function plot(w: W, spec: ChartSpec) {
  const k = spec.type;
  const el = isBar(k) ? "c:barChart" : `c:${k}Chart`;
  const stacked = k === "columnStacked" || k === "barStacked";
  const horizontal = k === "bar" || k === "barStacked";
  w.open(el);
  if (isBar(k)) {
    w.val("c:barDir", horizontal ? "bar" : "col");
    w.val("c:grouping", stacked ? "stacked" : "clustered");
  } else if (k === "line" || k === "area") w.val("c:grouping", "standard");
  else if (k === "scatter") w.val("c:scatterStyle", "lineMarker");
  w.val("c:varyColors", isRound(k) ? "1" : "0");
  // Pie and doughnut charts show the first series; the others are kept for other types.
  spec.series.forEach((s, i) => series(w, spec, i, s));
  if (spec.dataLabels) {
    w.open("c:dLbls");
    for (const [name, on] of [["c:showLegendKey", false], ["c:showVal", true], ["c:showCatName", false], ["c:showSerName", false], ["c:showPercent", false], ["c:showBubbleSize", false]] as const) w.val(name, on ? "1" : "0");
    w.close("c:dLbls");
  }
  if (isBar(k)) {
    w.val("c:gapWidth", "150");
    if (stacked) w.val("c:overlap", "100");
  } else if (k === "line") w.val("c:marker", "1");
  else if (k === "pie") w.val("c:firstSliceAng", "0");
  else if (k === "doughnut") {
    w.val("c:firstSliceAng", "0");
    w.val("c:holeSize", "50");
  }
  if (!isRound(k)) {
    w.val("c:axId", CAT_AX);
    w.val("c:axId", VAL_AX);
  }
  w.close(el);
  if (isRound(k)) return;
  if (k === "scatter") {
    valueAxis(w, CAT_AX, VAL_AX, "b", false, "midCat");
    valueAxis(w, VAL_AX, CAT_AX, "l", true, "midCat");
  } else {
    categoryAxis(w, horizontal ? "l" : "b");
    valueAxis(w, VAL_AX, CAT_AX, horizontal ? "b" : "l", true, "between");
  }
}

/** One c:ser in schema order: idx, order, tx, spPr, marker, cat/xVal, val/yVal, smooth. */
function series(w: W, spec: ChartSpec, i: number, s: ChartSeries) {
  const k = spec.type;
  const n = String(i);
  w.open("c:ser");
  w.val("c:idx", n);
  w.val("c:order", n);
  if (s.name) {
    w.open("c:tx");
    w.leaf("c:v", s.name);
    w.close("c:tx");
  }
  if (k === "scatter") {
    // Markers only: the series has no line.
    w.open("c:spPr");
    w.open("a:ln", [["w", "19050"]]);
    w.empty("a:noFill");
    w.close("a:ln");
    w.close("c:spPr");
  } else if (isBar(k)) w.val("c:invertIfNegative", "0");
  else if (k === "line") {
    w.open("c:marker");
    w.val("c:symbol", "none");
    w.close("c:marker");
  }
  if (k === "scatter") {
    const xs: (number | null)[] = [];
    let numeric = true;
    for (const c of spec.categories) {
      if (!c) { xs.push(null); continue; }
      const v = parseF64(c);
      if (v === undefined || !Number.isFinite(v)) { numeric = false; break; }
      xs.push(v);
    }
    w.open("c:xVal");
    if (numeric) numLit(w, xs);
    else strLit(w, spec.categories);
    w.close("c:xVal");
    w.open("c:yVal");
    numLit(w, s.values);
    w.close("c:yVal");
  } else {
    w.open("c:cat");
    strLit(w, spec.categories);
    w.close("c:cat");
    w.open("c:val");
    numLit(w, s.values);
    w.close("c:val");
  }
  if (k === "line" || k === "scatter") w.val("c:smooth", "0");
  w.close("c:ser");
}

function strLit(w: W, items: string[]) {
  w.open("c:strLit");
  w.val("c:ptCount", String(items.length));
  items.forEach((s, i) => {
    if (!s) return;
    w.open("c:pt", [["idx", String(i)]]);
    w.leaf("c:v", s);
    w.close("c:pt");
  });
  w.close("c:strLit");
}

function numLit(w: W, values: (number | null)[]) {
  w.open("c:numLit");
  w.leaf("c:formatCode", "General");
  w.val("c:ptCount", String(values.length));
  values.forEach((v, i) => {
    if (v === null || !Number.isFinite(v)) return;
    w.open("c:pt", [["idx", String(i)]]);
    w.leaf("c:v", v === 0 ? "0" : rustFloat(v));
    w.close("c:pt");
  });
  w.close("c:numLit");
}

function categoryAxis(w: W, pos: string) {
  w.open("c:catAx");
  w.val("c:axId", CAT_AX);
  w.open("c:scaling");
  w.val("c:orientation", "minMax");
  w.close("c:scaling");
  w.val("c:delete", "0");
  w.val("c:axPos", pos);
  w.val("c:majorTickMark", "none");
  w.val("c:minorTickMark", "none");
  w.val("c:tickLblPos", "nextTo");
  w.val("c:crossAx", VAL_AX);
  w.val("c:crosses", "autoZero");
  w.val("c:auto", "1");
  w.val("c:lblAlgn", "ctr");
  w.val("c:lblOffset", "100");
  w.val("c:noMultiLvlLbl", "0");
  w.close("c:catAx");
}

/** A value axis: with gridlines and no axis line (`grid`), else with its line and no gridlines. */
function valueAxis(w: W, id: string, cross: string, pos: string, grid: boolean, between: string) {
  w.open("c:valAx");
  w.val("c:axId", id);
  w.open("c:scaling");
  w.val("c:orientation", "minMax");
  w.close("c:scaling");
  w.val("c:delete", "0");
  w.val("c:axPos", pos);
  if (grid) w.empty("c:majorGridlines");
  w.val("c:majorTickMark", "none");
  w.val("c:minorTickMark", "none");
  w.val("c:tickLblPos", "nextTo");
  if (grid) {
    w.open("c:spPr");
    w.open("a:ln");
    w.empty("a:noFill");
    w.close("a:ln");
    w.close("c:spPr");
  }
  w.val("c:crossAx", cross);
  w.val("c:crosses", "autoZero");
  w.val("c:crossBetween", between);
  w.close("c:valAx");
}

// ---- reader ----
const kids = (e: Element | undefined): Element[] => (e ? Array.from(e.children) : []);
const child = (e: Element | undefined, name: string) => kids(e).find((c) => c.tagName === name);
const valOf = (e: Element | undefined, name: string) => child(e, name)?.getAttribute("val") ?? undefined;

/**
 * The chart a chart part holds, only when the part is exactly what `chartXml` writes for it (so
 * editing and saving loses nothing). Undefined for any other chart.
 */
export function specOf(xml: string): ChartSpec | undefined {
  const space = parseChartXml(xml);
  if (!space) return undefined;
  const chart = child(space, "c:chart");
  const pl = child(chart, "c:plotArea");
  const ty = kids(pl).find((e) => e.tagName.endsWith("Chart"));
  if (!chart || !ty) return undefined;
  const bar = (st: ChartType, cl: ChartType) => (valOf(ty, "c:grouping") === "stacked" ? st : cl);
  const map: Record<string, () => ChartType> = {
    "c:barChart": () => (valOf(ty, "c:barDir") === "bar" ? bar("barStacked", "bar") : bar("columnStacked", "column")),
    "c:lineChart": () => "line",
    "c:areaChart": () => "area",
    "c:pieChart": () => "pie",
    "c:doughnutChart": () => "doughnut",
    "c:scatterChart": () => "scatter",
  };
  if (!Object.prototype.hasOwnProperty.call(map, ty.tagName)) return undefined;
  const type = map[ty.tagName]!();
  const scatter = type === "scatter";
  const sers = kids(ty).filter((e) => e.tagName === "c:ser").slice(0, MAX_SERIES + 1);
  // More than the model holds can't be the same chart: don't build it to find out.
  const count = (s: Element) => parseUint(s.getElementsByTagName("c:ptCount")[0]?.getAttribute("val") ?? undefined, USIZE_MAX) ?? 0;
  const longest = Math.max(0, ...sers.map(count));
  if (sers.length > MAX_SERIES || longest > MAX_POINTS || sers.length * longest > MAX_CELLS) return undefined;
  const catEl = child(sers[0], scatter ? "c:xVal" : "c:cat");
  const categories = catEl ? literal(catEl) : [];
  const series = sers.map((s) => {
    const v = child(child(s, "c:tx"), "c:v");
    const vals = child(s, scatter ? "c:yVal" : "c:val");
    return {
      name: v ? ownText(v) : "",
      values: (vals ? literal(vals) : []).map((t) => { const x = parseF64(t); return x !== undefined && Number.isFinite(x) ? x : null; }),
    };
  });
  const at = child(chart, "c:title")?.getElementsByTagName("a:t")[0];
  const lg = child(chart, "c:legend");
  const lp = lg ? valOf(lg, "c:legendPos") : undefined;
  const raw: ChartSpec = {
    type,
    categories,
    series,
    title: at ? ownText(at) : undefined,
    legend: lg ? (lp === "l" ? "left" : lp === "t" ? "top" : lp === "b" ? "bottom" : "right") : undefined,
    dataLabels: child(child(ty, "c:dLbls"), "c:showVal")?.getAttribute("val") === "1",
  };
  const spec = sanitize(raw);
  // Editable only when nothing would be lost: writing the model again gives the same part.
  const again = parseChartXml(chartXml(spec));
  return again && same(space, again, 0) ? spec : undefined;
}

/** The points of a c:strLit / c:numLit as texts by index, ptCount long. */
function literal(e: Element): string[] {
  const lit = kids(e).find((c) => c.tagName === "c:strLit" || c.tagName === "c:numLit");
  if (!lit) return [];
  const n = Math.min(parseUint(valOf(lit, "c:ptCount"), USIZE_MAX) ?? 0, MAX_POINTS + 1);
  const out: string[] = new Array(n).fill("");
  for (const p of kids(lit).filter((c) => c.tagName === "c:pt")) {
    const i = parseUint(p.getAttribute("idx") ?? undefined, USIZE_MAX);
    if (i !== undefined && i < out.length) { const v = child(p, "c:v"); out[i] = v ? ownText(v) : ""; }
  }
  return out;
}

/** Whether two elements are the same markup: names, attributes in any order, child elements in order, non-blank text. */
function same(a: Element, b: Element, depth: number): boolean {
  if (depth > MAX_DEPTH || a.tagName !== b.tagName || a.attributes.length !== b.attributes.length) return false;
  for (const at of Array.from(a.attributes)) if (!b.hasAttribute(at.name) || b.getAttribute(at.name) !== at.value) return false;
  const text = (e: Element) => Array.from(e.childNodes).filter((n) => (n.nodeType === 3 || n.nodeType === 4) && (n.nodeValue ?? "").trim() !== "").map((n) => n.nodeValue).join("");
  if (text(a) !== text(b)) return false;
  const x = kids(a), y = kids(b);
  return x.length === y.length && x.every((p, i) => same(p, y[i]!, depth + 1));
}
