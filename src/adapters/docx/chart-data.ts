// A chart's data as a table (categories x series), read from the chart part's caches, and written
// back: into the caches (c:strCache / c:numCache / literal c:strLit / c:numLit) and, when the chart
// keeps an embedded workbook, into the worksheet cells its series formulas (c:f) point at, so
// Word's "Edit Data" opens the same numbers. The table keeps its shape: rows and series are edited
// in place, never added or removed (that would move the formulas' ranges).
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export interface ChartData {
  title?: string;
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
}

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const kids = (el: Element | undefined | null, local: string): Element[] => (el ? Array.from(el.children).filter((c) => c.localName === local) : []);
const kid = (el: Element | undefined | null, local: string): Element | undefined => kids(el, local)[0];
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const PLOTS = /^(bar|bar3D|line|line3D|area|area3D|pie|pie3D|doughnut|scatter|radar|bubble|ofPie|stock|surface|surface3D)Chart$/;

/** The data source element of a series part (c:cat, c:val, c:xVal, c:yVal, c:tx): its cache. */
function cacheOf(part: Element | undefined): Element | undefined {
  if (!part) return undefined;
  const ref = kid(part, "numRef") ?? kid(part, "strRef") ?? kid(part, "multiLvlStrRef");
  return ref ? kid(ref, "numCache") ?? kid(ref, "strCache") ?? kid(ref, "multiLvlStrCache") : kid(part, "numLit") ?? kid(part, "strLit");
}
function points(cache: Element | undefined): string[] {
  if (!cache) return [];
  const n = Number(kid(cache, "ptCount")?.getAttribute("val") ?? "0");
  // A multi-level category cache: the innermost level (the labels next to the values).
  const lvl = kid(cache, "lvl");
  const src = lvl ?? cache;
  const out: string[] = new Array(Math.max(0, Math.min(n || 0, 10000))).fill("");
  for (const pt of kids(src, "pt")) {
    const i = Number(pt.getAttribute("idx"));
    if (Number.isInteger(i) && i >= 0 && i < 10000) out[i] = kid(pt, "v")?.textContent ?? "";
  }
  return out;
}
const firstPlot = (doc: Document): Element | undefined => {
  const plotArea = doc.getElementsByTagNameNS(C, "plotArea")[0];
  return Array.from(plotArea?.children ?? []).find((c) => PLOTS.test(c.localName));
};
const valPart = (ser: Element) => kid(ser, "val") ?? kid(ser, "yVal");
const catPart = (ser: Element) => kid(ser, "cat") ?? kid(ser, "xVal");

/** The first plot's data: category labels from the first series, one column per series. */
export function readChartData(chartXml: string): ChartData | undefined {
  const doc = parse(chartXml);
  if (doc.getElementsByTagName("parsererror").length) return undefined;
  const plot = firstPlot(doc);
  const sers = kids(plot, "ser");
  if (!plot || !sers.length) return undefined;
  const categories = points(cacheOf(catPart(sers[0]!)));
  const series = sers.map((s, i) => {
    const vals = points(cacheOf(valPart(s)));
    return { name: points(cacheOf(kid(s, "tx")))[0] || kid(kid(s, "tx"), "v")?.textContent || `Series ${i + 1}`, values: vals.map((v) => (v === "" || !isFinite(Number(v)) ? null : Number(v))) };
  });
  const rows = Math.max(categories.length, ...series.map((s) => s.values.length));
  while (categories.length < rows) categories.push(String(categories.length + 1));
  for (const s of series) while (s.values.length < rows) s.values.push(null);
  const titleTs = Array.from(doc.getElementsByTagNameNS(C, "title")[0]?.getElementsByTagNameNS(A, "t") ?? []);
  return { title: titleTs.length ? titleTs.map((t) => t.textContent ?? "").join("") : undefined, categories, series };
}

/** Write values into a cache (keeping its ptCount and number format), point by point. */
function setPoints(cache: Element | undefined, values: string[]): void {
  if (!cache) return;
  const doc = cache.ownerDocument;
  const src = kid(cache, "lvl") ?? cache;
  const byIdx = new Map(kids(src, "pt").map((p) => [Number(p.getAttribute("idx")), p] as const));
  values.forEach((v, i) => {
    let pt = byIdx.get(i);
    if (v === "") { pt?.remove(); return; }
    if (!pt) {
      pt = doc.createElementNS(C, "c:pt");
      pt.setAttribute("idx", String(i));
      pt.appendChild(doc.createElementNS(C, "c:v"));
      // Points stay in index order.
      const after = kids(src, "pt").find((p) => Number(p.getAttribute("idx")) > i);
      src.insertBefore(pt, after ?? null);
    }
    const ve = kid(pt, "v")!;
    ve.textContent = v;
  });
}

/** Apply edited data to a chart part's caches and title, in place. */
export function writeChartData(chartXml: string, data: ChartData): string {
  const doc = parse(chartXml);
  const plot = firstPlot(doc);
  const sers = kids(plot, "ser");
  sers.forEach((s, i) => {
    const d = data.series[i];
    if (!d) return;
    setPoints(cacheOf(valPart(s)), d.values.map((v) => (v === null || !isFinite(v) ? "" : String(v))));
    const tx = kid(s, "tx");
    const txCache = cacheOf(tx);
    if (txCache) setPoints(txCache, [d.name]);
    else if (kid(tx, "v")) kid(tx, "v")!.textContent = d.name;
    const cat = catPart(s);
    const catCache = cacheOf(cat);
    // Scatter x values are numbers; text categories go in as written.
    if (catCache) setPoints(catCache, data.categories);
  });
  if (data.title !== undefined) {
    const ts = Array.from(doc.getElementsByTagNameNS(C, "title")[0]?.getElementsByTagNameNS(A, "t") ?? []);
    if (ts.length) { ts[0]!.textContent = data.title; for (const t of ts.slice(1)) t.textContent = ""; }
  }
  return new XMLSerializer().serializeToString(doc);
}

// ---- the embedded workbook ----
/** "Sheet1!$B$2:$B$5" -> sheet name and the cells, in order. */
function refCells(f: string): { sheet: string; cells: string[] } | undefined {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f.trim());
  if (!m) return undefined;
  const sheet = (m[1] ?? m[2]!).replace(/''/g, "'");
  const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const name = (c: number) => { let s = ""; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const c0 = col(m[3]!), r0 = Number(m[4]), c1 = m[5] ? col(m[5]) : c0, r1 = m[6] ? Number(m[6]) : r0;
  if ((c1 - c0 + 1) * (r1 - r0 + 1) > 10000) return undefined;
  const cells: string[] = [];
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) cells.push(`${name(c)}${r}`);
  return { sheet, cells };
}
const cellOrder = (ref: string): [number, number] => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  return [Number(m[2]), [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)];
};

/** Set cells of one worksheet XML (numbers as values, text as inline strings), creating rows/cells. */
function setCells(sheetXml: string, values: Map<string, string | number>): string {
  const doc = parse(sheetXml);
  const NS = doc.documentElement.namespaceURI;
  const data = doc.getElementsByTagNameNS(NS, "sheetData")[0];
  if (!data) return sheetXml;
  for (const [ref, v] of values) {
    const [r, c] = cellOrder(ref);
    let row = kids(data, "row").find((x) => Number(x.getAttribute("r")) === r);
    if (!row) {
      row = doc.createElementNS(NS, "row");
      row.setAttribute("r", String(r));
      data.insertBefore(row, kids(data, "row").find((x) => Number(x.getAttribute("r")) > r) ?? null);
    }
    let cell = kids(row, "c").find((x) => x.getAttribute("r") === ref);
    if (!cell) {
      cell = doc.createElementNS(NS, "c");
      cell.setAttribute("r", ref);
      row.insertBefore(cell, kids(row, "c").find((x) => cellOrder(x.getAttribute("r") ?? "A1")[1] > c) ?? null);
    }
    const style = cell.getAttribute("s");
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    for (const a of Array.from(cell.attributes)) if (a.name !== "r") cell.removeAttribute(a.name);
    if (style) cell.setAttribute("s", style);
    if (typeof v === "number") {
      const ve = doc.createElementNS(NS, "v");
      ve.textContent = String(v);
      cell.appendChild(ve);
    } else if (v !== "") {
      cell.setAttribute("t", "inlineStr");
      const is = doc.createElementNS(NS, "is");
      const t = doc.createElementNS(NS, "t");
      t.textContent = v;
      is.appendChild(t);
      cell.appendChild(is);
    }
  }
  return new XMLSerializer().serializeToString(doc);
}

/** Write edited data into the chart's embedded workbook at the cells its formulas name. */
export function updateWorkbook(xlsx: Uint8Array, chartXml: string, data: ChartData): Uint8Array {
  const doc = parse(chartXml);
  const sers = kids(firstPlot(doc), "ser");
  const writes = new Map<string, Map<string, string | number>>();
  const put = (f: string | null | undefined, vals: (string | number)[]) => {
    const ref = f ? refCells(f) : undefined;
    if (!ref) return;
    const m = writes.get(ref.sheet) ?? new Map();
    ref.cells.forEach((cell, i) => { if (i < vals.length) m.set(cell, vals[i]!); });
    writes.set(ref.sheet, m);
  };
  const fOf = (part: Element | undefined) => kid(kid(part, "numRef") ?? kid(part, "strRef") ?? kid(part, "multiLvlStrRef"), "f")?.textContent;
  sers.forEach((s, i) => {
    const d = data.series[i];
    if (!d) return;
    put(fOf(valPart(s)), d.values.map((v) => (v === null ? "" : v)));
    put(fOf(kid(s, "tx")), [d.name]);
    if (i === 0) put(fOf(catPart(s)), data.categories.map((c) => (c.trim() !== "" && isFinite(Number(c)) && kid(catPart(s), "numRef") ? Number(c) : c)));
  });
  if (!writes.size) return xlsx;
  const files = unzipSync(xlsx);
  const wb = files["xl/workbook.xml"] && parse(strFromU8(files["xl/workbook.xml"]));
  const rels = files["xl/_rels/workbook.xml.rels"] && parse(strFromU8(files["xl/_rels/workbook.xml.rels"]));
  if (!wb || !rels) return xlsx;
  for (const [sheet, vals] of writes) {
    const sh = Array.from(wb.getElementsByTagName("sheet")).find((x) => x.getAttribute("name") === sheet);
    const rid = sh?.getAttribute("r:id") ?? sh?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const target = Array.from(rels.getElementsByTagName("Relationship")).find((r) => r.getAttribute("Id") === rid)?.getAttribute("Target");
    if (!target) continue;
    const key = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    if (files[key]) files[key] = strToU8(setCells(strFromU8(files[key]!), vals));
  }
  // Cached formula results elsewhere in the workbook may be stale: ask Excel to recalculate.
  const wbXml = strFromU8(files["xl/workbook.xml"]!);
  if (/<calcPr\b/.test(wbXml) && !/fullCalcOnLoad/.test(wbXml)) files["xl/workbook.xml"] = strToU8(wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"'));
  return zipSync(files);
}
