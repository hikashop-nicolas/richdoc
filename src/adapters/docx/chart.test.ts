import { describe, expect, it } from "vitest";
import { categoryText, chartToSvg, formatValue, narrow, niceScale, textW } from "./chart";
import { CHART_TYPES, MAX_CELLS, MAX_POINTS, MAX_SERIES, MAX_TEXT, chartPart, chartXml, sampleChart, sanitize, specOf, type ChartSpec } from "./chart-spec";

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const THEME: Record<string, string> = { dk1: "000000", lt1: "FFFFFF", dk2: "0E2841", lt2: "E8E8E8", accent1: "156082", accent2: "E97132", accent3: "196B24", accent4: "0F9ED5", accent5: "A02B93", accent6: "4EA72E", hlink: "467886", folHlink: "96607D" };
const ACCENT = (k: number) => "#" + THEME[`accent${k + 1}`]!.toLowerCase();

const space = (plot: string, extra = "") => `<c:chartSpace ${NS}><c:chart>${extra}<c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`;
const withLegend = (plot: string, pos: string) => `<c:chartSpace ${NS}><c:chart><c:plotArea>${plot}</c:plotArea><c:legend><c:legendPos val="${pos}"/></c:legend></c:chart></c:chartSpace>`;

function ser(i: number, name: string, color: string | undefined, cats: string[], vals: string[]): string {
  const fill = color ? `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>` : "";
  const cp = cats.map((c, k) => `<c:pt idx="${k}"><c:v>${c}</c:v></c:pt>`).join("");
  const vp = vals.map((v, k) => `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>${fill}<c:cat><c:strRef><c:strCache><c:ptCount val="${cats.length}"/>${cp}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vp}</c:numCache></c:numRef></c:val></c:ser>`;
}
const AXES = '<c:axId val="1"/><c:axId val="2"/>';
const CART_AXES = '<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:majorGridlines/></c:valAx>';

interface Svg { root: Element; fills: { fill: string; lo: number; hi: number; d: string }[]; strokes: Element[]; texts: string[]; swatches: Element[]; ellipses: Element[] }
function render(xml: string, w = 300, h = 200): Svg {
  const s = chartToSvg(xml, THEME, w, h);
  expect(s, "renders").toBeTypeOf("string");
  const root = new DOMParser().parseFromString(s!, "image/svg+xml").documentElement;
  expect(root.getElementsByTagName("parsererror").length).toBe(0);
  const paths = Array.from(root.getElementsByTagName("path"));
  const fills = paths.filter((p) => p.getAttribute("fill") !== "none").map((p) => {
    const d = p.getAttribute("d")!;
    const ys = Array.from(d.matchAll(/[ML](-?[\d.e+-]+) (-?[\d.e+-]+)/g)).map((m) => Number(m[2]));
    return { fill: p.getAttribute("fill")!, lo: Math.min(...ys), hi: Math.max(...ys), d };
  });
  return {
    root,
    fills,
    strokes: paths.filter((p) => p.getAttribute("fill") === "none"),
    texts: Array.from(root.getElementsByTagName("text")).map((t) => t.textContent ?? ""),
    swatches: Array.from(root.getElementsByTagName("rect")).filter((r) => r.getAttribute("width") === "7"),
    ellipses: Array.from(root.getElementsByTagName("ellipse")),
  };
}
const segCount = (p: Element) => (p.getAttribute("d")!.match(/[MLCZ]/g) ?? []).length;

describe("chartToSvg", () => {
  it("writes a standalone SVG in points", () => {
    const s = chartToSvg(space(`<c:barChart>${ser(0, "A", undefined, ["Q1"], ["1"])}${AXES}</c:barChart>${CART_AXES}`), THEME, 300, 200)!;
    expect(s).toMatch(/^<svg xmlns="http:\/\/www.w3.org\/2000\/svg" viewBox="0 0 300 200" width="300pt" height="200pt" preserveAspectRatio="none"/);
    expect(s).toContain("Calibri");
    expect(chartToSvg(space(`<c:barChart>${ser(0, "A", undefined, ["Q1"], ["1"])}</c:barChart>`), THEME)).toContain('viewBox="0 0 432 252"');
  });

  it("clustered columns draw one bar per point", () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "North", "FF0000", ["Q1", "Q2", "Q3"], ["4", "8", "2"])}${ser(1, "South", "00FF00", ["Q1", "Q2", "Q3"], ["2", "6", "4"])}${AXES}</c:barChart>${CART_AXES}`;
    const r = render(space(plot));
    const bars = r.fills.filter((f) => f.fill !== "#ffffff");
    expect(bars.length).toBe(6);
    const h = bars.map((b) => b.hi - b.lo);
    expect(h[1]! / h[0]!).toBeCloseTo(2, 3);
    expect(h[3]! / h[1]!).toBeCloseTo(0.25, 3);
    expect(h[4]! / h[1]!).toBeCloseTo(0.75, 3);
    expect(bars[0]!.fill).toBe("#ff0000");
    expect(bars[3]!.fill).toBe("#00ff00");
    expect(r.texts).toContain("Q2");
    // Axis tick labels 0..8 by 1 (scale rounded to nice steps) and gridlines.
    expect(r.texts).toContain("0");
    expect(r.texts).toContain("8");
    expect(r.strokes.some((p) => p.getAttribute("stroke") === "#d9d9d9")).toBe(true);
  });

  it("stacked and percent bars add up per category", () => {
    const sers = ser(0, "A", "FF0000", ["Q1", "Q2"], ["2", "6"]) + ser(1, "B", "00FF00", ["Q1", "Q2"], ["2", "2"]);
    for (const grouping of ["stacked", "percentStacked"]) {
      const r = render(space(`<c:barChart><c:barDir val="bar"/><c:grouping val="${grouping}"/>${sers}<c:overlap val="100"/>${AXES}</c:barChart>${CART_AXES}`));
      const bars = r.fills.filter((f) => f.fill !== "#ffffff");
      expect(bars.length).toBe(4);
      // Horizontal: the two segments of a category share the same band (same y range).
      expect(bars[0]!.lo).toBeCloseTo(bars[2]!.lo, 3);
      if (grouping === "percentStacked") expect(r.texts).toContain("100%");
    }
  });

  it("data labels show the values", () => {
    const labels = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/></c:dLbls>';
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "North", undefined, ["Q1", "Q2"], ["4.25", "17"])}${labels}${AXES}</c:barChart>${CART_AXES}`;
    const shown = render(space(plot)).texts;
    expect(shown).toContain("4.25");
    expect(shown).toContain("17");
    const pie = `<c:pieChart><c:varyColors val="1"/>${ser(0, "S", undefined, ["A", "B"], ["1", "3"])}${labels}</c:pieChart>`;
    expect(render(space(pie), 200, 150).texts).toContain("3");
    expect(render(space(plot.replace(labels, ""))).texts).not.toContain("4.25");
  });

  it("pie draws one slice per point in the accent colours", () => {
    const plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, "Share", undefined, ["A", "B", "C", "D"], ["50", "25", "15", "10"])}</c:pieChart>`;
    const r = render(withLegend(plot, "r"), 200, 150);
    const slices = r.fills.filter((f) => f.d.includes("C"));
    expect(slices.length).toBe(4);
    expect(slices[0]!.fill).toBe(ACCENT(0));
    expect(slices[1]!.fill).toBe(ACCENT(1));
    // A pie's legend lists the categories.
    for (const c of ["A", "B", "C", "D"]) expect(r.texts).toContain(c);
    expect(r.swatches.length).toBe(4);
  });

  it("doughnut slices are annuli", () => {
    const plot = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, "S", undefined, ["A", "B"], ["1", "3"])}<c:holeSize val="50"/></c:doughnutChart>`;
    const r = render(space(plot), 200, 150);
    const slices = r.fills.filter((f) => f.d.includes("C"));
    expect(slices.length).toBe(2);
    // Outer arc, a line in to the hole, the inner arc back: two arcs in each slice.
    expect(slices[1]!.d.match(/L/g)!.length).toBe(1);
  });

  it("lines, stacked lines and areas", () => {
    const stacked = `<c:lineChart><c:grouping val="stacked"/>${ser(0, "A", "FF0000", ["Q1", "Q2"], ["0", "10"])}${ser(1, "B", "00FF00", ["Q1", "Q2"], ["0", "5"])}${AXES}</c:lineChart>`;
    const single = `<c:lineChart><c:grouping val="standard"/>${ser(0, "T", "0000FF", ["Q1", "Q2"], ["0", "15"])}${AXES}</c:lineChart>`;
    const endY = (r: Svg, col: string) => Number(r.strokes.find((p) => p.getAttribute("stroke") === col)!.getAttribute("d")!.split(" ").pop());
    expect(endY(render(space(stacked)), "#00ff00")).toBeCloseTo(endY(render(space(single)), "#0000ff"), 3);
    const area = `<c:areaChart><c:grouping val="standard"/>${ser(0, "A", "FF0000", ["Q1", "Q2"], ["4", "8"])}${ser(1, "Empty", "00FF00", ["Q1", "Q2"], [])}${AXES}</c:areaChart>`;
    const fills = render(space(area)).fills.filter((f) => f.fill !== "#ffffff");
    expect(fills.map((f) => f.fill)).toEqual(["#ff0000"]);
  });

  it("scatter: lines, markers and date x values", () => {
    const pts = (v: string[]) => v.map((x, k) => `<c:pt idx="${k}"><c:v>${x}</c:v></c:pt>`).join("");
    const sc = (spPr: string) => `<c:scatterChart><c:ser><c:idx val="0"/>${spPr}<c:xVal><c:numRef><c:numCache><c:formatCode>m/d/yyyy</c:formatCode><c:ptCount val="3"/>${pts(["45292", "45323", "45352"])}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${pts(["1", "2", "3"])}</c:numCache></c:numRef></c:yVal></c:ser>${AXES}</c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>`;
    const lined = render(space(sc("")));
    expect(lined.strokes.some((p) => p.getAttribute("stroke") === ACCENT(0))).toBe(true);
    expect(lined.texts.some((t) => t.endsWith("/2024"))).toBe(true);
    expect(lined.texts.some((t) => t.startsWith("45"))).toBe(false);
    const marked = render(space(sc('<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>')));
    expect(marked.ellipses.length).toBe(3);
  });

  const radar = (style: string, sers: string[], w = 300, h = 220) => render(space(`<c:radarChart><c:radarStyle val="${style}"/>${sers.join("")}</c:radarChart><c:valAx><c:majorGridlines/></c:valAx>`), w, h);

  it("radar charts draw spokes, grid and series polygons", () => {
    const cats = ["Speed", "Power", "Range", "Armor", "Luck"];
    const sers = [ser(0, "A", "FF0000", cats, ["3", "4", "5", "2", "1"]), ser(1, "B", "0000FF", cats, ["2", "2", "3", "5", "4"])];
    const std = radar("standard", sers);
    for (const c of cats) expect(std.texts).toContain(c);
    const pentagons = (r: Svg, col: string) => r.strokes.filter((p) => p.getAttribute("stroke") === col && segCount(p) === 6).length;
    expect(pentagons(std, "#ff0000")).toBe(1);
    expect(pentagons(std, "#0000ff")).toBe(1);
    expect(std.fills.filter((f) => f.fill !== "#ffffff").length).toBe(0);
    expect(std.ellipses.length).toBe(0);
    expect(std.strokes.filter((p) => p.getAttribute("stroke") === "#d9d9d9" && segCount(p) === 2).length).toBe(5);
    expect(pentagons(std, "#d9d9d9")).toBeGreaterThan(0);
    expect(radar("marker", sers).ellipses.length).toBe(10);
    const solid = radar("filled", sers).fills.filter((f) => f.fill !== "#ffffff").map((f) => f.fill);
    expect(solid.sort()).toEqual(["#0000ff", "#ff0000"]);
    const legend = render(withLegend(`<c:radarChart>${ser(0, "Team", undefined, ["a", "b", "c"], ["1", "2", "3"])}</c:radarChart>`, "r"));
    expect(legend.texts).toContain("Team");
  });

  it("hostile radar charts stay bounded", { timeout: 30_000 }, () => {
    expect(chartToSvg(space("<c:radarChart/>"), THEME, 300, 200)).toBeUndefined();
    for (const n of [1, 2]) {
      const cats = Array.from({ length: n }, (_, i) => `c${i}`);
      expect(radar("filled", [ser(0, "S", undefined, cats, cats.map(() => "5"))]).texts).toContain("c0");
    }
    radar("filled", [ser(0, "S", undefined, ["a", "b", "c"], ["NaN", "inf", "1e999"])]);
    chartToSvg(space(`<c:radarChart>${ser(0, "S", undefined, ["a", "b", "c"], ["1", "2", "3"])}</c:radarChart>`), THEME, 20, 20);
    const cats = Array.from({ length: 150 }, (_, i) => `c${i}`);
    const many = Array.from({ length: 60 }, (_, i) => ser(i, "S", undefined, cats, cats.map(() => "1")));
    // Counted on the string: a DOM of thousands of markers is slow under jsdom.
    const svg = chartToSvg(space(`<c:radarChart><c:radarStyle val="marker"/>${many.join("")}</c:radarChart><c:valAx><c:majorGridlines/></c:valAx>`), THEME, 300, 220)!;
    const paths = svg.match(/<path [^>]*>/g)!;
    expect(paths.filter((p) => p.includes('fill="none"') && !p.includes("#d9d9d9")).length).toBeLessThanOrEqual(50);
    expect(paths.filter((p) => p.includes("#d9d9d9") && /d="M[^MLCZ"]*L[^MLCZ"]*"/.test(p)).length).toBe(100);
  });

  it("title and axis labels", () => {
    const title = "<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r><a:r><a:t> 2024</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    const plot = `<c:barChart><c:barDir val="col"/>${ser(0, "A", undefined, ["Jan", "Feb"], ["1200", "3400"])}${AXES}</c:barChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:numFmt formatCode="#,##0" sourceLinked="0"/></c:valAx>`;
    const r = render(space(plot, title));
    expect(r.texts[0]).toBe("Sales 2024");
    expect(r.root.getElementsByTagName("text")[0]!.getAttribute("font-weight")).toBe("bold");
    expect(r.texts).toContain("Jan");
    expect(r.texts).toContain("4,000");
    // An empty title shows a lone series' name; a deleted automatic title doesn't.
    expect(render(space(plot, "<c:title/>")).texts[0]).toBe("A");
    expect(render(space(plot, '<c:title/><c:autoTitleDeleted val="1"/>')).texts).not.toContain("A");
  });

  it("unsupported chart types draw a placeholder", () => {
    const s1 = '<c:ser><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>';
    const title = "<c:title><c:tx><c:rich><a:p><a:r><a:t>Skills</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    const bubble = render(space(`<c:bubbleChart>${s1}</c:bubbleChart>`, title));
    expect(bubble.texts).toEqual(["Skills", "Bubble chart (not supported yet)"]);
    expect(Array.from(bubble.root.getElementsByTagName("rect")).some((r) => r.getAttribute("stroke") === "#bfbfbf" && r.getAttribute("width") === "300")).toBe(true);
    expect(render(space(`<c:stockChart>${s1}${s1}${s1}</c:stockChart>`)).texts).toEqual(["Stock chart (not supported yet)"]);
    expect(render(space(`<c:fancyChart>${s1}</c:fancyChart>`)).texts).toEqual(["Chart (not supported yet)"]);
    const combo = render(space(`<c:stockChart>${s1}</c:stockChart><c:lineChart>${s1}</c:lineChart>`));
    expect(combo.texts.some((t) => t.includes("not supported"))).toBe(false);
  });

  const columns = (names: string[]) => `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${names.map((n, i) => ser(i, n, "FF0000", ["Q1"], ["1"])).join("")}${AXES}</c:barChart>${CART_AXES}`;

  it("legends: side entries that fit, bottom rows, long names cut", () => {
    const side = render(withLegend(columns(Array.from({ length: 10 }, (_, i) => `S${i}`)), "r"), 300, 60);
    expect(side.swatches.length).toBeGreaterThan(0);
    expect(side.swatches.length).toBeLessThan(10);
    const names = Array.from({ length: 12 }, (_, i) => `Region number ${i}`);
    const rows = render(withLegend(columns(names), "b"), 200, 200);
    const ys = new Set(rows.swatches.map((s) => Math.trunc(Number(s.getAttribute("y")))));
    expect(ys.size).toBeGreaterThan(1);
    expect(ys.size).toBeLessThanOrEqual(3);
    expect(rows.swatches.every((s) => Number(s.getAttribute("x")) >= 0 && Number(s.getAttribute("x")) <= 200)).toBe(true);
    const long = "A very long series name that would push the legend past the left edge of the chart";
    const cut = render(withLegend(columns([long, "Short"]), "r"), 200, 150);
    const t = cut.texts.find((x) => x.endsWith("…"))!;
    expect(t).toBeDefined();
    expect(textW(t, 9)).toBeLessThanOrEqual(200 * 0.4);
    expect(cut.texts).toContain("Short");
  });

  it("automatic series colours follow c:idx", () => {
    const colours = (sers: string[]) => {
      const r = render(withLegend(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers.join("")}${AXES}</c:barChart>${CART_AXES}`, "r"));
      return [r.fills.filter((f) => f.fill !== "#ffffff").map((f) => f.fill), r.swatches.map((s) => s.getAttribute("fill"))];
    };
    const renum = (s: string, io: string) => s.slice(0, s.indexOf("<c:idx")) + io + s.slice(s.indexOf("<c:tx>"));
    const a = ser(0, "A", undefined, ["Q1"], ["1"]), b = ser(1, "B", undefined, ["Q1"], ["2"]);
    expect(colours([renum(a, '<c:idx val="3"/><c:order val="0"/>'), renum(b, '<c:idx val="0"/><c:order val="1"/>')])).toEqual([[ACCENT(3), ACCENT(0)], [ACCENT(3), ACCENT(0)]]);
    expect(colours([renum(a, '<c:idx val="2"/>')])[0]).toEqual([ACCENT(2)]);
    expect(colours([renum(a, ""), renum(b, "")])[0]).toEqual([ACCENT(0), ACCENT(1)]);
    for (const junk of ["x", "-1", "1e3", "99999999999999999999999"]) {
      expect(colours([renum(a, `<c:idx val="${junk}"/>`), renum(b, `<c:idx val="${junk}"/>`)])[0]).toEqual([ACCENT(0), ACCENT(1)]);
    }
    expect(colours([renum(ser(0, "A", "FF0000", ["Q1"], ["1"]), '<c:idx val="3"/>')])[0]).toEqual(["#ff0000"]);
  });

  it("theme colours resolve from a:schemeClr, and missing slots fall back to the Office theme", () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></c:spPr><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>${AXES}</c:barChart>${CART_AXES}`;
    const s = chartToSvg(space(plot), { accent2: "FF0000" }, 300, 200)!;
    expect(s).toContain('fill="#800000"');
    expect(chartToSvg(space(`<c:barChart>${ser(0, "A", undefined, ["Q1"], ["1"])}${AXES}</c:barChart>${CART_AXES}`), {}, 300, 200)).toContain(`fill="${ACCENT(0)}"`);
  });

  it("series are capped to the cell budget", () => {
    const sers = Array.from({ length: 256 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:val><c:numRef><c:numCache><c:ptCount val="10000"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`).join("");
    const r = render(space(`<c:lineChart><c:grouping val="standard"/>${sers}${AXES}</c:lineChart>`));
    expect(r.strokes.length).toBe(MAX_CELLS / 10_000);
  });

  it("never throws on junk input", { timeout: 30_000 }, () => {
    const bodies = [
      `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:gapWidth val="-9999999999999999999999"/><c:overlap val="NaN"/><c:ser><c:tx><c:v>x</c:v></c:tx><c:cat><c:strRef><c:strCache><c:ptCount val="99999999999"/><c:pt idx="18446744073709551615"><c:v>a</c:v></c:pt><c:pt idx="-3"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1e400</c:v></c:pt><c:pt idx="1"><c:v>NaN</c:v></c:pt><c:pt idx="2"><c:v>-1e308</c:v></c:pt><c:pt idx="99999"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1e308</c:v></c:pt><c:pt idx="0"><c:v>1e308</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/></c:barChart><c:valAx><c:axId val="1"/><c:scaling><c:max val="1e400"/><c:min val="-1e308"/></c:scaling></c:valAx>`,
      '<c:scatterChart><c:ser><c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>-1e308</c:v></c:pt><c:pt idx="1"><c:v>1e308</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser></c:scatterChart>',
      '<c:pieChart><c:ser><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>-4</c:v></c:pt><c:pt idx="1"><c:v>1e308</c:v></c:pt></c:numLit></c:val><c:dPt><c:idx val="9999999999999"/></c:dPt></c:ser><c:firstSliceAng val="-9"/></c:pieChart>',
    ];
    for (const b of bodies) {
      for (const [w, h] of [[0, 10], [NaN, 10], [-1, 1], [1e30, 1e30], [50, 40]] as const) {
        const s = chartToSvg(space(b), THEME, w, h);
        if (s !== undefined) expect(s).not.toMatch(/NaN|Infinity/);
      }
    }
    expect(chartToSvg(space(bodies[0]!), THEME, NaN, 10)).toBeUndefined();
    for (const junk of ["", "not xml", "<c:chartSpace", "<a/>", `<c:chartSpace ${NS}/>`, `<c:chartSpace ${NS}><c:chart/></c:chartSpace>`, "\u0000", "<x:y/>"]) {
      expect(chartToSvg(junk, THEME, 300, 200)).toBeUndefined();
    }
    expect(chartToSvg(undefined as unknown as string, undefined as unknown as Record<string, string>)).toBeUndefined();
    const deep = "<a:p>".repeat(2000) + "<a:t>x</a:t>" + "</a:p>".repeat(2000);
    chartToSvg(space(`<c:barChart>${ser(0, "A", undefined, ["Q1"], ["1"])}</c:barChart>`, `<c:title><c:tx><c:rich>${deep}</c:rich></c:tx></c:title>`), THEME);
  });
});

describe("formats and scales", () => {
  it("nice scales", () => {
    expect(niceScale(0, 22.3, true)).toEqual({ lo: 0, hi: 25, step: 5 });
    expect(niceScale(0, 100, true)).toEqual({ lo: 0, hi: 120, step: 20 });
    expect(niceScale(0, 22.3, true, 0, 30).hi).toBe(30);
    const flat = niceScale(0, 0, true);
    expect(flat.lo).toBe(0);
    expect(flat.hi).toBeCloseTo(1.2, 9);
    expect(flat.step).toBe(0.2);
    expect(niceScale(-10, -2, true)).toEqual({ lo: -12, hi: 0, step: 2 });
    expect(niceScale(-1050, -1000, false).hi).toBe(-1000);
    expect(narrow(1000, 1050) && narrow(-1050, -1000)).toBe(true);
    expect(narrow(10, 100) || narrow(0, 1)).toBe(false);
    expect(niceScale(1000, 1050, false).lo).toBe(1000);
    expect(niceScale(1000, 1050, true).lo).toBe(0);
  });

  it("number formats", () => {
    expect(formatValue(1234567, "#,##0", 1)).toBe("1,234,567");
    expect(formatValue(-1500, "#,##0", 500)).toBe("-1,500");
    expect(formatValue(0.25, "0%", 0.05)).toBe("25%");
    expect(formatValue(0.5, "General", 0.25)).toBe("0.50");
    expect(formatValue(10, '"€"0.00', 5)).toBe("€10.00");
    expect(formatValue(1234.5, "#,##0.00 [$€-407]", 0.25)).toBe("1,234.50 €");
    expect(formatValue(5, "[Red]0", 1)).toBe("5");
    expect(formatValue(3, '"Items: "0', 1)).toBe("Items: 3");
    expect(formatValue(-5, "#,##0;(#,##0)", 1)).toBe("(5)");
    expect(formatValue(-1500, "#,##0_);(#,##0)", 500)).toBe("(1,500)");
    expect(formatValue(-2, '0";"0;-0', 1)).toBe("-2");
    expect(formatValue(12345, "0.00E+00", 1)).toBe("1.23E+04");
    expect(formatValue(0.00012, "0.0E+0", 0.0001)).toBe("1.2E-4");
    expect(formatValue(0, "0.00E+00", 1)).toBe("0.00E+00");
    expect(formatValue(-12345, "0.0E-00", 1)).toBe("-1.2E04");
    // Ties round to even, as Rust prints them.
    expect(formatValue(2.5, "0", 1)).toBe("2");
    expect(formatValue(1e22, "0", 1)).toBe("10000000000000000000000");
  });

  it("dates and times", () => {
    const t = (v: number, code: string) => categoryText(v, code, 1, false);
    expect(t(45292, "m/d/yyyy")).toBe("1/1/2024");
    expect(t(45323, "d/m/yyyy")).toBe("1/2/2024");
    expect(t(45292, "mmm-yy")).toBe("Jan-24");
    expect(t(45292, "0.00")).toBe("45292.00");
    const at = 45292 + (13 * 3600 + 5 * 60 + 9) / 86_400;
    expect(t(at, "hh:mm:ss")).toBe("13:05:09");
    expect(t(at, "h:mm AM/PM")).toBe("1:05 PM");
    expect(t(at, "m:ss")).toBe("5:09");
    expect(t(45292, "h:mm a/p")).toBe("12:00 A");
    expect(t(45292, "dddd, mmmm d")).toBe("Monday, January 1");
    expect(t(45292, "ddd d mmm")).toBe("Mon 1 Jan");
    expect(t(0, "m/d/yyyy")).toBe("1/0/1900");
    expect(t(59, "m/d/yyyy")).toBe("2/28/1900");
    expect(t(60, "m/d/yyyy")).toBe("2/29/1900");
    expect(t(61, "m/d/yyyy dddd")).toBe("3/1/1900 Thursday");
    expect(t(-1, "m/d/yyyy")).toBe("-1");
    expect(categoryText(0, "m/d/yyyy", 1, true)).toBe("1/1/1904");
  });
});

describe("chart-spec", () => {
  it("every type reads back as itself and draws", () => {
    for (const type of CHART_TYPES) {
      const s = sampleChart(type);
      s.title = "Sales & <costs>";
      s.dataLabels = true;
      s.legend = "right";
      s.series[0]!.values[1] = null;
      s.categories[2] = "";
      const spec = sanitize(s);
      expect(specOf(chartXml(spec)), type).toEqual(spec);
      expect(chartToSvg(chartXml(spec), THEME, 300, 200), type).toBeTypeOf("string");
    }
    const plain: ChartSpec = { ...sampleChart("line"), legend: undefined };
    delete plain.legend;
    expect(specOf(chartXml(plain))).toEqual(plain);
  });

  it("renders the written chart with its title, legend and slices", () => {
    const pie = sampleChart("pie");
    pie.title = "Regions";
    const r = render(chartXml(pie));
    expect(r.texts[0]).toBe("Regions");
    for (const c of pie.categories) expect(r.texts).toContain(c);
    expect(r.fills.filter((f) => f.d.includes("C")).length).toBe(4);
    const col = render(chartXml(sampleChart("columnStacked")));
    expect(col.fills.filter((f) => f.fill !== "#ffffff").length).toBe(12);
    // Scatter writes markers only.
    expect(render(chartXml(sampleChart("scatter"))).ellipses.length).toBe(8);
  });

  it("a chart with more than the model stays read-only", () => {
    const x = chartXml(sampleChart("column"));
    expect(specOf(x)).toBeDefined();
    expect(specOf(x.replace('<c:roundedCorners val="0"/>', '<c:roundedCorners val="0"/><c:style val="2"/>'))).toBeUndefined();
    expect(specOf(x.replace("<c:v>3.2</c:v>", "<c:v>3.20</c:v>"))).toBeUndefined();
    expect(specOf(x.replace('<c:gapWidth val="150"/>', '<c:gapWidth val="150" extra="1"/>'))).toBeUndefined();
    // Attribute order and whitespace between elements don't matter.
    expect(specOf(x.replace(/></g, ">\n  <"))).toEqual(sampleChart("column"));
    for (const junk of ["", "junk", "<a/>", `<c:chartSpace ${NS}><c:chart><c:plotArea><c:bubbleChart/></c:plotArea></c:chart></c:chartSpace>`]) expect(specOf(junk)).toBeUndefined();
  });

  it("elements come in schema order", () => {
    const s = sampleChart("columnStacked");
    s.title = "T";
    s.dataLabels = true;
    const root = new DOMParser().parseFromString(chartXml(s), "application/xml").documentElement;
    const names = (e: Element | undefined) => Array.from(e!.children).map((c) => c.tagName);
    const ch = (e: Element | undefined, n: string) => Array.from(e!.children).find((c) => c.tagName === n);
    expect(names(root)).toEqual(["c:roundedCorners", "c:chart", "c:spPr"]);
    const chart = ch(root, "c:chart");
    expect(names(chart)).toEqual(["c:title", "c:autoTitleDeleted", "c:plotArea", "c:legend", "c:plotVisOnly", "c:dispBlanksAs"]);
    const plot = ch(chart, "c:plotArea");
    expect(names(plot)).toEqual(["c:layout", "c:barChart", "c:catAx", "c:valAx"]);
    const bar = ch(plot, "c:barChart");
    expect(names(bar)).toEqual(["c:barDir", "c:grouping", "c:varyColors", "c:ser", "c:ser", "c:ser", "c:dLbls", "c:gapWidth", "c:overlap", "c:axId", "c:axId"]);
    expect(names(ch(bar, "c:ser"))).toEqual(["c:idx", "c:order", "c:tx", "c:invertIfNegative", "c:cat", "c:val"]);
    const sc = new DOMParser().parseFromString(chartXml(sampleChart("scatter")), "application/xml").getElementsByTagName("c:ser")[0];
    expect(names(sc)).toEqual(["c:idx", "c:order", "c:tx", "c:spPr", "c:xVal", "c:yVal", "c:smooth"]);
  });

  it("a chart past the limits is written within them", () => {
    const spec: ChartSpec = {
      type: "column",
      categories: Array.from({ length: 4500 }, (_, i) => String(i)),
      series: Array.from({ length: 300 }, (_, i) => ({ name: `S${i}`, values: new Array(4500).fill(1) })),
      dataLabels: false,
    };
    const { xml, cut } = chartPart(spec);
    expect(cut).toBe(true);
    const series = xml.split("<c:ser>").length - 1;
    expect(series).toBeGreaterThanOrEqual(1);
    expect(series).toBeLessThanOrEqual(MAX_SERIES);
    expect(xml.split("<c:pt ").length - 1).toBeLessThanOrEqual(2 * MAX_CELLS);
    expect(xml).not.toContain(`<c:pt idx="${MAX_POINTS}"`);
  });

  it("sanitize caps sizes and drops non-finite values", () => {
    const c = sanitize({
      type: "column",
      categories: new Array(MAX_POINTS + 10).fill("a"),
      series: Array.from({ length: MAX_SERIES + 5 }, (_, i) => ({ name: `${"x".repeat(400)}\n${i}`, values: [NaN, Infinity, 1] })),
      title: "  ",
      dataLabels: false,
    });
    expect(c.categories.length).toBe(MAX_POINTS);
    expect(c.series.length * c.categories.length).toBeLessThanOrEqual(MAX_CELLS);
    expect(c.series.length).toBeGreaterThan(0);
    expect(c.series[0]!.values.length).toBe(MAX_POINTS);
    expect(c.series[0]!.values.slice(0, 3)).toEqual([null, null, 1]);
    expect(Array.from(c.series[0]!.name).length).toBe(MAX_TEXT);
    expect(c.series[0]!.name).not.toContain("\n");
    expect(c.title).toBeUndefined();
    const padded = sanitize({ type: "line", categories: ["a"], series: [{ name: "s", values: [1, 2] }], dataLabels: false });
    expect(padded.categories).toEqual(["a", ""]);
    expect(sanitize(null as unknown as ChartSpec)).toEqual({ type: "column", categories: [], series: [], dataLabels: false });
  });
});
