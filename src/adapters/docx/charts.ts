// Charts in the docx editor: a chart is a span (class docx-chart) holding its part's XML
// (data-chart-xml) and a rendered SVG. A chart richdoc can write from its model (specOf) is
// edited freely; any other chart keeps its shape and type, with its values, labels, names and
// title written back into its caches and embedded workbook.
import type { ChartHooks, ChartModel } from "../../core/types";
import { chartToSvg } from "./chart";
import { CHART_TYPES, LEGEND_POSITIONS, chartXml, sampleChart, sanitize, specOf } from "./chart-spec";
import type { ChartSpec, ChartType, LegendPos } from "./chart-spec";
import { readChartData, writeChartData } from "./chart-data";
import { escapeAttr } from "./shared";

const PX_PT = 0.75;

/** The chart's picture: an <img> of its SVG at the element's size. */
export function chartImg(xml: string, colors: Record<string, string>, wPx: number, hPx: number): string {
  const svg = chartToSvg(xml, colors, wPx * PX_PT, hPx * PX_PT);
  const src = svg ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}` : "";
  return `<img src="${escapeAttr(src)}" alt="" style="width:${Math.round(wPx)}px;height:${Math.round(hPx)}px">`;
}

/** The chart span for a chart part read from a file (its run passed through in `pass`). */
export function chartHtml(xml: string, part: string, pass: string, colors: Record<string, string>, wPx: number, hPx: number): string {
  return `<span class="docx-chart" contenteditable="false"${pass} data-chart-part="${escapeAttr(part)}" data-chart-xml="${escapeAttr(xml)}">${chartImg(xml, colors, wPx, hPx)}</span>`;
}

const sizeOf = (el: HTMLElement): [number, number] => {
  const img = el.querySelector("img");
  return [parseFloat(img?.style.width ?? "") || 576, parseFloat(img?.style.height ?? "") || 336];
};

export function chartHooks(colors: () => Record<string, string>): ChartHooks {
  const render = (el: HTMLElement, xml: string) => {
    const [w, h] = sizeOf(el);
    el.setAttribute("data-chart-xml", xml);
    el.innerHTML = chartImg(xml, colors(), w, h);
  };
  return {
    types: [...CHART_TYPES],
    legends: [...LEGEND_POSITIONS],
    read(el) {
      const xml = el.getAttribute("data-chart-xml");
      if (!xml) return undefined;
      const spec = specOf(xml);
      if (spec) return { title: spec.title ?? "", categories: [...spec.categories], series: spec.series.map((s) => ({ name: s.name, values: [...s.values] })), type: spec.type, legend: spec.legend, dataLabels: spec.dataLabels, fixed: false };
      const data = readChartData(xml);
      return data && { title: data.title ?? "", categories: data.categories, series: data.series, fixed: true };
    },
    update(el, m: ChartModel) {
      const xml = el.getAttribute("data-chart-xml") ?? "";
      let next: string;
      if (m.fixed) {
        const data = { title: m.title, categories: m.categories, series: m.series };
        next = writeChartData(xml, data);
        el.setAttribute("data-chart-data", JSON.stringify(data)); // for the embedded workbook on save
      } else {
        const spec: ChartSpec = { type: (CHART_TYPES.includes(m.type as ChartType) ? m.type : "column") as ChartType, categories: m.categories, series: m.series, title: m.title || undefined, legend: (m.legend || undefined) as LegendPos | undefined, dataLabels: !!m.dataLabels };
        next = chartXml(sanitize(spec));
      }
      el.setAttribute("data-chart-dirty", "1");
      render(el, next);
    },
    create() {
      const el = document.createElement("span");
      el.className = "docx-chart";
      el.contentEditable = "false";
      el.setAttribute("data-chart-new", "1");
      const xml = chartXml(sampleChart("column"));
      el.setAttribute("data-chart-xml", xml);
      el.innerHTML = chartImg(xml, colors(), 576, 336);
      return el;
    },
  };
}
