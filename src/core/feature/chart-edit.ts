// Charts: double-clicking a chart (or inserting one) opens a dialog with its title, type, legend,
// data labels and a data grid (categories x series). The adapter reads and applies the edit
// (ChartHooks); a chart it can only patch keeps its type and table size ("fixed").
import { t } from "../i18n";
import type { ChartHooks, ChartModel } from "../types";
import { makeDialogAccessible } from "./dialog-a11y";

export interface ChartEditDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  mark: () => void;
  getActiveEl: () => HTMLElement;
  hooks: ChartHooks;
}

const label = (key: string, fallback: string): string => {
  const v = t(key);
  return v && v !== key ? v : fallback;
};

export function setupChartEdit(deps: ChartEditDeps) {
  const { wrap, regions, mark, getActiveEl, hooks } = deps;
  const overlay = document.createElement("div");
  overlay.className = "docxedit-dialog-overlay";
  overlay.hidden = true;
  const panel = document.createElement("div");
  panel.className = "docxedit-dialog docxedit-chart-dialog";
  overlay.appendChild(panel);
  wrap.appendChild(overlay);
  makeDialogAccessible(overlay);

  let target: HTMLElement | null = null;
  let model: ChartModel | null = null;

  const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls = "", text = ""): HTMLElementTagNameMap[K] => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text) e.textContent = text;
    return e;
  };
  const button = (text: string, fn: () => void, cls = "docxedit-menu-item") => {
    const b = el("button", cls, text);
    b.type = "button";
    b.addEventListener("click", fn);
    return b;
  };

  const render = () => {
    if (!model) return;
    const m = model;
    panel.replaceChildren();
    panel.appendChild(el("div", "docxedit-dialog-title", t("chartEdit")));
    // Title, type, legend, data labels.
    const opts = el("div", "docxedit-chart-opts");
    const titleIn = el("input", "docxedit-name-input");
    titleIn.value = m.title;
    titleIn.placeholder = t("chartTitle");
    titleIn.addEventListener("input", () => { m.title = titleIn.value; });
    const field = (name: string, input: HTMLElement) => { const l = el("label", "docxedit-chart-field"); l.append(el("span", "", name), input); return l; };
    opts.appendChild(field(t("chartTitle"), titleIn));
    if (!m.fixed) {
      const type = el("select", "docxedit-dialog-font");
      for (const id of hooks.types) type.add(new Option(label(`chartType_${id}`, id), id));
      type.value = m.type ?? hooks.types[0]!;
      type.addEventListener("change", () => { m.type = type.value; });
      opts.appendChild(field(t("chartType"), type));
      const legend = el("select", "docxedit-dialog-font");
      legend.add(new Option(t("chartLegendNone"), ""));
      for (const id of hooks.legends) legend.add(new Option(label(`chartLegend_${id}`, id), id));
      legend.value = m.legend ?? "";
      legend.addEventListener("change", () => { m.legend = legend.value || undefined; });
      opts.appendChild(field(t("chartLegend"), legend));
      const dl = el("input");
      dl.type = "checkbox";
      dl.checked = !!m.dataLabels;
      dl.addEventListener("change", () => { m.dataLabels = dl.checked; });
      const dlField = el("label", "docxedit-pagesetup-check");
      dlField.append(dl, el("span", "", t("chartDataLabels")));
      opts.appendChild(dlField);
    } else {
      opts.appendChild(el("div", "docxedit-chart-note", t("chartFixedNote")));
    }
    panel.appendChild(opts);
    // The data grid.
    const grid = el("div", "docxedit-chart-grid");
    const table = el("table");
    const head = el("tr");
    head.appendChild(el("th", "", t("chartCategory")));
    m.series.forEach((s, j) => {
      const th = el("th");
      const nIn = el("input", "docxedit-chart-cell");
      nIn.value = s.name;
      nIn.setAttribute("aria-label", t("chartSeriesName"));
      nIn.addEventListener("input", () => { s.name = nIn.value; });
      th.appendChild(nIn);
      if (!m.fixed && m.series.length > 1) th.appendChild(button("✕", () => { m.series.splice(j, 1); render(); }, "docxedit-chart-x"));
      head.appendChild(th);
    });
    table.appendChild(head);
    m.categories.forEach((c, i) => {
      const tr = el("tr");
      const td0 = el("td");
      const cIn = el("input", "docxedit-chart-cell");
      cIn.value = c;
      cIn.setAttribute("aria-label", t("chartCategory"));
      cIn.addEventListener("input", () => { m.categories[i] = cIn.value; });
      td0.appendChild(cIn);
      if (!m.fixed && m.categories.length > 1) td0.appendChild(button("✕", () => { m.categories.splice(i, 1); for (const s of m.series) s.values.splice(i, 1); render(); }, "docxedit-chart-x"));
      tr.appendChild(td0);
      for (const s of m.series) {
        const td = el("td");
        const v = el("input", "docxedit-chart-cell docxedit-chart-num");
        v.inputMode = "decimal";
        v.value = s.values[i] === null || s.values[i] === undefined ? "" : String(s.values[i]);
        v.setAttribute("aria-label", `${s.name} ${cIn.value}`);
        v.addEventListener("input", () => {
          const n = Number(v.value.replace(",", "."));
          s.values[i] = v.value.trim() === "" || !isFinite(n) ? null : n;
        });
        td.appendChild(v);
        tr.appendChild(td);
      }
      table.appendChild(tr);
    });
    grid.appendChild(table);
    panel.appendChild(grid);
    if (!m.fixed) {
      const shape = el("div", "docxedit-dialog-row");
      shape.append(
        button(t("chartAddRow"), () => { m.categories.push(`${t("chartCategory")} ${m.categories.length + 1}`); for (const s of m.series) s.values.push(null); render(); }),
        button(t("chartAddSeries"), () => { m.series.push({ name: `${t("chartSeriesName")} ${m.series.length + 1}`, values: m.categories.map(() => null) }); render(); }),
      );
      panel.appendChild(shape);
    }
    const actions = el("div", "docxedit-dialog-row docxedit-dialog-actions");
    actions.append(button(t("cancel"), close), button(t("apply"), apply, "docxedit-menu-item docxedit-dialog-primary"));
    panel.appendChild(actions);
  };

  const close = () => {
    overlay.hidden = true;
    target = null;
    model = null;
  };
  const apply = () => {
    if (target && model) {
      hooks.update(target, model);
      mark();
    }
    close();
  };
  const open = (chart: HTMLElement) => {
    const m = hooks.read(chart);
    if (!m) return;
    target = chart;
    model = { ...m, categories: [...m.categories], series: m.series.map((s) => ({ name: s.name, values: [...s.values] })) };
    render();
    overlay.hidden = false;
    panel.querySelector<HTMLInputElement>("input")?.focus();
  };
  overlay.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); });
  overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) close(); });

  const onDbl = (e: MouseEvent) => {
    const chart = (e.target as Element | null)?.closest?.(".docx-chart") as HTMLElement | null;
    if (chart && regions.some((r) => r.contains(chart))) open(chart);
  };
  for (const r of regions) r.addEventListener("dblclick", onDbl);

  /** Insert a new chart at the caret and open its dialog. */
  const insert = () => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !regions.some((r) => r.contains(sel.anchorNode))) return;
    const range = sel.getRangeAt(0).cloneRange();
    getActiveEl().focus({ preventScroll: true });
    range.deleteContents();
    const chart = hooks.create();
    range.insertNode(chart);
    range.setStartAfter(chart);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
    open(chart);
  };

  const teardown = () => {
    for (const r of regions) r.removeEventListener("dblclick", onDbl);
    overlay.remove();
  };
  return { insert, open, teardown };
}
