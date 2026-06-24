// Table editing: a small floating toolbar shown while the caret is in a table cell, with
// insert/delete row and column. Each operation edits the DOM table and drops its
// data-*-xml skeleton, so on save the adapter rebuilds the table from the DOM (reusing the
// per-element tblPr/tblGrid/tcPr the read preserved, so styling survives the edit).
import { t } from "../i18n";

export interface TableEditDeps {
  wrap: HTMLElement;
  scroll: HTMLElement;
  mark: () => void;
  scheduleReflow: () => void;
}

export function setupTableEdit(deps: TableEditDeps) {
  const { wrap, scroll, mark, scheduleReflow } = deps;

  const bar = document.createElement("div");
  bar.className = "docxedit-table-bar";
  bar.hidden = true;

  let curTable: HTMLTableElement | null = null;
  let curCell: HTMLTableCellElement | null = null;

  const newCell = (): HTMLTableCellElement => {
    const td = document.createElement("td");
    const div = document.createElement("div");
    div.className = "docx-cell";
    div.contentEditable = "true";
    div.innerHTML = "<br>";
    td.appendChild(div);
    return td;
  };
  const dirty = (table: HTMLTableElement): void => {
    // Structurally changed: drop the per-format skeleton so the adapter rebuilds from the DOM.
    table.removeAttribute("data-docx-xml");
    table.removeAttribute("data-odt-xml");
    mark();
    scheduleReflow();
  };

  const insertRow = (below: boolean): void => {
    if (!curTable || !curCell) return;
    const tr = curCell.parentElement as HTMLTableRowElement;
    const nr = document.createElement("tr");
    for (let i = 0; i < tr.cells.length; i++) nr.appendChild(newCell());
    tr.parentElement?.insertBefore(nr, below ? tr.nextSibling : tr);
    dirty(curTable);
  };
  const insertCol = (right: boolean): void => {
    if (!curTable || !curCell) return;
    const idx = curCell.cellIndex;
    for (const tr of Array.from(curTable.rows)) {
      const ref = tr.cells[idx] ?? null;
      tr.insertBefore(newCell(), right ? (ref ? ref.nextSibling : null) : ref);
    }
    dirty(curTable);
  };
  // Grid layout of a DOM table: each td placed in every grid cell it spans, so neighbours
  // (right / below) can be found correctly even with existing spans.
  type Pos = { row: number; col: number; rowspan: number; colspan: number };
  const computeGrid = (table: HTMLTableElement): { pos: Map<HTMLTableCellElement, Pos>; at: (HTMLTableCellElement | undefined)[][] } => {
    const pos = new Map<HTMLTableCellElement, Pos>();
    const at: (HTMLTableCellElement | undefined)[][] = Array.from(table.rows, () => []);
    Array.from(table.rows).forEach((tr, r) => {
      let col = 0;
      for (const td of Array.from(tr.cells)) {
        while (at[r][col]) col++;
        const cs = td.colSpan || 1;
        const rs = td.rowSpan || 1;
        pos.set(td, { row: r, col, rowspan: rs, colspan: cs });
        for (let dr = 0; dr < rs; dr++) for (let dc = 0; dc < cs; dc++) if (at[r + dr]) at[r + dr][col + dc] = td;
        col += cs;
      }
    });
    return { pos, at };
  };
  const moveContent = (src: HTMLTableCellElement, dst: HTMLTableCellElement): void => {
    const sd = src.querySelector(".docx-cell");
    const dd = dst.querySelector(".docx-cell");
    if (sd && dd && (sd.textContent ?? "").trim()) for (const n of Array.from(sd.childNodes)) dd.appendChild(n);
  };
  const mergeRight = (): void => {
    if (!curTable || !curCell) return;
    const { pos, at } = computeGrid(curTable);
    const p = pos.get(curCell);
    const right = p ? at[p.row]?.[p.col + p.colspan] : undefined;
    const rp = right && pos.get(right);
    if (!p || !right || right === curCell || !rp || rp.row !== p.row || rp.rowspan !== p.rowspan) return;
    curCell.colSpan = p.colspan + rp.colspan;
    moveContent(right, curCell);
    right.remove();
    dirty(curTable);
  };
  const mergeDown = (): void => {
    if (!curTable || !curCell) return;
    const { pos, at } = computeGrid(curTable);
    const p = pos.get(curCell);
    const below = p ? at[p.row + p.rowspan]?.[p.col] : undefined;
    const bp = below && pos.get(below);
    if (!p || !below || !bp || bp.col !== p.col || bp.colspan !== p.colspan) return;
    curCell.rowSpan = p.rowspan + bp.rowspan;
    moveContent(below, curCell);
    below.remove();
    dirty(curTable);
  };
  const splitCell = (): void => {
    if (!curTable || !curCell) return;
    const { pos } = computeGrid(curTable);
    const p = pos.get(curCell);
    if (!p || (p.colspan <= 1 && p.rowspan <= 1)) return;
    const { colspan: cs, rowspan: rs, col, row } = p;
    curCell.colSpan = 1;
    curCell.rowSpan = 1;
    for (let k = 1; k < cs; k++) curCell.parentElement?.insertBefore(newCell(), curCell.nextSibling);
    const rows = Array.from(curTable.rows);
    for (let dr = 1; dr < rs; dr++) {
      const tr = rows[row + dr];
      if (!tr) continue;
      const ref = Array.from(tr.cells).find((td) => (pos.get(td)?.col ?? 0) >= col) ?? null;
      for (let k = 0; k < cs; k++) tr.insertBefore(newCell(), ref);
    }
    dirty(curTable);
  };

  const deleteRow = (): void => {
    if (!curTable || !curCell || curTable.rows.length <= 1) return;
    const table = curTable;
    (curCell.parentElement as HTMLTableRowElement).remove();
    hide();
    dirty(table);
  };
  const deleteCol = (): void => {
    if (!curTable || !curCell || curTable.rows[0].cells.length <= 1) return;
    const table = curTable;
    const idx = curCell.cellIndex;
    for (const tr of Array.from(table.rows)) tr.cells[idx]?.remove();
    hide();
    dirty(table);
  };

  const btn = (svg: string, title: string, fn: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.innerHTML = svg;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the cell's caret
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      fn();
    });
    return b;
  };
  // Clear SVG icons: a table-body box with a +/− or arrow showing which edge/axis is affected.
  const I = (body: string): string => `<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;
  const icRowAbove = I('<rect x="2.5" y="8" width="11" height="5.5" rx="1"/><path d="M8 2v4M6.1 3.9h3.8"/>');
  const icRowBelow = I('<rect x="2.5" y="2.5" width="11" height="5.5" rx="1"/><path d="M8 10v4M6.1 12.1h3.8"/>');
  const icColLeft = I('<rect x="8" y="2.5" width="5.5" height="11" rx="1"/><path d="M2 8h4M3.9 6.1v3.8"/>');
  const icColRight = I('<rect x="2.5" y="2.5" width="5.5" height="11" rx="1"/><path d="M10 8h4M12.1 6.1v3.8"/>');
  const icMergeDown = I('<rect x="3" y="2.5" width="10" height="4" rx="1"/><rect x="3" y="9.5" width="10" height="4" rx="1"/><path d="M8 6.5v3M6.6 8.2 8 9.6l1.4-1.4"/>');
  const icMergeRight = I('<rect x="2.5" y="3" width="4" height="10" rx="1"/><rect x="9.5" y="3" width="4" height="10" rx="1"/><path d="M6.5 8h3M8.2 6.6 9.6 8 8.2 9.4"/>');
  const icSplit = I('<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M8 2.5v11M2.5 8h11" stroke-dasharray="1.6 1.6"/>');
  const icDelRow = I('<rect x="2.5" y="6" width="11" height="4" rx="1"/><path d="M5.5 8h5"/>');
  const icDelCol = I('<rect x="6" y="2.5" width="4" height="7.5" rx="1"/><path d="M8 11v3M6.4 12.4 8 14l1.6-1.6"/>');
  bar.append(
    btn(icRowAbove, t("tableRowAbove"), () => insertRow(false)),
    btn(icRowBelow, t("tableRowBelow"), () => insertRow(true)),
    btn(icColLeft, t("tableColLeft"), () => insertCol(false)),
    btn(icColRight, t("tableColRight"), () => insertCol(true)),
    btn(icMergeDown, t("tableMergeDown"), mergeDown),
    btn(icMergeRight, t("tableMergeRight"), mergeRight),
    btn(icSplit, t("tableSplit"), splitCell),
    btn(icDelRow, t("tableDelRow"), deleteRow),
    btn(icDelCol, t("tableDelCol"), deleteCol),
  );
  wrap.appendChild(bar);

  const place = (): void => {
    if (!curTable) return;
    const tr = curTable.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    bar.hidden = false;
    bar.style.left = `${tr.left - wr.left}px`;
    bar.style.top = `${tr.top - wr.top - bar.offsetHeight - 4}px`;
  };
  function hide(): void {
    bar.hidden = true;
    curTable = null;
    curCell = null;
  }
  const refresh = (): void => {
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.anchorNode : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
    const cell = el?.closest?.(".docx-cell") as HTMLElement | null;
    if (cell && wrap.contains(cell)) {
      curCell = cell.closest("td") as HTMLTableCellElement | null;
      curTable = cell.closest("table.docx-table") as HTMLTableElement | null;
      if (curTable) place();
      else hide();
    } else {
      hide();
    }
  };
  let timer = 0;
  const sched = (): void => {
    window.clearTimeout(timer);
    timer = window.setTimeout(refresh, 30);
  };
  document.addEventListener("selectionchange", sched);
  scroll.addEventListener("scroll", place); // keep it pinned to the table while scrolling

  const teardown = (): void => {
    document.removeEventListener("selectionchange", sched);
    scroll.removeEventListener("scroll", place);
    window.clearTimeout(timer);
  };
  return { teardown };
}
