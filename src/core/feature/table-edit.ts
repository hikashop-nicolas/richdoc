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
    td.setAttribute("style", "border:1px solid #999;padding:0;vertical-align:top");
    const div = document.createElement("div");
    div.className = "docx-cell";
    div.contentEditable = "true";
    div.setAttribute("style", "padding:3px 6px;min-height:1.2em;outline:none");
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

  const btn = (label: string, title: string, fn: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the cell's caret
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      fn();
    });
    return b;
  };
  bar.append(
    btn("↥", t("tableRowAbove"), () => insertRow(false)),
    btn("↧", t("tableRowBelow"), () => insertRow(true)),
    btn("⇤", t("tableColLeft"), () => insertCol(false)),
    btn("⇥", t("tableColRight"), () => insertCol(true)),
    btn("⌦", t("tableDelRow"), deleteRow),
    btn("⌫", t("tableDelCol"), deleteCol),
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
