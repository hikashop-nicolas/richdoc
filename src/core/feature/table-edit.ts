// Table editing, Google-Docs style: hovering a table shows a row handle on its left and a
// column handle on top (a drag grip + a "+" menu to insert/delete), and a per-cell menu
// button for merge/split. Dragging a grip reorders the row/column. Every structural change
// edits the DOM and drops the per-format skeleton (data-*-xml) so the adapter rebuilds the
// table from the DOM on save (reusing the preserved tblPr/tblGrid/tcPr / cell styles).
import { t } from "../i18n";

export interface TableEditDeps {
  wrap: HTMLElement;
  scroll: HTMLElement;
  mark: () => void;
  scheduleReflow: () => void;
}

const GRIP =
  '<svg width="10" height="14" viewBox="0 0 10 14" fill="currentColor" aria-hidden="true">' +
  "<circle cx=3 cy=3 r=1.2/><circle cx=7 cy=3 r=1.2/><circle cx=3 cy=7 r=1.2/><circle cx=7 cy=7 r=1.2/><circle cx=3 cy=11 r=1.2/><circle cx=7 cy=11 r=1.2/></svg>";
const PLUS =
  '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><path d="M8 3v10M3 8h10"/></svg>';
const CARET =
  '<svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 6l4 4 4-4z"/></svg>';

export function setupTableEdit(deps: TableEditDeps) {
  const { wrap, scroll, mark, scheduleReflow } = deps;

  let curTable: HTMLTableElement | null = null;
  let curCell: HTMLTableCellElement | null = null;
  let menuOpen = false;

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
    table.removeAttribute("data-docx-xml");
    table.removeAttribute("data-odt-xml");
    mark();
    scheduleReflow();
  };

  // --- Grid layout (each cell placed in every grid cell it spans) ----------------------
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

  // --- Structural operations (act on curCell / curTable) -------------------------------
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
    hideHandles();
    dirty(table);
  };
  const deleteCol = (): void => {
    if (!curTable || !curCell || curTable.rows[0].cells.length <= 1) return;
    const table = curTable;
    const idx = curCell.cellIndex;
    for (const tr of Array.from(table.rows)) tr.cells[idx]?.remove();
    hideHandles();
    dirty(table);
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
  const moveRow = (table: HTMLTableElement, from: number, to: number): void => {
    const rows = Array.from(table.rows);
    const src = rows[from];
    const tgt = rows[to];
    if (!src || !tgt || src === tgt) return;
    if (to > from) tgt.after(src);
    else tgt.before(src);
    dirty(table);
  };
  const moveCol = (table: HTMLTableElement, from: number, to: number): void => {
    if (from === to) return;
    for (const tr of Array.from(table.rows)) {
      const cells = Array.from(tr.cells);
      const src = cells[from];
      const tgt = cells[to];
      if (!src || !tgt) continue;
      if (to > from) tgt.after(src);
      else tgt.before(src);
    }
    dirty(table);
  };

  // --- Floating controls ----------------------------------------------------------------
  const mkHandle = (cls: string, title: string, onPlus: () => void, onDrag: (e: PointerEvent) => void): HTMLElement => {
    const h = document.createElement("div");
    h.className = `docxedit-th ${cls}`;
    h.hidden = true;
    const grip = document.createElement("span");
    grip.className = "docxedit-th-grip";
    grip.innerHTML = GRIP;
    grip.title = title;
    grip.addEventListener("pointerdown", onDrag);
    const add = document.createElement("button");
    add.type = "button";
    add.className = "docxedit-th-add";
    add.innerHTML = PLUS;
    add.title = title;
    add.addEventListener("mousedown", (e) => e.preventDefault());
    add.addEventListener("click", (e) => {
      e.stopPropagation();
      onPlus();
    });
    h.append(grip, add);
    return h;
  };

  const menu = document.createElement("div");
  menu.className = "docxedit-menu";
  menu.hidden = true;
  const showMenu = (anchor: HTMLElement, items: { label: string; fn: () => void }[]): void => {
    menu.innerHTML = "";
    for (const it of items) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "docxedit-menu-item";
      b.textContent = it.label;
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        it.fn();
        closeMenu();
        hideHandles();
      });
      menu.appendChild(b);
    }
    const ar = anchor.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    menu.hidden = false;
    menu.style.left = `${ar.left - wr.left}px`;
    menu.style.top = `${ar.bottom - wr.top + 2}px`;
    menuOpen = true;
  };
  function closeMenu(): void {
    menu.hidden = true;
    menuOpen = false;
  }

  let drag: { axis: "row" | "col"; from: number; table: HTMLTableElement } | null = null;
  const onDragMove = (): void => {}; // drop target resolved on release
  const onDragUp = (e: PointerEvent): void => {
    document.removeEventListener("pointermove", onDragMove);
    document.removeEventListener("pointerup", onDragUp);
    const d = drag;
    drag = null;
    if (!d) return;
    const el = document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null;
    const td = el?.closest?.(".docx-cell")?.closest("td") as HTMLTableCellElement | null;
    if (!td || td.closest("table") !== d.table) return;
    const tp = computeGrid(d.table).pos.get(td);
    if (!tp) return;
    if (d.axis === "row") moveRow(d.table, d.from, tp.row);
    else moveCol(d.table, d.from, tp.col);
  };
  const startDrag = (axis: "row" | "col") => (e: PointerEvent): void => {
    if (!curTable || !curCell) return;
    e.preventDefault();
    const p = computeGrid(curTable).pos.get(curCell);
    if (!p) return;
    drag = { axis, from: axis === "row" ? p.row : p.col, table: curTable };
    document.addEventListener("pointermove", onDragMove);
    document.addEventListener("pointerup", onDragUp);
  };

  const rowHandle = mkHandle(
    "docxedit-th-row",
    t("tableRow"),
    () => showMenu(rowHandle, [
      { label: t("tableRowAbove"), fn: () => insertRow(false) },
      { label: t("tableRowBelow"), fn: () => insertRow(true) },
      { label: t("tableDelRow"), fn: deleteRow },
    ]),
    startDrag("row"),
  );
  const colHandle = mkHandle(
    "docxedit-th-col",
    t("tableColumn"),
    () => showMenu(colHandle, [
      { label: t("tableColLeft"), fn: () => insertCol(false) },
      { label: t("tableColRight"), fn: () => insertCol(true) },
      { label: t("tableDelCol"), fn: deleteCol },
    ]),
    startDrag("col"),
  );
  const cellBtn = document.createElement("button");
  cellBtn.type = "button";
  cellBtn.className = "docxedit-th-cell";
  cellBtn.hidden = true;
  cellBtn.innerHTML = CARET;
  cellBtn.title = t("tableCellMenu");
  cellBtn.addEventListener("mousedown", (e) => e.preventDefault());
  cellBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    showMenu(cellBtn, [
      { label: t("tableMergeDown"), fn: mergeDown },
      { label: t("tableMergeRight"), fn: mergeRight },
      { label: t("tableSplit"), fn: splitCell },
    ]);
  });
  wrap.append(rowHandle, colHandle, cellBtn, menu);

  // --- Hover positioning ----------------------------------------------------------------
  function hideHandles(): void {
    rowHandle.hidden = true;
    colHandle.hidden = true;
    cellBtn.hidden = true;
    curCell = null;
    curTable = null;
  }
  const position = (cellDiv: HTMLElement): void => {
    if (!curTable) return;
    const cr = cellDiv.getBoundingClientRect();
    const tr = curTable.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    rowHandle.hidden = false;
    colHandle.hidden = false;
    cellBtn.hidden = false;
    // row handle: left of the table, centred on the hovered row
    rowHandle.style.left = `${tr.left - wr.left - rowHandle.offsetWidth - 4}px`;
    rowHandle.style.top = `${cr.top - wr.top + cr.height / 2 - rowHandle.offsetHeight / 2}px`;
    // column handle: above the table, centred on the hovered column
    colHandle.style.left = `${cr.left - wr.left + cr.width / 2 - colHandle.offsetWidth / 2}px`;
    colHandle.style.top = `${tr.top - wr.top - colHandle.offsetHeight - 4}px`;
    // cell menu button: top-right of the hovered cell
    cellBtn.style.left = `${cr.right - wr.left - cellBtn.offsetWidth - 2}px`;
    cellBtn.style.top = `${cr.top - wr.top + 2}px`;
  };
  const onMove = (e: MouseEvent): void => {
    if (menuOpen || drag) return;
    const tgt = e.target as HTMLElement;
    if (tgt.closest?.(".docxedit-th, .docxedit-th-cell, .docxedit-menu")) return; // over a control: keep
    const cellDiv = tgt.closest?.(".docx-cell") as HTMLElement | null;
    const table = cellDiv?.closest("table.docx-table") as HTMLTableElement | null;
    if (!cellDiv || !table || !wrap.contains(cellDiv)) {
      hideHandles();
      return;
    }
    const td = cellDiv.closest("td") as HTMLTableCellElement | null;
    if (td === curCell) return;
    curCell = td;
    curTable = table;
    position(cellDiv);
  };
  scroll.addEventListener("mousemove", onMove);
  const onDocClick = (e: MouseEvent): void => {
    if (menuOpen && !menu.contains(e.target as Node)) closeMenu();
  };
  document.addEventListener("click", onDocClick);

  const teardown = (): void => {
    scroll.removeEventListener("mousemove", onMove);
    document.removeEventListener("click", onDocClick);
    document.removeEventListener("pointermove", onDragMove);
    document.removeEventListener("pointerup", onDragUp);
  };
  return { teardown };
}
