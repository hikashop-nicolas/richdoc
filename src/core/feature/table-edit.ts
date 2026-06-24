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

  // Multi-cell selection (drag across cells). Cell-menu operations act on the whole
  // selection when one exists, otherwise on the single hovered cell.
  const selected = new Set<HTMLTableCellElement>();
  let selAnchor: HTMLTableCellElement | null = null;
  let selDragging = false;
  const targets = (): HTMLTableCellElement[] => (selected.size ? [...selected] : curCell ? [curCell] : []);
  const clearSelection = (): void => {
    for (const c of selected) c.classList.remove("rdoc-sel");
    selected.clear();
  };

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
  // Column widths live in a <colgroup> (one <col> per grid column, px width). Creating it is
  // the signal that the table was resized, so the adapters only rewrite the grid then.
  const gridColCount = (table: HTMLTableElement): number => {
    let n = 0;
    for (const p of computeGrid(table).pos.values()) n = Math.max(n, p.col + p.colspan);
    return n;
  };
  const measureGridCols = (table: HTMLTableElement): number[] => {
    const { pos } = computeGrid(table);
    const n = gridColCount(table);
    const w = new Array(n).fill(0);
    for (const [td, p] of pos) if (p.colspan === 1) w[p.col] = Math.max(w[p.col], td.getBoundingClientRect().width);
    const known = w.filter((x) => x > 0);
    const avg = known.length ? known.reduce((a, b) => a + b, 0) / known.length : 64;
    return w.map((x) => Math.round(x || avg));
  };
  const ensureColgroup = (table: HTMLTableElement): HTMLElement => {
    const n = gridColCount(table);
    let cg = table.querySelector(":scope > colgroup") as HTMLElement | null;
    if (cg && cg.children.length === n) return cg;
    const widths = measureGridCols(table);
    if (cg) cg.remove();
    cg = document.createElement("colgroup");
    for (const wpx of widths) {
      const col = document.createElement("col");
      col.style.width = `${wpx}px`;
      cg.appendChild(col);
    }
    table.insertBefore(cg, table.firstChild);
    return cg;
  };
  const moveContent = (src: HTMLTableCellElement, dst: HTMLTableCellElement): void => {
    const sd = src.querySelector(".docx-cell");
    const dd = dst.querySelector(".docx-cell");
    if (sd && dd && (sd.textContent ?? "").trim()) for (const n of Array.from(sd.childNodes)) dd.appendChild(n);
  };
  // The rectangle of grid cells spanned by two corner cells, grown to fully enclose any
  // merged cell that straddles the edge (so the selection is always a clean rectangle).
  const selectRect = (table: HTMLTableElement, a: HTMLTableCellElement, b: HTMLTableCellElement): HTMLTableCellElement[] => {
    const { pos } = computeGrid(table);
    const pa = pos.get(a);
    const pb = pos.get(b);
    if (!pa || !pb) return [];
    let r1 = Math.min(pa.row, pb.row);
    let r2 = Math.max(pa.row + pa.rowspan - 1, pb.row + pb.rowspan - 1);
    let c1 = Math.min(pa.col, pb.col);
    let c2 = Math.max(pa.col + pa.colspan - 1, pb.col + pb.colspan - 1);
    const overlaps = (p: { row: number; col: number; rowspan: number; colspan: number }): boolean =>
      p.row <= r2 && p.row + p.rowspan - 1 >= r1 && p.col <= c2 && p.col + p.colspan - 1 >= c1;
    for (let changed = true; changed; ) {
      changed = false;
      for (const p of pos.values()) {
        if (!overlaps(p)) continue;
        if (p.row < r1) (r1 = p.row), (changed = true);
        if (p.row + p.rowspan - 1 > r2) (r2 = p.row + p.rowspan - 1), (changed = true);
        if (p.col < c1) (c1 = p.col), (changed = true);
        if (p.col + p.colspan - 1 > c2) (c2 = p.col + p.colspan - 1), (changed = true);
      }
    }
    const out: HTMLTableCellElement[] = [];
    for (const [td, p] of pos) if (overlaps(p)) out.push(td);
    return out;
  };
  const mergeSelection = (): void => {
    if (!curTable || selected.size < 2) return;
    const table = curTable;
    const { pos } = computeGrid(table);
    let r1 = Infinity, r2 = -1, c1 = Infinity, c2 = -1;
    for (const td of selected) {
      const p = pos.get(td);
      if (!p) continue;
      r1 = Math.min(r1, p.row);
      r2 = Math.max(r2, p.row + p.rowspan - 1);
      c1 = Math.min(c1, p.col);
      c2 = Math.max(c2, p.col + p.colspan - 1);
    }
    let keep: HTMLTableCellElement | null = null;
    for (const td of selected) {
      const p = pos.get(td);
      if (p && p.row === r1 && p.col === c1) (keep = td);
    }
    if (!keep) return;
    keep.colSpan = c2 - c1 + 1;
    keep.rowSpan = r2 - r1 + 1;
    for (const td of selected) if (td !== keep) (moveContent(td, keep), td.remove());
    clearSelection();
    curCell = keep;
    dirty(table);
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
  // Cell borders: side-classes (rdoc-bt/br/bb/bl) on the <td>, the source of truth for both
  // display (CSS) and save. Untouched cells keep their original borders (preserved tcPr /
  // cell style); a cell only switches to class-driven borders once edited here.
  // Per-side border state. A side is either off (the cell keeps its light editing-guide
  // border) or on with a chosen colour, drawn as an inset box-shadow so it survives the
  // border-collapse conflict; the side class suppresses the guide on that one edge.
  // Per-side border state lives in data-rdoc-b{t,r,b,l} = "<w>px <style> <#color>". A side is
  // off (the cell keeps its light editing-guide border) or on, drawn as a CSS gradient line in
  // a background layer: gradients give us colour, width and solid/dashed/dotted/double while
  // being immune to the border-collapse conflict that hands a shared edge to the neighbour.
  type SideKey = "t" | "r" | "b" | "l";
  const SIDE = { t: "rdoc-bt", r: "rdoc-br", b: "rdoc-bb", l: "rdoc-bl" } as const;
  const SIDES: readonly SideKey[] = ["t", "r", "b", "l"];
  let lastBorderColor = "#000000";
  let lastBorderStyle = "solid";
  let lastBorderWidth = 1;
  const buildSpec = (): string => `${lastBorderWidth}px ${lastBorderStyle} ${lastBorderColor}`;
  const borderIcon = (on: ReadonlySet<string>): string => {
    const line = (x1: number, y1: number, x2: number, y2: number, k: string): string =>
      `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${on.has(k) ? "#fff" : "#777"}" stroke-width="${on.has(k) ? 2 : 1}"${on.has(k) ? "" : ' stroke-dasharray="1.5 1.5"'}/>`;
    return `<svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">${line(3, 3, 13, 3, "t")}${line(13, 3, 13, 13, "r")}${line(3, 13, 13, 13, "b")}${line(3, 3, 3, 13, "l")}</svg>`;
  };
  const parseSpec = (v: string): { w: number; style: string; color: string } => {
    const m = v.match(/^([\d.]+)px\s+(\w+)\s+(#[0-9a-fA-F]{3,8}|[a-z]+)/i);
    return m ? { w: parseFloat(m[1]!), style: m[2]!.toLowerCase(), color: m[3]! } : { w: 1, style: "solid", color: "#000000" };
  };
  const lineGradient = (horizontal: boolean, style: string, color: string, w: number): string => {
    const along = horizontal ? "to right" : "to bottom";
    const cross = horizontal ? "to bottom" : "to right";
    if (style === "double") {
      if (w < 3) return `linear-gradient(${color}, ${color})`;
      const band = Math.max(1, Math.round(w / 3));
      return `linear-gradient(${cross}, ${color} 0, ${color} ${band}px, transparent ${band}px, transparent ${w - band}px, ${color} ${w - band}px)`;
    }
    if (style === "dashed" || style === "dotted") {
      const dash = style === "dotted" ? Math.max(1, w) : Math.max(2, w * 2);
      const gap = style === "dotted" ? Math.max(1, w) : Math.max(2, Math.round(w * 1.2));
      return `repeating-linear-gradient(${along}, ${color} 0, ${color} ${dash}px, transparent ${dash}px, transparent ${dash + gap}px)`;
    }
    return `linear-gradient(${color}, ${color})`; // solid
  };
  const paintBorders = (cell: HTMLTableCellElement): void => {
    const imgs: string[] = [];
    const poss: string[] = [];
    const sizes: string[] = [];
    let any = false;
    for (const s of SIDES) {
      const v = cell.getAttribute(`data-rdoc-b${s}`);
      cell.classList.toggle(SIDE[s], !!v); // suppress the guide on an on side
      if (!v) continue;
      any = true;
      const { w, style, color } = parseSpec(v);
      const horizontal = s === "t" || s === "b";
      imgs.push(lineGradient(horizontal, style, color, w));
      poss.push(s === "t" ? "left top" : s === "b" ? "left bottom" : s === "l" ? "left top" : "right top");
      sizes.push(horizontal ? `100% ${w}px` : `${w}px 100%`);
    }
    cell.classList.toggle("rdoc-bordered", any);
    if (any) {
      cell.style.backgroundImage = imgs.join(", ");
      cell.style.backgroundPosition = poss.join(", ");
      cell.style.backgroundSize = sizes.join(", ");
    } else {
      cell.style.removeProperty("background-image");
      cell.style.removeProperty("background-position");
      cell.style.removeProperty("background-size");
    }
  };
  const setSide = (cell: HTMLTableCellElement, s: SideKey, spec: string | null): void => {
    if (spec) cell.setAttribute(`data-rdoc-b${s}`, spec);
    else cell.removeAttribute(`data-rdoc-b${s}`);
    paintBorders(cell);
  };
  const applyBorder = (preset: "all" | "none" | SideKey, spec: string): void => {
    const cells = targets();
    if (!cells.length || !curTable) return;
    if (preset === "all") for (const c of cells) for (const s of SIDES) setSide(c, s, spec);
    else if (preset === "none") for (const c of cells) for (const s of SIDES) setSide(c, s, null);
    else {
      const allOn = cells.every((c) => c.hasAttribute(`data-rdoc-b${preset}`)); // group toggle
      for (const c of cells) setSide(c, preset, allOn ? null : spec);
    }
    dirty(curTable);
  };
  const BORDER_PRESETS = [
    { p: "all" as const, on: new Set(["t", "r", "b", "l"]), title: t("borderAll") },
    { p: "none" as const, on: new Set<string>(), title: t("borderNone") },
    { p: "t" as const, on: new Set(["t"]), title: t("borderTop") },
    { p: "b" as const, on: new Set(["b"]), title: t("borderBottom") },
    { p: "l" as const, on: new Set(["l"]), title: t("borderLeft") },
    { p: "r" as const, on: new Set(["r"]), title: t("borderRight") },
  ];
  const showCellMenu = (): void => {
    menu.innerHTML = "";
    const items = selected.size > 1
      ? [{ label: t("tableMergeCells"), fn: mergeSelection }]
      : [
          { label: t("tableMergeDown"), fn: mergeDown },
          { label: t("tableMergeRight"), fn: mergeRight },
          { label: t("tableSplit"), fn: splitCell },
        ];
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
    const opts = document.createElement("div");
    opts.className = "docxedit-border-opts";
    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.value = lastBorderColor;
    colorInput.title = t("borderColor");
    colorInput.setAttribute("aria-label", t("borderColor"));
    colorInput.addEventListener("input", () => (lastBorderColor = colorInput.value));
    const styleSel = document.createElement("select");
    styleSel.title = t("borderStyle");
    styleSel.setAttribute("aria-label", t("borderStyle"));
    for (const [v, key] of [["solid", "bsSolid"], ["dashed", "bsDashed"], ["dotted", "bsDotted"], ["double", "bsDouble"]] as const) styleSel.add(new Option(t(key), v));
    styleSel.value = lastBorderStyle;
    styleSel.addEventListener("change", () => (lastBorderStyle = styleSel.value));
    const widthSel = document.createElement("select");
    widthSel.title = t("borderWidth");
    widthSel.setAttribute("aria-label", t("borderWidth"));
    for (const w of [1, 2, 3, 4]) widthSel.add(new Option(`${w} px`, String(w)));
    widthSel.value = String(lastBorderWidth);
    widthSel.addEventListener("change", () => (lastBorderWidth = Number(widthSel.value)));
    opts.append(colorInput, styleSel, widthSel);
    menu.appendChild(opts);
    const grid = document.createElement("div");
    grid.className = "docxedit-border-grid";
    for (const bp of BORDER_PRESETS) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "docxedit-border-btn";
      b.innerHTML = borderIcon(bp.on);
      b.title = bp.title;
      b.setAttribute("aria-label", bp.title);
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        applyBorder(bp.p, buildSpec());
      });
      grid.appendChild(b);
    }
    menu.appendChild(grid);
    const ar = cellBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    menu.hidden = false;
    menu.style.left = `${ar.left - wr.left}px`;
    menu.style.top = `${ar.bottom - wr.top + 2}px`;
    menuOpen = true;
  };
  const cellBtn = document.createElement("button");
  cellBtn.type = "button";
  cellBtn.className = "docxedit-th-cell";
  cellBtn.hidden = true;
  cellBtn.innerHTML = CARET;
  cellBtn.title = t("tableCellMenu");
  cellBtn.addEventListener("mousedown", (e) => e.preventDefault());
  cellBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    showCellMenu();
  });
  // --- Column / row resize: a thin strip over a cell boundary, dragged to resize ---------
  const colResize = document.createElement("div");
  colResize.className = "docxedit-resize col";
  colResize.hidden = true;
  const rowResize = document.createElement("div");
  rowResize.className = "docxedit-resize row";
  rowResize.hidden = true;
  let colInfo: { table: HTMLTableElement; gridCol: number } | null = null;
  let rowInfo: { table: HTMLTableElement; row: number } | null = null;
  let rdrag: { axis: "col" | "row"; table: HTMLTableElement; start: number; a: HTMLElement; aSize: number; b: HTMLElement | null; bSize: number } | null = null;
  const updateResizers = (e: MouseEvent, cellDiv: HTMLElement, table: HTMLTableElement): void => {
    const td = cellDiv.closest("td") as HTMLTableCellElement | null;
    const { pos } = computeGrid(table);
    const p = td ? pos.get(td) : undefined;
    const cr = cellDiv.getBoundingClientRect();
    const tr = table.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    const NEAR = 4;
    // The outer right/bottom edges resize the table itself (no neighbour to give back to).
    if (p && e.clientX >= cr.right - NEAR) {
      colResize.hidden = false;
      colResize.style.left = `${cr.right - wr.left - 2}px`;
      colResize.style.top = `${tr.top - wr.top}px`;
      colResize.style.width = "4px";
      colResize.style.height = `${tr.height}px`;
      colInfo = { table, gridCol: p.col + p.colspan - 1 };
    } else colResize.hidden = true;
    if (p && e.clientY >= cr.bottom - NEAR) {
      rowResize.hidden = false;
      rowResize.style.left = `${tr.left - wr.left}px`;
      rowResize.style.top = `${cr.bottom - wr.top - 2}px`;
      rowResize.style.width = `${tr.width}px`;
      rowResize.style.height = "4px";
      rowInfo = { table, row: p.row + p.rowspan - 1 };
    } else rowResize.hidden = true;
  };
  const onResizeMove = (e: PointerEvent): void => {
    if (!rdrag) return;
    const wr = wrap.getBoundingClientRect();
    if (rdrag.axis === "col") {
      const dx = e.clientX - rdrag.start;
      const MIN = 24;
      // Each side clamps to its own minimum: the dragged column follows the cursor and, once
      // the neighbour hits its minimum, the table widens rather than the drag stalling.
      rdrag.a.style.width = `${Math.round(Math.max(MIN, rdrag.aSize + dx))}px`;
      if (rdrag.b) rdrag.b.style.width = `${Math.round(Math.max(MIN, rdrag.bSize - dx))}px`;
      colResize.style.left = `${e.clientX - wr.left - 2}px`;
    } else {
      const dy = e.clientY - rdrag.start;
      rdrag.a.style.height = `${Math.round(Math.max(16, rdrag.aSize + dy))}px`;
      rowResize.style.top = `${e.clientY - wr.top - 2}px`;
    }
  };
  const onResizeUp = (): void => {
    document.removeEventListener("pointermove", onResizeMove);
    document.removeEventListener("pointerup", onResizeUp);
    colResize.classList.remove("dragging");
    rowResize.classList.remove("dragging");
    const d = rdrag;
    rdrag = null;
    if (d) dirty(d.table);
  };
  colResize.addEventListener("pointerdown", (e) => {
    if (!colInfo) return;
    e.preventDefault();
    const cg = ensureColgroup(colInfo.table);
    const a = cg.children[colInfo.gridCol] as HTMLElement;
    const b = (cg.children[colInfo.gridCol + 1] as HTMLElement) ?? null;
    rdrag = { axis: "col", table: colInfo.table, start: e.clientX, a, aSize: parseFloat(a.style.width) || 64, b, bSize: b ? parseFloat(b.style.width) || 64 : 0 };
    colResize.classList.add("dragging");
    document.addEventListener("pointermove", onResizeMove);
    document.addEventListener("pointerup", onResizeUp);
  });
  rowResize.addEventListener("pointerdown", (e) => {
    if (!rowInfo) return;
    e.preventDefault();
    const trEl = rowInfo.table.rows[rowInfo.row] as HTMLElement;
    rdrag = { axis: "row", table: rowInfo.table, start: e.clientY, a: trEl, aSize: trEl.getBoundingClientRect().height, b: null, bSize: 0 };
    rowResize.classList.add("dragging");
    document.addEventListener("pointermove", onResizeMove);
    document.addEventListener("pointerup", onResizeUp);
  });
  wrap.append(rowHandle, colHandle, cellBtn, colResize, rowResize, menu);

  // --- Hover positioning ----------------------------------------------------------------
  function hideHandles(): void {
    rowHandle.hidden = true;
    colHandle.hidden = true;
    cellBtn.hidden = true;
    colResize.hidden = true;
    rowResize.hidden = true;
    curCell = null;
    curTable = null;
  }
  const position = (cellDiv: HTMLElement): void => {
    if (!curTable || !curCell) return;
    const cr = cellDiv.getBoundingClientRect();
    const tr = curTable.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    const p = computeGrid(curTable).pos.get(curCell);
    // cell menu button: top-right of the hovered cell (always reachable, shown for any cell)
    cellBtn.hidden = false;
    cellBtn.style.left = `${cr.right - wr.left - cellBtn.offsetWidth - 2}px`;
    cellBtn.style.top = `${cr.top - wr.top + 2}px`;
    // The row/column handles sit at the table edge; only show them for a cell on that edge,
    // overlapping the table slightly so there is no gap to cross when reaching for them.
    if (p && p.col === 0) {
      rowHandle.hidden = false;
      rowHandle.style.left = `${tr.left - wr.left - rowHandle.offsetWidth + 6}px`;
      rowHandle.style.top = `${cr.top - wr.top + cr.height / 2 - rowHandle.offsetHeight / 2}px`;
    } else rowHandle.hidden = true;
    if (p && p.row === 0) {
      colHandle.hidden = false;
      colHandle.style.left = `${cr.left - wr.left + cr.width / 2 - colHandle.offsetWidth / 2}px`;
      colHandle.style.top = `${tr.top - wr.top - colHandle.offsetHeight + 6}px`;
    } else colHandle.hidden = true;
  };
  const onMove = (e: MouseEvent): void => {
    if (menuOpen || drag || selDragging || rdrag) return;
    const tgt = e.target as HTMLElement;
    if (tgt.closest?.(".docxedit-th, .docxedit-th-cell, .docxedit-menu, .docxedit-resize")) return; // over a control: keep
    const cellDiv = tgt.closest?.(".docx-cell") as HTMLElement | null;
    const table = cellDiv?.closest("table.docx-table") as HTMLTableElement | null;
    if (!cellDiv || !table || !wrap.contains(cellDiv)) {
      hideHandles();
      return;
    }
    const td = cellDiv.closest("td") as HTMLTableCellElement | null;
    if (td !== curCell) {
      curCell = td;
      curTable = table;
      position(cellDiv);
    }
    updateResizers(e, cellDiv, table); // proximity to a boundary changes within one cell too
  };
  scroll.addEventListener("mousemove", onMove);

  // --- Multi-cell drag-select -----------------------------------------------------------
  const cellUnder = (e: MouseEvent): HTMLTableCellElement | null =>
    ((e.target as HTMLElement).closest?.(".docx-cell")?.closest("td") as HTMLTableCellElement | null) ?? null;
  const onSelDown = (e: MouseEvent): void => {
    if (e.button !== 0) return;
    if ((e.target as HTMLElement).closest?.(".docxedit-th, .docxedit-th-cell, .docxedit-menu, .docxedit-resize")) return;
    const td = cellUnder(e);
    clearSelection();
    selAnchor = td && wrap.contains(td) ? td : null;
    selDragging = false;
  };
  const onSelMove = (e: MouseEvent): void => {
    if (!selAnchor || (e.buttons & 1) === 0) return;
    const table = selAnchor.closest("table.docx-table") as HTMLTableElement | null;
    const td = cellUnder(e);
    if (!table || !td || td.closest("table") !== table) return;
    if (td === selAnchor && !selDragging) return; // still in the anchor cell: allow text selection
    selDragging = true;
    window.getSelection()?.removeAllRanges();
    clearSelection();
    for (const c of selectRect(table, selAnchor, td)) {
      c.classList.add("rdoc-sel");
      selected.add(c);
    }
    curTable = table;
    curCell = selAnchor;
    if (selected.size) position((selAnchor.querySelector(".docx-cell") as HTMLElement) ?? selAnchor);
  };
  const onSelUp = (): void => {
    selAnchor = null;
    selDragging = false;
  };
  scroll.addEventListener("mousedown", onSelDown);
  scroll.addEventListener("mousemove", onSelMove);
  document.addEventListener("mouseup", onSelUp);

  const onDocClick = (e: MouseEvent): void => {
    const tgt = e.target as HTMLElement;
    if (menuOpen && !menu.contains(tgt)) closeMenu();
    // Clicking outside any table cell and outside the chrome retracts the handles and selection.
    // Keep handles/selection while the click is anywhere inside a table (a drag-select ends
    // with a synthetic click on the table element, not on a cell) or on the chrome.
    const inTable = tgt.closest?.(".docx-table");
    const inChrome = tgt.closest?.(".docxedit-th, .docxedit-th-cell, .docxedit-menu, .docxedit-resize");
    if (!inTable && !inChrome) {
      hideHandles();
      clearSelection();
    }
  };
  document.addEventListener("click", onDocClick);

  const teardown = (): void => {
    scroll.removeEventListener("mousemove", onMove);
    scroll.removeEventListener("mousedown", onSelDown);
    scroll.removeEventListener("mousemove", onSelMove);
    document.removeEventListener("mouseup", onSelUp);
    document.removeEventListener("click", onDocClick);
    document.removeEventListener("pointermove", onDragMove);
    document.removeEventListener("pointerup", onDragUp);
    document.removeEventListener("pointermove", onResizeMove);
    document.removeEventListener("pointerup", onResizeUp);
  };
  return { teardown };
}
