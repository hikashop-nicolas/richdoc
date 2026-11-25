// Shapes and text boxes: their text is edited in place (each text box body is its own editing
// host), a selected shape gets a small bar (fill, outline, size, delete), and new text boxes /
// shapes are inserted at the caret. Edits are flagged on the shape (data-shape-dirty: "text",
// "look") so the adapter rewrites only what changed; untouched shapes save byte for byte.
import { t } from "../i18n";
import { paintShapeGeom } from "./shape-geom";
import type { ShapeHooks, ShapeLook, ShapeModel } from "../types";

export type ShapeKind = "textbox" | "rect" | "roundRect" | "ellipse";

export interface ShapeEditDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  mark: () => void;
  getActiveEl: () => HTMLElement;
  /** The view zoom, to turn a pointer's screen move into the document's px. */
  getZoom?: () => number;
  /** The adapter's drawn shapes: redrawn after a resize; ready-made ones to insert. */
  hooks?: ShapeHooks;
}

const BLOCK_SEL = "[class^='docx-tb-p'], [class^='docx-tb-h'], .docx-tb-li, .docx-tb-div";

/** Flag a shape as edited ("text" / "look"), keeping both flags once set. */
export function markShapeDirty(shape: Element, kind: "text" | "look" | "pos"): void {
  const cur = new Set((shape.getAttribute("data-shape-dirty") ?? "").split(/\s+/).filter(Boolean));
  cur.add(kind);
  shape.setAttribute("data-shape-dirty", [...cur].join(" "));
}

/** A new shape's markup: the box (size, look) with an editable text body. */
export function newShapeEl(kind: ShapeKind): HTMLElement {
  const shape = document.createElement("span");
  shape.className = "docx-shape";
  shape.contentEditable = "false";
  shape.setAttribute("data-shape-new", kind);
  const box = document.createElement("span");
  box.className = "docx-shape-box";
  const isBox = kind === "textbox";
  Object.assign(box.style, {
    width: isBox ? "192px" : "144px",
    height: isBox ? "96px" : "96px",
    backgroundColor: isBox ? "#ffffff" : "#4472c4",
    border: isBox ? "1px solid #000000" : "1px solid #2f528f",
    color: isBox ? "" : "#ffffff",
    borderRadius: kind === "ellipse" ? "50%" : kind === "roundRect" ? "16.7%" : "",
  });
  const body = document.createElement("span");
  body.className = "docx-shape-body";
  body.contentEditable = "true";
  body.style.padding = "5px 10px";
  if (!isBox) body.style.justifyContent = "center";
  const p = document.createElement("span");
  p.className = "docx-tb-p";
  if (!isBox) p.style.textAlign = "center";
  p.appendChild(isBox ? document.createTextNode(t("shapeTextPlaceholder")) : document.createElement("br"));
  body.appendChild(p);
  box.appendChild(body);
  shape.appendChild(box);
  return shape;
}

export function setupShapeEdit(deps: ShapeEditDeps) {
  const { wrap, regions, mark, getActiveEl, getZoom, hooks } = deps;
  const inRegion = (n: Node | null) => !!n && regions.some((r) => r.contains(n));

  // Text edits inside a text box flag its shape.
  const onInput = (e: Event) => {
    const shape = (e.target as Element | null)?.closest?.(".docx-shape");
    if (shape) markShapeDirty(shape, "text");
  };
  // Enter splits the text box paragraph (a span) itself: a browser-made <div> inside the
  // surrounding paragraph would not survive being saved.
  const onKey = (e: KeyboardEvent) => {
    if (e.key !== "Enter" || e.isComposing) return;
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    const body = (node?.nodeType === 1 ? (node as Element) : node?.parentElement)?.closest(".docx-shape-body");
    if (!sel || !sel.rangeCount || !body) return;
    e.preventDefault();
    const range = sel.getRangeAt(0);
    range.deleteContents();
    if (e.shiftKey) {
      const br = document.createElement("br");
      range.insertNode(br);
      range.setStartAfter(br);
    } else {
      let block = (range.startContainer.nodeType === 1 ? (range.startContainer as Element) : range.startContainer.parentElement)?.closest(BLOCK_SEL) as HTMLElement | null;
      if (!block || !body.contains(block)) {
        // Loose text: wrap it in a paragraph first.
        block = document.createElement("span");
        block.className = "docx-tb-p";
        while (body.firstChild) block.appendChild(body.firstChild);
        body.appendChild(block);
      }
      const tail = document.createRange();
      tail.setStart(range.startContainer, range.startOffset);
      tail.setEnd(block, block.childNodes.length);
      const next = block.cloneNode(false) as HTMLElement;
      next.appendChild(tail.extractContents());
      if (!next.textContent) next.replaceChildren(document.createElement("br"));
      if (!block.textContent && !block.querySelector("br")) block.appendChild(document.createElement("br"));
      block.after(next);
      range.setStart(next, 0);
    }
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    const shape = body.closest(".docx-shape");
    if (shape) markShapeDirty(shape, "text");
    mark();
  };
  for (const r of regions) {
    r.addEventListener("input", onInput);
    r.addEventListener("keydown", onKey);
  }

  // --- the shape bar --------------------------------------------------------------------------
  let shape: HTMLElement | null = null;
  const bar = document.createElement("div");
  bar.className = "docxedit-imgbar docxedit-shapebar";
  bar.hidden = true;
  const field = (label: string, input: HTMLElement): HTMLElement => {
    const l = document.createElement("label");
    l.className = "docxedit-shapebar-field";
    l.title = label;
    const s = document.createElement("span");
    s.textContent = label;
    l.append(s, input);
    return l;
  };
  const color = () => { const i = document.createElement("input"); i.type = "color"; return i; };
  const fillIn = color(), lineIn = color();
  const noFill = document.createElement("input");
  noFill.type = "checkbox";
  const lineW = document.createElement("select");
  for (const [v, l] of [["0", t("shapeNoOutline")], ["1", "0.75 pt"], ["2", "1.5 pt"], ["3", "2.25 pt"], ["4", "3 pt"]] as const) lineW.add(new Option(l, v));
  const num = () => { const i = document.createElement("input"); i.type = "number"; i.min = "0.2"; i.step = "0.1"; i.className = "docxedit-shapebar-num"; return i; };
  const wIn = num(), hIn = num();
  const del = document.createElement("button");
  del.type = "button";
  del.className = "docxedit-imgbar-btn";
  del.textContent = "✕";
  del.title = t("shapeDelete");
  del.setAttribute("aria-label", t("shapeDelete"));
  // A floating shape (behind or in front of the text) moves by dragging it, or this handle when
  // its text fills it; shapes in the text flow move with the text.
  const move = document.createElement("button");
  move.type = "button";
  move.className = "docxedit-imgbar-btn docxedit-shapebar-move";
  move.textContent = "✥";
  move.title = t("shapeMove");
  move.setAttribute("aria-label", t("shapeMove"));
  bar.append(move);
  bar.append(field(t("shapeFill"), fillIn), field(t("shapeNoFill"), noFill), field(t("shapeOutline"), lineIn), lineW, field(t("shapeWidthCm"), wIn), field(t("shapeHeightCm"), hIn), del);
  wrap.appendChild(bar);
  const PX_CM = 96 / 2.54;
  const rgbHex = (c: string): string => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
    return m ? `#${[m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("")}` : /^#[0-9a-f]{6}$/i.test(c) ? c : "#000000";
  };
  // A docx shape keeps its look as its box's CSS (what the writer reads); others go through the
  // adapter's model.
  const boxOf = (s: HTMLElement) => s.querySelector<HTMLElement>(":scope > .docx-shape-box");
  const cssModel: ShapeModel = {
    selector: ".docx-shape",
    get(s) {
      const b = boxOf(s);
      if (!b) return undefined;
      const bg = b.style.backgroundColor;
      const bw = parseFloat(b.style.borderTopWidth || "0");
      const lined = !!b.style.borderTopStyle && b.style.borderTopStyle !== "none" && bw > 0;
      return {
        fill: !bg || bg === "transparent" ? null : rgbHex(bg),
        line: lined ? rgbHex(b.style.borderTopColor || getComputedStyle(b).borderTopColor) : null,
        lineW: lined ? bw : 0,
        w: b.offsetWidth, h: b.offsetHeight,
        group: !!b.querySelector(":scope > .docx-shape-box, :scope > img"),
      };
    },
    set(s, patch) {
      const b = boxOf(s);
      if (!b) return;
      if ("fill" in patch) b.style.backgroundColor = patch.fill ?? "";
      if ("lineW" in patch || "line" in patch) {
        const w = patch.lineW ?? 0;
        b.style.border = w ? `${w}px solid ${patch.line ?? "#000000"}` : "none";
      }
      if (patch.w) b.style.width = `${Math.round(patch.w)}px`;
      if (patch.h) b.style.height = `${Math.round(patch.h)}px`;
      if (patch.w || patch.h) hooks?.redraw?.(s); // the outline is worked out again for the new size
      paintShapeGeom(b);
    },
    posEl: (s) => s,
    moved: () => {},
  };
  const SEL = hooks?.model ? `.docx-shape, ${hooks.model.selector}` : ".docx-shape";
  const modelOf = (s: HTMLElement): ShapeModel => (hooks?.model && s.matches(hooks.model.selector) ? hooks.model : cssModel);
  const place = () => {
    if (!shape) return;
    const r = shape.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
    bar.style.left = `${Math.max(4, Math.min(r.left - wr.left, wr.width - bar.offsetWidth - 4))}px`;
    bar.style.top = `${Math.max(4, r.top - wr.top - 40)}px`;
  };
  const sync = () => {
    if (!shape) return;
    const lk = modelOf(shape).get(shape);
    if (!lk) return;
    noFill.checked = !lk.fill;
    fillIn.value = lk.fill ?? "#ffffff";
    lineIn.value = lk.line ?? "#000000";
    lineW.value = !lk.line || !lk.lineW ? "0" : String(Math.min(4, Math.max(1, Math.round(lk.lineW))));
    wIn.value = (lk.w / PX_CM).toFixed(1);
    hIn.value = (lk.h / PX_CM).toFixed(1);
    move.hidden = !floating(shape);
    for (const el of [fillIn, noFill, lineIn, lineW]) (el as HTMLInputElement).disabled = !!lk.group;
  };
  const select = (s: HTMLElement | null) => {
    if (s === shape) return place();
    shape?.classList.remove("docx-shape-selected");
    shape = s;
    bar.hidden = !s;
    if (!s) return;
    s.classList.add("docx-shape-selected");
    sync();
    requestAnimationFrame(place); // after the bar has a width to clamp
  };
  const apply = (patch: Partial<ShapeLook>) => {
    if (!shape) return;
    modelOf(shape).set(shape, patch);
    markShapeDirty(shape, "look");
    mark();
    place();
  };
  fillIn.addEventListener("input", () => { noFill.checked = false; apply({ fill: fillIn.value }); });
  noFill.addEventListener("change", () => apply({ fill: noFill.checked ? null : fillIn.value }));
  const lineNow = (): Partial<ShapeLook> => { const w = Number(lineW.value); return { line: w ? lineIn.value : null, lineW: w }; };
  lineIn.addEventListener("input", () => { if (lineW.value === "0") lineW.value = "1"; apply(lineNow()); });
  lineW.addEventListener("change", () => apply(lineNow()));
  const setSize = () => {
    const w = parseFloat(wIn.value), h = parseFloat(hIn.value);
    apply({ ...(w > 0 ? { w: w * PX_CM } : {}), ...(h > 0 ? { h: h * PX_CM } : {}) });
  };
  wIn.addEventListener("change", setSize);
  hIn.addEventListener("change", setSize);
  del.addEventListener("click", () => {
    if (!shape) return;
    const s = shape;
    select(null);
    s.remove();
    mark();
  });
  const floating = (s: HTMLElement | null) => !!s && /^(behind|front)$/.test(s.getAttribute("data-shape-wrap") ?? "");
  let drag: { s: HTMLElement; el: HTMLElement; x0: number; y0: number; left: number; top: number; id: number; moved: boolean } | null = null;
  const startDrag = (s: HTMLElement, e: PointerEvent) => {
    e.preventDefault(); // no text selection while dragging; this also stops the mousedown that selects
    select(s);
    const el = modelOf(s).posEl(s);
    if (!el) return;
    const pos = shapePos(el);
    drag = { s, el, x0: e.clientX, y0: e.clientY, left: pos.left, top: pos.top, id: e.pointerId, moved: false };
  };
  const onPointerDown = (e: PointerEvent) => {
    const el = e.target as HTMLElement;
    if (el === move && shape && floating(shape)) return startDrag(shape, e);
    if (bar.contains(el) || el.closest?.(".docx-shape-body")) return;
    const s = el.closest?.(SEL) as HTMLElement | null;
    if (s && inRegion(s) && floating(s) && e.button === 0) startDrag(s, e);
  };
  const onPointerMove = (e: PointerEvent) => {
    if (!drag || e.pointerId !== drag.id) return;
    const z = getZoom?.() || 1;
    const dx = (e.clientX - drag.x0) / z, dy = (e.clientY - drag.y0) / z;
    if (!drag.moved && Math.hypot(dx, dy) < 3) return;
    drag.moved = true;
    setShapePos(drag.el, Math.round(drag.left + dx), Math.round(drag.top + dy));
    place();
  };
  const onPointerUp = (e: PointerEvent) => {
    if (!drag || e.pointerId !== drag.id) return;
    const d = drag;
    drag = null;
    if (!d.moved) return;
    modelOf(d.s).moved(d.s);
    markShapeDirty(d.s, "pos");
    mark();
    place();
  };
  document.addEventListener("pointerdown", onPointerDown);
  document.addEventListener("pointermove", onPointerMove);
  document.addEventListener("pointerup", onPointerUp);
  document.addEventListener("pointercancel", onPointerUp);

  // Clicking in a shape selects it (the bar shows while its text is edited); elsewhere deselects.
  const onDown = (e: MouseEvent) => {
    const el = e.target as HTMLElement;
    if (bar.contains(el)) return;
    // A click on a shape (its frame or its text) selects it; the text keeps the caret.
    const s = el.closest?.(SEL) as HTMLElement | null;
    select(s && inRegion(s) ? s : null);
  };
  document.addEventListener("mousedown", onDown);
  wrap.addEventListener("keydown", (e) => {
    if (!shape || bar.contains(e.target as Node)) return;
    if (e.key === "Escape") select(null);
    else if ((e.key === "Delete" || e.key === "Backspace") && !(e.target as Element).closest?.(".docx-shape-body")) {
      e.preventDefault();
      del.click();
    }
  });

  /** Insert a new text box or shape at the caret. */
  const insert = (kind: string) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !inRegion(sel.anchorNode)) return;
    const range = sel.getRangeAt(0).cloneRange();
    getActiveEl().focus({ preventScroll: true });
    range.deleteContents();
    const basic = ["textbox", "rect", "roundRect", "ellipse"].includes(kind);
    const el = basic || !hooks?.create ? newShapeEl(basic ? (kind as ShapeKind) : "rect") : hooks.create(kind);
    range.insertNode(el);
    range.setStartAfter(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
    select(el);
  };

  const teardown = () => {
    document.removeEventListener("mousedown", onDown);
    document.removeEventListener("pointerdown", onPointerDown);
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", onPointerUp);
    document.removeEventListener("pointercancel", onPointerUp);
    bar.remove();
  };
  return { insert, select, teardown };
}

/** A floating shape's offset (px, as its left / top). */
export function shapePos(s: HTMLElement): { left: number; top: number } {
  return { left: parseFloat(s.style.left) || 0, top: parseFloat(s.style.top) || 0 };
}

export function setShapePos(s: HTMLElement, left: number, top: number): void {
  s.style.left = `${left}px`;
  s.style.top = `${top}px`;
}
