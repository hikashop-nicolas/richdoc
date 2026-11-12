// Shapes and text boxes: their text is edited in place (each text box body is its own editing
// host), a selected shape gets a small bar (fill, outline, size, delete), and new text boxes /
// shapes are inserted at the caret. Edits are flagged on the shape (data-shape-dirty: "text",
// "look") so the adapter rewrites only what changed; untouched shapes save byte for byte.
import { t } from "../i18n";

export type ShapeKind = "textbox" | "rect" | "roundRect" | "ellipse";

export interface ShapeEditDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  mark: () => void;
  getActiveEl: () => HTMLElement;
}

const BLOCK_SEL = "[class^='docx-tb-p'], [class^='docx-tb-h'], .docx-tb-li, .docx-tb-div";

/** Flag a shape as edited ("text" / "look"), keeping both flags once set. */
export function markShapeDirty(shape: Element, kind: "text" | "look"): void {
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
  const { wrap, regions, mark, getActiveEl } = deps;
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
  bar.append(field(t("shapeFill"), fillIn), field(t("shapeNoFill"), noFill), field(t("shapeOutline"), lineIn), lineW, field(t("shapeWidthCm"), wIn), field(t("shapeHeightCm"), hIn), del);
  wrap.appendChild(bar);
  const PX_CM = 96 / 2.54;
  const box = () => shape?.querySelector<HTMLElement>(":scope > .docx-shape-box") ?? null;
  const rgbHex = (c: string): string => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
    return m ? `#${[m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("")}` : /^#[0-9a-f]{6}$/i.test(c) ? c : "#000000";
  };
  const place = () => {
    if (!shape) return;
    const r = shape.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
    bar.style.left = `${Math.max(4, Math.min(r.left - wr.left, wr.width - bar.offsetWidth - 4))}px`;
    bar.style.top = `${Math.max(4, r.top - wr.top - 40)}px`;
  };
  const sync = () => {
    const b = box();
    if (!b) return;
    const cs = getComputedStyle(b);
    const bg = b.style.backgroundColor;
    noFill.checked = !bg || bg === "transparent";
    fillIn.value = noFill.checked ? "#ffffff" : rgbHex(bg);
    const bw = parseFloat(b.style.borderTopWidth || "0");
    lineIn.value = rgbHex(b.style.borderTopColor || cs.borderTopColor);
    lineW.value = !b.style.borderTopStyle || b.style.borderTopStyle === "none" || !bw ? "0" : String(Math.min(4, Math.max(1, Math.round(bw))));
    wIn.value = (b.offsetWidth / PX_CM).toFixed(1);
    hIn.value = (b.offsetHeight / PX_CM).toFixed(1);
    const isGroup = !!b.querySelector(":scope > .docx-shape-box, :scope > img");
    for (const el of [fillIn, noFill, lineIn, lineW]) (el as HTMLInputElement).disabled = isGroup;
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
  const apply = (fn: (b: HTMLElement) => void) => {
    const b = box();
    if (!b || !shape) return;
    fn(b);
    markShapeDirty(shape, "look");
    mark();
    place();
  };
  fillIn.addEventListener("input", () => apply((b) => { noFill.checked = false; b.style.backgroundColor = fillIn.value; }));
  noFill.addEventListener("change", () => apply((b) => { b.style.backgroundColor = noFill.checked ? "" : fillIn.value; }));
  const setLine = (b: HTMLElement) => {
    const w = Number(lineW.value);
    b.style.border = w ? `${w}px solid ${lineIn.value}` : "none";
  };
  lineIn.addEventListener("input", () => apply((b) => { if (lineW.value === "0") lineW.value = "1"; setLine(b); }));
  lineW.addEventListener("change", () => apply(setLine));
  const setSize = () => apply((b) => {
    const w = parseFloat(wIn.value), h = parseFloat(hIn.value);
    if (w > 0) b.style.width = `${Math.round(w * PX_CM)}px`;
    if (h > 0) b.style.height = `${Math.round(h * PX_CM)}px`;
  });
  wIn.addEventListener("change", setSize);
  hIn.addEventListener("change", setSize);
  del.addEventListener("click", () => {
    if (!shape) return;
    const s = shape;
    select(null);
    s.remove();
    mark();
  });
  // Clicking in a shape selects it (the bar shows while its text is edited); elsewhere deselects.
  const onDown = (e: MouseEvent) => {
    const el = e.target as HTMLElement;
    if (bar.contains(el)) return;
    // A click on a shape (its frame or its text) selects it; the text keeps the caret.
    const s = el.closest?.(".docx-shape") as HTMLElement | null;
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
  const insert = (kind: ShapeKind) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !inRegion(sel.anchorNode)) return;
    const range = sel.getRangeAt(0).cloneRange();
    getActiveEl().focus({ preventScroll: true });
    range.deleteContents();
    const el = newShapeEl(kind);
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
    bar.remove();
  };
  return { insert, select, teardown };
}
