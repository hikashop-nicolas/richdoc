// Shapes and text boxes: their text is edited in place (each text box body is its own editing
// host), a selected shape gets a small bar (layout, fill, outline, size, delete) and moves by
// dragging when it floats, and new text boxes / shapes are inserted at the caret. Edits are
// flagged on the shape (data-shape-dirty: "text", "look", "pos", "wrap", "z") so the adapter
// rewrites only what changed; untouched shapes save byte for byte.
import { t } from "../i18n";
import { paintShapeGeom } from "./shape-geom";
import { ALIGN_ICONS, ORDER_ICONS, WRAP_ICONS } from "./image-layout";
import { behindAt as floatBehind, dropWrapped, ghostOf, markShapeDirty, originOf, restack as restackFloat, setZ, stack, topZ } from "./stacking";
import type { ShapeHooks, ShapeLook, ShapeModel, ShapeWrap } from "../types";

export { markShapeDirty };

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
  const num = (min: string, step: string) => { const i = document.createElement("input"); i.type = "number"; i.min = min; i.step = step; i.className = "docxedit-shapebar-num"; return i; };
  const lineW = num("0", "0.25");
  lineW.max = "20";
  const wIn = num("0.2", "0.1"), hIn = num("0.2", "0.1");
  const btn = (text: string, title: string, cls = ""): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `docxedit-imgbar-btn ${cls}`.trim();
    b.innerHTML = text;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the caret in the shape's text
    return b;
  };
  const del = btn("✕", t("shapeDelete"));
  // A floating or wrapped shape moves by dragging it, or this handle; one in line with the text
  // moves with the text.
  const move = btn("✥", t("shapeMove"), "docxedit-shapebar-move");
  // Layout: how the shape sits in the text, its alignment and its stacking order, in a menu.
  const layoutBtn = btn(WRAP_ICONS.inline, t("shapeLayout"), "docxedit-shapebar-layout");
  layoutBtn.setAttribute("aria-haspopup", "menu");
  layoutBtn.setAttribute("aria-expanded", "false");
  const menu = document.createElement("div");
  menu.className = "docxedit-menu docxedit-shapebar-menu";
  menu.setAttribute("role", "menu");
  menu.hidden = true;
  const WRAPS: [ShapeWrap, string][] = [["inline", "wrapInline"], ["square", "wrapSquare"], ["tight", "wrapTight"], ["topbottom", "wrapTopBottom"], ["behind", "wrapBehind"], ["front", "wrapFront"]];
  const item = (icon: string, label: string, fn: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item docxedit-shapebar-menu-item";
    b.setAttribute("role", "menuitemradio");
    b.innerHTML = `${icon}<span></span>`;
    b.lastElementChild!.textContent = label;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => { e.stopPropagation(); fn(); });
    return b;
  };
  const sepEl = () => { const s = document.createElement("div"); s.className = "docxedit-shapebar-menu-sep"; return s; };
  const wrapItems = new Map<ShapeWrap, HTMLButtonElement>();
  for (const [w, key] of WRAPS) {
    const b = item(WRAP_ICONS[w], t(key), () => setWrap(w));
    wrapItems.set(w, b);
    menu.appendChild(b);
  }
  const alignSep = sepEl();
  const alignRow = document.createElement("div");
  alignRow.className = "docxedit-shapebar-menu-row";
  const alignBtns = new Map<"left" | "center" | "right", HTMLButtonElement>();
  for (const [a, key] of [["left", "alignLeft"], ["center", "alignCenter"], ["right", "alignRight"]] as const) {
    const b = btn(ALIGN_ICONS[a], t(key));
    b.addEventListener("click", (e) => { e.stopPropagation(); setAlign(a); });
    alignBtns.set(a, b);
    alignRow.appendChild(b);
  }
  const orderSep = sepEl();
  const front = item(ORDER_ICONS.front, t("shapeBringFront"), () => restack(1));
  const back = item(ORDER_ICONS.back, t("shapeSendBack"), () => restack(-1));
  for (const b of [front, back]) b.setAttribute("role", "menuitem");
  menu.append(alignSep, alignRow, orderSep, front, back);
  bar.append(move, layoutBtn);
  bar.append(field(t("shapeFill"), fillIn), field(t("shapeNoFill"), noFill), field(t("shapeOutline"), lineIn), field(t("shapeOutlineWidthPt"), lineW), field(t("shapeWidthCm"), wIn), field(t("shapeHeightCm"), hIn), del);
  wrap.append(bar, menu);
  const PX_CM = 96 / 2.54, PX_PT = 4 / 3;
  const rgbHex = (c: string): string => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
    return m ? `#${[m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("")}` : /^#[0-9a-f]{6}$/i.test(c) ? c : "#000000";
  };
  const wrapOf = (s: HTMLElement): ShapeWrap => (s.getAttribute("data-shape-wrap") as ShapeWrap | null) ?? "inline";
  const floating = (s: HTMLElement | null) => !!s && /^(behind|front)$/.test(s.getAttribute("data-shape-wrap") ?? "");
  const zoom = () => getZoom?.() || 1;

  // A docx shape keeps its look as its box's CSS (what the writer reads) and its place as
  // data-shape-x / -y: the offset from its reference (the page or its margins across, the page,
  // its margins or its paragraph down), drawn as left / top from wherever it is positioned.
  const boxOf = (s: HTMLElement) => s.querySelector<HTMLElement>(":scope > .docx-shape-box");
  const origin = (s: HTMLElement) => originOf(s, zoom());
  /** Place a floating shape from its offsets. A docx axis aligned rather than offset (no
   *  data-shape-x) is placed by its alignment across its reference. */
  const placeCss = (s: HTMLElement) => {
    if (!floating(s)) return;
    const o = origin(s);
    if (!s.matches(".docx-shape")) {
      // An odt drawing: its box sits at its offsets (data-shape-bx / -by) from its references.
      const box = s.querySelector<HTMLElement>(":scope > .odt-shape-box");
      const bx = parseFloat(s.getAttribute("data-shape-bx") ?? ""), by = parseFloat(s.getAttribute("data-shape-by") ?? "");
      if (box && Number.isFinite(bx) && Number.isFinite(by)) setShapePos(box, Math.round((bx + o.x) * 100) / 100, Math.round((by + o.y) * 100) / 100);
      return;
    }
    let x = parseFloat(s.getAttribute("data-shape-x") ?? "");
    if (!Number.isFinite(x)) {
      const docEl = s.closest<HTMLElement>(".docxedit-doc");
      const ds = docEl ? getComputedStyle(docEl) : null;
      const page = s.getAttribute("data-shape-relh") === "page";
      const fw = docEl ? (page ? docEl.clientWidth : docEl.clientWidth - (parseFloat(ds!.paddingLeft) || 0) - (parseFloat(ds!.paddingRight) || 0)) : 0;
      const w = boxOf(s)?.offsetWidth ?? 0;
      const a = s.getAttribute("data-shape-alignh");
      x = a === "center" ? (fw - w) / 2 : a === "right" ? fw - w : 0;
    }
    const y = parseFloat(s.getAttribute("data-shape-y") ?? "") || 0;
    setShapePos(s, Math.round(x + o.x), Math.round(y + o.y));
  };
  // A group's members are boxes of their own; its look is theirs, edited all at once.
  const members = (b: HTMLElement) => Array.from(b.querySelectorAll<HTMLElement>(".docx-shape-box")).filter((m) => !m.querySelector(":scope > .docx-shape-box, :scope > img"));
  const isGroup = (b: HTMLElement) => !!b.querySelector(":scope > .docx-shape-box, :scope > img");
  const lookOf = (b: HTMLElement) => {
    const bg = b.style.backgroundColor;
    const bw = parseFloat(b.style.borderTopWidth || "0");
    const lined = !!b.style.borderTopStyle && b.style.borderTopStyle !== "none" && bw > 0;
    return {
      fill: !bg || bg === "transparent" ? null : rgbHex(bg),
      line: lined ? rgbHex(b.style.borderTopColor || getComputedStyle(b).borderTopColor) : null,
      lineW: lined ? bw : 0,
    };
  };
  const paint = (b: HTMLElement, patch: Partial<ShapeLook>) => {
    if ("fill" in patch) b.style.backgroundColor = patch.fill ?? "";
    if ("lineW" in patch || "line" in patch) {
      const w = patch.lineW ?? 0;
      b.style.border = w ? `${Math.round(w * 100) / 100}px solid ${patch.line ?? "#000000"}` : "none";
    }
    paintShapeGeom(b);
  };
  const cssModel: ShapeModel = {
    selector: ".docx-shape",
    get(s) {
      const b = boxOf(s);
      if (!b) return undefined;
      if (isGroup(b)) {
        const ms = members(b);
        const first = ms.find((m) => lookOf(m).fill) ?? ms[0];
        return { ...(first ? lookOf(first) : { fill: null, line: null, lineW: 0 }), w: b.offsetWidth, h: b.offsetHeight, group: true, groupLook: ms.length > 0 };
      }
      return { ...lookOf(b), w: b.offsetWidth, h: b.offsetHeight };
    },
    set(s, patch) {
      const b = boxOf(s);
      if (!b) return;
      if (isGroup(b)) {
        for (const m of members(b)) paint(m, patch);
        if (patch.w || patch.h) {
          // Its members stretch with it (their places and sizes), and are drawn again.
          const sx = patch.w ? patch.w / (parseFloat(b.style.width) || b.offsetWidth || patch.w) : 1;
          const sy = patch.h ? patch.h / (parseFloat(b.style.height) || b.offsetHeight || patch.h) : 1;
          for (const m of Array.from(b.querySelectorAll<HTMLElement>(".docx-shape-box, img"))) {
            if (m.style.position !== "absolute") continue;
            for (const [k, f] of [["left", sx], ["top", sy], ["width", sx], ["height", sy]] as const) {
              const v = parseFloat(m.style[k]);
              if (Number.isFinite(v)) m.style[k] = `${Math.round(v * f * 100) / 100}px`;
            }
          }
          if (patch.w) b.style.width = `${Math.round(patch.w)}px`;
          if (patch.h) b.style.height = `${Math.round(patch.h)}px`;
          hooks?.redraw?.(s);
        }
        return;
      }
      paint(b, patch);
      if (patch.w) b.style.width = `${Math.round(patch.w)}px`;
      if (patch.h) b.style.height = `${Math.round(patch.h)}px`;
      if (patch.w || patch.h) { hooks?.redraw?.(s); paintShapeGeom(b); } // the outline is worked out again for the new size
    },
    posEl: (s) => s,
    moved(s) {
      const o = origin(s), p = shapePos(s);
      s.setAttribute("data-shape-x", String(Math.round(p.left - o.x)));
      s.setAttribute("data-shape-y", String(Math.round(p.top - o.y)));
      s.removeAttribute("data-shape-alignh");
    },
    setWrap(s, w, align) {
      s.style.left = s.style.top = s.style.zIndex = "";
      if (w === "inline") {
        for (const a of ["wrap", "align", "alignh", "relh", "relv", "x", "y"]) s.removeAttribute(`data-shape-${a}`);
        return;
      }
      s.setAttribute("data-shape-wrap", w);
      s.setAttribute("data-shape-align", align);
      if (w === "behind" || w === "front") {
        // Placed from its column and paragraph, as Word places a shape it floats.
        if (!s.hasAttribute("data-shape-relv")) { s.setAttribute("data-shape-relh", "margin"); s.setAttribute("data-shape-relv", "para"); }
        if (!s.hasAttribute("data-shape-x") && !s.hasAttribute("data-shape-alignh")) s.setAttribute("data-shape-x", "0");
        if (!s.hasAttribute("data-shape-y")) s.setAttribute("data-shape-y", "0");
        placeCss(s);
      } else {
        for (const a of ["alignh", "relh", "relv", "x", "y"]) s.removeAttribute(`data-shape-${a}`);
      }
    },
    wraps: () => WRAPS.map(([w]) => w),
  };
  const SEL = hooks?.model ? `.docx-shape, ${hooks.model.selector}` : ".docx-shape";
  const modelOf = (s: HTMLElement): ShapeModel => (hooks?.model && s.matches(hooks.model.selector) ? hooks.model : cssModel);
  const allShapes = () => regions.flatMap((r) => Array.from(r.querySelectorAll<HTMLElement>(SEL))).filter((s) => !s.parentElement?.closest(SEL));

  /** Lay the floating shapes out from their offsets, and stack every float (shapes and
   *  pictures) in its order. */
  const layout = () => {
    for (const s of allShapes()) placeCss(s);
    stack(regions);
  };
  requestAnimationFrame(layout);

  const place = () => {
    if (!shape) return;
    const r = (shape.querySelector(":scope > .odt-shape-box, :scope > .docx-shape-box") ?? shape).getBoundingClientRect(), wr = wrap.getBoundingClientRect();
    bar.style.left = `${Math.max(4, Math.min(r.left - wr.left, wr.width - bar.offsetWidth - 4))}px`;
    // Above the shape, or below it when there is no room above.
    const above = r.top - wr.top - bar.offsetHeight - 6;
    bar.style.top = `${above >= 4 ? above : Math.min(r.bottom - wr.top + 6, wr.height - bar.offsetHeight - 4)}px`;
    if (!menu.hidden) placeMenu();
  };
  const placeMenu = () => {
    const r = layoutBtn.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
    menu.style.left = `${Math.max(4, Math.min(r.left - wr.left, wr.width - menu.offsetWidth - 4))}px`;
    menu.style.top = `${r.bottom - wr.top + 4}px`;
  };
  const sync = () => {
    if (!shape) return;
    const m = modelOf(shape);
    const lk = m.get(shape);
    if (!lk) return;
    noFill.checked = !lk.fill;
    fillIn.value = lk.fill ?? "#ffffff";
    lineIn.value = lk.line ?? "#000000";
    lineW.value = !lk.line || !lk.lineW ? "0" : String(Math.round((lk.lineW / PX_PT) * 100) / 100);
    wIn.value = (lk.w / PX_CM).toFixed(1);
    hIn.value = (lk.h / PX_CM).toFixed(1);
    move.hidden = !floating(shape) && !wrapped(shape);
    for (const el of [fillIn, noFill, lineIn, lineW]) (el as HTMLInputElement).disabled = !!lk.group && !lk.groupLook;
    for (const el of [wIn, hIn]) el.disabled = !!lk.fixedSize;
    const w = wrapOf(shape), can = new Set(m.wraps(shape));
    layoutBtn.innerHTML = WRAP_ICONS[w];
    for (const [k, b] of wrapItems) {
      b.hidden = !can.has(k);
      b.setAttribute("aria-checked", String(k === w));
      b.classList.toggle("is-on", k === w);
    }
    const aligned = w === "square" || w === "tight" || w === "topbottom";
    alignSep.hidden = alignRow.hidden = !aligned;
    const a = shape.getAttribute("data-shape-align") ?? "left";
    for (const [k, b] of alignBtns) b.classList.toggle("is-on", k === a);
    orderSep.hidden = front.hidden = back.hidden = !floating(shape);
  };
  const closeMenu = () => { menu.hidden = true; layoutBtn.setAttribute("aria-expanded", "false"); };
  layoutBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    menu.hidden = !menu.hidden;
    layoutBtn.setAttribute("aria-expanded", String(!menu.hidden));
    if (!menu.hidden) { sync(); placeMenu(); }
  });
  const select = (s: HTMLElement | null) => {
    if (s === shape) return place();
    shape?.classList.remove("docx-shape-selected");
    shape = s;
    bar.hidden = !s;
    closeMenu();
    if (!s) return;
    s.classList.add("docx-shape-selected");
    sync();
    requestAnimationFrame(place); // after the bar has a width to clamp
  };
  const apply = (patch: Partial<ShapeLook>) => {
    if (!shape) return;
    modelOf(shape).set(shape, patch);
    markShapeDirty(shape, "look");
    layout(); // a redrawn shape is placed and stacked again
    mark();
    place();
  };
  fillIn.addEventListener("input", () => { noFill.checked = false; apply({ fill: fillIn.value }); });
  noFill.addEventListener("change", () => apply({ fill: noFill.checked ? null : fillIn.value }));
  const lineNow = (): Partial<ShapeLook> => { const w = Math.max(0, parseFloat(lineW.value) || 0) * PX_PT; return { line: w ? lineIn.value : null, lineW: w }; };
  lineIn.addEventListener("input", () => { if (!(parseFloat(lineW.value) > 0)) lineW.value = "0.75"; apply(lineNow()); });
  lineW.addEventListener("change", () => apply(lineNow()));
  const setSize = () => {
    const w = parseFloat(wIn.value), h = parseFloat(hIn.value);
    apply({ ...(w > 0 ? { w: w * PX_CM } : {}), ...(h > 0 ? { h: h * PX_CM } : {}) });
    layout();
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

  /** Move a floating shape by (dx, dy) document px. */
  const nudge = (s: HTMLElement, dx: number, dy: number) => {
    const m = modelOf(s);
    const el = m.posEl(s);
    if (!el || (!dx && !dy)) return;
    const p = shapePos(el);
    setShapePos(el, Math.round(p.left + dx), Math.round(p.top + dy));
    m.moved(s);
    markShapeDirty(s, "pos");
  };
  /** Change how a shape sits in the text. A shape that starts floating stays where it was drawn. */
  const setWrap = (w: ShapeWrap, s = shape) => {
    if (!s || !modelOf(s).wraps(s).includes(w)) return;
    const m = modelOf(s);
    const was = wrapOf(s);
    const before = (m.posEl(s) ?? s).getBoundingClientRect();
    const align = (s.getAttribute("data-shape-align") as "left" | "center" | "right" | null) ?? "left";
    m.setWrap(s, w, align);
    if ((w === "behind" || w === "front") && !(was === "behind" || was === "front")) {
      if (!(Number(s.getAttribute("data-shape-z")) > 0)) setZ(s, topZ(regions)); // a shape newly floating goes on top
      const after = (m.posEl(s) ?? s).getBoundingClientRect();
      nudge(s, (before.left - after.left) / zoom(), (before.top - after.top) / zoom());
    }
    markShapeDirty(s, "wrap");
    layout();
    mark();
    if (s === shape) { sync(); place(); }
  };
  const setAlign = (a: "left" | "center" | "right") => {
    if (!shape) return;
    modelOf(shape).setWrap(shape, wrapOf(shape), a);
    markShapeDirty(shape, "wrap");
    layout();
    mark();
    sync();
    place();
  };
  const restack = (dir: 1 | -1) => {
    if (!shape || !floating(shape)) return;
    restackFloat(shape, dir, regions);
    mark();
  };

  // A floating shape moves itself; a wrapped one (square, tight, top and bottom) shows where it
  // will land (flow) and moves on the drop, to the paragraph there.
  let drag: { s: HTMLElement; el: HTMLElement; x0: number; y0: number; left: number; top: number; id: number; moved: boolean; text?: { x: number; y: number }; flow?: ReturnType<typeof ghostOf> | null } | null = null;
  const wrapped = (s: HTMLElement) => /^(square|tight|topbottom)$/.test(s.getAttribute("data-shape-wrap") ?? "");
  const startDrag = (s: HTMLElement, e: PointerEvent) => {
    e.preventDefault(); // no text selection while dragging; this also stops the mousedown that selects
    select(s);
    if (wrapped(s)) {
      drag = { s, el: s, x0: e.clientX, y0: e.clientY, left: 0, top: 0, id: e.pointerId, moved: false, flow: null };
      return;
    }
    const el = modelOf(s).posEl(s);
    if (!el) return;
    const pos = shapePos(el);
    drag = { s, el, x0: e.clientX, y0: e.clientY, left: pos.left, top: pos.top, id: e.pointerId, moved: false };
  };
  const onPointerDown = (e: PointerEvent) => {
    const el = e.target as HTMLElement;
    if (el === move && shape && (floating(shape) || wrapped(shape))) return startDrag(shape, e);
    if (bar.contains(el) || menu.contains(el)) return;
    if (!menu.hidden) closeMenu();
    if (e.button !== 0) return;
    const s = el.closest?.(SEL) as HTMLElement | null;
    // A press on a floating shape picks it up to drag it; a click in its text without moving
    // places the caret there instead (on the pointer's release).
    if (s && inRegion(s) && (floating(s) || wrapped(s))) {
      startDrag(s, e);
      if (drag && el.closest?.(".docx-shape-body")) drag.text = { x: e.clientX, y: e.clientY };
      return;
    }
    if (el.closest?.(".docx-shape-body")) return;
    // A click on bare page over a shape behind the text (not on a letter) picks that shape.
    if (!s && inRegion(el) && !e.shiftKey) {
      const under = floatBehind(e.clientX, e.clientY, regions);
      if (under && under.matches(SEL)) startDrag(under, e);
    }
  };
  const caretAt = (x: number, y: number) => {
    const r = document.caretRangeFromPoint?.(x, y);
    const body = (r?.startContainer.nodeType === 1 ? (r.startContainer as Element) : r?.startContainer.parentElement)?.closest<HTMLElement>(".docx-shape-body");
    if (!r || !body) return;
    body.focus({ preventScroll: true });
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
  };
  const onPointerMove = (e: PointerEvent) => {
    if (!drag || e.pointerId !== drag.id) return;
    const z = zoom();
    const dx = (e.clientX - drag.x0) / z, dy = (e.clientY - drag.y0) / z;
    if (!drag.moved && Math.hypot(dx, dy) < 3) return;
    drag.moved = true;
    if (drag.flow !== undefined) {
      drag.flow ??= ghostOf(drag.s);
      drag.flow.move(e.clientX - drag.x0, e.clientY - drag.y0);
      return;
    }
    setShapePos(drag.el, Math.round(drag.left + dx), Math.round(drag.top + dy));
    place();
  };
  const onPointerUp = (e: PointerEvent) => {
    if (!drag || e.pointerId !== drag.id) return;
    const d = drag;
    drag = null;
    d.flow?.remove();
    if (!d.moved) {
      if (d.text) caretAt(d.text.x, d.text.y);
      return;
    }
    if (d.flow !== undefined) {
      // Dropped where its outline's corner is: it leads that paragraph, aligned as dropped.
      const box = (d.s.querySelector<HTMLElement>(":scope > .docx-shape-box, :scope > .odt-shape-box") ?? d.s).getBoundingClientRect();
      const align = dropWrapped(d.s, box.left + e.clientX - d.x0, box.top + e.clientY - d.y0, box.width, regions);
      if (!align) return;
      modelOf(d.s).setWrap(d.s, wrapOf(d.s), align);
      markShapeDirty(d.s, "wrap");
      layout();
      mark();
      sync();
      place();
      return;
    }
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
    if (bar.contains(el) || menu.contains(el)) return;
    // A click on a shape (its frame or its text) selects it; the text keeps the caret.
    const s = el.closest?.(SEL) as HTMLElement | null;
    select(s && inRegion(s) ? s : null);
  };
  document.addEventListener("mousedown", onDown);
  wrap.addEventListener("keydown", (e) => {
    if (!shape || bar.contains(e.target as Node) || menu.contains(e.target as Node)) return;
    const inText = !!(e.target as Element).closest?.(".docx-shape-body");
    if (e.key === "Escape") { if (!menu.hidden) closeMenu(); else select(null); }
    else if ((e.key === "Delete" || e.key === "Backspace") && !inText) {
      e.preventDefault();
      del.click();
    } else if (!inText && floating(shape) && /^Arrow(Left|Right|Up|Down)$/.test(e.key)) {
      // Arrow keys nudge a floating shape (Shift: ten times as far).
      e.preventDefault();
      const k = e.shiftKey ? 10 : 1;
      nudge(shape, e.key === "ArrowLeft" ? -k : e.key === "ArrowRight" ? k : 0, e.key === "ArrowUp" ? -k : e.key === "ArrowDown" ? k : 0);
      mark();
      place();
    }
  });

  /** Insert a new text box or shape at the caret. A shape floats in front of the text where it
   *  was inserted (it can be dragged from there); a text box sits beside the text. */
  const insert = (kind: string) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !inRegion(sel.anchorNode)) return;
    const range = sel.getRangeAt(0).cloneRange();
    getActiveEl().focus({ preventScroll: true });
    range.deleteContents();
    const basic = ["textbox", "rect", "roundRect", "ellipse"].includes(kind);
    const el = hooks?.create?.(kind) ?? (basic ? newShapeEl(kind as ShapeKind) : undefined);
    if (!el) return;
    range.insertNode(el);
    range.setStartAfter(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    select(el);
    setWrap(kind === "textbox" ? "square" : "front", el);
    sync();
    mark();
  };

  const onScroll = () => { if (shape) place(); };
  wrap.addEventListener("scroll", onScroll, true);

  const teardown = () => {
    document.removeEventListener("mousedown", onDown);
    document.removeEventListener("pointerdown", onPointerDown);
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", onPointerUp);
    document.removeEventListener("pointercancel", onPointerUp);
    wrap.removeEventListener("scroll", onScroll, true);
    bar.remove();
    menu.remove();
  };
  return { insert, select, layout, teardown };
}

/** A floating shape's offset (px, as its left / top). */
export function shapePos(s: HTMLElement): { left: number; top: number } {
  return { left: parseFloat(s.style.left) || 0, top: parseFloat(s.style.top) || 0 };
}

export function setShapePos(s: HTMLElement, left: number, top: number): void {
  s.style.left = `${left}px`;
  s.style.top = `${top}px`;
}
