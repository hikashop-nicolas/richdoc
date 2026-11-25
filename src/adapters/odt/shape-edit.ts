// Edits the shape bar makes to an odt drawing (draw:line, draw:custom-shape, draw:g ...): its
// size, its place and its fill / outline. Kept on the shape element as data-shape-edit (JSON)
// and applied to a copy of the original element, both to redraw it and to save it.

const SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/** The arrowhead new lines end with: LibreOffice's own "Arrow" marker. */
export const ARROW_MARKER = { name: "Arrow", vb: "0 0 20 30", d: "M10 0l-10 30h20z" };

const TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/** Colours #rrggbb or null for none; lengths in px; dx / dy the move since it was read; how it
 *  sits in the text and its stacking order (draw:z-index). */
export interface OdtShapeEdit {
  fill?: string | null; line?: string | null; lineW?: number; w?: number; h?: number; dx?: number; dy?: number;
  wrap?: "inline" | "square" | "tight" | "topbottom" | "behind" | "front"; align?: "left" | "center" | "right"; z?: number;
  /** A new line's arrowheads (LibreOffice's "Arrow" marker). */
  ends?: "tail" | "both";
}

export function editOf(shape: HTMLElement): OdtShapeEdit {
  try {
    return JSON.parse(shape.getAttribute("data-shape-edit") || "{}") as OdtShapeEdit;
  } catch {
    return {};
  }
}

export function setEditOf(shape: HTMLElement, e: OdtShapeEdit): void {
  shape.setAttribute("data-shape-edit", JSON.stringify(e));
}

const UNIT_PX: Record<string, number> = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 4 / 3, pc: 16, px: 1 };
const lenPx = (v: string | null): number => {
  const m = /^\s*(-?[\d.]+)\s*([a-z]*)\s*$/i.exec(v ?? "");
  return m ? Number(m[1]) * (UNIT_PX[(m[2] || "cm").toLowerCase()] ?? 1) : 0;
};
const cm = (px: number): string => `${Math.round((px * 2.54 / 96) * 1000) / 1000 || 0}cm`; // to 0.01 mm
const setLen = (el: Element, name: string, px: number) => el.setAttributeNS(SVG, name, cm(px));

function move(el: Element, dx: number, dy: number): void {
  if (el.tagName === "draw:g") {
    for (const c of Array.from(el.children)) if (c.namespaceURI === DRAW && c.tagName !== "draw:glue-point") move(c, dx, dy);
    return;
  }
  if (el.hasAttribute("svg:x1")) {
    for (const [a, d] of [["svg:x1", dx], ["svg:x2", dx], ["svg:y1", dy], ["svg:y2", dy]] as const) setLen(el, a, lenPx(el.getAttribute(a)) + d);
    return;
  }
  const tf = el.getAttribute("draw:transform");
  if (tf && /translate/.test(tf)) {
    el.setAttributeNS(DRAW, "draw:transform", tf.replace(/translate\s*\(\s*([^\s,)]+)[\s,]+([^\s,)]+)\s*\)/, (_m, x: string, y: string) => `translate (${cm(lenPx(x) + dx)} ${cm(lenPx(y) + dy)})`));
    return;
  }
  setLen(el, "svg:x", lenPx(el.getAttribute("svg:x")) + dx);
  setLen(el, "svg:y", lenPx(el.getAttribute("svg:y")) + dy);
}

const isMember = (c: Element) => c.namespaceURI === DRAW && c.tagName !== "draw:glue-point" && c.tagName !== "draw:enhanced-geometry";
const translateOf = (el: Element): [number, number] | undefined => {
  const m = /translate\s*\(\s*([^\s,)]+)[\s,]+([^\s,)]+)\s*\)/.exec(el.getAttribute("draw:transform") ?? "");
  return m ? [lenPx(m[1]!), lenPx(m[2]!)] : undefined;
};

/** A group's box (px): its members' places and sizes, their rotation aside. */
export function groupBox(g: Element): { x: number; y: number; w: number; h: number } | undefined {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  const add = (x: number, y: number) => { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); };
  const walk = (el: Element) => {
    for (const c of Array.from(el.children)) {
      if (!isMember(c)) continue;
      if (c.tagName === "draw:g") { walk(c); continue; }
      if (c.hasAttribute("svg:x1")) {
        add(lenPx(c.getAttribute("svg:x1")), lenPx(c.getAttribute("svg:y1")));
        add(lenPx(c.getAttribute("svg:x2")), lenPx(c.getAttribute("svg:y2")));
        continue;
      }
      const [x, y] = translateOf(c) ?? [lenPx(c.getAttribute("svg:x")), lenPx(c.getAttribute("svg:y"))];
      add(x, y);
      add(x + lenPx(c.getAttribute("svg:width")), y + lenPx(c.getAttribute("svg:height")));
    }
  };
  walk(g);
  return x1 > x0 || y1 > y0 ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : undefined;
}

/** Stretch a group's members by sx, sy from its corner (ox, oy). */
function scaleMembers(g: Element, ox: number, oy: number, sx: number, sy: number): void {
  const at = (v: number, o: number, k: number) => o + (v - o) * k;
  for (const c of Array.from(g.children)) {
    if (!isMember(c)) continue;
    if (c.tagName === "draw:g") { scaleMembers(c, ox, oy, sx, sy); continue; }
    if (c.hasAttribute("svg:x1")) {
      for (const [a, o, k] of [["svg:x1", ox, sx], ["svg:x2", ox, sx], ["svg:y1", oy, sy], ["svg:y2", oy, sy]] as const) setLen(c, a, at(lenPx(c.getAttribute(a)), o, k));
      continue;
    }
    const t = translateOf(c);
    if (t) c.setAttributeNS(DRAW, "draw:transform", (c.getAttribute("draw:transform") ?? "").replace(/translate\s*\([^)]*\)/, `translate (${cm(at(t[0], ox, sx))} ${cm(at(t[1], oy, sy))})`));
    else {
      setLen(c, "svg:x", at(lenPx(c.getAttribute("svg:x")), ox, sx));
      setLen(c, "svg:y", at(lenPx(c.getAttribute("svg:y")), oy, sy));
    }
    if (c.hasAttribute("svg:width")) setLen(c, "svg:width", lenPx(c.getAttribute("svg:width")) * sx);
    if (c.hasAttribute("svg:height")) setLen(c, "svg:height", lenPx(c.getAttribute("svg:height")) * sy);
  }
}

function resize(el: Element, w: number | undefined, h: number | undefined): void {
  if (el.tagName === "draw:g") {
    // A group stretches its members to its new box.
    const b = groupBox(el);
    if (!b) return;
    scaleMembers(el, b.x, b.y, w !== undefined && b.w ? w / b.w : 1, h !== undefined && b.h ? h / b.h : 1);
    return;
  }
  if (el.hasAttribute("svg:x1")) {
    // A line keeps its start and direction; its extent becomes w x h.
    for (const [a1, a2, size] of [["svg:x1", "svg:x2", w], ["svg:y1", "svg:y2", h]] as const) {
      if (size === undefined) continue;
      const p1 = lenPx(el.getAttribute(a1)), p2 = lenPx(el.getAttribute(a2));
      setLen(el, a2, p1 + (p2 >= p1 ? size : -size));
    }
    return;
  }
  if (w !== undefined) setLen(el, "svg:width", w);
  if (h !== undefined) setLen(el, "svg:height", h);
}

/** Apply the size, place, anchor and order of an edit to a drawing element. */
export function applyGeometry(el: Element, e: OdtShapeEdit): void {
  if (e.w !== undefined || e.h !== undefined) resize(el, e.w, e.h);
  if (e.wrap === "inline") {
    // In line: anchored as a character, its place is the text's. A line or a group keeps its own
    // coordinates, from 0,0.
    el.setAttributeNS(TEXT, "text:anchor-type", "as-char");
    if (el.hasAttribute("svg:x1")) {
      const v = (a: string) => lenPx(el.getAttribute(a));
      move(el, -Math.min(v("svg:x1"), v("svg:x2")), -Math.min(v("svg:y1"), v("svg:y2")));
    } else if (el.tagName === "draw:g") {
      const b = groupBox(el);
      if (b) move(el, -b.x, -b.y);
    } else {
      el.removeAttribute("svg:x");
      el.removeAttribute("svg:y");
    }
  } else {
    if (e.wrap && el.getAttribute("text:anchor-type") === "as-char") el.setAttributeNS(TEXT, "text:anchor-type", "paragraph");
    if (e.dx || e.dy) move(el, e.dx ?? 0, e.dy ?? 0);
  }
  if (e.z !== undefined) el.setAttributeNS(DRAW, "draw:z-index", String(Math.max(0, Math.round(e.z))));
}

/** Whether an edit changes the graphic style: its look or how it wraps. */
export const hasLook = (e: OdtShapeEdit) => "fill" in e || "line" in e || "lineW" in e || !!e.ends || (!!e.wrap && e.wrap !== "inline");

const WRAP_VALUE = { square: "parallel", tight: "dynamic", topbottom: "none", behind: "run-through", front: "run-through" } as const;

/** Write an edit's wrap onto style:graphic-properties: wrapped beside the text at its alignment,
 *  or running through it (behind / in front) from its offsets. */
function applyWrapProps(props: Element, e: OdtShapeEdit): void {
  if (!e.wrap || e.wrap === "inline") return;
  props.setAttributeNS(STYLE, "style:wrap", WRAP_VALUE[e.wrap]);
  const through = e.wrap === "behind" || e.wrap === "front";
  if (through) props.setAttributeNS(STYLE, "style:run-through", e.wrap === "behind" ? "background" : "foreground");
  else props.removeAttribute("style:run-through");
  if (through) {
    if (!/^from-/.test(props.getAttribute("style:horizontal-pos") ?? "")) {
      props.setAttributeNS(STYLE, "style:horizontal-pos", "from-left");
      props.setAttributeNS(STYLE, "style:horizontal-rel", "paragraph");
    }
    if (!/^from-/.test(props.getAttribute("style:vertical-pos") ?? "")) {
      props.setAttributeNS(STYLE, "style:vertical-pos", "from-top");
      props.setAttributeNS(STYLE, "style:vertical-rel", "paragraph");
    }
  } else {
    props.setAttributeNS(STYLE, "style:horizontal-pos", e.align ?? "left");
    props.setAttributeNS(STYLE, "style:horizontal-rel", "paragraph");
    props.setAttributeNS(STYLE, "style:vertical-pos", "top");
    props.setAttributeNS(STYLE, "style:vertical-rel", "paragraph");
  }
}

/** Write an edit's wrap, fill and outline onto style:graphic-properties. */
export function applyGraphicProps(props: Element, e: OdtShapeEdit): void {
  applyWrapProps(props, e);
  if (e.ends) {
    props.setAttributeNS(DRAW, "draw:marker-end", ARROW_MARKER.name);
    props.setAttributeNS(DRAW, "draw:marker-end-width", "0.3cm");
    if (e.ends === "both") {
      props.setAttributeNS(DRAW, "draw:marker-start", ARROW_MARKER.name);
      props.setAttributeNS(DRAW, "draw:marker-start-width", "0.3cm");
    }
  }
  if ("fill" in e) {
    props.setAttributeNS(DRAW, "draw:fill", e.fill ? "solid" : "none");
    if (e.fill) props.setAttributeNS(DRAW, "draw:fill-color", e.fill);
  }
  if ("line" in e || "lineW" in e) {
    if (!e.line || !e.lineW) props.setAttributeNS(DRAW, "draw:stroke", "none");
    else {
      if (props.getAttribute("draw:stroke") !== "dash") props.setAttributeNS(DRAW, "draw:stroke", "solid");
      props.setAttributeNS(SVG, "svg:stroke-color", e.line);
      props.setAttributeNS(SVG, "svg:stroke-width", cm(e.lineW));
    }
  }
}

/** A new automatic graphic style for an edited look: a copy of the shape's own automatic style,
 *  or one based on its named style. Returns the new style's name. */
export function editedGraphicStyle(auto: Element, base: string, e: OdtShapeEdit, taken: (n: string) => boolean): string {
  const doc = auto.ownerDocument;
  let n = 1;
  while (taken(`rdocgr${n}`)) n++;
  const name = `rdocgr${n}`;
  const own = Array.from(auto.children).find((c) => c.tagName === "style:style" && c.getAttribute("style:name") === base && c.getAttribute("style:family") === "graphic");
  let st: Element;
  if (own) {
    st = own.cloneNode(true) as Element;
  } else {
    st = doc.createElementNS(STYLE, "style:style");
    st.setAttributeNS(STYLE, "style:family", "graphic");
    if (base) st.setAttributeNS(STYLE, "style:parent-style-name", base);
  }
  st.setAttributeNS(STYLE, "style:name", name);
  let props = Array.from(st.children).find((c) => c.tagName === "style:graphic-properties");
  if (!props) {
    props = doc.createElementNS(STYLE, "style:graphic-properties");
    st.appendChild(props);
  }
  applyGraphicProps(props, e);
  auto.appendChild(st);
  return name;
}

/** LibreOffice also keeps a shape's fill on its text style (loext:graphic-properties on the
 *  draw:text-style-name paragraph style), and reads it from there first. A fill edit gets a copy
 *  of that style with the new fill; `undefined` when it carries none. */
export function editedTextStyle(auto: Element, base: string, e: OdtShapeEdit, taken: (n: string) => boolean): string | undefined {
  if (!("fill" in e) || !base) return undefined;
  const own = Array.from(auto.children).find((c) => c.tagName === "style:style" && c.getAttribute("style:name") === base);
  const props = own && Array.from(own.children).find((c) => c.localName === "graphic-properties");
  if (!own || !props) return undefined;
  let n = 1;
  while (taken(`rdocsp${n}`)) n++;
  const st = own.cloneNode(true) as Element;
  st.setAttributeNS(STYLE, "style:name", `rdocsp${n}`);
  const p = Array.from(st.children).find((c) => c.localName === "graphic-properties")!;
  p.setAttributeNS(DRAW, "draw:fill", e.fill ? "solid" : "none");
  if (e.fill) p.setAttributeNS(DRAW, "draw:fill-color", e.fill);
  auto.appendChild(st);
  return `rdocsp${n}`;
}
