// Edits the shape bar makes to an odt drawing (draw:line, draw:custom-shape, draw:g ...): its
// size, its place and its fill / outline. Kept on the shape element as data-shape-edit (JSON)
// and applied to a copy of the original element, both to redraw it and to save it.

const SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/** Colours #rrggbb or null for none; lengths in px; dx / dy the move since it was read. */
export interface OdtShapeEdit { fill?: string | null; line?: string | null; lineW?: number; w?: number; h?: number; dx?: number; dy?: number }

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

function resize(el: Element, w: number | undefined, h: number | undefined): void {
  if (el.tagName === "draw:g") return;
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

/** Apply the size and place of an edit to a drawing element. */
export function applyGeometry(el: Element, e: OdtShapeEdit): void {
  if (e.w !== undefined || e.h !== undefined) resize(el, e.w, e.h);
  if (e.dx || e.dy) move(el, e.dx ?? 0, e.dy ?? 0);
}

export const hasLook = (e: OdtShapeEdit) => "fill" in e || "line" in e || "lineW" in e;

/** Write an edit's fill and outline onto style:graphic-properties. */
export function applyGraphicProps(props: Element, e: OdtShapeEdit): void {
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
