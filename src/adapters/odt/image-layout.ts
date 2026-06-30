// Floating-image layout for odt: the bridge between a draw:frame's anchor + wrap (carried on a
// graphic style) and the format-agnostic ImageLayout the editor records on an <img>. The reader
// calls readOdtLayout to set the data attrs; the writer calls applyFrameLayout to anchor a frame
// and reference the matching graphic style. The shared ImageLayout type lives in core.
import { NS } from "./shared";
import type { ImageLayout, ImageWrap } from "../../core/types";

const pxToCm = (px: number): string => `${Math.round((px / (96 / 2.54)) * 1000) / 1000}cm`;

/** A graphic style's wrap-relevant properties, resolved from style:graphic-properties. */
export interface GraphicStyleInfo {
  wrap?: string; // style:wrap
  runThrough?: string; // style:run-through (background = behind, foreground = in front)
  hpos?: string; // style:horizontal-pos
  vpos?: string; // style:vertical-pos
  dist?: { t: number; r: number; b: number; l: number }; // fo:margin-* (wrap padding, px)
}

/** Index every graphic-family style in a document by name -> its wrap properties. */
export function collectGraphicStyles(doc: Document, lenToPx: (v: string | null) => number | undefined): Map<string, GraphicStyleInfo> {
  const map = new Map<string, GraphicStyleInfo>();
  for (const st of Array.from(doc.getElementsByTagName("style:style"))) {
    if (st.getAttribute("style:family") !== "graphic") continue;
    const name = st.getAttribute("style:name");
    if (!name) continue;
    const gp = st.getElementsByTagName("style:graphic-properties")[0];
    const hasMargin = gp && ["fo:margin-top", "fo:margin-right", "fo:margin-bottom", "fo:margin-left"].some((a) => gp.hasAttribute(a));
    map.set(name, {
      wrap: gp?.getAttribute("style:wrap") ?? undefined,
      runThrough: gp?.getAttribute("style:run-through") ?? undefined,
      hpos: gp?.getAttribute("style:horizontal-pos") ?? undefined,
      vpos: gp?.getAttribute("style:vertical-pos") ?? undefined,
      dist: hasMargin
        ? { t: Math.round(lenToPx(gp!.getAttribute("fo:margin-top")) ?? 0), r: Math.round(lenToPx(gp!.getAttribute("fo:margin-right")) ?? 0), b: Math.round(lenToPx(gp!.getAttribute("fo:margin-bottom")) ?? 0), l: Math.round(lenToPx(gp!.getAttribute("fo:margin-left")) ?? 0) }
        : undefined,
    });
  }
  return map;
}

/** Read a draw:frame's floating layout, or null when it is inline (anchored as-char). */
export function readOdtLayout(frame: Element, gmap: Map<string, GraphicStyleInfo>, lenToPx: (v: string | null) => number | undefined): ImageLayout | null {
  const anchor = frame.getAttribute("text:anchor-type");
  if (!anchor || anchor === "as-char") return null; // inline
  const gs = gmap.get(frame.getAttribute("draw:style-name") ?? "");
  let wrap: ImageWrap;
  if (gs?.wrap === "none") wrap = "topbottom";
  else if (gs?.wrap === "run-through") wrap = gs.runThrough === "background" ? "behind" : "front";
  else if (gs?.wrap === "dynamic") wrap = "tight";
  else wrap = "square";
  const hpos = gs?.hpos ?? "";
  const align = hpos === "right" || hpos === "center" ? hpos : "left";
  const x = lenToPx(frame.getAttribute("svg:x")) ?? 0;
  const y = lenToPx(frame.getAttribute("svg:y")) ?? 0;
  // A wrapped frame positioned from an edge (style:horizontal/vertical-pos="from-*") keeps its
  // offset per axis (the svg:x / svg:y values).
  const wrapped = wrap !== "behind" && wrap !== "front";
  const absX = wrapped && /^from-/.test(hpos) ? true : undefined;
  const absY = wrapped && /^from-/.test(gs?.vpos ?? "") ? true : undefined;
  return { wrap, align, x, y, dist: gs?.dist, absX, absY };
}

/** Create (once) an automatic graphic style for a wrap mode + alignment; return its name. */
export function graphicStyleFor(doc: Document, auto: Element, created: Map<string, string>, layout: ImageLayout): string {
  const runThrough = layout.wrap === "behind" || layout.wrap === "front"; // text flows over/under
  const offH = runThrough || !!layout.absX; // H placed by svg:x offset rather than alignment
  const offV = runThrough || !!layout.absY; // V placed by svg:y offset rather than "top"
  const d = layout.dist;
  const key = `g_${layout.wrap}_${offH ? "ax" : layout.align}_${offV ? "ay" : "top"}_${d ? `${d.t},${d.r},${d.b},${d.l}` : ""}`;
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_fr${created.size}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "graphic");
  const gp = doc.createElementNS(NS.style, "style:graphic-properties");
  // The wrap value follows the wrap mode (an offset-placed square is still parallel wrap).
  const wrapVal = layout.wrap === "topbottom" ? "none" : layout.wrap === "tight" ? "dynamic" : runThrough ? "run-through" : "parallel";
  gp.setAttributeNS(NS.style, "style:wrap", wrapVal);
  if (runThrough) gp.setAttributeNS(NS.style, "style:run-through", layout.wrap === "behind" ? "background" : "foreground");
  gp.setAttributeNS(NS.style, "style:horizontal-pos", offH ? "from-left" : layout.align);
  gp.setAttributeNS(NS.style, "style:horizontal-rel", "paragraph");
  gp.setAttributeNS(NS.style, "style:vertical-pos", offV ? "from-top" : "top");
  gp.setAttributeNS(NS.style, "style:vertical-rel", "paragraph");
  if (d) {
    gp.setAttributeNS(NS.fo, "fo:margin-top", pxToCm(d.t));
    gp.setAttributeNS(NS.fo, "fo:margin-right", pxToCm(d.r));
    gp.setAttributeNS(NS.fo, "fo:margin-bottom", pxToCm(d.b));
    gp.setAttributeNS(NS.fo, "fo:margin-left", pxToCm(d.l));
  }
  st.appendChild(gp);
  auto.appendChild(st);
  created.set(key, name);
  return name;
}

/** Anchor a draw:frame for the given layout (inline = as-char), referencing the graphic style
 *  and, for behind/front, setting the svg:x / svg:y absolute offset. */
export function applyFrameLayout(doc: Document, frame: Element, layout: ImageLayout | null, auto: Element, created: Map<string, string>): void {
  if (!layout) {
    frame.setAttributeNS(NS.text, "text:anchor-type", "as-char");
    frame.removeAttributeNS(NS.draw, "style-name");
    frame.removeAttributeNS(NS.svg, "x");
    frame.removeAttributeNS(NS.svg, "y");
    return;
  }
  frame.setAttributeNS(NS.text, "text:anchor-type", "paragraph");
  frame.setAttributeNS(NS.draw, "draw:style-name", graphicStyleFor(doc, auto, created, layout));
  const runThrough = layout.wrap === "behind" || layout.wrap === "front";
  if (runThrough || layout.absX) frame.setAttributeNS(NS.svg, "svg:x", pxToCm(layout.x)); else frame.removeAttributeNS(NS.svg, "x");
  if (runThrough || layout.absY) frame.setAttributeNS(NS.svg, "svg:y", pxToCm(layout.y)); else frame.removeAttributeNS(NS.svg, "y");
}
