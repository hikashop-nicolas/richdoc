// Floating-image layout for docx: the bridge between a wp:drawing's wrap/position and the
// format-agnostic ImageLayout the editor records on an <img>. The reader calls readLayout to
// set the data attrs; the writer calls makeContainer to build the right wp:inline or wp:anchor
// around the picture's a:graphic. The shared ImageLayout type / data-attr reader live in core.
import { NS_DECLS } from "./shared";
import type { ImageLayout, ImageWrap } from "../../core/types";

const WP_NS = NS_DECLS["xmlns:wp"]!;

export const EMU_PER_PX = 9525;
const pxToEmu = (px: number): number => Math.round(px * EMU_PER_PX);
const emuToPx = (v: string | null | undefined): number => {
  const n = v ? Number(v) : NaN;
  return Number.isFinite(n) ? Math.round(n / EMU_PER_PX) : 0;
};

/** Read a drawing-bearing run's floating layout, or null when it is inline (no wp:anchor). */
export function readLayout(run: Element): ImageLayout | null {
  const anchor = run.getElementsByTagName("wp:anchor")[0];
  if (!anchor) return null;
  const behind = anchor.getAttribute("behindDoc") === "1";
  let wrap: ImageWrap;
  if (run.getElementsByTagName("wp:wrapTopAndBottom")[0]) wrap = "topbottom";
  else if (run.getElementsByTagName("wp:wrapSquare")[0]) wrap = "square";
  else if (run.getElementsByTagName("wp:wrapTight")[0] || run.getElementsByTagName("wp:wrapThrough")[0]) wrap = "tight";
  else wrap = behind ? "behind" : "front"; // wp:wrapNone (or none declared)
  const posH = anchor.getElementsByTagName("wp:positionH")[0];
  const posV = anchor.getElementsByTagName("wp:positionV")[0];
  const alignH = posH?.getElementsByTagName("wp:align")[0]?.textContent?.trim();
  const align = alignH === "right" || alignH === "center" ? alignH : "left";
  const x = emuToPx(posH?.getElementsByTagName("wp:posOffset")[0]?.textContent);
  const y = emuToPx(posV?.getElementsByTagName("wp:posOffset")[0]?.textContent);
  const hasDist = ["distT", "distR", "distB", "distL"].some((a) => anchor.hasAttribute(a));
  const dist = hasDist
    ? { t: emuToPx(anchor.getAttribute("distT")), r: emuToPx(anchor.getAttribute("distR")), b: emuToPx(anchor.getAttribute("distB")), l: emuToPx(anchor.getAttribute("distL")) }
    : undefined;
  // A wrapped image axis positioned by posOffset (not wp:align) keeps that exact offset on save;
  // each axis is independent (a square image is commonly aligned in H but offset in V).
  const wrapped = wrap !== "behind" && wrap !== "front";
  const absX = wrapped && !!posH?.getElementsByTagName("wp:posOffset")[0] && !posH.getElementsByTagName("wp:align")[0] ? true : undefined;
  const absY = wrapped && !!posV?.getElementsByTagName("wp:posOffset")[0] && !posV.getElementsByTagName("wp:align")[0] ? true : undefined;
  return { wrap, align, x, y, dist, absX, absY };
}

const el = (doc: Document, name: string): Element => doc.createElementNS(WP_NS, name);

/** Build the wp:positionH / wp:positionV pair for an anchor. */
function position(doc: Document, tag: "wp:positionH" | "wp:positionV", relativeFrom: string, align: string | null, offsetPx: number): Element {
  const p = el(doc, tag);
  p.setAttribute("relativeFrom", relativeFrom);
  if (align) {
    const a = el(doc, "wp:align");
    a.textContent = align;
    p.appendChild(a);
  } else {
    const o = el(doc, "wp:posOffset");
    o.textContent = String(pxToEmu(offsetPx));
    p.appendChild(o);
  }
  return p;
}

/** Build the wrap element (square / tight / top-and-bottom / none) for a wrap mode. */
function wrapElement(doc: Document, wrap: ImageWrap): Element {
  if (wrap === "topbottom") return el(doc, "wp:wrapTopAndBottom");
  if (wrap === "behind" || wrap === "front") return el(doc, "wp:wrapNone");
  if (wrap === "tight") {
    const w = el(doc, "wp:wrapTight");
    w.setAttribute("wrapText", "bothSides");
    const poly = el(doc, "wp:wrapPolygon");
    poly.setAttribute("edited", "0");
    const pt = (tag: string, x: number, y: number): Element => {
      const p = el(doc, tag);
      p.setAttribute("x", String(x));
      p.setAttribute("y", String(y));
      return p;
    };
    poly.append(pt("wp:start", 0, 0), pt("wp:lineTo", 0, 21600), pt("wp:lineTo", 21600, 21600), pt("wp:lineTo", 21600, 0), pt("wp:lineTo", 0, 0));
    w.appendChild(poly);
    return w;
  }
  const w = el(doc, "wp:wrapSquare");
  w.setAttribute("wrapText", "bothSides");
  return w;
}

/** Build a wp:inline or wp:anchor wrapping the picture's a:graphic + wp:docPr at the given
 *  size and layout. The graphic and docPr come from the caller (fresh, or lifted from a
 *  preserved drawing so the blip relationship and name survive). */
export function makeContainer(
  doc: Document,
  graphic: Element,
  docPr: Element,
  cx: number,
  cy: number,
  layout: ImageLayout | null,
): Element {
  const extent = el(doc, "wp:extent");
  extent.setAttribute("cx", String(cx));
  extent.setAttribute("cy", String(cy));
  if (!layout) {
    const inline = el(doc, "wp:inline");
    inline.append(extent, docPr, graphic);
    return inline;
  }
  const anchor = el(doc, "wp:anchor");
  const d = layout.dist;
  const dist = { distT: String(d ? pxToEmu(d.t) : 0), distB: String(d ? pxToEmu(d.b) : 0), distL: String(d ? pxToEmu(d.l) : 114300), distR: String(d ? pxToEmu(d.r) : 114300) };
  for (const [k, v] of Object.entries({ ...dist, simplePos: "0", relativeHeight: "251658240", behindDoc: layout.wrap === "behind" ? "1" : "0", locked: "0", layoutInCell: "1", allowOverlap: "1" })) anchor.setAttribute(k, v);
  const simple = el(doc, "wp:simplePos");
  simple.setAttribute("x", "0");
  simple.setAttribute("y", "0");
  // behind/front are always offset-positioned; a wrapped image is offset per axis when read as
  // absX / absY (positioned by posOffset rather than alignment), else H uses align and V is "top".
  const float = layout.wrap !== "behind" && layout.wrap !== "front";
  const offH = !float || layout.absX;
  const offV = !float || layout.absY;
  const posH = position(doc, "wp:positionH", "column", offH ? null : layout.align, layout.x);
  const posV = position(doc, "wp:positionV", "paragraph", offV ? null : "top", layout.y);
  anchor.append(simple, posH, posV, extent, wrapElement(doc, layout.wrap), docPr, graphic);
  return anchor;
}
