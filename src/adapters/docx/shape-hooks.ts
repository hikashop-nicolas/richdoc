// Drawn shapes in the editor: redraw an outline after a resize, and the ready-made shapes that
// Insert > Shape offers (Word's preset shapes, grouped as Word groups them).
import type { ShapeHooks } from "../../core/types";
import { t } from "../../core/i18n";
import { newShapeEl } from "../../core/feature/shape-edit";
import { paintShapeGeom } from "../../core/feature/shape-geom";
import { adjustToward, presetGeometry, presetHandles } from "./drawingml-geom";
import { geomSvg, lineEnd } from "./shapes";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

/** A gallery entry: its preset, and the line ends a line or connector gets. */
interface Entry { prst: string; ends?: "tail" | "both" }

const LINES: Record<string, Entry> = {
  line: { prst: "line" },
  arrowLine: { prst: "straightConnector1", ends: "tail" },
  doubleArrowLine: { prst: "straightConnector1", ends: "both" },
  elbowArrow: { prst: "bentConnector3", ends: "tail" },
  curvedArrow: { prst: "curvedConnector3", ends: "tail" },
};

const GROUPS: [string, string[]][] = [
  ["shapeGroupLines", Object.keys(LINES)],
  ["shapeGroupBasic", ["rect", "roundRect", "ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid", "diamond", "pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "pie", "chord", "teardrop", "frame", "halfFrame", "corner", "diagStripe", "plus", "plaque", "can", "cube", "bevel", "donut", "noSmoking", "blockArc", "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun", "moon", "cloud", "arc", "bracketPair", "bracePair", "leftBracket", "rightBracket", "leftBrace", "rightBrace"]],
  ["shapeGroupArrows", ["rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "leftRightUpArrow", "bentArrow", "uturnArrow", "leftUpArrow", "bentUpArrow", "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "stripedRightArrow", "notchedRightArrow", "homePlate", "chevron", "rightArrowCallout", "downArrowCallout", "leftArrowCallout", "upArrowCallout", "leftRightArrowCallout", "quadArrowCallout", "circularArrow"]],
  ["shapeGroupEquation", ["mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"]],
  ["shapeGroupFlowchart", ["flowChartProcess", "flowChartAlternateProcess", "flowChartDecision", "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument", "flowChartTerminator", "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation", "flowChartConnector", "flowChartOffpageConnector", "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr", "flowChartCollate", "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOnlineStorage", "flowChartDelay", "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay"]],
  ["shapeGroupStars", ["irregularSeal1", "irregularSeal2", "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32", "ribbon2", "ribbon", "ellipseRibbon2", "ellipseRibbon", "verticalScroll", "horizontalScroll", "wave", "doubleWave"]],
  ["shapeGroupCallouts", ["wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "borderCallout1", "borderCallout2", "borderCallout3", "accentCallout1", "accentCallout2", "accentCallout3", "callout1", "callout2", "callout3", "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3"]],
];

const entry = (id: string): Entry => LINES[id] ?? { prst: id };
/** The preset and line ends a gallery id is saved with. */
export const galleryPreset = entry;
/** A gallery shape's name, as Word names it in the editor's language. */
const label = (id: string): string => t(`shp_${id}`);

/** A preset's spPr geometry (and line ends), as the outline drawing reads it. */
function spPrFor(e: Entry): Element {
  const ends = e.ends ? `${e.ends === "both" ? '<a:headEnd type="triangle"/>' : ""}<a:tailEnd type="triangle"/>` : "";
  const xml = `<wps:spPr xmlns:wps="${WPS}" xmlns:a="${A}"><a:prstGeom prst="${e.prst}"><a:avLst/></a:prstGeom><a:ln>${ends}</a:ln></wps:spPr>`;
  return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

/** The spPr each drawn box of a shape is drawn from, with its box: a gallery shape's preset, the
 *  preserved run's shape, or a group's members in order. */
function drawnBoxes(shape: HTMLElement): [HTMLElement, Element][] {
  const box = shape.querySelector<HTMLElement>(":scope > .docx-shape-box");
  if (!box) return [];
  const id = shape.getAttribute("data-shape-prst");
  if (id) return [[box, withAdjust(spPrFor(entry(id)), shape)]];
  const xml = shape.getAttribute("data-docx-xml");
  if (!xml) return [];
  const run = new DOMParser().parseFromString(xml, "application/xml");
  const spPrs = Array.from(run.getElementsByTagName("wps:wsp")).map((w) => Array.from(w.children).find((c) => c.tagName === "wps:spPr"));
  if (spPrs.length === 1) return spPrs[0] ? [[box, withAdjust(spPrs[0], shape)]] : [];
  // A group: its member boxes (not its sub-groups' boxes) are its shapes, in order.
  const boxes = Array.from(box.querySelectorAll<HTMLElement>(".docx-shape-box")).filter((m) => !m.querySelector(":scope > .docx-shape-box, :scope > img"));
  return boxes.length === spPrs.length ? boxes.flatMap((b, i) => (spPrs[i] ? [[b, spPrs[i]!] as [HTMLElement, Element]] : [])) : [];
}

const hex = (c: string): string | undefined => {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  return /^#[0-9a-f]{6}$/i.test(c) ? c.slice(1).toUpperCase() : undefined;
};

/** A shape's edited adjust values (its handles moved in the editor), {name: "val N"}. */
export function shapeAdjust(shape: HTMLElement): Record<string, string> | undefined {
  try {
    const a = shape.getAttribute("data-shape-adj");
    return a ? (JSON.parse(a) as Record<string, string>) : undefined;
  } catch {
    return undefined;
  }
}

/** Write adjust values into a prstGeom's a:avLst (an spPr's or the prstGeom itself). */
export function setAvLst(el: Element, adj: Record<string, string>): void {
  const prst = el.localName === "prstGeom" ? el : Array.from(el.children).find((c) => c.localName === "prstGeom");
  if (!prst) return;
  const ns = prst.namespaceURI ?? A;
  let av = Array.from(prst.children).find((c) => c.localName === "avLst");
  if (!av) { av = prst.ownerDocument.createElementNS(ns, "a:avLst"); prst.appendChild(av); }
  while (av.firstChild) av.removeChild(av.firstChild);
  for (const [name, fmla] of Object.entries(adj)) {
    const gd = prst.ownerDocument.createElementNS(ns, "a:gd");
    gd.setAttribute("name", name);
    gd.setAttribute("fmla", fmla);
    av.appendChild(gd);
  }
}

/** An spPr with the shape's edited adjust values, when it has any. */
function withAdjust(spPr: Element, shape: HTMLElement): Element {
  const adj = shapeAdjust(shape);
  if (!adj) return spPr;
  const copy = spPr.cloneNode(true) as Element;
  setAvLst(copy, adj);
  return copy;
}

/** A single preset shape's handles' context: its preset, adjust values, size and flips. */
function presetOf(shape: HTMLElement) {
  const [only, ...more] = drawnBoxes(shape);
  if (!only || more.length) return undefined;
  const [box, spPr] = only;
  if (box !== shape.querySelector(":scope > .docx-shape-box") || box.style.transform) return undefined; // not a group member; not turned
  const geom = Array.from(spPr.children).find((c) => c.localName === "prstGeom");
  const xfrm = Array.from(spPr.children).find((c) => c.localName === "xfrm");
  const prst = geom?.getAttribute("prst");
  if (!prst || Number(xfrm?.getAttribute("rot") ?? 0)) return undefined;
  const adj: Record<string, string> = {};
  for (const gd of Array.from(Array.from(geom!.children).find((c) => c.localName === "avLst")?.children ?? [])) {
    const n = gd.getAttribute("name");
    if (n) adj[n] = gd.getAttribute("fmla") ?? "";
  }
  const w = parseFloat(box.style.width), h = parseFloat(box.style.height);
  if (!(w > 0) || !(h >= 0)) return undefined;
  const flip = (a: string) => xfrm?.getAttribute(a) === "1" || xfrm?.getAttribute(a) === "true";
  return { box, prst, adj, w, h, flipH: flip("flipH"), flipV: flip("flipV") };
}

/** Draw a shape's outline (each member's, for a group) again at its box's current size. */
function redraw(shape: HTMLElement): void {
  for (const [box, spPr] of drawnBoxes(shape)) {
    if (!box.hasAttribute("data-geom")) continue;
    const w = parseFloat(box.style.width), h = parseFloat(box.style.height);
    if (!(w >= 0 && h >= 0)) continue;
    const bw = box.style.borderTopStyle && box.style.borderTopStyle !== "none" ? parseFloat(box.style.borderTopWidth) || 0 : 0;
    const svg = geomSvg(spPr, w, h, { fill: hex(box.style.backgroundColor), line: bw ? hex(box.style.borderTopColor) ?? "000000" : undefined, lineW: bw });
    if (!svg) continue;
    box.querySelector(":scope > svg.rdoc-shape-geom")?.remove();
    box.insertAdjacentHTML("afterbegin", svg);
    paintShapeGeom(box);
  }
}

function icon(id: string): string {
  const e = entry(id);
  const W = 28, H = 20;
  const paths = presetGeometry(e.prst, W, H) ?? [];
  let body = "";
  for (const p of paths) {
    const fill = p.fill === "none" || LINES[id] ? "none" : "currentColor";
    body += `<path d="${p.d}" fill="${fill}" fill-opacity=".18"${p.stroke ? ' stroke="currentColor"' : ""}/>`;
  }
  const open = paths.find((p) => !p.closed && p.start && p.end);
  if (open && e.ends) {
    const head = (from: [number, number], tip: [number, number]) => lineEnd(spPrFor(e).getElementsByTagName("a:tailEnd")[0], from, tip, 0.8);
    body += head(open.end![0], open.end![1]).replace("<path", '<path fill="currentColor"');
    if (e.ends === "both") body += head(open.start![1], open.start![0]).replace("<path", '<path fill="currentColor"');
  }
  return `<svg width="32" height="24" viewBox="-2 -2 32 24" aria-hidden="true" stroke-width="1.2" stroke-linejoin="round">${body}</svg>`;
}

let galleryCache: ReturnType<NonNullable<ShapeHooks["gallery"]>> | undefined;

export const docxShapeHooks: ShapeHooks = {
  redraw,
  handles(shape) {
    const p = presetOf(shape);
    if (!p) return undefined;
    const points = presetHandles(p.prst, p.w, p.h, p.adj).map((hd) => ({ x: p.flipH ? p.w - hd.x : hd.x, y: p.flipV ? p.h - hd.y : hd.y }));
    return points.length ? { frame: p.box, points } : undefined;
  },
  moveHandle(shape, i, x, y) {
    const p = presetOf(shape);
    if (!p) return;
    const adj = adjustToward(p.prst, p.w, p.h, p.adj, i, p.flipH ? p.w - x : x, p.flipV ? p.h - y : y);
    shape.setAttribute("data-shape-adj", JSON.stringify(adj));
    // A box drawn by CSS (a rounded rectangle) is drawn from its outline from now on.
    if (!p.box.hasAttribute("data-geom")) { p.box.setAttribute("data-geom", ""); p.box.style.borderRadius = ""; }
    redraw(shape);
  },
  gallery() {
    galleryCache ??= GROUPS.map(([group, ids]) => ({ group: t(group), items: ids.map((id) => ({ id, label: label(id), icon: icon(id) })) }));
    return galleryCache;
  },
  textBox: true,
  create(id: string): HTMLElement {
    // Text boxes and plain boxes stay CSS boxes, as when they are read.
    if (id === "textbox") return newShapeEl(id);
    if (id === "rect" || id === "roundRect" || id === "ellipse") {
      const el = newShapeEl(id);
      el.setAttribute("data-shape-prst", id); // its handles (a rounded rectangle's corners)
      return el;
    }
    const shape = newShapeEl("rect");
    const box = shape.querySelector<HTMLElement>(":scope > .docx-shape-box")!;
    shape.setAttribute("data-shape-prst", id);
    box.setAttribute("data-geom", "");
    box.style.borderRadius = "";
    if (LINES[id]) {
      // A line: no fill, no text, a heavier outline in the accent colour.
      box.style.backgroundColor = "";
      box.style.border = "2px solid #4472c4";
      box.style.height = "72px";
      box.querySelector(".docx-shape-body")?.remove();
    }
    redraw(shape);
    return shape;
  },
};
