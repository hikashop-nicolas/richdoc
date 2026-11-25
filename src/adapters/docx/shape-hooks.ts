// Drawn shapes in the editor: redraw an outline after a resize, and the ready-made shapes that
// Insert > Shape offers (Word's preset shapes, grouped as Word groups them).
import type { ShapeHooks } from "../../core/types";
import { t } from "../../core/i18n";
import { newShapeEl } from "../../core/feature/shape-edit";
import { paintShapeGeom } from "../../core/feature/shape-geom";
import { presetGeometry } from "./drawingml-geom";
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
/** "flowChartMagneticDisk" -> "Flow chart magnetic disk": the preset's own name, as a tooltip. */
const label = (id: string): string => {
  const words = id.replace(/([a-z])([A-Z0-9])/g, "$1 $2").toLowerCase();
  return words.charAt(0).toUpperCase() + words.slice(1);
};

/** A preset's spPr geometry (and line ends), as the outline drawing reads it. */
function spPrFor(e: Entry): Element {
  const ends = e.ends ? `${e.ends === "both" ? '<a:headEnd type="triangle"/>' : ""}<a:tailEnd type="triangle"/>` : "";
  const xml = `<wps:spPr xmlns:wps="${WPS}" xmlns:a="${A}"><a:prstGeom prst="${e.prst}"><a:avLst/></a:prstGeom><a:ln>${ends}</a:ln></wps:spPr>`;
  return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

/** The spPr a shape is drawn from: a gallery shape's preset, or the preserved run's first shape. */
function spPrOf(shape: HTMLElement): Element | undefined {
  const id = shape.getAttribute("data-shape-prst");
  if (id) return spPrFor(entry(id));
  const xml = shape.getAttribute("data-docx-xml");
  if (!xml) return undefined;
  const run = new DOMParser().parseFromString(xml, "application/xml");
  const wsps = run.getElementsByTagName("wps:wsp");
  if (wsps.length !== 1) return undefined; // groups are drawn member by member and not redrawn
  return Array.from(wsps[0]!.children).find((c) => c.tagName === "wps:spPr");
}

const hex = (c: string): string | undefined => {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  return /^#[0-9a-f]{6}$/i.test(c) ? c.slice(1).toUpperCase() : undefined;
};

function redraw(shape: HTMLElement): void {
  const box = shape.querySelector<HTMLElement>(":scope > .docx-shape-box");
  if (!box || !box.hasAttribute("data-geom")) return;
  const spPr = spPrOf(shape);
  const w = parseFloat(box.style.width), h = parseFloat(box.style.height);
  if (!spPr || !(w >= 0 && h >= 0)) return;
  const bw = box.style.borderTopStyle && box.style.borderTopStyle !== "none" ? parseFloat(box.style.borderTopWidth) || 0 : 0;
  const svg = geomSvg(spPr, w, h, { fill: hex(box.style.backgroundColor), line: bw ? hex(box.style.borderTopColor) ?? "000000" : undefined, lineW: bw });
  if (!svg) return;
  box.querySelector(":scope > svg.rdoc-shape-geom")?.remove();
  box.insertAdjacentHTML("afterbegin", svg);
  paintShapeGeom(box);
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

let galleryCache: ReturnType<ShapeHooks["gallery"]> | undefined;

export const docxShapeHooks: ShapeHooks = {
  redraw,
  gallery() {
    galleryCache ??= GROUPS.map(([group, ids]) => ({ group: t(group), items: ids.map((id) => ({ id, label: label(id), icon: icon(id) })) }));
    return galleryCache;
  },
  create(id: string): HTMLElement {
    // Plain boxes stay CSS boxes, as when they are read.
    if (id === "rect" || id === "roundRect" || id === "ellipse") return newShapeEl(id);
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
