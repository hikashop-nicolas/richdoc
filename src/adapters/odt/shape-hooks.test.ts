import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { odtToHtml } from "./index";
import { odtShapeHooks } from "./shape-hooks";

const ROOT = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"';
const CONTENT = `<?xml version="1.0"?><office:document-content ${ROOT}><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties svg:stroke-color="#000080" svg:stroke-width="0.05cm" draw:fill="solid" draw:fill-color="#ffff00" style:wrap="run-through" style:run-through="foreground"/></style:style></office:automatic-styles>` +
  '<office:body><office:text><text:p>Text<draw:ellipse text:anchor-type="paragraph" draw:style-name="gr1" svg:x="1cm" svg:y="0cm" svg:width="2cm" svg:height="1cm"/></text:p></office:text></office:body></office:document-content>';

function mount() {
  const bytes = zipSync({ mimetype: strToU8("application/vnd.oasis.opendocument.text"), "content.xml": strToU8(CONTENT), "styles.xml": strToU8("<x/>") });
  document.body.innerHTML = odtToHtml(bytes);
  const hooks = odtShapeHooks(() => unzipSync(bytes));
  const shape = document.querySelector<HTMLElement>(".odt-shape")!;
  return { hooks, shape, model: hooks.model! };
}

describe("odt shape hooks", () => {
  it("reads the look the bar shows", () => {
    const { shape, model } = mount();
    expect(shape.matches(model.selector)).toBe(true);
    expect(model.get(shape)).toEqual({ fill: "#ffff00", line: "#000080", lineW: 1.89, w: 75.59, h: 37.8 });
  });

  it("redraws with a new fill, outline and size, keeping the original for the save", () => {
    const { shape, model } = mount();
    const xml = shape.getAttribute("data-odt-xml");
    model.set(shape, { fill: "#ff0000" });
    model.set(shape, { line: null, lineW: 0 });
    model.set(shape, { w: 151.18 });
    const svg = shape.querySelector("svg")!.outerHTML;
    expect(svg).toContain('fill="#ff0000"');
    expect(svg).not.toContain('stroke="#000080"');
    expect(svg).toMatch(/<ellipse cx="75\.59"/);
    expect(model.get(shape)).toMatchObject({ fill: "#ff0000", line: null, lineW: 0, w: 151.18 });
    expect(shape.getAttribute("data-odt-xml")).toBe(xml);
  });

  it("records a drag as a move", () => {
    const { shape, model } = mount();
    const box = model.posEl(shape)!;
    const left = parseFloat(box.style.left), top = parseFloat(box.style.top);
    box.style.left = `${left + 20}px`;
    box.style.top = `${top + 10}px`;
    model.moved(shape);
    expect(JSON.parse(shape.getAttribute("data-shape-edit")!)).toEqual({ dx: 20, dy: 10 });
    // A later redraw keeps it where it was dropped.
    model.set(shape, { fill: "#00ff00" });
    expect(parseFloat(model.posEl(shape)!.style.left)).toBeCloseTo(left + 20, 1);
  });

  it("changes how it sits in the text", () => {
    const { shape, model } = mount();
    expect(model.wraps(shape)).toContain("inline");
    model.setWrap(shape, "square", "right");
    expect(shape.getAttribute("data-shape-wrap")).toBe("square");
    expect(shape.getAttribute("data-shape-align")).toBe("right");
    model.setWrap(shape, "inline", "left");
    expect(shape.getAttribute("data-shape-wrap")).toBe("inline");
    model.setWrap(shape, "behind", "left");
    expect(shape.getAttribute("data-shape-wrap")).toBe("behind");
    expect(JSON.parse(shape.getAttribute("data-shape-edit")!)).toMatchObject({ wrap: "behind" });
  });

  it("inserts a gallery shape and a line with its arrowhead", () => {
    const { hooks } = mount();
    const star = hooks.create!("star5")!;
    expect(star.getAttribute("data-odt-xml")).toMatch(/^<draw:custom-shape [^>]*text:anchor-type="as-char"[^>]*><draw:enhanced-geometry svg:viewBox="0 0 0 0" draw:type="ooxml-star5" draw:modifiers="19098 105146 110557" /);
    expect(star.querySelector("svg path")?.getAttribute("fill")).toBe("#4472c4");
    const arrow = hooks.create!("arrowLine")!;
    expect(arrow.getAttribute("data-odt-xml")).toMatch(/^<draw:path /);
    expect(arrow.querySelectorAll("svg path").length).toBeGreaterThan(2); // the line, its hit area and its head
    expect(hooks.model!.get(arrow)).toMatchObject({ fill: null, line: "#4472c4", lineW: 2 });
  });

  it("moves the handles of a Word preset and of a LibreOffice shape, as modifiers", () => {
    const { hooks } = mount();
    const arrow = hooks.create!("rightArrow")!;
    document.body.appendChild(arrow);
    // In its picture's px: the shaft's edge, then the head's start 96 px across (of 144).
    const h = hooks.handles!(arrow)!;
    expect(h.points).toHaveLength(2);
    expect(Math.round(h.points[1]!.x - h.points[0]!.x)).toBe(96);
    hooks.moveHandle!(arrow, 1, h.points[1]!.x + 24, h.points[1]!.y);
    expect(JSON.parse(arrow.getAttribute("data-shape-edit")!).mods).toEqual([50000, 25000]);
    // A LibreOffice shape: its handle moves $0 across, within its range.
    const lo = document.createElement("span");
    lo.className = "odt-shape";
    lo.setAttribute("data-shape-wrap", "inline");
    lo.setAttribute("data-odt-xml", '<draw:custom-shape xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:anchor-type="as-char" svg:width="2cm" svg:height="1cm"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:modifiers="5400" draw:enhanced-path="M $0 0 L 21600 10800 $0 21600 0 10800 Z N"><draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/></draw:enhanced-geometry></draw:custom-shape>');
    document.body.appendChild(lo);
    hooks.redraw!(lo);
    const lh = hooks.handles!(lo)!;
    expect(lh.points).toHaveLength(1);
    hooks.moveHandle!(lo, 0, lh.points[0]!.x + 1000, lh.points[0]!.y); // far right: held at 10800
    expect(JSON.parse(lo.getAttribute("data-shape-edit")!).mods).toEqual([10800]);
  });

  it("gives a turned shape and each shape in a group their handles", () => {
    const { hooks } = mount();
    const NS = 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
    const geo = '<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:modifiers="5400" draw:enhanced-path="M $0 0 L 21600 10800 $0 21600 0 10800 Z N"><draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/></draw:enhanced-geometry>';
    const mk = (xml: string) => {
      const el = document.createElement("span");
      el.className = "odt-shape";
      el.setAttribute("data-shape-wrap", "front");
      el.setAttribute("data-odt-xml", xml);
      document.body.appendChild(el);
      hooks.redraw!(el);
      return el;
    };
    // Turned a quarter: its handle sits where the turned outline puts it, and a drag there keeps $0.
    const turned = mk(`<draw:custom-shape ${NS} text:anchor-type="paragraph" svg:width="2cm" svg:height="1cm" draw:transform="rotate (1.5707963267949) translate (1cm 3cm)">${geo}</draw:custom-shape>`);
    const tp = hooks.handles!(turned)!.points[0]!;
    hooks.moveHandle!(turned, 0, tp.x, tp.y);
    expect(JSON.parse(turned.getAttribute("data-shape-edit")!).mods[0]).toBeCloseTo(5400, 0);
    // A group: its second member's handle moves that member's modifiers only.
    const g = mk(`<draw:g ${NS} text:anchor-type="paragraph"><draw:custom-shape svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="1cm">${geo}</draw:custom-shape><draw:custom-shape svg:x="3cm" svg:y="0cm" svg:width="2cm" svg:height="1cm">${geo}</draw:custom-shape></draw:g>`);
    const gp = hooks.handles!(g)!.points;
    expect(gp).toHaveLength(2);
    hooks.moveHandle!(g, 1, gp[1]!.x + 1000, gp[1]!.y);
    expect(JSON.parse(g.getAttribute("data-shape-edit")!).memberMods).toEqual({ 1: [10800] });
  });
});
