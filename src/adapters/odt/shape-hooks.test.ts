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
    expect(star.getAttribute("data-odt-xml")).toMatch(/^<draw:custom-shape [^>]*text:anchor-type="as-char"[^>]*><draw:enhanced-geometry svg:viewBox="0 0 14400 9600" draw:type="non-primitive" draw:enhanced-path="M /);
    expect(star.querySelector("svg path")?.getAttribute("fill")).toBe("#4472c4");
    const arrow = hooks.create!("arrowLine")!;
    expect(arrow.getAttribute("data-odt-xml")).toMatch(/^<draw:path /);
    expect(arrow.querySelectorAll("svg path").length).toBeGreaterThan(2); // the line, its hit area and its head
    expect(hooks.model!.get(arrow)).toMatchObject({ fill: null, line: "#4472c4", lineW: 2 });
  });
});
