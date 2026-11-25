import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync, zipSync, strToU8 } from "fflate";
import { docxShapeHooks } from "./shape-hooks";
import { docxToParts, htmlToDocx } from "./index";

const DOC = (body: string) => zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "_rels/.rels": strToU8("<Relationships/>"),
  "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`),
  "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
});

describe("docx shape hooks", () => {
  it("offers Word's groups of ready-made shapes with pictures", () => {
    const g = docxShapeHooks.gallery!();
    expect(g.map((x) => x.items.length).reduce((a, b) => a + b, 0)).toBeGreaterThan(140);
    expect(g[0]!.items.map((i) => i.id)).toContain("arrowLine");
    for (const grp of g) for (const it of grp.items) expect(it.icon).toMatch(/<path d="M/);
  });

  it("creates a drawn shape and redraws it at a new size", () => {
    const s = docxShapeHooks.create!("rightArrow");
    const box = s.querySelector<HTMLElement>(".docx-shape-box")!;
    expect(box.hasAttribute("data-geom")).toBe(true);
    const before = box.querySelector("svg path")!.getAttribute("d");
    expect(before).toBe("M0,24 L96,24 L96,0 L144,48 L96,96 L96,72 L0,72 Z");
    box.style.width = "288px";
    docxShapeHooks.redraw!(s);
    expect(box.querySelectorAll("svg").length).toBe(1);
    expect(box.querySelector("svg path")!.getAttribute("d")).toBe("M0,24 L240,24 L240,0 L288,48 L240,96 L240,72 L0,72 Z");
    expect(box.querySelector("svg path")!.getAttribute("fill")).toBe("rgb(68, 114, 196)");
  });

  it("moves an adjust handle, redraws, and saves the adjust value", () => {
    const s = docxShapeHooks.create!("rightArrow");
    // The shaft's edge (adj1) and the head's start (adj2), px in the 144 x 96 box.
    expect(docxShapeHooks.handles!(s)!.points).toEqual([{ x: 0, y: 24 }, { x: 96, y: 0 }]);
    docxShapeHooks.moveHandle!(s, 1, 120, 0); // a shorter head
    expect(JSON.parse(s.getAttribute("data-shape-adj")!)).toEqual({ adj1: "val 50000", adj2: "val 25000" });
    expect(s.querySelector("svg path")!.getAttribute("d")).toBe("M0,24 L120,24 L120,0 L144,48 L120,96 L120,72 L0,72 Z");
    expect(docxShapeHooks.handles!(s)!.points[1]).toEqual({ x: 120, y: 0 });
    const d = DOC("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    const xml = strFromU8(unzipSync(htmlToDocx(`<p>x${s.outerHTML}</p>`, d))["word/document.xml"]!);
    expect(xml).toContain('<a:prstGeom prst="rightArrow"><a:avLst><a:gd name="adj1" fmla="val 50000"/><a:gd name="adj2" fmla="val 25000"/></a:avLst></a:prstGeom>');
  });

  it("places a turned shape's handles turned with it, and reads a drag back through the turn", () => {
    const s = docxShapeHooks.create!("rightArrow");
    s.querySelector<HTMLElement>(".docx-shape-box")!.style.transform = "rotate(90deg)";
    // 144 x 96 turned a quarter clockwise about its centre (72, 48): (96, 0) lands at (120, 72).
    const p = docxShapeHooks.handles!(s)!.points[1]!;
    expect([Math.round(p.x), Math.round(p.y)]).toEqual([120, 72]);
    docxShapeHooks.moveHandle!(s, 1, 120, 96); // 24 px further down the turned shape: the head starts 120 across
    expect(JSON.parse(s.getAttribute("data-shape-adj")!).adj2).toBe("val 25000");
  });

  it("gives each preset in a group its handles, kept and saved per member", () => {
    const member = (x: number, prst: string) => `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>`;
    const grp = `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>${member(0, "rect")}${member(952500, "star5")}</wpg:wgp>`;
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
    const d = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"), "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">${grp}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
    });
    document.body.innerHTML = docxToParts(d).body;
    const s = document.querySelector<HTMLElement>(".docx-shape")!;
    const hs = docxShapeHooks.handles!(s)!.points;
    expect(hs).toHaveLength(1); // the star's (a rectangle has none), in the group: 100 px across
    expect(Math.round(hs[0]!.x)).toBe(150);
    docxShapeHooks.moveHandle!(s, 0, 150, 30);
    expect(Object.keys(JSON.parse(s.getAttribute("data-shape-adjs")!))).toEqual(["1"]);
    s.setAttribute("data-shape-dirty", "adj");
    const xml = strFromU8(unzipSync(htmlToDocx(document.body.innerHTML, d))["word/document.xml"]!);
    expect(xml).toMatch(/<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom>/);
    expect(xml).toMatch(/<a:prstGeom prst="star5"><a:avLst><a:gd name="adj" fmla="val \d+"\/>/);
  });

  it("gives a rounded rectangle its corner handle, drawn from its outline once moved", () => {
    const s = docxShapeHooks.create!("roundRect");
    const box = s.querySelector<HTMLElement>(".docx-shape-box")!;
    box.style.width = "144px"; box.style.height = "96px";
    expect(docxShapeHooks.handles!(s)!.points).toHaveLength(1);
    docxShapeHooks.moveHandle!(s, 0, 40, 0);
    expect(box.hasAttribute("data-geom")).toBe(true);
    expect(box.style.borderRadius).toBe("");
    expect(box.querySelector("svg path")!.getAttribute("d")).toMatch(/^M0,40 /);
  });

  it("creates plain boxes as CSS boxes", () => {
    expect(docxShapeHooks.create!("ellipse").querySelector(".docx-shape-box")!.hasAttribute("data-geom")).toBe(false);
  });

  it("saves an inserted preset with its geometry, and a line with its arrowheads and no text", () => {
    const d = DOC("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    const star = docxShapeHooks.create!("star5").outerHTML;
    const arrow = docxShapeHooks.create!("doubleArrowLine").outerHTML;
    const xml = strFromU8(unzipSync(htmlToDocx(`<p>x${star}${arrow}</p>`, d))["word/document.xml"]!);
    expect(xml).toContain('<a:prstGeom prst="star5">');
    expect(xml).toMatch(/<a:prstGeom prst="straightConnector1"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln w="19050"><a:solidFill><a:srgbClr val="4472C4"\/><\/a:solidFill><a:headEnd type="triangle"\/><a:tailEnd type="triangle"\/><\/a:ln>/);
    expect((xml.match(/<wps:txbx>/g) ?? []).length).toBe(1);
    // Read back, both are drawn from their presets.
    const back = docxToParts(htmlToDocx(`<p>x${star}${arrow}</p>`, d)).body;
    expect((back.match(/data-geom/g) ?? []).length).toBe(2);
    expect((back.match(/data-head=/g) ?? []).length).toBe(2);
  });
});
