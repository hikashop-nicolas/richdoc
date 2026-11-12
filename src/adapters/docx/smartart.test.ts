import { describe, expect, it } from "vitest";
import { strToU8 } from "fflate";
import { MAX_ITEMS, findSmartArtDrawing, smartArtItems, smartArtSvg, type SmartArtItem } from "./smartart";
import { Warp, readTextEffects, readTextWarp, wordArtSvg } from "./wordart";

// wordcraft's default theme colours (accent1 156082), as readThemeColors returns them.
const COLORS: Record<string, string> = { dk1: "000000", lt1: "FFFFFF", dk2: "0E2841", lt2: "E8E8E8", accent1: "156082", accent2: "E97132" };
const NS = `xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const drawingXml = (body: string) => `<dsp:drawing ${NS}><dsp:spTree>${body}</dsp:spTree></dsp:drawing>`;
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").documentElement;
const items = (body: string, w: number, h: number): SmartArtItem[] => smartArtItems(parse(drawingXml(body)), COLORS, w, h, (id) => (id === "rId1" ? "image1.png" : undefined));
// A dsp:sp of 952500 EMU (100 px) square with the given xfrm attributes, preset and extra spPr XML.
const sp = (xfrm: string, prst: string, rest: string) =>
  `<dsp:sp><dsp:spPr><a:xfrm ${xfrm}><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${rest}</dsp:spPr></dsp:sp>`;
const texts = (out: SmartArtItem[]) => out.flatMap((i) => (i.k === "text" ? [i] : []));
const wellFormed = (svg: string | undefined) => !!svg && new DOMParser().parseFromString(svg, "image/svg+xml").getElementsByTagName("parsererror").length === 0;

describe("SmartArt drawing", () => {
  it("draws a round rect in its theme colour with its text", () => {
    const body = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></dsp:spPr><dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2000" b="1"/><a:t>Foo</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    const out = items(body, 100, 50);
    expect(out[0]).toMatchObject({ k: "geom", prst: "roundRect", fill: "156082", stroke: undefined, rect: [0, 0, 100, 50] });
    const t = out[1]!;
    expect(t).toMatchObject({ k: "text", text: "Foo", color: "FFFFFF", bold: true, align: "center" });
    expect(t.k === "text" && t.size).toBeCloseTo(20 * (4 / 3), 5);
    const svg = smartArtSvg(drawingXml(body), COLORS, 100, 50)!;
    expect(wellFormed(svg)).toBe(true);
    expect(svg).toContain('fill="#156082"');
    expect(svg).toContain(">Foo</text>");
    expect(svg).toContain('font-weight="bold"');
  });

  it("draws a chevron as a closed six-point path", () => {
    const svg = smartArtSvg(drawingXml(sp("", "chevron", '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>')), COLORS, 100, 100)!;
    const d = /<path d="([^"]+)"/.exec(svg)![1]!;
    expect(d.match(/[ML]/g)!.length).toBe(6);
    expect(d.endsWith("Z")).toBe(true);
    // The tip is at the middle of the right edge.
    expect(d).toContain("L100 50");
  });

  it("draws other presets from their outlines, unknown ones as rectangles", () => {
    for (const prst of ["homePlate", "rightArrow", "triangle", "diamond", "hexagon", "pentagon", "ellipse", "star5", "heart", "pie", "blockArc", "leftRightArrow", "madeUp"]) {
      const svg = smartArtSvg(drawingXml(sp("", prst, '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')), COLORS, 100, 100);
      expect(wellFormed(svg), prst).toBe(true);
      expect(svg, prst).toMatch(/<path d="M[^"]+Z" fill="#FF0000"/);
    }
    expect(smartArtSvg(drawingXml(sp("", "madeUp", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')), COLORS, 100, 100)).toContain('d="M0 0L100 0L100 100L0 100Z"');
    // A line is open: no fill.
    expect(smartArtSvg(drawingXml(sp("", "line", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')), COLORS, 100, 100)).toContain('d="M0 0L100 100" fill="none"');
  });

  it("turns a rotated shape about its centre", () => {
    const out = items(sp('rot="2700000"', "rect", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'), 100, 100);
    expect(out[0]).toMatchObject({ k: "geom", rot: 45, fill: "FF0000" });
    const svg = smartArtSvg(drawingXml(sp('rot="2700000"', "rect", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')), COLORS, 100, 100)!;
    expect(svg).toContain('transform="rotate(45 50 50)"');
    // A 100 px square turned 45 degrees: its top-left corner lands at (50, -20.71).
    const a = Math.PI / 4, [dx, dy] = [-50, -50];
    expect(50 + dx * Math.cos(a) - dy * Math.sin(a)).toBeCloseTo(50, 2);
    expect(50 + dx * Math.sin(a) + dy * Math.cos(a)).toBeCloseTo(-20.71, 2);
    expect(smartArtSvg(drawingXml(sp('flipH="1"', "rightArrow", "")), COLORS, 100, 100)).toContain("scale(-1 1)");
  });

  it("darkens a theme colour with lumMod", () => {
    const out = items(sp("", "rect", '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'), 10, 10);
    const fill = (out[0] as { fill: string }).fill;
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(fill.slice(i, i + 2), 16));
    expect(r! < 0x15 && g! < 0x60 && b! < 0x82).toBe(true);
  });

  it("takes the style's colours when the shape has none, and lines from a:ln", () => {
    const body = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></dsp:spPr><dsp:style><a:lnRef idx="2"><a:schemeClr val="lt1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef></dsp:style></dsp:sp>`;
    expect(items(body, 100, 100)[0]).toMatchObject({ fill: "E97132", stroke: "112233", strokeWidth: 2 });
    const none = body.replace('idx="1"', 'idx="0"');
    expect(items(none, 100, 100)[0]).toMatchObject({ fill: undefined });
  });

  it("draws a picture fill as an image", () => {
    const out = items(sp("", "ellipse", '<a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'), 10, 10);
    expect(out[0]).toMatchObject({ k: "image", href: "image1.png" });
  });

  it("scales the drawing to the frame and maps group children", () => {
    // A 100 pt drawing in a 96 px (72 pt) frame: a group at the origin holds a half-size child.
    const grp = `<dsp:grpSp><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="2540000"/></a:xfrm></dsp:grpSpPr><dsp:sp><dsp:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr></dsp:sp></dsp:grpSp>`;
    const r = (items(grp, 96, 96)[0] as { rect: number[] }).rect;
    r.forEach((v) => expect(v).toBeCloseTo(48, 5));
    // A drawing smaller than its frame keeps its own size.
    expect((items(sp("", "rect", ""), 500, 500)[0] as { rect: number[] }).rect).toEqual([0, 0, 100, 100]);
  });

  it("wraps long text into several lines", () => {
    const body = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr lIns="0" rIns="0"/><a:p><a:r><a:rPr sz="1000"/><a:t>one two three four five six</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    const t = texts(items(body, 40, 40)).map((x) => x.text);
    expect(t.length).toBeGreaterThan(1);
    expect(t.join(" ")).toBe("one two three four five six");
  });

  it("sizes text by its autofit and the drawing's fit", () => {
    const body = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr><a:p><a:r><a:rPr sz="2000"/><a:t>Big</a:t></a:r></a:p><a:p><a:r><a:rPr/><a:t>Small</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    const sizes = (w: number) => texts(items(body, w, w)).map((x) => [x.text, Math.round(x.size * 0.75 * 100) / 100]);
    expect(sizes(100)).toEqual([["Big", 10], ["Small", 9]]);
    expect(sizes(50)).toEqual([["Big", 5], ["Small", 4.5]]);
  });

  it("falls back to the theme's dark or light text colour", () => {
    const t = (rest: string) => texts(items(`<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>${rest}</dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:pPr algn="r"/><a:r><a:t>x</a:t></a:r></a:p></dsp:txBody></dsp:sp>`, 100, 100))[0]!;
    expect(t("")).toMatchObject({ color: "000000", align: "right" });
    expect(t('<a:solidFill><a:srgbClr val="333333"/></a:solidFill>').color).toBe("FFFFFF");
  });

  it("survives hostile numbers and deep nesting", () => {
    const huge = `<dsp:sp><dsp:spPr><a:xfrm rot="zzz" flipH="1"><a:off x="99999999999999999999999" y="-1e400"/><a:ext cx="-5" cy="NaN"/></a:xfrm><a:prstGeom prst="leftRightArrow"><a:avLst><a:gd name="adj1" fmla="val 1e999"/><a:gd name="adj2" fmla="val -7"/></a:avLst></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr lIns="99999999999999999999" tIns="-9"/><a:p><a:r><a:rPr sz="99999999999999"/><a:t>x</a:t></a:r></a:p></dsp:txBody></dsp:sp>${sp('rot="2700000"', "chevron", "")}`;
    let nested = "";
    for (let i = 0; i < 40; i++) nested = `<dsp:grpSp><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></dsp:grpSpPr>${nested}${sp("", "rect", "")}</dsp:grpSp>`;
    expect(() => items(huge, NaN, -3)).not.toThrow();
    expect(() => items(huge, Infinity, 1e30)).not.toThrow();
    expect(items(nested, 100, 100).length).toBeLessThanOrEqual(8);
    expect(wellFormed(smartArtSvg(drawingXml(huge), {}, NaN, -3))).toBe(true);
    expect(smartArtItems(parse(drawingXml(huge)), {}, 10, 10).length).toBeLessThanOrEqual(MAX_ITEMS);
  });

  it("stops one huge text body at the item limit", () => {
    const words = "word ".repeat(MAX_ITEMS + 100);
    const body = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr lIns="0" rIns="0"/><a:p><a:r><a:t>${words}</a:t></a:r></a:p><a:p><a:r><a:t>${words}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    expect(items(body, 1, 1).length).toBe(MAX_ITEMS);
  });

  it("keeps text far outside nested groups finite", () => {
    let tree = `<dsp:sp><dsp:spPr><a:xfrm><a:off x="-1000000000" y="0"/><a:ext cx="1000000000" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Far</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    for (let i = 0; i < 8; i++) tree = `<dsp:grpSp><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000000" cy="1"/><a:chOff x="0" y="0"/><a:chExt cx="1" cy="1"/></a:xfrm></dsp:grpSpPr>${tree}</dsp:grpSp>`;
    const rects = texts(items(tree, 100, 100)).map((t) => t.rect);
    expect(rects.length).toBeGreaterThan(0);
    expect(rects.flat().every(Number.isFinite)).toBe(true);
  });

  it("returns nothing for junk", () => {
    expect(smartArtSvg("not xml at all <", COLORS, 100, 100)).toBeUndefined();
    expect(smartArtSvg("<a/>", COLORS, 100, 100)).toBeUndefined();
    expect(smartArtSvg(drawingXml(""), COLORS, 100, 100)).toBeUndefined();
    expect(smartArtSvg(drawingXml("<dsp:sp/>"), COLORS, 100, 100)).toBeUndefined();
  });
});

describe("SmartArt drawing lookup", () => {
  const gd = (dm = 'r:dm="rId4"') =>
    parse(`<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds ${dm} r:lo="rId5" r:qs="rId6" r:cs="rId7"/></a:graphicData>`);
  const data = (ext: string) =>
    strToU8(`<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">${ext}</a:ext></dgm:extLst></dgm:dataModel>`);
  const dsp = (n: number) => drawingXml(sp("", "rect", `<a:solidFill><a:srgbClr val="00000${n}"/></a:solidFill>`));
  const files = (ext: string): Record<string, Uint8Array> => ({
    "word/diagrams/data1.xml": data(ext),
    "word/diagrams/drawing1.xml": strToU8(dsp(1)),
    "word/diagrams/drawing2.xml": strToU8(dsp(2)),
  });
  const ext = (id: string) => `<dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="${id}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`;

  it("follows the data part's dataModelExt relId through the document rels", () => {
    const rels = new Map([["rId4", "diagrams/data1.xml"], ["rId8", "diagrams/drawing1.xml"], ["rId9", "diagrams/drawing2.xml"]]);
    expect(findSmartArtDrawing(gd(), rels, files(ext("rId9")))).toBe(dsp(2));
  });

  it("falls back to the drawing numbered like the data part", () => {
    const rels = new Map([["rId4", "/word/diagrams/data1.xml"], ["rId9", "diagrams/drawing2.xml"], ["rId8", "diagrams/drawing1.xml"]]);
    expect(findSmartArtDrawing(gd(), rels, files(""))).toBe(dsp(1));
    // A relId that isn't a drawing part is ignored.
    rels.set("rId3", "styles.xml");
    expect(findSmartArtDrawing(gd(), rels, { ...files(ext("rId3")), "word/styles.xml": strToU8("<w:styles xmlns:w='x'/>") })).toBe(dsp(1));
  });

  it("returns nothing without a data part or drawing", () => {
    const rels = new Map([["rId4", "diagrams/data1.xml"]]);
    expect(findSmartArtDrawing(gd('r:lo="rId5"'), rels, files(""))).toBeUndefined();
    expect(findSmartArtDrawing(gd(), new Map(), files(""))).toBeUndefined();
    expect(findSmartArtDrawing(gd(), rels, files(ext("rId8")))).toBeUndefined();
    expect(findSmartArtDrawing(gd(), new Map([["rId4", "diagrams/data1.xml"], ["rId8", "diagrams/drawing1.xml"]]), { "word/diagrams/data1.xml": strToU8("junk <"), "word/diagrams/drawing1.xml": strToU8("junk") })).toBeUndefined();
  });

  it("draws what it finds", () => {
    const rels = new Map([["rId4", "diagrams/data1.xml"], ["rId8", "diagrams/drawing1.xml"]]);
    expect(smartArtSvg(findSmartArtDrawing(gd(), rels, files(ext("rId8")))!, COLORS, 100, 100)).toContain('fill="#000001"');
  });
});

describe("WordArt", () => {
  const W = `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
  const wsp = (rpr: string, warp: string, extra = "") =>
    parse(`<wps:wsp ${W}><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${extra}</wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr>${rpr}</w:rPr><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="none">${warp}<a:spAutoFit/></wps:bodyPr></wps:wsp>`);
  const fx = (inner: string) => readTextEffects(parse(`<w:rPr ${W}>${inner}</w:rPr>`), COLORS);

  it("lays arch baselines on the arc", () => {
    const frame = { x: 10, y: 20, w: 200, h: 100 };
    const w = Warp.create(readTextWarp(parse(`<a:prstTxWarp xmlns:a="x" prst="textArchUp"/>`))!, frame, 0.25)!;
    const [cx, cy, rx, ry] = [110, 70, 100, 50];
    for (const v of [0, 0.8]) {
      const rho = Array.from({ length: 11 }, (_, i) => {
        const [x, y] = w.map(i / 10, v);
        expect(y).toBeLessThanOrEqual(cy + 1e-3);
        return Math.hypot((x - cx) / rx, (y - cy) / ry);
      });
      rho.forEach((r) => expect(Math.abs(r - rho[0]!)).toBeLessThan(1e-3));
    }
    const [x0] = w.map(0, 0), [xm, ym] = w.map(0.5, 0), [x1] = w.map(1, 0);
    expect(x0 < xm && xm < x1).toBe(true);
    expect(ym).toBeCloseTo(20, 3);
    const d = Warp.create({ preset: "textArchDown", adj: {} }, frame, 0.25)!;
    expect(d.map(0.5, 1)[1]).toBeGreaterThan(119);
  });

  it("draws plain and unknown warps flat and keeps hostile values in the frame", () => {
    const frame = { x: 0, y: 0, w: 100, h: 50 };
    expect(Warp.create({ preset: "textPlain", adj: {} }, frame, 0.2)).toBeUndefined();
    expect(Warp.create({ preset: "textCanUp", adj: {} }, frame, 0.2)).toBeUndefined();
    expect(readTextWarp(parse(`<a:prstTxWarp xmlns:a="x" prst="textMadeUp"/>`))).toBeUndefined();
    const t = readTextWarp(parse(`<a:prstTxWarp xmlns:a="x" prst="textSlantUp"><a:avLst><a:gd name="adj" fmla="val 9223372036854775807"/><a:gd name="bad name!" fmla="val 1"/></a:avLst></a:prstTxWarp>`))!;
    expect(t.adj).toEqual({ adj: 100_000_000 });
    const w = Warp.create(t, frame, NaN)!;
    for (const [u, v] of [[0, 0], [1, 1], [NaN, 0.5]] as const) {
      const [x, y] = w.map(u, v);
      expect(x >= 0 && x <= 100 && y >= 0 && y <= 50).toBe(true);
    }
  });

  it("reads w14 fills and outlines", () => {
    const a = fx(`<w14:textOutline w14:w="9525"><w14:solidFill><w14:schemeClr w14:val="accent1"/></w14:solidFill></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="ff0000"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:solidFill></w14:textFill>`);
    expect(a.outline).toEqual({ color: "156082", width: 0.75, transparency: 0 });
    expect(a.fill).toEqual({ t: "solid", color: "FF0000", transparency: 40 });
    expect(fx(`<w14:textOutline><w14:noFill/></w14:textOutline><w14:textFill><w14:noFill/></w14:textFill>`)).toEqual({ outline: { color: null, width: 0, transparency: 0 }, fill: { t: "none" } });
    const gs = Array.from({ length: 50 }, (_, i) => `<w14:gs w14:pos="${100_000 - i * 1000}"><w14:srgbClr w14:val="000000"/></w14:gs>`).join("");
    const g = fx(`<w14:textFill><w14:gradFill><w14:gsLst>${gs}</w14:gsLst><w14:lin w14:ang="5400000"/></w14:gradFill></w14:textFill>`).fill!;
    expect(g.t === "grad" && g.stops.length).toBe(10);
    expect(g.t === "grad" && g.stops[0]!.pos).toBe(91);
    expect(g.t === "grad" && g.angle).toBe(90);
    const dark = fx(`<w14:textFill><w14:solidFill><w14:schemeClr w14:val="accent1"><w14:lumMod w14:val="50000"/></w14:schemeClr></w14:solidFill></w14:textFill>`).fill;
    expect(dark?.t === "solid" && dark.color !== "156082").toBe(true);
  });

  it("bends warped letters along the warp", () => {
    const svg = wordArtSvg(wsp(`<w:sz w:val="72"/>`, `<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>`), COLORS, 300, 150)!;
    expect(wellFormed(svg)).toBe(true);
    const tr = svg.match(/transform="matrix\(([^)]+)\)"/g)!;
    expect(tr.length).toBe(5);
    // The arch turns the first and last letters in opposite directions.
    const b = tr.map((m) => Number(/matrix\(([^ ]+) ([^ ]+)/.exec(m)![2]));
    expect(b[0]! < 0 && b[4]! > 0).toBe(true);
    expect(svg).toContain(">H</text>");
  });

  it("draws effects flat when the warp isn't bent", () => {
    const rpr = `<w:b/><w:sz w:val="72"/><w14:textOutline w14:w="12700"><w14:solidFill><w14:srgbClr w14:val="0000FF"/></w14:solidFill></w14:textOutline><w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:schemeClr w14:val="accent1"/></w14:gs><w14:gs w14:pos="100000"><w14:schemeClr w14:val="accent2"/></w14:gs></w14:gsLst><w14:lin w14:ang="5400000"/></w14:gradFill></w14:textFill><w14:shadow w14:blurRad="38100" w14:dist="25400" w14:dir="2700000"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow><w14:reflection w14:stA="50000" w14:endPos="40000"/>`;
    const svg = wordArtSvg(wsp(rpr, ""), COLORS, 300, 80)!;
    expect(wellFormed(svg)).toBe(true);
    expect(svg).not.toMatch(/<text[^>]*matrix/);
    expect(svg).toContain('text-anchor="middle"');
    expect(svg).toMatch(/<linearGradient[^>]*><stop offset="0" stop-color="#156082"\/><stop offset="1" stop-color="#E97132"\/>/);
    expect(svg).toContain("<clipPath");
    expect(svg).toContain('stroke="#0000FF"');
    expect(svg).toContain("feGaussianBlur");
    expect(svg).toContain('matrix(1 0 0 -1');
    expect(svg).toContain('font-weight="bold"');
  });

  it("draws the run colour when only the warp makes it WordArt", () => {
    const svg = wordArtSvg(wsp(`<w:color w:val="C00000"/>`, `<a:prstTxWarp prst="textPlain"><a:avLst/></a:prstTxWarp>`), COLORS, 200, 50)!;
    expect(svg).toContain('fill="#C00000"');
    expect(wordArtSvg(wsp(`<w:color w:val="C00000"/>`, `<a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp>`, `<a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>`), COLORS, 200, 50)).toContain('fill="#EEEEEE"');
  });

  it("returns nothing for ordinary or junk shapes", () => {
    expect(wordArtSvg(wsp(`<w:b/>`, ""), COLORS, 100, 50)).toBeUndefined();
    expect(wordArtSvg(parse(`<wps:wsp ${W}><wps:bodyPr><a:prstTxWarp prst="textArchUp"/></wps:bodyPr></wps:wsp>`), COLORS, 100, 50)).toBeUndefined();
    expect(wordArtSvg(wsp("", `<a:prstTxWarp prst="textArchUp"/>`), COLORS, NaN, 50)).toBeUndefined();
    expect(wordArtSvg(parse("<junk/>"), COLORS, 100, 50)).toBeUndefined();
  });
});
