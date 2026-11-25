import { describe, expect, it } from "vitest";
import { ooxmlGeometryXml } from "./ooxml-shape";
import { evaluateGeometry } from "./geometry";
import { presetGeometry } from "../docx/drawingml-geom";
import { PRESETS } from "../docx/preset-shapes";

const NS = 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:drawooo="http://openoffice.org/2010/draw"';
const parse = (xml: string) => new DOMParser().parseFromString(xml.replace("<draw:enhanced-geometry ", `<draw:enhanced-geometry ${NS} `), "application/xml").documentElement;
const nums = (d: string) => (d.match(/-?\d+(\.\d+)?(e[-+]?\d+)?/gi) ?? []).map(Number);

describe("Word presets as odt ooxml shapes", () => {
  it("writes the guides, adjust values, path and handles as LibreOffice does", () => {
    const xml = ooxmlGeometryXml("rightArrow")!;
    expect(xml).toMatch(/^<draw:enhanced-geometry svg:viewBox="0 0 0 0" draw:type="ooxml-rightArrow" draw:modifiers="50000 50000"/);
    expect(xml).toContain('<draw:equation draw:name="f0" draw:formula="100000*(logwidth)/(min(logwidth,logheight))"/>'); // maxAdj2
    expect(xml).toMatch(/<draw:handle draw:handle-position="0 \?f\d+" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="100000"\/>/);
  });

  it("draws every preset as Word draws it, at any size", () => {
    for (const prst of Object.keys(PRESETS)) {
      for (const [w, h] of [[144, 96], [60, 200]] as const) {
        const g = evaluateGeometry(parse(ooxmlGeometryXml(prst)!), w, h);
        const want = presetGeometry(prst, w, h) ?? [];
        const got = g?.paths ?? [];
        expect(got.length, prst).toBe(want.length);
        got.forEach((p, k) => {
          const a = nums(p.d), b = nums(want[k]!.d);
          expect(a.length, `${prst} path ${k}`).toBe(b.length);
          // The view box is the shape in 1/100 mm: back to px, pair by pair.
          const sx = w / g!.vb[2], sy = h / g!.vb[3];
          a.forEach((v, i) => expect(Math.abs(v * (i % 2 ? sy : sx) - b[i]!), `${prst} path ${k} #${i}`).toBeLessThan(0.05));
          expect(p.fill, prst).toBe(want[k]!.fill !== "none");
          expect(p.stroke, prst).toBe(want[k]!.stroke);
        });
      }
    }
  });
});
