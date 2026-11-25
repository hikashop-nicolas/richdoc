import { describe, expect, it } from "vitest";
import { evaluateGeometry } from "./geometry";

const geo = (attrs: string, eqs: Record<string, string> = {}) =>
  new DOMParser().parseFromString(
    `<draw:enhanced-geometry xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ${attrs}>` +
      Object.entries(eqs).map(([n, f]) => `<draw:equation draw:name="${n}" draw:formula="${f}"/>`).join("") +
      "</draw:enhanced-geometry>",
    "application/xml",
  ).documentElement;

describe("odt enhanced geometry", () => {
  it("evaluates functions, constants and chained equations", () => {
    const g = evaluateGeometry(geo('draw:modifiers="100" draw:enhanced-path="M ?f2 ?f3 L right bottom"', { f0: "$0 *2", f1: "if(?f0 -150,max(?f0 ,10),0)", f2: "?f1 +abs(-5)", f3: "sqrt(16)*cos(0)" }), 100, 100)!;
    expect(g.paths[0]!.d).toBe("M205,4 L21600,21600");
  });

  it("splits subpaths and keeps their no-fill and no-stroke marks", () => {
    const g = evaluateGeometry(geo('draw:enhanced-path="M 0 0 L 10 0 Z N F M 0 5 L 5 5 N S M 1 1 L 2 2 N"'), 10, 10)!;
    expect(g.paths.map((p) => [p.fill, p.stroke])).toEqual([[true, true], [false, true], [true, false]]);
  });

  it("draws an angle ellipse as Bézier quarters round its centre", () => {
    const g = evaluateGeometry(geo('svg:viewBox="0 0 20 20" draw:enhanced-path="U 10 10 10 10 0 360 Z N"'), 20, 20)!;
    expect(g.paths[0]!.d.startsWith("M20,10 C")).toBe(true);
    expect(g.paths[0]!.d.match(/C/g)).toHaveLength(4);
    expect(g.paths[0]!.d).toContain(" 10,0 C"); // 90 degrees is up: y grows downward
  });

  it("mirrors horizontally inside the view box", () => {
    const g = evaluateGeometry(geo('svg:viewBox="0 0 100 100" draw:mirror-horizontal="true" draw:enhanced-path="M 0 0 L 30 50"'), 10, 10)!;
    expect(g.paths[0]!.d).toBe("M100,0 L70,50");
  });

  it("gives nothing for a geometry without a path", () => {
    expect(evaluateGeometry(geo('draw:type="mso-spt100"'), 10, 10)).toBeUndefined();
  });
});
