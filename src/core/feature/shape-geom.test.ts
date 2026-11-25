import { describe, expect, it } from "vitest";
import { paintShapeGeom, shadeColor } from "./shape-geom";

describe("drawn shape outlines", () => {
  it("shades a colour toward black or white", () => {
    expect(shadeColor("#808080", -0.5)).toBe("#404040");
    expect(shadeColor("rgb(0, 0, 0)", 0.5)).toBe("#808080");
    expect(shadeColor("#123456", 0)).toBe("#123456");
  });

  it("repaints the outline from the box's fill and border after a shape bar edit", () => {
    const box = document.createElement("span");
    box.innerHTML = '<svg class="rdoc-shape-geom"><path data-fill="norm" data-stroke="1"/><path data-fill="darken" data-stroke="1"/><path data-head="solid"/></svg>';
    box.style.backgroundColor = "#808080";
    box.style.border = "2px solid #ff0000";
    paintShapeGeom(box);
    const [a, b, head] = Array.from(box.querySelectorAll("path"));
    expect(a!.getAttribute("fill")).toBe("rgb(128, 128, 128)");
    expect(b!.getAttribute("fill")).toBe("#4d4d4d");
    expect(a!.getAttribute("stroke")).toBe("rgb(255, 0, 0)");
    expect(a!.getAttribute("stroke-width")).toBe("2");
    expect(head!.getAttribute("fill")).toBe("rgb(255, 0, 0)");
    box.style.border = "none";
    box.style.backgroundColor = "";
    paintShapeGeom(box);
    expect(a!.getAttribute("fill")).toBe("none");
    expect(a!.getAttribute("stroke")).toBe("none");
  });
});
