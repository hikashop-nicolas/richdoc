import { describe, expect, it } from "vitest";
import { adjustToward, adjustValues, presetHandles } from "./drawingml-geom";

describe("preset adjust handles", () => {
  it("places a preset's handles from its adjust values", () => {
    // A right arrow 144 x 96: the shaft's top edge (adj1 50%) and the head's start (adj2 50% of 96).
    expect(presetHandles("rightArrow", 144, 96)).toEqual([
      { x: 0, y: 24, refs: [{ name: "adj1", axis: "y", min: 0, max: 100000 }] },
      { x: 96, y: 0, refs: [{ name: "adj2", axis: "x", min: 0, max: 150000 }] },
    ]);
    expect(adjustValues("rightArrow", { adj2: "val 25000" })).toEqual({ adj1: 50000, adj2: 25000 });
  });

  it("finds the adjust value that brings a handle under the pointer", () => {
    for (const [prst, i, px, py] of [["rightArrow", 0, 0, 40], ["rightArrow", 1, 60, 0], ["star5", 0, 72, 30], ["roundRect", 0, 20, 0], ["blockArc", 1, 30, 48], ["wedgeRectCallout", 0, 130, 110]] as const) {
      const adj = adjustToward(prst, 144, 96, {}, i, px, py);
      const h = presetHandles(prst, 144, 96, adj)[i]!;
      const err = h.refs.some((r) => r.axis === "r" || r.axis === "ang") ? Math.hypot(h.x - px, h.y - py) : Math.max(h.refs.some((r) => r.axis === "x") ? Math.abs(h.x - px) : 0, h.refs.some((r) => r.axis === "y") ? Math.abs(h.y - py) : 0);
      expect(err, `${prst} ${i}: ${JSON.stringify(adj)} -> ${h.x},${h.y}`).toBeLessThan(prst === "blockArc" ? 2 : 1);
    }
  });

  it("keeps a value inside its range", () => {
    const adj = adjustToward("rightArrow", 144, 96, {}, 0, 0, -500);
    expect(adj.adj1).toBe("val 100000");
  });
});
