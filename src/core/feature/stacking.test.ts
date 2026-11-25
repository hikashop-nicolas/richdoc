import { describe, expect, it } from "vitest";
import { restack, stack, topZ } from "./stacking";

function mount() {
  document.body.innerHTML = `<div id="doc"><p>Text
    <span class="docx-shape" id="a" data-shape-wrap="front" data-shape-z="10"></span>
    <img id="b" data-rdoc-wrap="front" data-rdoc-z="20">
    <span class="odt-shape" id="c" data-shape-wrap="behind" data-shape-z="0"><span class="odt-shape-box"></span></span>
    <img id="d" data-rdoc-wrap="behind" data-rdoc-z="5">
    <span class="docx-shape" id="e" data-shape-wrap="square" data-shape-z="30"></span></p></div>`;
  const doc = document.getElementById("doc")!;
  const el = (id: string) => document.getElementById(id) as HTMLElement;
  return { regions: [doc], el };
}

describe("stacking floating pictures and shapes", () => {
  it("paints each layer by order, in front of the text above it and behind it below", () => {
    const { regions, el } = mount();
    stack(regions);
    expect([el("a").style.zIndex, el("b").style.zIndex]).toEqual(["4", "5"]);
    expect([(el("c").firstElementChild as HTMLElement).style.zIndex, el("d").style.zIndex]).toEqual(["-2", "-1"]);
    expect(el("e").style.zIndex).toBe(""); // wrapped in the text: not stacked
  });

  it("brings one to the front and sends one to the back", () => {
    const { regions, el } = mount();
    expect(topZ(regions)).toBe(31); // above the wrapped shape's order too, for when it floats
    restack(el("a"), 1, regions);
    expect(el("a").getAttribute("data-shape-z")).toBe("31");
    expect(el("a").getAttribute("data-shape-dirty")).toBe("z");
    expect(el("a").style.zIndex).toBe("5");
    // Nothing below 0: the others are lifted a step, the odt drawing's edit too.
    restack(el("d"), -1, regions);
    expect(el("d").getAttribute("data-rdoc-z")).toBe("0");
    expect(JSON.parse(el("c").getAttribute("data-shape-edit")!)).toEqual({ z: 1 });
    expect(el("d").style.zIndex).toBe("-2");
  });
});
