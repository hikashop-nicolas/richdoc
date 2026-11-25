import { describe, expect, it } from "vitest";
import { colorChooser, setupPopovers } from "./popover";

describe("toolbar panels", () => {
  it("opens one panel at a time from its button, and closes on Escape or a press elsewhere", () => {
    document.body.innerHTML = '<div id="w"><div id="bar"></div></div><p id="out">x</p>';
    const wrap = document.getElementById("w")!, bar = document.getElementById("bar")!;
    const pops = setupPopovers(wrap, bar);
    const mk = (n: string) => { const b = document.createElement("button"); b.textContent = n; bar.appendChild(b); const c = document.createElement("div"); c.innerHTML = `<input aria-label="${n}">`; return { b, c }; };
    const a = mk("a"), b = mk("b");
    let refreshed = 0;
    pops.add(a.b, "A", a.c, () => refreshed++);
    pops.add(b.b, "B", b.c);
    const panel = (x: { c: HTMLElement }) => x.c.parentElement!;
    expect(a.b.getAttribute("aria-haspopup")).toBe("dialog");
    a.b.click();
    expect(panel(a).hidden).toBe(false);
    expect(a.b.getAttribute("aria-expanded")).toBe("true");
    expect(refreshed).toBe(1);
    expect(panel(a).getAttribute("role")).toBe("dialog");
    b.b.click();
    expect(panel(a).hidden).toBe(true);
    expect(panel(b).hidden).toBe(false);
    panel(b).dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(panel(b).hidden).toBe(true);
    expect(document.activeElement).toBe(b.b);
    a.b.click();
    document.getElementById("out")!.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true }));
    expect(panel(a).hidden).toBe(true);
    pops.teardown();
  });

  it("offers swatches, none and any colour, and shows the current one", () => {
    const picked: (string | null)[] = [];
    const c = colorChooser({ noneLabel: "No fill", otherLabel: "Other", onPick: (x) => picked.push(x) });
    c.set("#C00000");
    expect(c.el.querySelector('[aria-label="#c00000"]')!.getAttribute("aria-pressed")).toBe("true");
    (c.el.querySelector('[aria-label="#4472c4"]') as HTMLButtonElement).click();
    (c.el.querySelector('[aria-label="No fill"]') as HTMLButtonElement).click();
    expect(picked).toEqual(["#4472c4", null]);
  });
});
