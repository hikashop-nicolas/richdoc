// The toolbar picker wiring over fontshelf (detection and catalog tests live in fontshelf).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setupFontPicker } from "./font-picker";

describe("font picker", () => {
  beforeEach(() => {
    (globalThis as unknown as Record<symbol, unknown>)[Symbol.for("fontshelf.state.v1")] = undefined;
    const mem = new Map<string, string>();
    vi.stubGlobal("localStorage", { getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v) });
  });
  afterEach(() => { vi.unstubAllGlobals(); document.body.replaceChildren(); });
  const mount = (fontOptions?: Parameters<typeof setupFontPicker>[0]["fontOptions"]) => {
    const wrap = document.createElement("div");
    document.body.appendChild(wrap);
    const apply = vi.fn();
    const p = setupFontPicker({ wrap, docFonts: ["Calibri"], fontOptions, getActiveEl: () => wrap, captureSel: () => {}, restoreSel: () => {}, apply });
    return { p, apply, wrap, values: () => Array.from(p.select.options).map((o) => o.value) };
  };
  it("offers web fonts only when the host turns them on", () => {
    expect(mount().values().some((v) => v.endsWith("remote"))).toBe(false);
    expect(mount({ remote: false }).values().some((v) => v.endsWith("remote"))).toBe(false);
    expect(mount({ remote: { provider: "google" } }).values().some((v) => v.endsWith("remote"))).toBe(true);
    expect(mount({ local: false }).values().some((v) => v.endsWith("local"))).toBe(false);
  });
  it("lists document fonts and applies a pick", () => {
    const { p, apply } = mount();
    expect(p.has("Calibri")).toBe(true);
    p.select.value = "Calibri";
    p.select.dispatchEvent(new Event("change"));
    expect(apply).toHaveBeenCalledWith("Calibri");
  });
  it("runs an action entry without applying it, restoring the previous choice", () => {
    const { p, apply, wrap } = mount({ remote: { provider: "google" } });
    p.select.value = "Calibri";
    p.select.dispatchEvent(new Event("focus"));
    p.select.value = p.select.options[p.select.options.length - 1]!.value;
    p.select.dispatchEvent(new Event("change"));
    expect(apply).not.toHaveBeenCalled();
    expect(p.select.value).toBe("Calibri");
    const consent = wrap.querySelector(".fontshelf-consent"); // consent comes first, inside the editor
    expect(consent?.textContent).toContain("GitHub");
  });
  it("ignores an action reached by arrowing through a closed select", () => {
    const { p, wrap } = mount({ remote: { provider: "google" } });
    p.select.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown" }));
    p.select.value = p.select.options[p.select.options.length - 1]!.value;
    p.select.dispatchEvent(new Event("change"));
    expect(wrap.querySelector(".fontshelf-consent")).toBeNull();
  });
});
