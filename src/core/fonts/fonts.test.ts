import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  BASE_FONTS, LOCAL_CANDIDATES, WEB_FONTS, axisFallbacks, buildFontGroups, detectFonts, detectFontsAsync,
  fontPresent, googleCssUrl, loadWebFont, mergeNames, readFontPrefs, writeFontPrefs, type Measure,
} from "./index";
import { setupFontPicker } from "../feature/toolbar/font-picker";

// A fake canvas: each family has its own widths per script; unknown families draw as the fallback.
const LATIN: Record<string, number> = { monospace: 10, serif: 11, "sans-serif": 12, Helvetica: 12, "Fake Sans": 13 };
const CJK_ONLY = new Set(["Fake Mincho"]);
const fakeMeasure: Measure = (font) => {
  const fams = font.replace(/^\d+px\s+/, "").split(",").map((f) => f.trim().replace(/^"|"$/g, ""));
  const generic = fams[fams.length - 1]!;
  const first = fams[0]!;
  const latin = LATIN[first] ?? LATIN[generic]!;
  const cjk = CJK_ONLY.has(first) ? 99 : 50;
  return [latin, cjk];
};

describe("font detection", () => {
  it("finds installed fonts and rejects unknown ones", () => {
    expect(fontPresent("Fake Sans", fakeMeasure)).toBe(true);
    expect(fontPresent("Nope Font", fakeMeasure)).toBe(false);
  });
  it("counts a font equal to one generic as present when another generic differs", () => {
    expect(fontPresent("Helvetica", fakeMeasure)).toBe(true); // same as sans-serif, not as serif
  });
  it("catches a font that only changes the CJK part of the probe", () => {
    expect(fontPresent("Fake Mincho", fakeMeasure)).toBe(true);
  });
  it("never reports a generic family", () => {
    expect(fontPresent("serif", fakeMeasure)).toBe(false);
  });
  it("treats an unmeasurable font as absent", () => {
    expect(fontPresent("Fake Sans", () => null)).toBe(false);
  });
  it("caches per name", () => {
    const m = vi.fn(fakeMeasure);
    const cache = new Map<string, boolean>();
    fontPresent("Fake Sans", m, cache);
    const calls = m.mock.calls.length;
    fontPresent("fake sans", m, cache);
    expect(m.mock.calls.length).toBe(calls);
  });
  it("detects in time slices, keeping input order", async () => {
    const names = ["Nope", "Fake Sans", "Fake Mincho", ...Array.from({ length: 50 }, (_, i) => `X${i}`)];
    expect(await detectFontsAsync(names, fakeMeasure, undefined, 0)).toEqual(["Fake Sans", "Fake Mincho"]);
    expect(detectFonts(names, fakeMeasure)).toEqual(["Fake Sans", "Fake Mincho"]);
  });
});

describe("font lists", () => {
  it("merges case-insensitively, sorted", () => {
    expect(mergeNames(["b", "Arial"], ["arial", "C", " a "])).toEqual(["a", "Arial", "b", "C"]);
  });
  it("splits document fonts from the rest and keeps only drawable base fonts", () => {
    const g = buildFontGroups({
      docFonts: ["Calibri", "Arial"],
      base: ["Arial", "Century", "Georgia"],
      extra: ["Klee One", "georgia"],
      canDraw: (n) => n !== "Calibri" && n !== "Century",
    });
    expect(g.doc).toEqual([{ name: "Arial", missing: false }, { name: "Calibri", missing: true }]);
    expect(g.fonts).toEqual(["Georgia", "Klee One"]);
  });
  it("has the built-in lists the picker relies on", () => {
    expect(BASE_FONTS).toContain("Arial");
    expect(LOCAL_CANDIDATES.length).toBeGreaterThan(300);
    for (const f of ["Hiragino Mincho ProN", "Yu Gothic", "Meiryo", "Noto Sans CJK JP", "Microsoft YaHei", "Malgun Gothic", "Carlito"]) expect(LOCAL_CANDIDATES).toContain(f);
    const fams = WEB_FONTS.map((w) => w.family);
    expect(new Set(fams).size).toBe(fams.length);
    for (const f of ["Arimo", "Tinos", "Cousine", "Carlito", "Caladea", "Klee One", "Noto Sans JP", "Nanum Myeongjo", "Amiri"]) expect(fams).toContain(f);
  });
});

describe("Google Fonts", () => {
  it("builds CSS2 URLs", () => {
    expect(googleCssUrl([{ family: "Klee One", v: "" }])).toBe("https://fonts.googleapis.com/css2?family=Klee+One&display=swap");
    expect(googleCssUrl([{ family: "Arimo", v: "bi" }])).toBe("https://fonts.googleapis.com/css2?family=Arimo:ital,wght@0,400;0,700;1,400;1,700&display=swap");
    expect(googleCssUrl([{ family: "A", v: "b" }, { family: "B", v: "i" }], "AB")).toBe("https://fonts.googleapis.com/css2?family=A:wght@400;700&family=B:ital@0;1&text=AB");
  });
  it("tries the full axis set, then the declared one, then regular", () => {
    expect(axisFallbacks("b")).toEqual(["bi", "b", ""]);
    expect(axisFallbacks("")).toEqual(["bi", ""]);
  });
  it("falls back on a refused axis set and injects the CSS", async () => {
    const urls: string[] = [];
    const ok = await loadWebFont("Yuji Mai", async (u) => {
      urls.push(u);
      return { ok: !u.includes("ital"), text: async () => "@font-face{font-family:'Yuji Mai';}" };
    });
    expect(ok).toBe(true);
    expect(urls).toHaveLength(2);
    expect(document.head.querySelector('style[data-rdoc-font="Yuji Mai"]')).not.toBeNull();
  });
  it("gives up offline and refuses families outside the catalog", async () => {
    expect(await loadWebFont("Yuji Syuku", async () => { throw new TypeError("offline"); })).toBe(false);
    expect(await loadWebFont("Not In Catalog", async () => ({ ok: true, text: async () => "" }))).toBe(false);
  });
});

describe("remembered font choices", () => {
  // An in-memory Storage: the runtime's own localStorage may be absent or partial under test.
  const mem = new Map<string, string>();
  const storage = { getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v) };
  beforeEach(() => { mem.clear(); vi.stubGlobal("localStorage", storage); });
  afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });
  it("round-trips and merges", () => {
    writeFontPrefs({ local: true });
    writeFontPrefs({ added: ["Klee One"] });
    expect(readFontPrefs()).toEqual({ local: true, added: ["Klee One"] });
  });
  it("survives blocked storage", () => {
    vi.stubGlobal("localStorage", { getItem: () => { throw new Error("blocked"); }, setItem: () => { throw new Error("blocked"); } });
    expect(readFontPrefs()).toEqual({});
    expect(() => writeFontPrefs({ consent: true })).not.toThrow();
  });
});

describe("font picker", () => {
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
    expect(wrap.querySelector(".docxedit-fonts-consent")).not.toBeNull(); // consent comes first
  });
  it("ignores an action reached by arrowing through a closed select", () => {
    const { p, wrap } = mount({ remote: { provider: "google" } });
    p.select.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown" }));
    p.select.value = p.select.options[p.select.options.length - 1]!.value;
    p.select.dispatchEvent(new Event("change"));
    expect(wrap.querySelector(".docxedit-fonts-consent")).toBeNull();
  });
});
