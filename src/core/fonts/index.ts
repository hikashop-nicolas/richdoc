// Font availability: which families the browser can draw (canvas detection or the Local Font
// Access API), the remembered choices, and loading families from Google Fonts.
import { GROUP_SAMPLE, WEB_FONTS, type WebFont } from "./catalog";

export { BASE_FONTS, LOCAL_CANDIDATES, WEB_FONTS, GROUP_SAMPLE } from "./catalog";
export type { WebFont, WebFontGroup } from "./catalog";

/** Metrics of `text` drawn in a CSS `font`, or null when they cannot be measured. */
export type Measure = (font: string, text: string) => readonly number[] | null;

// Latin plus CJK/Hangul so a font covering only one of them still changes the result.
export const PROBE_TEXT = "mmmmmmmmmmlli WwQq@#&0123 永あア한中";
const FALLBACKS = ["monospace", "serif", "sans-serif"] as const;
const SIZE = "72px";
const GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "math", "emoji", "fangsong", "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded"]);

export const cssFamily = (name: string): string => `"${name.replace(/["\\]/g, "")}"`;
const same = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => Math.abs(v - b[i]!) < 0.01);

/** A name -> present map; one per measure function. */
export type DetectCache = Map<string, boolean>;

// Baseline metrics of each bare generic, per measure function.
const baselines = new WeakMap<Measure, Map<string, readonly number[] | null>>();
const baseline = (measure: Measure, g: string): readonly number[] | null => {
  let m = baselines.get(measure);
  if (!m) baselines.set(measure, (m = new Map()));
  if (!m.has(g)) m.set(g, measure(`${SIZE} ${g}`, PROBE_TEXT));
  return m.get(g)!;
};

/** True when `name` draws differently from a bare generic family for any of the fallbacks. */
export function fontPresent(name: string, measure: Measure, cache: DetectCache = new Map()): boolean {
  const key = name.toLowerCase();
  const hit = cache.get(key);
  if (hit !== undefined) return hit;
  let found = false;
  if (!GENERIC.has(key)) {
    for (const g of FALLBACKS) {
      const base = baseline(measure, g);
      const m = measure(`${SIZE} ${cssFamily(name)}, ${g}`, PROBE_TEXT);
      if (base && m && !same(base, m)) { found = true; break; }
    }
  }
  cache.set(key, found);
  return found;
}

/** The present names, in their input order. */
export function detectFonts(names: readonly string[], measure: Measure, cache?: DetectCache): string[] {
  return names.filter((n) => fontPresent(n, measure, cache));
}

/** detectFonts, yielding to the page whenever a slice runs past `budgetMs`, so input stays live. */
export async function detectFontsAsync(names: readonly string[], measure: Measure, cache?: DetectCache, budgetMs = 12): Promise<string[]> {
  const out: string[] = [];
  let start = performance.now();
  for (const n of names) {
    if (fontPresent(n, measure, cache)) out.push(n);
    if (performance.now() - start > budgetMs) {
      await idle();
      start = performance.now();
    }
  }
  return out;
}

// Yield one task (input and rendering run in between); idle callbacks starve on a busy page.
const idle = (): Promise<void> => {
  const sched = (globalThis as { scheduler?: { yield?: () => Promise<void> } }).scheduler;
  return typeof sched?.yield === "function" ? sched.yield() : new Promise((res) => setTimeout(res, 0));
};

let canvasMeasureFn: Measure | null | undefined;
/** A canvas measureText probe (null where there is no real canvas, e.g. jsdom). */
export function canvasMeasure(): Measure | null {
  if (canvasMeasureFn !== undefined) return canvasMeasureFn;
  canvasMeasureFn = null;
  if (typeof document === "undefined" || /jsdom/i.test(navigator.userAgent)) return null;
  let ctx: CanvasRenderingContext2D | null = null;
  try { ctx = document.createElement("canvas").getContext("2d"); } catch { ctx = null; }
  if (!ctx) return null;
  const c = ctx;
  canvasMeasureFn = (font, text) => {
    c.font = "10px serif";
    c.font = font;
    if (c.font === "10px serif") return null; // the font string was rejected
    const m = c.measureText(text);
    return [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent, m.actualBoundingBoxLeft, m.actualBoundingBoxRight];
  };
  return canvasMeasureFn;
}

// Session-wide detection cache shared by every editor.
const sessionCache: DetectCache = new Map();

/** Fonts defined by @font-face in this page (embedded document fonts, loaded web fonts). */
export function faceFamilies(): Set<string> {
  const out = new Set<string>();
  const fs = typeof document !== "undefined" ? (document.fonts as unknown as Iterable<FontFace> | undefined) : undefined;
  if (!fs || typeof (fs as { [Symbol.iterator]?: unknown })[Symbol.iterator] !== "function") return out;
  for (const f of fs) out.add(f.family.replace(/^["']|["']$/g, "").toLowerCase());
  return out;
}

/** Whether the browser can draw `name`: a page font face or an installed one. Unknown = true. */
export function canDraw(name: string, measure: Measure | null = canvasMeasure()): boolean {
  if (faceFamilies().has(name.toLowerCase())) return true;
  if (!measure) return true; // cannot tell: do not hide anything
  return fontPresent(name, measure, sessionCache);
}

/** Probe the built-in candidate list (cached for the session). */
let localDetect: Promise<string[]> | null = null;
export function detectInstalled(candidates: readonly string[], measure: Measure | null = canvasMeasure()): Promise<string[]> {
  if (!measure) return Promise.resolve([]);
  if (!localDetect) localDetect = detectFontsAsync([...new Set(candidates)], measure, sessionCache);
  return localDetect;
}

interface LocalFontData { family: string }
type QueryLocalFonts = () => Promise<LocalFontData[]>;

/** The Local Font Access API is present (Chromium desktop). */
export const hasLocalFontAccess = (): boolean =>
  typeof window !== "undefined" && typeof (window as unknown as { queryLocalFonts?: unknown }).queryLocalFonts === "function";

/** "granted" | "prompt" | "denied", or null when it cannot be asked. Never prompts. */
export async function localFontsPermission(): Promise<PermissionState | null> {
  try {
    const s = await navigator.permissions.query({ name: "local-fonts" as PermissionName });
    return s.state;
  } catch {
    return null;
  }
}

/** Installed family names via queryLocalFonts (prompts unless already granted); null when refused. */
export async function queryLocalFamilies(): Promise<string[] | null> {
  if (!hasLocalFontAccess()) return null;
  try {
    const list = await (window as unknown as { queryLocalFonts: QueryLocalFonts }).queryLocalFonts();
    return mergeNames(list.map((f) => f.family).filter((f) => f && !f.startsWith(".")));
  } catch {
    return null;
  }
}

/** De-duplicated (case-insensitive, first spelling kept) and sorted. */
export function mergeNames(...lists: readonly (readonly string[])[]): string[] {
  const seen = new Map<string, string>();
  for (const l of lists) for (const n of l) {
    const name = n.trim();
    if (name && !seen.has(name.toLowerCase())) seen.set(name.toLowerCase(), name);
  }
  return [...seen.values()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

export interface FontGroups {
  /** Fonts the document uses, with whether the browser can draw them. */
  doc: { name: string; missing: boolean }[];
  /** Every other usable font. */
  fonts: string[];
}

/** Split the picker into the document's fonts and the rest; base fonts are kept only if drawable. */
export function buildFontGroups(input: {
  docFonts: readonly string[];
  base: readonly string[];
  extra: readonly string[]; // already known to be usable (local, added web fonts)
  canDraw: (name: string) => boolean;
}): FontGroups {
  const doc = mergeNames(input.docFonts).map((name) => ({ name, missing: !input.canDraw(name) }));
  const inDoc = new Set(doc.map((d) => d.name.toLowerCase()));
  const fonts = mergeNames(input.base.filter((f) => input.canDraw(f)), input.extra).filter((f) => !inDoc.has(f.toLowerCase()));
  return { doc, fonts };
}

// --- Remembered choices ---------------------------------------------------------------------

export interface FontPrefs {
  local?: boolean; // the user asked for their installed fonts
  consent?: boolean; // the user agreed to download web fonts
  added?: string[]; // web font families they added
}
const PREFS_KEY = "richdoc.fonts";

export function readFontPrefs(): FontPrefs {
  try {
    const raw = globalThis.localStorage?.getItem(PREFS_KEY);
    const p = raw ? (JSON.parse(raw) as FontPrefs) : {};
    return p && typeof p === "object" ? p : {};
  } catch {
    return {};
  }
}

export function writeFontPrefs(patch: FontPrefs): FontPrefs {
  const next = { ...readFontPrefs(), ...patch };
  try { globalThis.localStorage?.setItem(PREFS_KEY, JSON.stringify(next)); } catch { /* storage blocked */ }
  return next;
}

// --- Google Fonts ---------------------------------------------------------------------------

const GOOGLE_CSS = "https://fonts.googleapis.com/css2";
const AXES: Record<WebFont["v"], string> = {
  bi: ":ital,wght@0,400;0,700;1,400;1,700",
  b: ":wght@400;700",
  i: ":ital@0;1",
  "": "",
};
const familyParam = (f: string, v: WebFont["v"]): string => `family=${encodeURIComponent(f).replace(/%20/g, "+")}${AXES[v]}`;

/** The CSS2 API URL for one or more families. */
export function googleCssUrl(families: readonly { family: string; v: WebFont["v"] }[], text?: string): string {
  const q = families.map((f) => familyParam(f.family, f.v));
  if (text) q.push(`text=${encodeURIComponent(text)}`);
  else q.push("display=swap");
  return `${GOOGLE_CSS}?${q.join("&")}`;
}

/** The axis sets to try, richest first (a wrong guess makes the API answer 400). */
export function axisFallbacks(v: WebFont["v"]): WebFont["v"][] {
  return [...new Set<WebFont["v"]>(["bi", v, ""])];
}

export const findWebFont = (family: string): WebFont | undefined =>
  WEB_FONTS.find((w) => w.family.toLowerCase() === family.toLowerCase());

export type Fetcher = (url: string) => Promise<{ ok: boolean; text(): Promise<string> }>;
const defaultFetch: Fetcher = (url) => fetch(url, { mode: "cors", credentials: "omit" });

const injected = new Map<string, Promise<boolean>>();
const injectCss = (css: string, tag: string): void => {
  const s = document.createElement("style");
  s.dataset.rdocFont = tag;
  s.textContent = css;
  document.head.appendChild(s);
};

/** Load a catalog family (its @font-face rules) from Google Fonts. Resolves false on failure. */
export function loadWebFont(family: string, fetcher: Fetcher = defaultFetch): Promise<boolean> {
  const wf = findWebFont(family);
  if (!wf) return Promise.resolve(false);
  const key = wf.family.toLowerCase();
  const prev = injected.get(key);
  if (prev) return prev;
  const p = (async () => {
    for (const v of axisFallbacks(wf.v)) {
      try {
        const res = await fetcher(googleCssUrl([{ family: wf.family, v }]));
        if (!res.ok) continue;
        injectCss(await res.text(), wf.family);
        try { await document.fonts?.load(`16px ${cssFamily(wf.family)}`, GROUP_SAMPLE[wf.group] ?? wf.family); } catch { /* still usable */ }
        return true;
      } catch {
        return false; // offline or blocked: no point trying other axes
      }
    }
    return false;
  })();
  injected.set(key, p);
  p.then((ok) => { if (!ok) injected.delete(key); });
  return p;
}

export const previewFamily = (family: string): string => `rdoc-preview ${family}`;

/** Load tiny name-only subsets for previews, renamed so they never stand in for the real font. */
const previewed = new Set<string>();
export async function loadWebFontPreviews(fonts: readonly WebFont[], fetcher: Fetcher = defaultFetch, batch = 12): Promise<void> {
  const todo = fonts.filter((f) => !previewed.has(f.family));
  for (let i = 0; i < todo.length; i += batch) {
    const part = todo.slice(i, i + batch);
    const text = [...new Set(part.map((f) => f.family + (GROUP_SAMPLE[f.group] ?? "")).join(""))].join("");
    try {
      const res = await fetcher(googleCssUrl(part.map((f) => ({ family: f.family, v: "" })), text));
      if (!res.ok) continue;
      const css = (await res.text()).replace(/font-family:\s*'([^']+)'/g, (_, f: string) => `font-family: '${previewFamily(f)}'`);
      injectCss(css, "preview");
      for (const f of part) previewed.add(f.family);
    } catch {
      return; // offline
    }
  }
}
