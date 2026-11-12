// docx font resolution shared by the reader and writer: theme fonts (theme1.xml), w:rFonts
// with its *Theme attributes, and the document defaults (docDefaults + the default paragraph style).
import { strFromU8 } from "fflate";
import { W } from "./shared";

/** The theme's major (headings) and minor (body) Latin and East Asian typefaces. */
export interface ThemeFonts {
  major?: string;
  minor?: string;
  majorEa?: string;
  minorEa?: string;
}

/** The document's default run font and size, as Word resolves them for an unstyled run. */
export interface DocDefaults {
  font?: string;
  eastAsia?: string;
  sizeHalfPt?: number;
}

const wv = (el: Element | undefined | null, name: string): string | null => el?.getAttributeNS(W, name) ?? el?.getAttribute("w:" + name) ?? null;
const parse = (bytes: Uint8Array): Document => new DOMParser().parseFromString(strFromU8(bytes), "application/xml");

/** The main document's theme part key, via its relationship or Word's usual path. */
function themeKey(files: Record<string, Uint8Array>): string | undefined {
  const rels = files["word/_rels/document.xml.rels"];
  if (rels) {
    for (const r of Array.from(parse(rels).getElementsByTagName("Relationship"))) {
      if (!(r.getAttribute("Type") ?? "").endsWith("/theme")) continue;
      const tgt = (r.getAttribute("Target") ?? "").replace(/^\.\//, "");
      const key = tgt.startsWith("/") ? tgt.slice(1) : "word/" + tgt;
      if (files[key]) return key;
    }
  }
  return files["word/theme/theme1.xml"] ? "word/theme/theme1.xml" : undefined;
}

export function readThemeFonts(files: Record<string, Uint8Array>): ThemeFonts {
  const key = themeKey(files);
  if (!key) return {};
  const doc = parse(files[key]!);
  const face = (scheme: string, tag: string): string | undefined => {
    const s = doc.getElementsByTagName(`a:${scheme}`)[0];
    const el = s ? Array.from(s.children).find((c) => c.tagName === tag) : undefined;
    return el?.getAttribute("typeface") || undefined;
  };
  return { major: face("majorFont", "a:latin"), minor: face("minorFont", "a:latin"), majorEa: face("majorFont", "a:ea"), minorEa: face("minorFont", "a:ea") };
}

const themeFace = (theme: ThemeFonts, ref: string | null, ea = false): string | undefined => {
  if (!ref) return undefined;
  const major = /^major/i.test(ref);
  return ea ? (major ? theme.majorEa : theme.minorEa) : major ? theme.major : theme.minor;
};

/** A w:rFonts element's Latin typeface: a theme reference wins over the literal name, as in Word. */
export function rFontsLatin(rf: Element | undefined | null, theme: ThemeFonts): string | undefined {
  if (!rf) return undefined;
  return themeFace(theme, wv(rf, "asciiTheme")) || wv(rf, "ascii") || themeFace(theme, wv(rf, "hAnsiTheme")) || wv(rf, "hAnsi") || undefined;
}

/** A w:rFonts element's East Asian typeface. */
export function rFontsEastAsia(rf: Element | undefined | null, theme: ThemeFonts): string | undefined {
  if (!rf) return undefined;
  return themeFace(theme, wv(rf, "eastAsiaTheme"), true) || wv(rf, "eastAsia") || undefined;
}

const childEl = (el: Element | undefined | null, tag: string): Element | undefined => (el ? Array.from(el.children).find((c) => c.tagName === tag) : undefined);

/** Resolve the document default font/size: docDefaults, then the default paragraph style
    (Word's "Normal") with its basedOn chain, theme references resolved. */
export function readDocDefaults(files: Record<string, Uint8Array>, theme: ThemeFonts = readThemeFonts(files)): DocDefaults {
  const sx = files["word/styles.xml"];
  if (!sx) return {};
  const doc = parse(sx);
  const out: DocDefaults = {};
  const apply = (rPr: Element | undefined) => {
    if (!rPr) return;
    const rf = childEl(rPr, "w:rFonts");
    const latin = rFontsLatin(rf, theme);
    if (latin) out.font = latin;
    const ea = rFontsEastAsia(rf, theme);
    if (ea) out.eastAsia = ea;
    const sz = wv(childEl(rPr, "w:sz"), "val");
    if (sz && /^\d+$/.test(sz) && Number(sz) > 0) out.sizeHalfPt = Number(sz);
  };
  apply(childEl(childEl(doc.getElementsByTagName("w:docDefaults")[0], "w:rPrDefault"), "w:rPr"));
  const paras = new Map<string, Element>();
  let def: Element | undefined;
  for (const st of Array.from(doc.getElementsByTagName("w:style"))) {
    if ((wv(st, "type") ?? "paragraph") !== "paragraph") continue;
    const id = wv(st, "styleId");
    if (id) paras.set(id, st);
    if (!def && (wv(st, "default") === "1" || wv(st, "default") === "true")) def = st;
  }
  const chain: Element[] = [];
  for (let st = def, seen = new Set<Element>(); st && !seen.has(st); st = paras.get(wv(st, "basedOn") ?? "")) {
    seen.add(st);
    chain.unshift(st);
  }
  for (const st of chain) apply(childEl(st, "w:rPr"));
  return out;
}

/** The theme's colour scheme (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink) as RRGGBB. */
export function readThemeColors(files: Record<string, Uint8Array>): Record<string, string> {
  const key = themeKey(files);
  const out: Record<string, string> = {};
  if (!key) return out;
  const scheme = parse(files[key]!).getElementsByTagName("a:clrScheme")[0];
  for (const c of Array.from(scheme?.children ?? [])) {
    const name = c.tagName.replace(/^a:/, "");
    const v = c.getElementsByTagName("a:srgbClr")[0]?.getAttribute("val") ?? c.getElementsByTagName("a:sysClr")[0]?.getAttribute("lastClr");
    if (v && /^[0-9a-f]{6}$/i.test(v)) out[name] = v.toUpperCase();
  }
  return out;
}

// w:themeColor / a:schemeClr names that alias the scheme slots.
const SCHEME_ALIAS: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2", text1: "dk1", background1: "lt1", text2: "dk2", background2: "lt2", dark1: "dk1", light1: "lt1", dark2: "dk2", light2: "lt2", hyperlink: "hlink", followedHyperlink: "folHlink" };
export const themeColor = (colors: Record<string, string>, name: string | null | undefined): string | undefined => (name ? colors[SCHEME_ALIAS[name] ?? name] : undefined);

const hexRgb = (h: string): [number, number, number] => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)) as [number, number, number];
const rgbHex = (c: number[]): string => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
function rgbToHsl([r, g, b]: number[]): [number, number, number] {
  const [R, G, B] = [r! / 255, g! / 255, b! / 255];
  const max = Math.max(R, G, B), min = Math.min(R, G, B), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === R ? (G - B) / d + (G < B ? 6 : 0) : max === G ? (B - R) / d + 2 : (R - G) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb([h, s, l]: number[]): number[] {
  if (!s) return [l! * 255, l! * 255, l! * 255];
  const q = l! < 0.5 ? l! * (1 + s) : l! + s - l! * s, p = 2 * l! - q;
  const f = (t: number) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return [f(h! + 1 / 3) * 255, f(h!) * 255, f(h! - 1 / 3) * 255];
}
/** Apply DrawingML colour modifiers (lumMod / lumOff / shade / tint, in 1/1000 %) to RRGGBB. */
export function modColor(hex: string, mods: { lumMod?: number; lumOff?: number; shade?: number; tint?: number }): string {
  let rgb: number[] = hexRgb(hex);
  if (mods.shade !== undefined) rgb = rgb.map((v) => v * (mods.shade! / 100000));
  if (mods.tint !== undefined) rgb = rgb.map((v) => v + (255 - v) * (1 - mods.tint! / 100000));
  if (mods.lumMod !== undefined || mods.lumOff !== undefined) {
    const [h, s, l] = rgbToHsl(rgb);
    rgb = hslToRgb([h, s, Math.min(1, Math.max(0, l * ((mods.lumMod ?? 100000) / 100000) + (mods.lumOff ?? 0) / 100000))]);
  }
  return rgbHex(rgb);
}
/** A DrawingML colour element's parent (a:solidFill, a:fillRef, ...) resolved to RRGGBB. */
export function drawingColor(parent: Element | undefined, colors: Record<string, string>): string | undefined {
  const c = Array.from(parent?.children ?? []).find((e) => /^a:(srgbClr|schemeClr|sysClr|prstClr)$/.test(e.tagName));
  if (!c) return undefined;
  const base = c.tagName === "a:srgbClr" ? c.getAttribute("val") : c.tagName === "a:sysClr" ? c.getAttribute("lastClr") : c.tagName === "a:schemeClr" ? themeColor(colors, c.getAttribute("val")) : c.getAttribute("val") === "black" ? "000000" : c.getAttribute("val") === "white" ? "FFFFFF" : undefined;
  if (!base || !/^[0-9a-f]{6}$/i.test(base)) return undefined;
  const num = (tag: string) => { const v = Array.from(c.children).find((e) => e.tagName === tag)?.getAttribute("val"); return v ? Number(v) : undefined; };
  return modColor(base.toUpperCase(), { lumMod: num("a:lumMod"), lumOff: num("a:lumOff"), shade: num("a:shade"), tint: num("a:tint") });
}
