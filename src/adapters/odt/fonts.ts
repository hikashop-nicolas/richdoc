// odt fonts: font-face declarations (style:font-name -> real family), embedded font files
// (svg:font-face-uri into Fonts/), the per-script (Western vs East Asian) run fonts and sizes,
// and the document defaults. LibreOffice picks a run's font and size per character script, so
// East Asian text uses style:font-name-asian / style:font-size-asian, not the fo:* ones.
import { strFromU8 } from "fflate";

export interface FontDecl {
  family: string; // svg:font-family, unquoted
  generic?: string; // CSS generic family from style:font-family-generic
  srcs: { href: string; weight: string; style: string }[];
}
export type FontFaces = Map<string, FontDecl>;

const GENERIC: Record<string, string> = { roman: "serif", swiss: "sans-serif", modern: "monospace", script: "cursive", decorative: "fantasy" };
const unquote = (v: string): string => v.trim().replace(/^['"]|['"]$/g, "").trim();

/** Index every style:font-face (styles.xml and content.xml) by its style:name. */
export function collectFontFaces(...docs: (Document | undefined)[]): FontFaces {
  const map: FontFaces = new Map();
  for (const doc of docs) {
    if (!doc) continue;
    for (const ff of Array.from(doc.getElementsByTagName("style:font-face"))) {
      const name = ff.getAttribute("style:name");
      if (!name) continue;
      const family = unquote((ff.getAttribute("svg:font-family") ?? name).split(",")[0] ?? name) || name;
      const srcs = Array.from(ff.getElementsByTagName("svg:font-face-uri")).map((u) => ({
        href: (u.getAttribute("xlink:href") ?? "").replace(/^\.\//, ""),
        weight: u.getAttribute("loext:font-weight") === "bold" ? "bold" : "normal",
        style: u.getAttribute("loext:font-style") === "italic" ? "italic" : "normal",
      })).filter((s) => s.href);
      const prev = map.get(name);
      // The same face is often declared in both parts; keep whichever carries the font files.
      if (!prev || (!prev.srcs.length && srcs.length)) map.set(name, { family, generic: GENERIC[ff.getAttribute("style:font-family-generic") ?? ""], srcs });
    }
  }
  return map;
}

/** @font-face rules (blob URLs) for the fonts embedded in the package. */
export function loadOdtEmbeddedFonts(files: Record<string, Uint8Array>, faces: FontFaces): { css: string; urls: string[] } {
  const out = { css: "", urls: [] as string[] };
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return out;
  const seen = new Set<string>();
  for (const f of faces.values()) {
    for (const s of f.srcs) {
      const key = `${f.family}|${s.href}`;
      const bytes = files[s.href];
      if (!bytes || seen.has(key)) continue;
      seen.add(key);
      const ext = (s.href.split(".").pop() ?? "").toLowerCase();
      const mime = ext === "otf" ? "font/otf" : ext === "woff2" ? "font/woff2" : ext === "woff" ? "font/woff" : "font/ttf";
      const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
      out.urls.push(url);
      out.css += `@font-face{font-family:"${f.family.replace(/"/g, "")}";font-weight:${s.weight};font-style:${s.style};font-display:swap;src:url(${url});}\n`;
    }
  }
  return out;
}

/** A style:font-name reference resolved to its declared family (falls back to the name). */
export const faceFamily = (faces: FontFaces, name: string | null | undefined): string | undefined =>
  name ? faces.get(name)?.family ?? name : undefined;

/** The Western run font of a style:text-properties: fo:font-family wins over style:font-name. */
export function westernFont(tp: Element | undefined | null, faces: FontFaces): string | undefined {
  if (!tp) return undefined;
  const fam = tp.getAttribute("fo:font-family");
  if (fam) return unquote(fam.split(",")[0] ?? "") || undefined;
  return faceFamily(faces, tp.getAttribute("style:font-name"));
}

/** The East Asian run font: style:font-family-asian wins over style:font-name-asian. */
export function asianFont(tp: Element | undefined | null, faces: FontFaces): string | undefined {
  if (!tp) return undefined;
  const fam = tp.getAttribute("style:font-family-asian");
  if (fam) return unquote(fam.split(",")[0] ?? "") || undefined;
  return faceFamily(faces, tp.getAttribute("style:font-name-asian"));
}

const pt = (v: string | null | undefined): number | undefined => {
  const m = v ? /^([\d.]+)pt$/.exec(v.trim()) : null;
  return m ? Math.round(parseFloat(m[1]!) * 100) / 100 : undefined;
};
export const asianSizePt = (tp: Element | undefined | null): number | undefined => pt(tp?.getAttribute("style:font-size-asian"));

/** CSS value for --rdoc-ea-font: the face, then the document font as the fallback (an East
    Asian font absent from this machine should fall back like LibreOffice, not to the UA serif). */
export const eaFontCss = (family: string): string => `'${family.replace(/'/g, "")}', var(--docxedit-doc-font, serif)`;

/** CSS declarations for a style:text-properties element: Western font/size/weight/colour plus
    the East Asian font and size as custom properties the East Asian script spans read. */
export function textPropsCss(tp: Element | undefined | null, faces: FontFaces): string[] {
  if (!tp) return [];
  const css: string[] = [];
  const wf = westernFont(tp, faces);
  if (wf) css.push(`font-family:'${wf.replace(/'/g, "")}'`);
  const sz = pt(tp.getAttribute("fo:font-size"));
  if (sz) css.push(`font-size:${sz}pt`);
  const fw = tp.getAttribute("fo:font-weight");
  if (fw) css.push(`font-weight:${/bold|[6-9]00/.test(fw) ? "bold" : "normal"}`);
  const fs = tp.getAttribute("fo:font-style");
  if (fs) css.push(`font-style:${fs === "italic" ? "italic" : "normal"}`);
  const color = tp.getAttribute("fo:color");
  if (color && /^#[0-9a-f]{6}$/i.test(color)) css.push(`color:${color}`);
  const ef = asianFont(tp, faces);
  if (ef) css.push(`--rdoc-ea-font:${eaFontCss(ef)}`);
  const es = asianSizePt(tp);
  if (es) css.push(`--rdoc-ea-size:${es}pt`);
  return css;
}

export interface OdtDefaults {
  font?: string;
  eaFont?: string;
  sizePt?: number;
  eaSizePt?: number;
}

/** The default run font and size: the paragraph default-style, then the "Standard" style. */
export function readOdtDefaults(stylesXml: Uint8Array | undefined, faces: FontFaces): OdtDefaults {
  const out: OdtDefaults = {};
  if (!stylesXml) return out;
  const doc = new DOMParser().parseFromString(strFromU8(stylesXml), "application/xml");
  const apply = (st: Element | undefined) => {
    const tp = st ? Array.from(st.children).find((c) => c.tagName === "style:text-properties") : undefined;
    if (!tp) return;
    out.font = westernFont(tp, faces) ?? out.font;
    out.eaFont = asianFont(tp, faces) ?? out.eaFont;
    out.sizePt = pt(tp.getAttribute("fo:font-size")) ?? out.sizePt;
    out.eaSizePt = asianSizePt(tp) ?? out.eaSizePt;
  };
  apply(Array.from(doc.getElementsByTagName("style:default-style")).find((s) => s.getAttribute("style:family") === "paragraph"));
  apply(Array.from(doc.getElementsByTagName("style:style")).find((s) => s.getAttribute("style:name") === "Standard" && s.getAttribute("style:family") === "paragraph"));
  return out;
}

/** Base CSS for an odt page: ODF paragraphs have no default spacing and single line spacing
    follows the font's own metrics; East Asian script spans take the East Asian font/size. */
export function odtBaseCss(d: OdtDefaults): string {
  const root: string[] = [];
  if (d.eaFont) root.push(`--rdoc-ea-font:${eaFontCss(d.eaFont)}`);
  if (d.eaSizePt) root.push(`--rdoc-ea-size:${d.eaSizePt}pt`);
  return (
    (root.length ? `.docxedit-page{${root.join(";")}}\n` : "") +
    ".docxedit-page .docxedit-doc,.docxedit-page .docxedit-header,.docxedit-page .docxedit-footer{line-height:normal}\n" +
    ".docxedit-doc p,.docxedit-header p,.docxedit-footer p{margin:0}\n" +
    `[${EA_ATTR}]{font-family:var(--rdoc-ea-font,inherit);font-size:var(--rdoc-ea-size,inherit)}\n`
  );
}

// Character script, as LibreOffice classifies it: East Asian blocks, other letters (Western;
// complex scripts are folded in), and weak characters (digits, punctuation, symbols, spaces)
// that take the script of the text before them.
export const EA_ATTR = "data-rdoc-script";
type Script = "a" | "l" | "w";
function scriptOf(c: number): Script {
  if ((c >= 0x1100 && c <= 0x11ff) || (c >= 0x2e80 && c <= 0xa4cf) || (c >= 0xac00 && c <= 0xd7af) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0xfe30 && c <= 0xfe4f) || (c >= 0xff00 && c <= 0xffef) || (c >= 0x20000 && c <= 0x3ffff)) return "a";
  return /\p{L}/u.test(String.fromCodePoint(c)) ? "l" : "w";
}

/** The script a block starts in: its first strong character's (weak lead-ins follow it). */
export function leadingScript(text: string): "a" | "l" {
  for (const ch of text) {
    const s = scriptOf(ch.codePointAt(0)!);
    if (s !== "w") return s;
  }
  return "l";
}

/** Split text into script runs, carrying the running script across calls (one paragraph). */
export function splitScripts(text: string, state: { script: "a" | "l" }): { ea: boolean; text: string }[] {
  const out: { ea: boolean; text: string }[] = [];
  for (const ch of text) {
    const s = scriptOf(ch.codePointAt(0)!);
    if (s !== "w") state.script = s;
    const ea = state.script === "a";
    const last = out[out.length - 1];
    if (last && last.ea === ea) last.text += ch;
    else out.push({ ea, text: ch });
  }
  return out;
}
