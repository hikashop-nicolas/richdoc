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
