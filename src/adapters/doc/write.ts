import { writeCfb } from "./cfb";
import { FC } from "./fib";
import { B64, b64 } from "./templates";

// Write half of the .doc adapter: richdoc's edited HTML -> a from-scratch Word 97-2003
// binary. We reuse a known-good FIB and the content-independent stylesheet / font table /
// property-set streams captured from a real `textutil` file, and build the content-
// dependent structures ourselves: the text stream, the piece table (Clx), the CHPX/PAPX
// formatted-disk-page runs, their PlcfBte locators, and the section table. Validated by
// having macOS `textutil` (Apple's real Word engine) read the output back.

const TEXT_START = 1536; // fc where the FIB template ends and the text begins

interface Run {
  text: string;
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  sizeHalf?: number; // font size in half-points
  color?: number; // packed RR | GG<<8 | BB<<16
  font?: string; // font family name
  highlight?: number; // highlight palette index (1..16)
}
interface Para {
  align: number; // 0 left, 1 center, 2 right, 3 justify
  runs: Run[];
}

// ---------------------------------------------------------------------------
// HTML -> paragraph/run model
// ---------------------------------------------------------------------------

interface Fmt {
  b: boolean;
  i: boolean;
  u: boolean;
  strike: boolean;
  sizeHalf?: number;
  color?: number;
  font?: string;
  highlight?: number;
}

// Word's 16-entry highlight palette (index -> #rrggbb), for nearest-colour matching.
const HIGHLIGHT_PALETTE: [number, number][] = [
  [1, 0x000000], [2, 0x0000ff], [3, 0x00ffff], [4, 0x00ff00], [5, 0xff00ff], [6, 0xff0000],
  [7, 0xffff00], [8, 0xffffff], [9, 0x000080], [10, 0x008080], [11, 0x008000], [12, 0x800080],
  [13, 0x800000], [14, 0x808000], [15, 0x808080], [16, 0xc0c0c0],
];
function nearestHighlight(rgb: number): number {
  const r = (rgb >> 16) & 0xff;
  const g = (rgb >> 8) & 0xff;
  const b = rgb & 0xff;
  let best = 7;
  let bestD = Infinity;
  for (const [idx, c] of HIGHLIGHT_PALETTE) {
    const d = ((c >> 16) & 0xff) - r;
    const d2 = ((c >> 8) & 0xff) - g;
    const d3 = (c & 0xff) - b;
    const dist = d * d + d2 * d2 + d3 * d3;
    if (dist < bestD) {
      bestD = dist;
      best = idx;
    }
  }
  return best;
}

function parseColor(v: string): number | undefined {
  const m = v.trim().match(/^#?([0-9a-f]{6})$/i);
  if (m) {
    const n = parseInt(m[1], 16);
    return ((n >> 16) & 0xff) | (n & 0xff00) | ((n & 0xff) << 16); // -> RR,GG,BB order below
  }
  const rgb = v.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i);
  if (rgb) return Number(rgb[1]) | (Number(rgb[2]) << 8) | (Number(rgb[3]) << 16);
  return undefined;
}

function applyInlineStyle(f: Fmt, el: Element): Fmt {
  const out = { ...f };
  const tag = el.tagName.toLowerCase();
  if (tag === "b" || tag === "strong") out.b = true;
  if (tag === "i" || tag === "em") out.i = true;
  if (tag === "u") out.u = true;
  if (tag === "s" || tag === "strike" || tag === "del") out.strike = true;
  const style = el.getAttribute("style") || "";
  const fw = /font-weight:\s*(bold|[6-9]00)/i.test(style);
  if (fw) out.b = true;
  if (/font-style:\s*italic/i.test(style)) out.i = true;
  if (/text-decoration:[^;]*underline/i.test(style)) out.u = true;
  if (/text-decoration:[^;]*line-through/i.test(style)) out.strike = true;
  const size = style.match(/font-size:\s*([\d.]+)pt/i);
  if (size) out.sizeHalf = Math.round(parseFloat(size[1]) * 2);
  const col = style.match(/(?:^|[;\s])color:\s*([^;]+)/i);
  if (col) {
    const c = parseColor(col[1]);
    if (c !== undefined) out.color = c;
  }
  const fam = style.match(/font-family:\s*([^;]+)/i);
  if (fam) out.font = fam[1].trim().replace(/^['"]|['"]$/g, "").split(",")[0].trim();
  const bg = style.match(/(?:background-color|background):\s*([^;]+)/i);
  if (bg) {
    const c = parseColor(bg[1]);
    if (c !== undefined) out.highlight = nearestHighlight(((c & 0xff) << 16) | (c & 0xff00) | ((c >> 16) & 0xff));
  }
  if (tag === "mark") out.highlight = out.highlight ?? 7;
  return out;
}

function mkRun(text: string, f: Fmt): Run {
  return { text, b: f.b, i: f.i, u: f.u, strike: f.strike, sizeHalf: f.sizeHalf, color: f.color, font: f.font, highlight: f.highlight };
}

function collectRuns(node: Node, f: Fmt, runs: Run[]): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      const text = child.textContent ?? "";
      if (text) runs.push(mkRun(text, f));
    } else if (child.nodeType === 1) {
      const el = child as Element;
      const tag = el.tagName.toLowerCase();
      if (tag === "br") {
        runs.push({ text: "", b: f.b, i: f.i, u: f.u, strike: f.strike, sizeHalf: f.sizeHalf, color: f.color });
        continue;
      }
      if (tag === "a" && el.getAttribute("href")) {
        const href = el.getAttribute("href") || "";
        runs.push(mkRun("\x13", f));
        runs.push(mkRun(`HYPERLINK "${href.replace(/"/g, "%22")}" `, f));
        runs.push(mkRun("\x14", f));
        collectRuns(el, { ...f, u: true, color: 0xee0000 }, runs);
        runs.push(mkRun("\x15", f));
        continue;
      }
      collectRuns(el, applyInlineStyle(f, el), runs);
    }
  }
}

function blockAlign(el: Element): number {
  const style = el.getAttribute("style") || "";
  const m = style.match(/text-align:\s*(left|center|right|justify)/i);
  const a = (m?.[1] || el.getAttribute("align") || "").toLowerCase();
  return a === "center" ? 1 : a === "right" ? 2 : a === "justify" ? 3 : 0;
}

function parseHtml(bodyHtml: string): Para[] {
  const doc = new DOMParser().parseFromString(`<body>${bodyHtml}</body>`, "text/html");
  const paras: Para[] = [];
  const blockTags = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "div", "blockquote", "pre"]);
  const walk = (el: Element): void => {
    for (const child of Array.from(el.children)) {
      const tag = child.tagName.toLowerCase();
      if (tag === "ul" || tag === "ol" || tag === "table" || tag === "tbody" || tag === "tr") {
        walk(child);
      } else if (blockTags.has(tag)) {
        const base: Fmt = { b: /^h[1-6]$/.test(tag), i: false, u: false, strike: false };
        if (/^h[1-6]$/.test(tag)) base.sizeHalf = { h1: 48, h2: 36, h3: 28, h4: 24, h5: 22, h6: 20 }[tag] ?? 24;
        const runs: Run[] = [];
        collectRuns(child, base, runs);
        const prefix = tag === "li" ? [{ text: "\t•\t", b: false, i: false, u: false, strike: false } as Run] : [];
        paras.push({ align: blockAlign(child), runs: [...prefix, ...runs] });
      }
    }
  };
  walk(doc.body);
  if (paras.length === 0) paras.push({ align: 0, runs: [] });
  return paras;
}

// ---------------------------------------------------------------------------
// A growable byte buffer
// ---------------------------------------------------------------------------

class Buf {
  private a: number[] = [];
  u8(v: number): void {
    this.a.push(v & 0xff);
  }
  u16(v: number): void {
    this.a.push(v & 0xff, (v >>> 8) & 0xff);
  }
  u32(v: number): void {
    this.a.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  }
  bytes(b: Uint8Array): void {
    for (const x of b) this.a.push(x);
  }
  get length(): number {
    return this.a.length;
  }
  done(): Uint8Array {
    return new Uint8Array(this.a);
  }
}

// Character sprms (grpprl) for a run's formatting. fontFtc resolves a font name to its
// index in the font table (or -1 to omit).
function chpxGrpprl(r: Run, fontFtc: (name: string) => number): Uint8Array {
  const b = new Buf();
  if (r.b) {
    b.u16(0x0835);
    b.u8(1);
  }
  if (r.i) {
    b.u16(0x0836);
    b.u8(1);
  }
  if (r.u) {
    b.u16(0x2a3e);
    b.u8(1); // single underline
  }
  if (r.strike) {
    b.u16(0x0837);
    b.u8(1);
  }
  if (r.font) {
    const ftc = fontFtc(r.font);
    if (ftc >= 0) {
      b.u16(0x4a4f);
      b.u16(ftc);
      b.u16(0x4a51);
      b.u16(ftc);
    }
  }
  if (r.sizeHalf) {
    b.u16(0x4a43);
    b.u16(r.sizeHalf);
  }
  if (r.highlight) {
    b.u16(0x2a0c);
    b.u8(r.highlight);
  }
  if (r.color !== undefined) {
    b.u16(0x6870);
    b.u8(r.color & 0xff); // RR
    b.u8((r.color >>> 8) & 0xff); // GG
    b.u8((r.color >>> 16) & 0xff); // BB
    b.u8(0);
  }
  return b.done();
}

// Paragraph sprms for alignment.
function papxGrpprl(p: Para): Uint8Array {
  const b = new Buf();
  if (p.align) {
    b.u16(0x2403);
    b.u8(p.align);
  }
  return b.done();
}

// Build a Sttbfffn (font table) from font names, with a name->ftc resolver.
function buildFontTable(names: string[]): { bytes: Uint8Array; ftc: (name: string) => number } {
  const lower = names.map((n) => n.toLowerCase());
  const b = new Buf();
  b.u16(names.length);
  b.u16(0);
  for (const name of names) {
    b.u8(39 + (name.length + 1) * 2);
    b.u8(0);
    b.u16(400);
    b.u8(0);
    b.u8(0);
    for (let i = 0; i < 34; i++) b.u8(0);
    for (const ch of name) b.u16(ch.charCodeAt(0));
    b.u16(0);
  }
  return { bytes: b.done(), ftc: (name: string) => lower.indexOf(name.toLowerCase()) };
}

// ---------------------------------------------------------------------------
// FKP builders (512-byte pages in the WordDocument stream)
// ---------------------------------------------------------------------------

interface CharRun {
  fcStart: number;
  fcEnd: number;
  grpprl: Uint8Array;
}
interface ParaRun {
  fcStart: number;
  fcEnd: number;
  grpprl: Uint8Array;
}

function buildChpxFkp(runs: CharRun[]): Uint8Array {
  const page = new Uint8Array(512);
  const dv = new DataView(page.buffer);
  const crun = runs.length;
  for (let i = 0; i <= crun; i++) dv.setUint32(i * 4, i < crun ? runs[i].fcStart : runs[crun - 1].fcEnd, true);
  const offArr = 4 * (crun + 1);
  // Pack CHPX blobs downward from the end (before the crun byte at 511).
  let pos = 511;
  for (let i = crun - 1; i >= 0; i--) {
    const g = runs[i].grpprl;
    pos -= 1 + g.length;
    if (pos % 2 !== 0) pos -= 1; // CHPX must start at an even offset
    page[pos] = g.length;
    page.set(g, pos + 1);
    page[offArr + i] = pos / 2;
  }
  page[511] = crun;
  return page;
}

function buildPapxFkp(paras: ParaRun[]): Uint8Array {
  const page = new Uint8Array(512);
  const dv = new DataView(page.buffer);
  const cpara = paras.length;
  for (let i = 0; i <= cpara; i++) dv.setUint32(i * 4, i < cpara ? paras[i].fcStart : paras[cpara - 1].fcEnd, true);
  const bxArr = 4 * (cpara + 1); // aBxPap: 13 bytes each
  let pos = 511;
  for (let i = cpara - 1; i >= 0; i--) {
    // grpprl = istd (2 bytes, 0 = Normal) + sprms
    const g = paras[i].grpprl;
    const grpprl = new Uint8Array(2 + g.length);
    grpprl.set(g, 2); // istd stays 0
    // PapxInFkp: 1-byte cb where the stored grpprl length is 2*cb-1 (pad if even).
    const stored = grpprl.length % 2 === 0 ? new Uint8Array([...grpprl, 0]) : grpprl;
    const cb = (stored.length + 1) / 2;
    pos -= 1 + stored.length;
    if (pos % 2 !== 0) pos -= 1;
    page[pos] = cb;
    page.set(stored, pos + 1);
    page[bxArr + i * 13] = pos / 2;
  }
  page[511] = cpara;
  return page;
}

// ---------------------------------------------------------------------------
// Assemble
// ---------------------------------------------------------------------------

export function htmlToDoc(bodyHtml: string): Uint8Array {
  const paras = parseHtml(bodyHtml);

  // Collect fonts (base four match the template order) + any run fonts, build the table.
  const fontNames = ["Times New Roman", "Symbol", "Arial", "Times"];
  for (const p of paras) for (const r of p.runs) {
    if (r.font && !fontNames.some((n) => n.toLowerCase() === r.font!.toLowerCase())) fontNames.push(r.font);
  }
  const fontTable = buildFontTable(fontNames);

  // 1. Build the text stream (UTF-16LE) and record run/para fc boundaries.
  const chars: number[] = [];
  const charRuns: CharRun[] = [];
  const paraRuns: ParaRun[] = [];
  const fcAt = (cp: number) => TEXT_START + cp * 2;
  for (const p of paras) {
    const paraStartCp = chars.length;
    const runList = p.runs.length ? p.runs : [{ text: "", b: false, i: false, u: false, strike: false } as Run];
    for (const r of runList) {
      const startCp = chars.length;
      for (const ch of r.text) chars.push(ch.codePointAt(0)!);
      charRuns.push({ fcStart: fcAt(startCp), fcEnd: fcAt(chars.length), grpprl: chpxGrpprl(r, fontTable.ftc) });
    }
    chars.push(0x0d); // paragraph mark
    // The paragraph's CHPX must also cover its mark; extend the last run to include it.
    charRuns[charRuns.length - 1].fcEnd = fcAt(chars.length);
    paraRuns.push({ fcStart: fcAt(paraStartCp), fcEnd: fcAt(chars.length), grpprl: papxGrpprl(p) });
  }
  const ccpText = chars.length;
  const textBytes = new Uint8Array(ccpText * 2);
  {
    const dv = new DataView(textBytes.buffer);
    for (let i = 0; i < ccpText; i++) dv.setUint16(i * 2, chars[i] > 0xffff ? 0x3f : chars[i], true);
  }

  // 2. WordDocument = FIB template + text + CHPX FKP + PAPX FKP.
  const fib = b64(B64.FIB).slice(); // 1536 bytes
  const chpxFkp = buildChpxFkp(charRuns);
  const papxFkp = buildPapxFkp(paraRuns);
  const textEnd = TEXT_START + textBytes.length;
  const fkpBase = Math.ceil(textEnd / 512) * 512;
  const chpxPage = fkpBase / 512;
  const papxPage = chpxPage + 1;
  const wdLen = Math.max(4096, (papxPage + 1) * 512);
  const wd = new Uint8Array(wdLen);
  wd.set(fib, 0);
  wd.set(textBytes, TEXT_START);
  wd.set(chpxFkp, chpxPage * 512);
  wd.set(papxFkp, papxPage * 512);

  // 3. 1Table = STSH + PlcfSed + PlcfBteChpx + PlcfBtePapx + Clx + Sttbfffn.
  const stsh = b64(B64.STSH);
  const sttb = fontTable.bytes;
  const tbl = new Buf();
  const fcStshf = 0;
  tbl.bytes(stsh);

  const fcSed = tbl.length;
  // PlcfSed: aCP[2] = {0, ccpText}; aSed[1] = {fn:0, fcSepx:0xFFFFFFFF, fnMpr:0, fcMpr:0xFFFFFFFF}
  tbl.u32(0);
  tbl.u32(ccpText);
  tbl.u16(0);
  tbl.u32(0xffffffff);
  tbl.u16(0);
  tbl.u32(0xffffffff);
  const lcbSed = tbl.length - fcSed;

  const fcChpxBte = tbl.length;
  tbl.u32(TEXT_START);
  tbl.u32(textEnd);
  tbl.u32(chpxPage);
  const lcbChpxBte = tbl.length - fcChpxBte;

  const fcPapxBte = tbl.length;
  tbl.u32(TEXT_START);
  tbl.u32(textEnd);
  tbl.u32(papxPage);
  const lcbPapxBte = tbl.length - fcPapxBte;

  const fcClx = tbl.length;
  // Clx = Pcdt(0x02) + lcb(4) + PlcPcd{ aCP[2]={0,ccpText}, aPcd[1]{flags:0, fc, prm:0} }
  const plcPcd = new Buf();
  plcPcd.u32(0);
  plcPcd.u32(ccpText);
  plcPcd.u16(0); // PCD flags
  plcPcd.u32(TEXT_START); // FcCompressed: fCompressed=0, fc = byte offset of text
  plcPcd.u16(0); // prm
  const plc = plcPcd.done();
  tbl.u8(0x02);
  tbl.u32(plc.length);
  tbl.bytes(plc);
  const lcbClx = tbl.length - fcClx;

  const fcSttb = tbl.length;
  tbl.bytes(sttb);
  const lcbSttb = tbl.length - fcSttb;

  const table = tbl.done();

  // 4. Patch the FIB: ccpText, cbMac, and the fc/lcb pointers into the new table stream.
  patchFib(wd, {
    ccpText,
    cbMac: wdLen,
    pointers: [
      [FC.stshf, fcStshf, stsh.length],
      [FC.plcfSed, fcSed, lcbSed],
      [FC.plcfBteChpx, fcChpxBte, lcbChpxBte],
      [FC.plcfBtePapx, fcPapxBte, lcbPapxBte],
      [FC.clx, fcClx, lcbClx],
      [FC.sttbfffn, fcSttb, lcbSttb],
    ],
  });

  return writeCfb([
    { name: "WordDocument", data: wd },
    { name: "1Table", data: table },
    { name: "SummaryInformation", data: b64(B64.SUMMARY) },
    { name: "DocumentSummaryInformation", data: b64(B64.DOCSUMMARY) },
  ]);
}

function patchFib(
  wd: Uint8Array,
  opts: { ccpText: number; cbMac: number; pointers: [number, number, number][] },
): void {
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  const csw = dv.getUint16(32, true);
  const rgLwOff = 34 + csw * 2 + 2;
  const cslw = dv.getUint16(34 + csw * 2, true);
  const blobOffset = rgLwOff + cslw * 4 + 2;
  dv.setUint32(rgLwOff + 0, opts.cbMac, true); // cbMac (fibRgLw[0])
  dv.setInt32(rgLwOff + 12, opts.ccpText, true); // ccpText (fibRgLw[3])
  for (const [idx, fc, lcb] of opts.pointers) {
    dv.setUint32(blobOffset + idx * 8, fc, true);
    dv.setUint32(blobOffset + idx * 8 + 4, lcb, true);
  }
}
