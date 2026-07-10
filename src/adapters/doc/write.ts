import { writeCfb } from "./cfb";
import { FC } from "./fib";
import { B64, b64 } from "./templates";
import type { Note, PageGeometry } from "../../core/types";

// A comment reduced to what the binary stores: an author name plus plain-text body.
export interface DocComment {
  id: string;
  author: string;
  text: string;
}

// The subdocuments appended after the main text, in the CP order Word lays them out.
type SubKind = "footnote" | "comment" | "endnote";

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
  special?: boolean; // sprmCFSpec: a special character (footnote/endnote ref auto-number)
  fnRef?: { id: string; kind: SubKind }; // an inline reference (note or comment) in the main text
}
interface Para {
  align: number; // 0 left, 1 center, 2 right, 3 justify
  runs: Run[];
  istd: number; // paragraph style index (0 = Normal, 1..6 = Heading 1..6)
  indentTwips?: number; // left indent in twips
  endChar?: number; // paragraph terminator (0x0D normally, 0x07 for table cells/rows)
  table?: { cell?: boolean; ttp?: boolean; cols?: number };
  noteBoundary?: SubKind; // first para of a note/comment (or its subdoc's trailing mark)
}

const HEADING_SIZE = [48, 36, 28, 24, 22, 20]; // h1..h6 in half-points

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
      if (el.getAttribute("data-docx-pagebreak") || el.classList.contains("docx-pagebreak")) {
        runs.push(mkRun("\f", f)); // 0x0C page break
        continue;
      }
      if (tag === "sup" && el.classList.contains("docx-fnref")) {
        const kind = el.getAttribute("data-fn-kind") === "endnote" ? "endnote" : "footnote";
        const id = el.getAttribute("data-fn-id") || "";
        runs.push({ ...mkRun("\x02", f), special: true, fnRef: { id, kind } });
        continue;
      }
      if (el.classList.contains("docx-comment-ref")) {
        const id = el.getAttribute("data-comment-id") || "";
        runs.push({ ...mkRun("\x05", f), special: true, fnRef: { id, kind: "comment" } });
        continue; // the emoji glyph inside the ref is not part of the text
      }
      if (el.classList.contains("docx-cmark")) continue; // bookmark markers carry no text
      if (el.classList.contains("docx-comment")) {
        collectRuns(el, f, runs); // a comment range wrapper: keep the text it spans
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
  const walk = (el: Element, list: { ordered: boolean; n: number } | null): void => {
    for (const child of Array.from(el.children)) {
      const tag = child.tagName.toLowerCase();
      if (tag === "ul") walk(child, { ordered: false, n: 0 });
      else if (tag === "ol") walk(child, { ordered: true, n: 0 });
      else if (tag === "table" || tag === "tbody" || tag === "thead") walk(child, list);
      else if (tag === "tr") {
        const cells = Array.from(child.children).filter((c) => /^t[dh]$/.test(c.tagName.toLowerCase()));
        for (const td of cells) {
          const cr: Run[] = [];
          collectRuns(td, { b: false, i: false, u: false, strike: false }, cr);
          paras.push({ align: 0, runs: cr, istd: 0, endChar: 0x07, table: { cell: true } });
        }
        paras.push({ align: 0, runs: [], istd: 0, endChar: 0x07, table: { ttp: true, cols: cells.length } });
      }
      else if (blockTags.has(tag)) {
        const hMatch = /^h([1-6])$/.exec(tag);
        const level = hMatch ? Number(hMatch[1]) : 0;
        const base: Fmt = { b: level > 0, i: false, u: false, strike: false };
        if (level > 0) base.sizeHalf = HEADING_SIZE[level - 1];
        const runs: Run[] = [];
        collectRuns(child, base, runs);
        let prefix: Run[] = [];
        let indent = 0;
        if (tag === "li") {
          const marker = list?.ordered ? `\t${++list.n}.\t` : "\t•\t";
          prefix = [{ text: marker, b: false, i: false, u: false, strike: false }];
          indent = 360; // 0.25" hanging indent, like textutil
        }
        paras.push({ align: blockAlign(child), runs: [...prefix, ...runs], istd: level, indentTwips: indent });
      }
    }
  };
  walk(doc.body, null);
  if (paras.length === 0) paras.push({ align: 0, runs: [], istd: 0 });
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
  if (r.special) {
    b.u16(0x0855); // sprmCFSpec
    b.u8(1);
  }
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
  if (p.indentTwips) {
    b.u16(0x840f); // sprmPDxaLeft
    b.u16(p.indentTwips);
  }
  if (p.table?.cell || p.table?.ttp) {
    b.u16(0x2416); // sprmPFInTable
    b.u8(1);
  }
  if (p.table?.ttp) {
    b.u16(0x2417); // sprmPFTtp
    b.u8(1);
    const operand = buildTDef(p.table.cols || 1);
    b.u16(0xd608); // sprmTDefTable (2-byte length prefix)
    b.u16(operand.length);
    b.bytes(operand);
  }
  return b.done();
}

// sprmTDefTable operand: itcMac + rgdxaCenter (column boundaries, twips) + a TC80 per
// column (single-line borders), spanning a 6" (8640-twip) table width.
function buildTDef(cols: number): Uint8Array {
  const b = new Buf();
  b.u8(cols); // itcMac
  const total = 8640;
  for (let i = 0; i <= cols; i++) b.u16(Math.round((i * total) / cols)); // rgdxaCenter
  for (let i = 0; i < cols; i++) {
    b.u16(0x0680); // TC80 grfcotc
    b.u16(0); // wWidth
    for (let brc = 0; brc < 4; brc++) {
      b.u8(0x08); // dptLineWidth
      b.u8(0x01); // brcType single
      b.u8(0x01);
      b.u8(0x00);
    }
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
  istd: number;
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
    // grpprl = istd (2 bytes) + sprms
    const g = paras[i].grpprl;
    const grpprl = new Uint8Array(2 + g.length);
    grpprl[0] = paras[i].istd & 0xff;
    grpprl[1] = (paras[i].istd >> 8) & 0xff;
    grpprl.set(g, 2);
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

// Build a paragraph "Heading N" STD (sti=N, based on Normal), carrying bold + the heading
// font size so it renders in Word even without direct formatting.
function buildHeadingStd(n: number): Uint8Array {
  const name = `heading ${n}`;
  const b = new Buf();
  b.u16(n); // sti = n, flags 0
  b.u16(0x0001); // stk = 1 (paragraph), istdBase = 0 (Normal)
  b.u16(0x0002); // cupx = 2, istdNext = 0
  b.u16(0); // bchUpe (patched below)
  b.u16(0); // grfstd
  b.u16(name.length);
  for (const ch of name) b.u16(ch.charCodeAt(0));
  b.u16(0); // null terminator
  // LPUpxPapx: cbUPX = 2, UpxPapx = istd(2) = n
  b.u16(2);
  b.u16(n);
  // LPUpxChpx: grpprl = bold + heading size
  const chpx = new Buf();
  chpx.u16(0x0835);
  chpx.u8(1);
  chpx.u16(0x4a43);
  chpx.u16(HEADING_SIZE[n - 1]);
  const cg = chpx.done();
  b.u16(cg.length);
  b.bytes(cg);
  if (cg.length % 2 !== 0) b.u8(0); // pad UPX to even
  const std = b.done();
  new DataView(std.buffer).setUint16(6, std.length, true); // bchUpe = cbStd (like Normal)
  return std;
}

// Rebuild the stylesheet, filling the template's empty istd 1..6 slots with Heading 1..6.
function buildStshWithHeadings(): Uint8Array {
  const tpl = b64(B64.STSH);
  const dv = new DataView(tpl.buffer, tpl.byteOffset, tpl.byteLength);
  const cbStshi = dv.getUint16(0, true);
  const cstd = dv.getUint16(2, true); // STSHI.cstd
  const stds: (Uint8Array | null)[] = [];
  let i = 2 + cbStshi;
  for (let istd = 0; istd < cstd; istd++) {
    const cbStd = dv.getUint16(i, true);
    i += 2;
    if (cbStd === 0) stds.push(null);
    else {
      stds.push(tpl.subarray(i, i + cbStd));
      i += cbStd;
    }
  }
  for (let n = 1; n <= 6; n++) stds[n] = buildHeadingStd(n);
  const out = new Buf();
  out.bytes(tpl.subarray(0, 2 + cbStshi));
  for (const std of stds) {
    if (!std) out.u16(0);
    else {
      out.u16(std.length);
      out.bytes(std);
    }
  }
  return out.done();
}

// Build a section-properties exception (SEPX): cb + grpprl of section sprms, from
// richdoc's page geometry (px). Covers page size/margins, multi-column, vertical
// (tategaki) and right-to-left flow.
function buildSepx(page: PageGeometry): Uint8Array {
  const g = new Buf();
  const tw = (px: number) => Math.round(px * 15); // px @96dpi -> twips
  g.u16(0xb01f); // sprmSXaPage
  g.u16(tw(page.widthPx));
  g.u16(0xb020); // sprmSYaPage
  g.u16(tw(page.heightPx));
  g.u16(0xb021); // sprmSDxaLeft
  g.u16(tw(page.margin.left));
  g.u16(0xb022); // sprmSDxaRight
  g.u16(tw(page.margin.right));
  g.u16(0x9023); // sprmSDyaTop
  g.u16(tw(page.margin.top));
  g.u16(0x9024); // sprmSDyaBottom
  g.u16(tw(page.margin.bottom));
  if (page.columns && page.columns > 1) {
    g.u16(0x500b); // sprmSCcolumns (cCols - 1)
    g.u16(page.columns - 1);
    g.u16(0x900c); // sprmSDxaColumns (spacing)
    g.u16(tw(page.columnGapPx ?? 36));
    g.u16(0x3005); // sprmSFEvenlySpaced
    g.u8(1);
  }
  if (page.vertical) {
    g.u16(0x5453); // sprmSTextFlow = 1 (top-to-bottom, right-to-left: tategaki)
    g.u16(1);
  }
  if (page.rtl) {
    g.u16(0x5228); // sprmSFBiDi
    g.u16(1);
  }
  const grpprl = g.done();
  const b = new Buf();
  b.u16(grpprl.length); // cb
  b.bytes(grpprl);
  return b.done();
}

// Build a note's subdocument paragraphs from its body HTML: the note text as paragraphs,
// with the first prefixed by the auto-number reference char (0x02, special) and a tab, and
// tagged as a note boundary so the assembler can record its text-span start.
function buildNoteParas(html: string, kind: "footnote" | "endnote"): Para[] {
  const sub = parseHtml(html || "<p><br></p>");
  if (!sub.length) sub.push({ align: 0, runs: [], istd: 0 });
  const noNo: Fmt = { b: false, i: false, u: false, strike: false };
  sub[0].runs = [{ ...mkRun("\x02", noNo), special: true }, mkRun("\t", noNo), ...sub[0].runs];
  sub[0].noteBoundary = kind;
  return sub;
}

// A comment's annotation-subdocument paragraph: the reference mark (0x05, special) then the
// plain comment text. One paragraph per comment (Word 97 comments are unthreaded plain text).
function buildCommentParas(text: string): Para[] {
  const noNo: Fmt = { b: false, i: false, u: false, strike: false };
  const runs: Run[] = [{ ...mkRun("\x05", noNo), special: true }];
  if (text) runs.push(mkRun(text, noNo));
  return [{ align: 0, runs, istd: 0, noteBoundary: "comment" }];
}

// ---------------------------------------------------------------------------
// Assemble
// ---------------------------------------------------------------------------

export function htmlToDoc(bodyHtml: string, page?: PageGeometry, notes?: Note[], comments?: DocComment[]): Uint8Array {
  const mainParas = parseHtml(bodyHtml);

  // Subdocuments follow the main text in a fixed CP order (footnote, comment, endnote). Each
  // one's bodies are ordered to match the order their references appear in the body, so ref[i]
  // lines up with body[i] in the PLCFs. Orphan refs (no matching body) get an empty body.
  const refSeq: { id: string; kind: SubKind }[] = [];
  for (const p of mainParas) for (const r of p.runs) if (r.fnRef) refSeq.push(r.fnRef);
  const noteById = new Map((notes ?? []).map((n) => [n.id, n]));
  const cmtById = new Map((comments ?? []).map((c) => [c.id, c]));
  const groupsFor = (kind: SubKind): Para[][] =>
    refSeq.filter((r) => r.kind === kind).map((r) =>
      kind === "comment" ? buildCommentParas(cmtById.get(r.id)?.text ?? "") : buildNoteParas(noteById.get(r.id)?.html ?? "", kind));
  // Each present subdocument = its bodies plus a trailing empty paragraph (Word's subdoc mark).
  const mkRegion = (kind: SubKind): { kind: SubKind; paras: Para[] } | null => {
    const groups = groupsFor(kind);
    return groups.length ? { kind, paras: [...groups.flat(), { align: 0, runs: [], istd: 0, noteBoundary: kind }] } : null;
  };
  const regions = (["footnote", "comment", "endnote"] as SubKind[]).map(mkRegion).filter((r): r is { kind: SubKind; paras: Para[] } => !!r);
  const paras = [...mainParas];
  const regionStartIdx = new Map<SubKind, number>();
  for (const reg of regions) { regionStartIdx.set(reg.kind, paras.length); paras.push(...reg.paras); }

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
  const refCps: Record<SubKind, number[]> = { footnote: [], comment: [], endnote: [] }; // ref char CPs in the main doc
  const txtCps: Record<SubKind, number[]> = { footnote: [], comment: [], endnote: [] }; // text-span starts within each subdoc
  const regionBaseCp = new Map<SubKind, number>(); // CP where each subdocument starts
  for (let pi = 0; pi < paras.length; pi++) {
    const p = paras[pi];
    const paraStartCp = chars.length;
    for (const [kind, idx] of regionStartIdx) if (pi === idx) regionBaseCp.set(kind, paraStartCp);
    if (p.noteBoundary) txtCps[p.noteBoundary].push(paraStartCp);
    const runList = p.runs.length ? p.runs : [{ text: "", b: false, i: false, u: false, strike: false } as Run];
    for (const r of runList) {
      const startCp = chars.length;
      if (r.fnRef) refCps[r.fnRef.kind].push(startCp);
      for (const ch of r.text) chars.push(ch.codePointAt(0)!);
      charRuns.push({ fcStart: fcAt(startCp), fcEnd: fcAt(chars.length), grpprl: chpxGrpprl(r, fontTable.ftc) });
    }
    chars.push(p.endChar ?? 0x0d); // paragraph / cell / row mark
    // The paragraph's CHPX must also cover its mark; extend the last run to include it.
    charRuns[charRuns.length - 1].fcEnd = fcAt(chars.length);
    paraRuns.push({ fcStart: fcAt(paraStartCp), fcEnd: fcAt(chars.length), grpprl: papxGrpprl(p), istd: p.istd });
  }
  const total = chars.length;
  // Each subdocument's char count is the gap to the next one (main, then footnote, comment,
  // endnote in that order); a missing subdocument contributes nothing.
  const baseOf = (kind: SubKind) => regionBaseCp.get(kind) ?? -1;
  const ftnBase = baseOf("footnote");
  const cmtBase = baseOf("comment");
  const ednBase = baseOf("endnote");
  const ccpText = ftnBase >= 0 ? ftnBase : cmtBase >= 0 ? cmtBase : ednBase >= 0 ? ednBase : total;
  const nextAfter = (cp: number) => [cmtBase, ednBase, total].filter((c) => c > cp).sort((a, b) => a - b)[0] ?? total;
  const ccpFtn = ftnBase >= 0 ? nextAfter(ftnBase) - ftnBase : 0;
  const ccpAtn = cmtBase >= 0 ? (ednBase >= 0 ? ednBase : total) - cmtBase : 0;
  const ccpEdn = ednBase >= 0 ? total - ednBase : 0;
  const relTo = (base: number, cps: number[]) => cps.map((c) => c - base);
  const ftnRefCps = refCps.footnote, ednRefCps = refCps.endnote, atnRefCps = refCps.comment;
  const ftnTxtCps = relTo(ftnBase, txtCps.footnote), ednTxtCps = relTo(ednBase, txtCps.endnote), atnTxtCps = relTo(cmtBase, txtCps.comment);
  // PLCF CPs: refs are absolute in the main doc (terminated at ccpText); text spans are
  // already relative to each subdocument and terminate at its char count.
  const plcffndRef = ccpFtn ? [...ftnRefCps, ccpText] : [];
  const plcfendRef = ccpEdn ? [...ednRefCps, ccpText] : [];
  const plcfandRef = ccpAtn ? [...atnRefCps, ccpText] : [];
  const plcffndTxt = ccpFtn ? [...ftnTxtCps, ccpFtn] : [];
  const plcfendTxt = ccpEdn ? [...ednTxtCps, ccpEdn] : [];
  const plcfandTxt = ccpAtn ? [...atnTxtCps, ccpAtn] : [];
  // Comment authors: a group of Xst (cch + UTF-16) indexed by each ATRD's ibst.
  const commentList = ccpAtn ? atnRefCps.map((_, i) => cmtById.get(refSeq.filter((r) => r.kind === "comment")[i]!.id)) : [];
  const authorNames = [...new Set(commentList.map((c) => c?.author || "Author"))];
  const authorIbst = commentList.map((c) => Math.max(0, authorNames.indexOf(c?.author || "Author")));
  const textBytes = new Uint8Array(total * 2);
  {
    const dv = new DataView(textBytes.buffer);
    for (let i = 0; i < total; i++) dv.setUint16(i * 2, chars[i] > 0xffff ? 0x3f : chars[i], true);
  }

  // 2. WordDocument = FIB template + text + CHPX FKP + PAPX FKP.
  const fib = b64(B64.FIB).slice(); // 1536 bytes
  const chpxFkp = buildChpxFkp(charRuns);
  const papxFkp = buildPapxFkp(paraRuns);
  const textEnd = TEXT_START + textBytes.length;
  const fkpBase = Math.ceil(textEnd / 512) * 512;
  const chpxPage = fkpBase / 512;
  const papxPage = chpxPage + 1;
  const sepx = page ? buildSepx(page) : null;
  const sepxOffset = (papxPage + 1) * 512;
  const wdEnd = sepx ? sepxOffset + sepx.length : (papxPage + 1) * 512;
  const wdLen = Math.max(4096, Math.ceil(wdEnd / 512) * 512);
  const wd = new Uint8Array(wdLen);
  wd.set(fib, 0);
  wd.set(textBytes, TEXT_START);
  wd.set(chpxFkp, chpxPage * 512);
  wd.set(papxFkp, papxPage * 512);
  if (sepx) wd.set(sepx, sepxOffset);

  // 3. 1Table = STSH + PlcfSed + PlcfBteChpx + PlcfBtePapx + Clx + Sttbfffn.
  const stsh = buildStshWithHeadings();
  const sttb = fontTable.bytes;
  const tbl = new Buf();
  const fcStshf = 0;
  tbl.bytes(stsh);

  const fcSed = tbl.length;
  // PlcfSed: aCP[2] = {0, ccpText}; aSed[1] = {fn:0, fcSepx:0xFFFFFFFF, fnMpr:0, fcMpr:0xFFFFFFFF}
  tbl.u32(0);
  tbl.u32(ccpText);
  tbl.u16(0);
  tbl.u32(sepx ? sepxOffset : 0xffffffff);
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

  // Footnote / endnote PLCFs: reference CPs (each with a 2-byte FRD of 0 = auto-numbered)
  // and text-span CPs (no per-element data). Empty when the doc has no notes of that kind.
  const writeRefPlc = (cps: number[]): { fc: number; lcb: number } => {
    const fc = tbl.length;
    if (!cps.length) return { fc: 0, lcb: 0 };
    for (const cp of cps) tbl.u32(cp);
    for (let k = 0; k < cps.length - 1; k++) tbl.u16(0); // FRD per reference
    return { fc, lcb: tbl.length - fc };
  };
  const writeTxtPlc = (cps: number[]): { fc: number; lcb: number } => {
    const fc = tbl.length;
    if (!cps.length) return { fc: 0, lcb: 0 };
    for (const cp of cps) tbl.u32(cp);
    return { fc, lcb: tbl.length - fc };
  };
  const fndRef = writeRefPlc(plcffndRef);
  const fndTxt = writeTxtPlc(plcffndTxt);
  const endRef = writeRefPlc(plcfendRef);
  const endTxt = writeTxtPlc(plcfendTxt);
  // Comment reference PLCF: CPs of the 0x05 marks + a 30-byte ATRD each (ibst = author index,
  // lTagBkmk = -1 for no bookmark range; all other fields zero).
  const andRef = ((): { fc: number; lcb: number } => {
    const fc = tbl.length;
    if (!plcfandRef.length) return { fc: 0, lcb: 0 };
    for (const cp of plcfandRef) tbl.u32(cp);
    for (let k = 0; k < plcfandRef.length - 1; k++) {
      const atrd = new Uint8Array(30);
      new DataView(atrd.buffer).setUint16(20, authorIbst[k] ?? 0, true); // ibst
      atrd[26] = atrd[27] = atrd[28] = atrd[29] = 0xff; // lTagBkmk = -1
      tbl.bytes(atrd);
    }
    return { fc, lcb: tbl.length - fc };
  })();
  const andTxt = writeTxtPlc(plcfandTxt);
  // GrpXstAtnOwners: the author names as back-to-back Xst (cch u16 + UTF-16 chars).
  const grpXst = ((): { fc: number; lcb: number } => {
    const fc = tbl.length;
    if (!authorNames.length) return { fc: 0, lcb: 0 };
    for (const name of authorNames) {
      tbl.u16(name.length);
      for (const ch of name) tbl.u16(ch.charCodeAt(0));
    }
    return { fc, lcb: tbl.length - fc };
  })();

  const fcClx = tbl.length;
  // Clx = Pcdt(0x02) + lcb(4) + PlcPcd{ aCP[2]={0,total}, aPcd[1]{flags:0, fc, prm:0} }. The
  // single piece spans the whole text (main + note subdocuments).
  const plcPcd = new Buf();
  plcPcd.u32(0);
  plcPcd.u32(total);
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

  // 4. Patch the FIB: ccpText, the subdocument char counts, cbMac, and the fc/lcb pointers.
  patchFib(wd, {
    ccpText,
    ccpFtn,
    ccpAtn,
    ccpEdn,
    cbMac: wdLen,
    pointers: [
      [FC.stshf, fcStshf, stsh.length],
      [FC.plcffndRef, fndRef.fc, fndRef.lcb],
      [FC.plcffndTxt, fndTxt.fc, fndTxt.lcb],
      [FC.plcfandRef, andRef.fc, andRef.lcb],
      [FC.plcfandTxt, andTxt.fc, andTxt.lcb],
      [FC.plcfSed, fcSed, lcbSed],
      [FC.plcfBteChpx, fcChpxBte, lcbChpxBte],
      [FC.plcfBtePapx, fcPapxBte, lcbPapxBte],
      [FC.clx, fcClx, lcbClx],
      [FC.sttbfffn, fcSttb, lcbSttb],
      [FC.grpXstAtnOwners, grpXst.fc, grpXst.lcb],
      [FC.plcfendRef, endRef.fc, endRef.lcb],
      [FC.plcfendTxt, endTxt.fc, endTxt.lcb],
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
  opts: { ccpText: number; ccpFtn?: number; ccpAtn?: number; ccpEdn?: number; cbMac: number; pointers: [number, number, number][] },
): void {
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  const csw = dv.getUint16(32, true);
  const rgLwOff = 34 + csw * 2 + 2;
  const cslw = dv.getUint16(34 + csw * 2, true);
  const blobOffset = rgLwOff + cslw * 4 + 2;
  dv.setUint32(rgLwOff + 0, opts.cbMac, true); // cbMac (fibRgLw[0])
  dv.setInt32(rgLwOff + 12, opts.ccpText, true); // ccpText (fibRgLw[3])
  dv.setInt32(rgLwOff + 16, opts.ccpFtn ?? 0, true); // ccpFtn (fibRgLw[4])
  dv.setInt32(rgLwOff + 28, opts.ccpAtn ?? 0, true); // ccpAtn (fibRgLw[7])
  dv.setInt32(rgLwOff + 32, opts.ccpEdn ?? 0, true); // ccpEdn (fibRgLw[8])
  for (const [idx, fc, lcb] of opts.pointers) {
    dv.setUint32(blobOffset + idx * 8, fc, true);
    dv.setUint32(blobOffset + idx * 8 + 4, lcb, true);
  }
}
