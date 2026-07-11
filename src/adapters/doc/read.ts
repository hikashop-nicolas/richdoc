import { readCfb } from "./cfb";
import { parseFib, parsePieceTable, readPieceText, FC, type Piece } from "./fib";
import { mtefToMathml } from "./mtef";
import { bytesToBase64 } from "../../core/util";
import type { CommentThread, Note, PageGeometry } from "../../core/types";

// A floating drawing anchored in the main text: a positioned image (logo, banner, page watermark).
// Offsets are px relative to its anchor paragraph (Word anchors these to the paragraph, not the
// page), so buildHtml places it inside that paragraph, absolutely positioned behind the text.
interface DocFloat {
  img: string; // data URL
  cp: number; // the anchor char position (a 0x08 drawn-object char) in the main text
  dx: number;
  dy: number;
  w: number;
  h: number;
  reserve: boolean; // wrap "top and bottom" / "square": reserve the shape's vertical space in the flow
}

// Read half of the .doc adapter: bytes -> HTML in richdoc's vocabulary. It extracts the
// text (piece table) and the character/paragraph formatting (CHPX/PAPX formatted-disk
// pages), then emits paragraphs with inline runs the engine edits like docx/odt.

export interface DocParts {
  body: string;
  page?: PageGeometry;
  notes?: Note[];
  comments?: CommentThread[];
  header?: string;
  footer?: string;
}

/** A footnote / endnote reference found in the main text, keyed by its character position. */
interface NoteRef {
  id: string;
  kind: "footnote" | "endnote";
}

interface CharProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  sizeHalf?: number;
  color?: string; // #rrggbb
  font?: string; // font family name
  highlight?: string; // background colour #rrggbb
  picOffset?: number; // sprmCPicLocation: byte offset of this run's picture in the Data stream
  rev?: "ins" | "del"; // a tracked insertion / deletion
  rmIbst?: number; // revision author index into SttbfRMark
  rmDate?: string; // revision date (YYYY-MM-DD)
}

// Decode a Word DTTM (packed 32-bit date) to YYYY-MM-DD; empty if zero.
function decodeDttm(v: Uint8Array): string {
  const n = v[0]! | (v[1]! << 8) | (v[2]! << 16) | (v[3]! << 24);
  if (!n) return "";
  const dom = (n >>> 11) & 0x1f;
  const mon = (n >>> 16) & 0xf;
  const yr = ((n >>> 20) & 0x1ff) + 1900;
  if (!mon || !dom) return "";
  return `${yr}-${String(mon).padStart(2, "0")}-${String(dom).padStart(2, "0")}`;
}

// Word's 16-entry text-highlight palette (sprmCHighlight ico).
const HIGHLIGHT: Record<number, string> = {
  1: "#000000", 2: "#0000ff", 3: "#00ffff", 4: "#00ff00", 5: "#ff00ff", 6: "#ff0000",
  7: "#ffff00", 8: "#ffffff", 9: "#000080", 10: "#008080", 11: "#008000", 12: "#800080",
  13: "#800000", 14: "#808000", 15: "#808080", 16: "#c0c0c0",
};

// Parse the font table (Sttbfffn) into a name-by-index array. Each entry is an FFN whose
// name (UTF-16, null-terminated) sits at offset 40 after the 1-byte cbFfnM1.
function parseFontTable(table: Uint8Array, fc: number, lcb: number): string[] {
  const names: string[] = [];
  if (lcb < 6) return names;
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  const cData = dv.getUint16(0, true); // font count
  let pos = 4; // cData(2) + cbExtra(2)
  for (let f = 0; f < cData && pos < lcb; f++) {
    const cb = table[fc + pos];
    const nameStart = fc + pos + 1 + 39; // after cbFfnM1 + fixed FFN fields
    let name = "";
    for (let j = nameStart; j + 1 < fc + pos + 1 + cb; j += 2) {
      const ch = table[j] | (table[j + 1] << 8);
      if (ch === 0) break;
      name += String.fromCharCode(ch);
    }
    names.push(name);
    pos += 1 + cb;
  }
  return names;
}
interface ParaProps {
  align?: number; // 0..3
  indentTwips?: number;
  istd?: number;
  headingLevel?: number; // 1..6 when the paragraph's style is a heading
  styleChp?: CharProps; // character formatting inherited from the paragraph's named style
  inTable?: boolean; // sprmPFInTable
  ttp?: boolean; // sprmPFTtp (table-terminating / row-end paragraph)
  ilfo?: number; // sprmPIlfo: list-format index (>0 = a list item)
  ilvl?: number; // sprmPIlvl: list nesting level (0-based)
  pageBreakBefore?: boolean; // sprmPFPageBreakBefore: start this paragraph on a new page
  spaceBeforeTw?: number; // sprmPDyaBefore: space above the paragraph (twips)
  spaceAfterTw?: number; // sprmPDyaAfter: space below the paragraph (twips)
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Decode a CHPX grpprl (a run of sprms) into character properties.
function decodeCharSprms(g: Uint8Array, fonts: string[] = []): CharProps {
  const dv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  const p: CharProps = {};
  let i = 0;
  while (i + 2 <= g.length) {
    const op = dv.getUint16(i, true);
    i += 2;
    const spra = (op >> 13) & 7;
    const len = spra === 6 ? g[i++] : [1, 1, 2, 4, 2, 2, 0, 3][spra];
    const v = g.subarray(i, i + len);
    i += len;
    switch (op) {
      case 0x0835:
        p.b = v[0] !== 0;
        break;
      case 0x0836:
        p.i = v[0] !== 0;
        break;
      case 0x0837:
        p.strike = v[0] !== 0;
        break;
      case 0x2a3e:
        p.u = v[0] !== 0;
        break;
      case 0x4a43:
        p.sizeHalf = v[0] | (v[1] << 8);
        break;
      case 0x4a4f: // sprmCRgFtc0: font index into the font table
        p.font = fonts[v[0] | (v[1] << 8)] || undefined;
        break;
      case 0x2a0c: // sprmCHighlight: highlight palette index
        if (v[0] && HIGHLIGHT[v[0]]) p.highlight = HIGHLIGHT[v[0]];
        break;
      case 0x6870: // sprmCCv: RR GG BB shade
        p.color = "#" + [v[0], v[1], v[2]].map((x) => x.toString(16).padStart(2, "0")).join("");
        break;
      case 0x6a03: // sprmCPicLocation: offset into the Data stream of this picture char's PICF
        p.picOffset = v[0] | (v[1] << 8) | (v[2] << 16) | (v[3] << 24);
        break;
      case 0x0801: // sprmCFRMark: a tracked insertion
        if (v[0]) p.rev = "ins";
        break;
      case 0x0800: // sprmCFRMarkDel: a tracked deletion
        if (v[0]) p.rev = "del";
        break;
      case 0x4804: // sprmCIbstRMark: insertion author index into SttbfRMark
      case 0x4863: // sprmCIbstRMarkDel: deletion author index into SttbfRMark
        p.rmIbst = v[0] | (v[1] << 8);
        break;
      case 0x6805: // sprmCDttmRMark: insertion date
      case 0x6864: // sprmCDttmRMarkDel: deletion date
        p.rmDate = decodeDttm(v);
        break;
    }
  }
  return p;
}

function decodeParaSprms(g: Uint8Array): ParaProps {
  const dv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  const p: ParaProps = {};
  if (g.length >= 2) p.istd = dv.getUint16(0, true);
  let i = 2;
  while (i + 2 <= g.length) {
    const op = dv.getUint16(i, true);
    i += 2;
    const spra = (op >> 13) & 7;
    const len = spra === 6 ? g[i++] : [1, 1, 2, 4, 2, 2, 0, 3][spra];
    const v = g.subarray(i, i + len);
    i += len;
    if (op === 0x2403 || op === 0x2461) p.align = v[0];
    else if (op === 0x840f) p.indentTwips = dv.getInt16(i - len, true);
    else if (op === 0x2416) p.inTable = v[0] !== 0;
    else if (op === 0x2417) p.ttp = v[0] !== 0;
    else if (op === 0x460b) p.ilfo = dv.getInt16(i - len, true); // sprmPIlfo: list index
    else if (op === 0x260a) p.ilvl = v[0]; // sprmPIlvl: list level
    else if (op === 0x2407) p.pageBreakBefore = v[0] !== 0; // sprmPFPageBreakBefore
    else if (op === 0xa413) p.spaceBeforeTw = dv.getUint16(i - len, true); // sprmPDyaBefore
    else if (op === 0xa414) p.spaceAfterTw = dv.getUint16(i - len, true); // sprmPDyaAfter
  }
  return p;
}

// A property span over a byte range [fcStart, fcEnd).
interface Span<T> {
  fcStart: number;
  fcEnd: number;
  props: T;
}

// Walk the FKP pages referenced by a PlcfBte, decoding each run's grpprl.
function readFkpSpans<T>(
  wd: Uint8Array,
  table: Uint8Array,
  fcPlc: number,
  lcbPlc: number,
  isPapx: boolean,
  decode: (g: Uint8Array) => T,
): Span<T>[] {
  if (lcbPlc < 8) return [];
  const plc = table.subarray(fcPlc, fcPlc + lcbPlc);
  const pdv = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
  const n = (lcbPlc - 4) / 8;
  const pnBase = (n + 1) * 4;
  const spans: Span<T>[] = [];
  for (let k = 0; k < n; k++) {
    const pn = pdv.getUint32(pnBase + k * 4, true) & 0x3fffff;
    const fkp = wd.subarray(pn * 512, pn * 512 + 512);
    if (fkp.length < 512) continue;
    const fdv = new DataView(fkp.buffer, fkp.byteOffset, fkp.byteLength);
    const crun = fkp[511];
    const fcs: number[] = [];
    for (let r = 0; r <= crun; r++) fcs.push(fdv.getUint32(r * 4, true));
    const wordOffBase = 4 * (crun + 1);
    for (let r = 0; r < crun; r++) {
      let props: T;
      if (isPapx) {
        const bOff = fkp[wordOffBase + r * 13];
        if (bOff === 0) props = decode(new Uint8Array(0));
        else {
          const cb = fkp[bOff * 2];
          const grpprlLen = cb === 0 ? fkp[bOff * 2 + 1] * 2 : cb * 2 - 1;
          const start = bOff * 2 + (cb === 0 ? 2 : 1);
          props = decode(fkp.subarray(start, start + grpprlLen)); // includes istd
        }
      } else {
        const off = fkp[wordOffBase + r];
        if (off === 0) props = decode(new Uint8Array(0));
        else {
          const cb = fkp[off * 2];
          props = decode(fkp.subarray(off * 2 + 1, off * 2 + 1 + cb));
        }
      }
      spans.push({ fcStart: fcs[r], fcEnd: fcs[r + 1], props });
    }
  }
  return spans;
}

/** Map a logical character position to its byte offset, using the piece table. */
function makeCpToFc(pieces: Piece[]): (cp: number) => number {
  return (cp: number): number => {
    for (const p of pieces) {
      if (cp >= p.cpStart && cp < p.cpEnd) return p.fc + (cp - p.cpStart) * (p.unicode ? 2 : 1);
    }
    const last = pieces[pieces.length - 1];
    return last ? last.fc + (cp - last.cpStart) * (last.unicode ? 2 : 1) : 0;
  };
}

function lookup<T>(spans: Span<T>[], fc: number): T | undefined {
  for (const s of spans) if (fc >= s.fcStart && fc < s.fcEnd) return s.props;
  return undefined;
}

// Map each style index (istd) to a heading level 1..6, by reading each STD's sti (the
// low 12 bits of its first u16; sti 1..9 = Heading 1..9).
function parseStyleHeadings(table: Uint8Array, fc: number, lcb: number): Map<number, number> {
  const map = new Map<number, number>();
  if (lcb < 4) return map;
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  const cbStshi = dv.getUint16(0, true);
  const cstd = dv.getUint16(2, true);
  let i = 2 + cbStshi;
  for (let istd = 0; istd < cstd && i + 2 <= lcb; istd++) {
    const cbStd = dv.getUint16(i, true);
    i += 2;
    if (cbStd === 0) continue;
    const sti = dv.getUint16(i, true) & 0xfff;
    if (sti >= 1 && sti <= 9) map.set(istd, Math.min(sti, 6));
    i += cbStd;
  }
  return map;
}

// Resolve each style index (istd) to the character formatting it defines (bold/size/font/...),
// following the istdBase chain so a style inherits from the one it is based on. This is what lets
// a paragraph in a named style (e.g. a bold "Title") render styled even when its runs carry no
// direct character formatting. Each STD holds a UpxChpx (a CHPX grpprl) after its STDF + name +
// UpxPapx; toggle sprms (bold value 0x81) read as on, which matches heading/title styles.
interface StyleProps { chp: CharProps; pap: ParaProps }
function parseStyleChps(table: Uint8Array, fc: number, lcb: number, fonts: string[]): Map<number, StyleProps> {
  const resolved = new Map<number, StyleProps>();
  if (lcb < 6) return resolved;
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  const cbStshi = dv.getUint16(0, true);
  const cstd = dv.getUint16(2, true);
  const cbBase = dv.getUint16(4, true); // cbSTDBaseInFile: the fixed STDF length
  const raw = new Map<number, { istdBase: number; chp: CharProps; pap: ParaProps }>();
  let i = 2 + cbStshi;
  for (let istd = 0; istd < cstd && i + 2 <= lcb; istd++) {
    const cbStd = dv.getUint16(i, true);
    i += 2;
    const std = i;
    i += cbStd;
    if (cbStd < cbBase + 2 || std + cbStd > lcb) continue;
    const istdBase = dv.getUint16(std + 2, true) >> 4; // STDF: sgc(4) + istdBase(12) at offset 2
    let p = std + cbBase;
    const cch = dv.getUint16(p, true); // style name: cch UTF-16 units + a null terminator
    p += 2 + cch * 2 + 2;
    if (p % 2) p++;
    if (p + 2 > std + cbStd) continue;
    // UpxPapx = istd (2) + a PAPX grpprl; decodeParaSprms reads exactly that shape (istd, then sprms).
    const cbUpxPapx = dv.getUint16(p, true);
    const pap = p + 2 + cbUpxPapx <= std + cbStd ? decodeParaSprms(table.subarray(fc + p + 2, fc + p + 2 + cbUpxPapx)) : {};
    p += 2 + cbUpxPapx;
    if (cbUpxPapx % 2) p++;
    if (p + 2 > std + cbStd) { raw.set(istd, { istdBase, chp: {}, pap }); continue; }
    const cbUpxChpx = dv.getUint16(p, true);
    const chpxStart = p + 2;
    const chp = chpxStart + cbUpxChpx <= std + cbStd ? decodeCharSprms(table.subarray(fc + chpxStart, fc + chpxStart + cbUpxChpx), fonts) : {};
    raw.set(istd, { istdBase, chp, pap });
  }
  const resolve = (istd: number, seen: Set<number>): StyleProps => {
    if (resolved.has(istd)) return resolved.get(istd)!;
    const r = raw.get(istd);
    if (!r || seen.has(istd)) return { chp: {}, pap: {} };
    seen.add(istd);
    const base = r.istdBase !== 0xfff && r.istdBase !== istd ? resolve(r.istdBase, seen) : { chp: {}, pap: {} };
    const merged: StyleProps = { chp: { ...base.chp, ...r.chp }, pap: { ...base.pap, ...r.pap } };
    resolved.set(istd, merged);
    return merged;
  };
  for (const istd of raw.keys()) resolve(istd, new Set());
  return resolved;
}

// Parse one SEPX (section properties) at fcSepx into richdoc page geometry (undefined if it
// declares nothing recognisable). Exported for unit testing of the margin resolution.
export function parseSepx(wd: Uint8Array, fcSepx: number): PageGeometry | undefined {
  if (fcSepx === 0xffffffff || fcSepx + 2 > wd.length) return undefined;
  const cb = wd[fcSepx] | (wd[fcSepx + 1] << 8);
  const g = wd.subarray(fcSepx + 2, fcSepx + 2 + cb);
  const gdv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  const px = (tw: number) => Math.round(tw / 15);
  const page: PageGeometry = { widthPx: 816, heightPx: 1056, margin: { top: 96, right: 96, bottom: 96, left: 96 } };
  let any = false;
  let i = 0;
  // Word stores the top/bottom margin as signed twips; a NEGATIVE value means the margin is
  // measured to the header/footer, so the effective body margin is the header/footer distance.
  // The binding gutter (dzaGutter) adds to the left margin. Collect these, resolve after the loop.
  let topTw = 1440, bottomTw = 1440, leftTw = 1440, gutterTw = 0, hdrTopTw = 720, ftrBottomTw = 720;
  while (i + 2 <= g.length) {
    const op = gdv.getUint16(i, true);
    i += 2;
    const spra = (op >> 13) & 7;
    const len = spra === 6 ? g[i++] : [1, 1, 2, 4, 2, 2, 0, 3][spra];
    const start = i;
    i += len;
    const u16 = () => g[start] | (g[start + 1] << 8);
    if (op === 0xb01f) { page.widthPx = px(u16()); any = true; }
    else if (op === 0xb020) { page.heightPx = px(u16()); any = true; }
    else if (op === 0xb021) { leftTw = u16(); any = true; }
    else if (op === 0xb022) page.margin.right = px(u16());
    else if (op === 0x9023) topTw = gdv.getInt16(start, true);
    else if (op === 0x9024) bottomTw = gdv.getInt16(start, true);
    else if (op === 0xb025) gutterTw = u16(); // sprmSDzaGutter: binding space added to the left
    else if (op === 0xb017) hdrTopTw = u16(); // sprmSDyaHdrTop: header distance from the top edge
    else if (op === 0xb018) ftrBottomTw = u16(); // sprmSDyaHdrBottom: footer distance from the bottom
    else if (op === 0x500b) { page.columns = u16() + 1; any = true; }
    else if (op === 0x900c) page.columnGapPx = px(u16());
    else if (op === 0x5453) { if (u16() !== 0) { page.vertical = true; any = true; } }
    else if (op === 0x5228) { if (u16() !== 0) { page.rtl = true; any = true; } }
  }
  page.margin.top = px(topTw < 0 ? hdrTopTw : topTw);
  page.margin.bottom = px(bottomTw < 0 ? ftrBottomTw : bottomTw);
  page.margin.left = px(leftTw + gutterTw);
  return any ? page : undefined;
}

// Parse every section from the PlcfSed: its end CP and its geometry. A single-section document
// yields one entry; a multi-section one yields a geometry per section, in document order.
function parseSections(wd: Uint8Array, table: Uint8Array, fcSed: number, lcbSed: number): { endCp: number; geom?: PageGeometry }[] {
  if (lcbSed < 16) return [];
  const nSed = (lcbSed - 4) / 16; // PLCF: 4*(n+1) CPs + 12*n Sed = 16n+4
  const sedBase = 4 * (nSed + 1);
  const sdv = new DataView(table.buffer, table.byteOffset + fcSed, lcbSed);
  const out: { endCp: number; geom?: PageGeometry }[] = [];
  for (let k = 0; k < nSed; k++) {
    const endCp = sdv.getUint32((k + 1) * 4, true);
    const fcSepx = sdv.getUint32(sedBase + k * 12 + 2, true); // Sed: fn(2) then fcSepx(4)
    out.push({ endCp, geom: parseSepx(wd, fcSepx) });
  }
  return out;
}

// PageGeometry -> the engine's SecGeom JSON carried on a section-boundary paragraph.
function secGeomJson(g: PageGeometry): string {
  const s: Record<string, unknown> = { w: g.widthPx, h: g.heightPx, mt: g.margin.top, mr: g.margin.right, mb: g.margin.bottom, ml: g.margin.left };
  if (g.columns) { s.cols = g.columns; s.colGap = g.columnGapPx; }
  if (g.vertical) s.vertical = true;
  if (g.rtl) s.rtl = true;
  return JSON.stringify(s);
}

// Read a PLCF's CPs: (lcb-4)/(4+dataSize) elements, so n+1 CPs of 4 bytes each. Footnote/
// endnote ref PLCFs carry a 2-byte FRD per element; the text PLCFs carry none.
function readPlcfCps(table: Uint8Array, fc: number, lcb: number, dataSize: number): number[] {
  if (lcb < 8) return [];
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  const n = Math.floor((lcb - 4) / (4 + dataSize));
  const cps: number[] = [];
  for (let k = 0; k <= n; k++) cps.push(dv.getUint32(k * 4, true));
  return cps;
}

// Render one note's subdocument text (a CP range in `full`) into note-body HTML: strip the
// leading auto-number ref char (0x02) and its tab, split on paragraph marks, and carry each
// character's run styling. Produces one <p> per paragraph, matching the editor's note bands.
function renderNoteBody(
  full: string,
  start: number,
  end: number,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  imageAt: (offset: number) => string | null = () => null,
): string {
  const paras: string[] = [];
  let runHtml = "";
  let curStyle: string | null = null;
  let curText = "";
  const flushRun = () => {
    if (!curText) return;
    const body = esc(curText).replace(/\n/g, "<br>");
    runHtml += curStyle ? `<span style="${curStyle}">${body}</span>` : body;
    curText = "";
  };
  const flushPara = () => {
    flushRun();
    paras.push(`<p>${runHtml || "<br>"}</p>`);
    runHtml = "";
    curStyle = null;
  };
  let leadStripped = false; // the leading 0x02 + optional tab is the auto-number, not body text
  let inFieldInstr = false; // between a field begin (0x13) and separator (0x14): hidden field code
  let instr = "";
  let suppressResult = false; // drop a live field's cached result (a docx-field span replaces it)
  for (let cp = start; cp < end && cp < full.length; cp++) {
    const c = full.charCodeAt(cp);
    if (!leadStripped) {
      if (c === 0x02) continue;
      if (c === 0x09) { leadStripped = true; continue; }
      leadStripped = true;
    }
    // Field codes: drop the markers and hidden instruction. A PAGE / NUMPAGES field (common in a
    // running footer) becomes a live field span the engine fills; other fields keep their cached
    // result (so an EMBED field doesn't print "EMBED Word.Picture").
    if (c === 0x13) { inFieldInstr = true; instr = ""; suppressResult = false; continue; }
    if (c === 0x14) {
      inFieldInstr = false;
      const k = fieldKind(instr);
      if (k === "PAGE" || k === "NUMPAGES") { flushRun(); runHtml += `<span class="docx-field" data-field="${k}" contenteditable="false"></span>`; suppressResult = true; }
      continue;
    }
    if (c === 0x15) { suppressResult = false; continue; }
    if (inFieldInstr) { instr += full[cp]; continue; }
    if (suppressResult) continue;
    if (c === 0x01) {
      // A picture char (e.g. an embedded Word.Picture in this story): render its raster blip.
      const off = lookup(charSpans, cpToFc(cp))?.picOffset;
      const img = off !== undefined ? imageAt(off) : null;
      if (img) { flushRun(); runHtml += img; }
      continue;
    }
    if (c === PARA) { flushPara(); continue; }
    if (c === 0x09) { curText += "\t"; continue; }
    if (c === 0x0b) { curText += "\n"; continue; }
    if (c === 0x02 || c < 0x09) continue;
    const style = runStyle(lookup(charSpans, cpToFc(cp))) || null;
    if (style !== curStyle) { flushRun(); curStyle = style; }
    curText += full[cp];
  }
  if (curText || runHtml) flushPara();
  while (paras.length > 1 && paras[paras.length - 1] === "<p><br></p>") paras.pop();
  return paras.join("") || "<p><br></p>";
}

// Parse one kind of note subdocument: match each reference CP in the main text to the note
// whose text span sits at the same ordinal in the text PLCF. Returns the notes plus a
// cp -> ref map so buildHtml can drop an inline reference marker at each ref position.
function parseNotes(
  full: string,
  fib: ReturnType<typeof parseFib>,
  table: Uint8Array,
  kind: "footnote" | "endnote",
  subStart: number,
  ccpSub: number,
  refIdx: number,
  txtIdx: number,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  refMap: Map<number, NoteRef>,
  imageAt: (offset: number) => string | null = () => null,
): Note[] {
  if (ccpSub <= 0) return [];
  const ref = fib.fc(refIdx);
  const txt = fib.fc(txtIdx);
  const refCps = readPlcfCps(table, ref.fc, ref.lcb, 2); // n+1 CPs; last is a terminator
  const txtCps = readPlcfCps(table, txt.fc, txt.lcb, 0); // cNotes+2 CPs (spans + trailing mark)
  const cNotes = Math.max(0, Math.min(refCps.length - 1, txtCps.length - 2));
  const notes: Note[] = [];
  for (let i = 0; i < cNotes; i++) {
    const id = `${kind === "footnote" ? "fn" : "en"}${i + 1}`;
    refMap.set(refCps[i], { id, kind });
    notes.push({ id, kind, html: renderNoteBody(full, subStart + txtCps[i], subStart + txtCps[i + 1], cpToFc, charSpans, imageAt) });
  }
  return notes;
}

// Parse an (extended) Sttbf of double-byte strings, e.g. SttbfRMark revision author names.
function parseSttbStrings(table: Uint8Array, fc: number, lcb: number): string[] {
  if (lcb < 4) return [];
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  const ext = dv.getUint16(0, true) === 0xffff;
  const cData = ext ? dv.getUint16(2, true) : dv.getUint16(0, true);
  const cbExtra = ext ? dv.getUint16(4, true) : dv.getUint16(2, true);
  let i = ext ? 6 : 4;
  const out: string[] = [];
  for (let k = 0; k < cData && i + 2 <= lcb; k++) {
    const cch = dv.getUint16(i, true);
    i += 2;
    let s = "";
    for (let j = 0; j < cch && i + 2 <= lcb; j++, i += 2) s += String.fromCharCode(dv.getUint16(i, true));
    i += cbExtra;
    out.push(s);
  }
  return out;
}

// Parse GrpXstAtnOwners: back-to-back Xst (cch u16 + UTF-16 chars) of comment author names.
function parseAtnOwners(table: Uint8Array, fc: number, lcb: number): string[] {
  const names: string[] = [];
  if (lcb < 2) return names;
  const dv = new DataView(table.buffer, table.byteOffset + fc, lcb);
  let i = 0;
  while (i + 2 <= lcb) {
    const cch = dv.getUint16(i, true);
    i += 2;
    let name = "";
    for (let j = 0; j < cch && i + 2 <= lcb; j++, i += 2) name += String.fromCharCode(dv.getUint16(i, true));
    names.push(name);
  }
  return names;
}

// Parse the comment (annotation) subdocument: PlcfandRef gives the 0x05 reference CPs in the
// main text plus a 30-byte ATRD each (ibst = author index); PlcfandTxt gives each comment's
// text span in the subdocument. Returns flat threads (Word 97 comments are unthreaded) plus a
// cp -> id map so buildHtml drops an inline reference marker at each anchor.
function parseComments(
  full: string,
  fib: ReturnType<typeof parseFib>,
  table: Uint8Array,
  subStart: number,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  refMap: Map<number, string>,
): CommentThread[] {
  if (fib.ccpAtn <= 0) return [];
  const ref = fib.fc(FC.plcfandRef);
  const txt = fib.fc(FC.plcfandTxt);
  const refCps = readPlcfCps(table, ref.fc, ref.lcb, 30);
  const txtCps = readPlcfCps(table, txt.fc, txt.lcb, 0);
  const owners = parseAtnOwners(table, fib.fc(FC.grpXstAtnOwners).fc, fib.fc(FC.grpXstAtnOwners).lcb);
  const cAtn = Math.max(0, Math.min(refCps.length - 1, txtCps.length - 2));
  const atrdBase = ref.fc + refCps.length * 4;
  const dv = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const out: CommentThread[] = [];
  for (let i = 0; i < cAtn; i++) {
    const id = `dc${i + 1}`;
    const ibst = dv.getUint16(atrdBase + i * 30 + 20, true);
    refMap.set(refCps[i], id);
    const raw = renderNoteBody(full, subStart + txtCps[i], subStart + txtCps[i + 1], cpToFc, charSpans);
    const text = raw.replace(/<[^>]+>/g, "").trim(); // comment bodies are stored as plain text
    out.push({ id, author: owners[ibst] ?? "Author", date: "", text, reactions: [], paraId: id, replies: [], resolved: false });
  }
  return out;
}

// The header/footer subdocument (HDD) stores stories delimited by PlcfHdd: 6 footnote/endnote
// separator stories, then 6 per section [even header, odd header, even footer, odd footer,
// first header, first footer]. The primary header/footer are section 0's odd stories (index 7
// and 9). Returns their HTML (empty when the story is empty).
function parseHeaderFooter(
  full: string,
  fib: ReturnType<typeof parseFib>,
  table: Uint8Array,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  imageAt: (offset: number) => string | null = () => null,
): { header: string; footer: string } {
  const empty = { header: "", footer: "" };
  if (fib.ccpHdd <= 0) return empty;
  const hdd = fib.fc(FC.plcfHdd);
  if (hdd.lcb < 44) return empty; // need at least the 6 separators + section-0 header/footer
  const dv = new DataView(table.buffer, table.byteOffset + hdd.fc, hdd.lcb);
  const n = hdd.lcb / 4;
  const cp = (k: number) => (k < n ? dv.getUint32(k * 4, true) : dv.getUint32((n - 1) * 4, true));
  const base = fib.ccpText + fib.ccpFtn;
  const story = (i: number) => (cp(i + 1) > cp(i) ? renderNoteBody(full, base + cp(i), base + cp(i + 1), cpToFc, charSpans, imageAt) : "");
  const clean = (h: string) => (h === "<p><br></p>" ? "" : h);
  // Each section has its own 6 stories after the 6 note separators: [evenHdr, oddHdr, evenFtr,
  // oddFtr, firstHdr, firstFtr]. The cover section's footer is often just a page number while a
  // later section carries the real running footer, so pick the section story with the most text.
  const letters = (h: string) => h.replace(/<[^>]+>/g, "").replace(/[^\p{L}]/gu, "").length;
  const sections = Math.max(1, Math.floor((n - 6) / 6));
  const pick = (slot: number): string => {
    let best = "";
    for (let s = 0; s < sections; s++) {
      const h = story(6 + s * 6 + slot);
      if (letters(h) > letters(best)) best = h;
    }
    return best;
  };
  return { header: clean(pick(1)), footer: clean(pick(3)) };
}

// Parse the textbox subdocument: the story text of every drawn text box, sitting after the
// endnotes in CP order. PlcfTxbxTxt delimits each box (its FTXBXS-per-box CPs plus a trailing
// reserved sentinel, so real boxes = elements - 1). Returns each box's rendered HTML in the
// order Word stores them, which matches the order of the 0x08 drawn-object anchors in the main
// text; buildHtml pairs them positionally. Anchoring to the exact shape (via FSPA/SPID) is not
// modelled, so multiple boxes on one paragraph keep document order but not their float layout.
function parseTextboxes(
  full: string,
  fib: ReturnType<typeof parseFib>,
  table: Uint8Array,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  imageAt: (offset: number) => string | null = () => null,
): string[] {
  if (fib.ccpTxbx <= 0) return [];
  const txbxStart = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn + fib.ccpEdn;
  const plc = fib.fc(FC.plcftxbxTxt);
  const cps = readPlcfCps(table, plc.fc, plc.lcb, 22); // FTXBXS is 22 bytes per element
  const boxes: string[] = [];
  // Elements = cps.length - 1; the last element is the reserved sentinel, so drop it.
  for (let i = 0; i < cps.length - 2; i++) {
    const a = txbxStart + cps[i];
    const b = txbxStart + cps[i + 1];
    if (b > a) boxes.push(renderNoteBody(full, a, b, cpToFc, charSpans, imageAt));
  }
  return boxes;
}

// Raster image magic numbers we can surface as a browser <img> (Word metafiles, WMF/EMF, are
// not browser-renderable and are skipped).
const IMAGE_SIGS: { mime: string; sig: number[] }[] = [
  { mime: "image/png", sig: [0x89, 0x50, 0x4e, 0x47] },
  { mime: "image/jpeg", sig: [0xff, 0xd8, 0xff] },
  { mime: "image/gif", sig: [0x47, 0x49, 0x46, 0x38] },
  { mime: "image/bmp", sig: [0x42, 0x4d] },
];
function indexOfSig(buf: Uint8Array, sig: number[]): number {
  for (let i = 0; i + sig.length <= buf.length; i++) {
    let ok = true;
    for (let j = 0; j < sig.length; j++) if (buf[i + j] !== sig[j]) { ok = false; break; }
    if (ok) return i;
  }
  return -1;
}
// Extract the raster blip a picture char points at: at `offset` in the Data stream sits a PICF
// (its first uint32 is the total byte length), and the embedded image bytes follow the header,
// wrapped in OfficeArt records. We locate the raster magic within the PICF and take it to the
// PICF's end. Returns an <img> data-URL, or null for a metafile / unrecognised blip.
function extractImageHtml(data: Uint8Array | undefined, offset: number): string | null {
  if (!data || offset < 0 || offset + 4 > data.length) return null;
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const lcb = dv.getUint32(offset, true);
  const end = lcb >= 8 && offset + lcb <= data.length ? offset + lcb : data.length;
  const region = data.subarray(offset, end);
  for (const { mime, sig } of IMAGE_SIGS) {
    const at = indexOfSig(region, sig);
    if (at >= 0) return `<img src="data:${mime};base64,${bytesToBase64(region.subarray(at))}">`;
  }
  return null;
}

// The OfficeArt blip store: each BSE record (0xF007) names an image whose bytes live in the delay
// stream (WordDocument) at foDelay. Returns a raster data URL per blip, indexed 1-based by pib
// (the order shapes reference them). Metafile blips (WMF/EMF) are skipped (not browser-raster).
function parseBlipStore(table: Uint8Array, wd: Uint8Array): (string | undefined)[] {
  const blips: (string | undefined)[] = [undefined]; // pib is 1-based; index 0 is unused
  const tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const wdv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  for (let i = 0; i + 44 <= table.length; i++) {
    if (tdv.getUint16(i + 2, true) !== 0xf007) continue; // msofbtBSE
    if (tdv.getUint32(i + 4, true) < 36) continue;
    const btWin = table[i + 8];
    if (btWin < 2 || btWin > 8) continue; // blip type (EMF..TIFF); guards against false matches
    const size = tdv.getUint32(i + 8 + 20, true);
    const foDelay = tdv.getUint32(i + 8 + 28, true);
    let url: string | undefined;
    if (foDelay !== 0xffffffff && foDelay + 8 <= wd.length) {
      const brl = wdv.getUint32(foDelay + 4, true); // the blip record's own length bounds the data
      const end = Math.min(foDelay + 8 + (brl > 0 && brl < wd.length ? brl : size), wd.length);
      const region = wd.subarray(foDelay, end);
      for (const { mime, sig } of IMAGE_SIGS) {
        const at = indexOfSig(region, sig);
        if (at >= 0) { url = `data:${mime};base64,${bytesToBase64(region.subarray(at))}`; break; }
      }
    }
    blips.push(url);
  }
  return blips;
}

// Map each shape id (spid) to the blip it displays (pib), by walking the OfficeArt records: an Sp
// record (0xF00A) carries the spid, the following OPT record (0xF00B) may hold the pib property
// (0x0104) in its property table (6 bytes each). -1 when the shape shows no picture (e.g. a textbox).
function parseShapePibs(table: Uint8Array): Map<number, number> {
  const map = new Map<number, number>();
  const tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
  let lastSpid = -1;
  for (let i = 0; i + 8 <= table.length; i++) {
    const rt = tdv.getUint16(i + 2, true);
    if (rt === 0xf00a) lastSpid = tdv.getUint32(i + 8, true);
    else if (rt === 0xf00b) {
      const nProp = (tdv.getUint16(i, true) >> 4) & 0xfff;
      let p = i + 8;
      let pib = -1;
      for (let q = 0; q < nProp && p + 6 <= table.length; q++) {
        if ((tdv.getUint16(p, true) & 0x3fff) === 0x0104) pib = tdv.getUint32(p + 2, true);
        p += 6;
      }
      if (lastSpid >= 0) { map.set(lastSpid, pib); lastSpid = -1; }
    }
  }
  return map;
}

// Parse the main document's floating drawings: each FSPA (26-byte shape anchor) gives a shape id and
// a page-relative rectangle in twips; resolving spid -> pib -> blip yields a positioned picture (a
// logo, banner or page watermark). Returns floats in px, on page 1 (cover-page graphics; anchoring to
// a specific later page is not modelled). Metafile and non-picture shapes are skipped.
function parseFloats(wd: Uint8Array, table: Uint8Array, fib: ReturnType<typeof parseFib>): DocFloat[] {
  const spa = fib.fc(FC.plcfspaMom);
  if (spa.lcb < 34 || (spa.lcb - 4) % 30 !== 0) return [];
  const blips = parseBlipStore(table, wd);
  if (blips.every((b) => !b)) return [];
  const spidPib = parseShapePibs(table);
  const dv = new DataView(table.buffer, table.byteOffset + spa.fc, spa.lcb);
  const n = (spa.lcb - 4) / 30; // FSPA is 26 bytes: lcb = 4*(n+1) + 26*n = 30n + 4
  const base = (n + 1) * 4;
  const px = (tw: number) => Math.round(tw / 15);
  const floats: DocFloat[] = [];
  for (let k = 0; k < n; k++) {
    const cp = dv.getUint32(k * 4, true); // the anchor CP (a 0x08 drawn-object char)
    const o = base + k * 26;
    const pib = spidPib.get(dv.getUint32(o, true));
    const img = pib && pib > 0 ? blips[pib] : undefined;
    if (!img) continue;
    const xl = dv.getInt32(o + 4, true), yt = dv.getInt32(o + 8, true);
    const xr = dv.getInt32(o + 12, true), yb = dv.getInt32(o + 16, true);
    if (xr <= xl || yb <= yt) continue;
    // FSPA flags (2-byte bitfield at o+20): wr = bits 5-8 = text-wrap mode. wr 2 = "none"
    // (floats in front of / behind text, no space reserved); every other mode (square,
    // top-and-bottom, tight) reserves the shape's vertical extent in the text flow.
    const wr = (dv.getUint16(o + 20, true) >> 5) & 0xf;
    floats.push({ img, cp, dx: px(xl), dy: px(yt), w: px(xr - xl), h: px(yb - yt), reserve: wr !== 2 });
  }
  return floats;
}

export function docToParts(bytes: Uint8Array): DocParts {
  const cfb = readCfb(bytes);
  const wd = cfb.get("WordDocument");
  if (!wd) throw new Error("no WordDocument stream");
  const fib = parseFib(wd);
  const table = cfb.get(fib.tableStream) ?? cfb.get("1Table") ?? cfb.get("0Table");
  if (!table) throw new Error("no table stream");

  const clx = fib.fc(FC.clx);
  const pieces = parsePieceTable(table, clx.fc, clx.lcb);
  const full = readPieceText(wd, pieces);
  const ccp = fib.ccpText > 0 ? fib.ccpText : full.length;
  const text = full.slice(0, ccp);

  const sttb = fib.fc(FC.sttbfffn);
  const fonts = parseFontTable(table, sttb.fc, sttb.lcb);
  const chpxPlc = fib.fc(FC.plcfBteChpx);
  const papxPlc = fib.fc(FC.plcfBtePapx);
  const charSpans = readFkpSpans(wd, table, chpxPlc.fc, chpxPlc.lcb, false, (g) => decodeCharSprms(g, fonts));
  const paraSpans = readFkpSpans(wd, table, papxPlc.fc, papxPlc.lcb, true, decodeParaSprms);
  const stshFc = fib.fc(FC.stshf);
  const headings = parseStyleHeadings(table, stshFc.fc, stshFc.lcb);
  const styleChps = parseStyleChps(table, stshFc.fc, stshFc.lcb, fonts);
  // The default style (istd 0 = Normal) IS the document default, so only carry a style's DELTA
  // from it: a plain Normal paragraph then adds no formatting, while e.g. a bold Title style does.
  const defaultChp = styleChps.get(0)?.chp ?? {};
  const CHP_KEYS = ["b", "i", "u", "strike", "sizeHalf", "color", "font", "highlight"] as const;
  const deltaChp = (chp: CharProps): CharProps | undefined => {
    const d: CharProps = {};
    let has = false;
    for (const k of CHP_KEYS) if (chp[k] !== undefined && chp[k] !== defaultChp[k]) { (d as Record<string, unknown>)[k] = chp[k]; has = true; }
    return has ? d : undefined;
  };
  for (const sp of paraSpans) {
    if (sp.props.istd != null) {
      sp.props.headingLevel = headings.get(sp.props.istd);
      const style = styleChps.get(sp.props.istd);
      if (style) {
        sp.props.styleChp = deltaChp(style.chp);
        // Inherit list membership / page-break / spacing from the style when the paragraph is silent.
        if (sp.props.ilfo == null) sp.props.ilfo = style.pap.ilfo;
        if (sp.props.ilvl == null) sp.props.ilvl = style.pap.ilvl;
        if (sp.props.pageBreakBefore == null) sp.props.pageBreakBefore = style.pap.pageBreakBefore;
        if (sp.props.spaceBeforeTw == null) sp.props.spaceBeforeTw = style.pap.spaceBeforeTw;
        if (sp.props.spaceAfterTw == null) sp.props.spaceAfterTw = style.pap.spaceAfterTw;
      }
    }
  }
  const cpToFc = makeCpToFc(pieces);
  const dataStream = cfb.get("Data");
  const imageAt = (offset: number) => extractImageHtml(dataStream, offset);

  const sedFc = fib.fc(FC.plcfSed);
  const sections = parseSections(wd, table, sedFc.fc, sedFc.lcb);
  // The last section's geometry is the document page; earlier sections mark a break (with their
  // own geometry) on the paragraph holding their section-break char (at endCp - 1).
  const page = sections.length ? sections[sections.length - 1].geom : undefined;
  const pageJson = page ? secGeomJson(page) : "";
  const sectionBreaks = new Map<number, string>();
  for (let k = 0; k < sections.length - 1; k++) {
    const g = sections[k].geom;
    // Only mark a break when the section's page geometry actually differs from the document's: a
    // same-geometry (typically continuous) section break needs no separate page layout, and forcing
    // one would switch on the multi-section renderer (per-section page cards) unnecessarily.
    if (g && secGeomJson(g) !== pageJson) sectionBreaks.set(sections[k].endCp - 1, secGeomJson(g));
  }

  // Footnote / endnote subdocuments follow the main text in CP order: main, footnote,
  // header, comment, endnote. Their references are 0x02 chars in the main text.
  const refMap = new Map<number, NoteRef>();
  const cmtRefMap = new Map<number, string>();
  const ftnStart = ccp;
  const atnStart = ccp + fib.ccpFtn + fib.ccpHdd;
  const ednStart = ccp + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn;
  const notes = [
    ...parseNotes(full, fib, table, "footnote", ftnStart, fib.ccpFtn, FC.plcffndRef, FC.plcffndTxt, cpToFc, charSpans, refMap, imageAt),
    ...parseNotes(full, fib, table, "endnote", ednStart, fib.ccpEdn, FC.plcfendRef, FC.plcfendTxt, cpToFc, charSpans, refMap, imageAt),
  ];
  const comments = parseComments(full, fib, table, atnStart, cpToFc, charSpans, cmtRefMap);
  const { header, footer } = parseHeaderFooter(full, fib, table, cpToFc, charSpans, imageAt);
  const textboxes = parseTextboxes(full, fib, table, cpToFc, charSpans, imageAt);
  // An MS Equation 3.0 object stores its formula as MTEF in the "Equation Native" stream. We
  // can recover MathML for the common constructs; multiple equations can't be told apart from
  // the flat stream map, so only the first renders as math and the rest as a placeholder.
  const eqNative = cfb.get("Equation Native");
  const equationMathml = eqNative ? mtefToMathml(eqNative) : null;
  const rmFc = fib.fc(FC.sttbfRMark);
  const rmAuthors = parseSttbStrings(table, rmFc.fc, rmFc.lcb);
  const floats = parseFloats(wd, table, fib);
  return {
    body: buildHtml(text, cpToFc, charSpans, paraSpans, refMap, cmtRefMap, imageAt, rmAuthors, sectionBreaks, textboxes, { mathml: equationMathml, present: !!eqNative }, floats),
    page,
    notes: notes.length ? notes : undefined,
    comments: comments.length ? comments : undefined,
    header: header || undefined,
    footer: footer || undefined,
  };
}

export function docToHtml(bytes: Uint8Array): string {
  return docToParts(bytes).body;
}

const PARA = 0x0d;
const LINEBREAK = 0x0b;
const CELL = 0x07;
const SECTION = 0x0c;
const FIELD_BEGIN = 0x13;
const FIELD_SEP = 0x14;
const FIELD_END = 0x15;

// Word ruby (phonetic guide) is an EQ field: EQ \* ... \o\a?(\s\up N(reading),base). Pull the
// reading and base back out into an HTML <ruby>.
function rubyFromEq(instr: string): { base: string; reading: string } | null {
  if (!/\bEQ\b/.test(instr)) return null;
  const m = instr.match(/\\up\s*\d+\s*\(([^)]*)\)\s*,\s*([^)]*)\)/);
  return m ? { reading: m[1]!, base: m[2]! } : null;
}

// The kind of a field from its instruction: a live field the engine fills (PAGE / NUMPAGES),
// a table of contents, or null (its result text is shown as-is).
function fieldKind(instr: string): "PAGE" | "NUMPAGES" | "TOC" | null {
  const t = instr.trim().toUpperCase();
  if (/^PAGE(\s|$|\\)/.test(t)) return "PAGE";
  if (/^NUMPAGES(\s|$|\\)/.test(t)) return "NUMPAGES";
  if (/^TOC(\s|$|\\)/.test(t)) return "TOC";
  return null;
}

function runStyle(p: CharProps | undefined, isHeading = false): string {
  if (!p) return "";
  const s: string[] = [];
  if (p.b && !isHeading) s.push("font-weight:bold");
  if (p.i) s.push("font-style:italic");
  const deco = [p.u ? "underline" : "", p.strike ? "line-through" : ""].filter(Boolean).join(" ");
  if (deco) s.push(`text-decoration:${deco}`);
  if (p.sizeHalf && !isHeading) s.push(`font-size:${p.sizeHalf / 2}pt`);
  if (p.color && p.color !== "#000000") s.push(`color:${p.color}`);
  if (p.font) s.push(`font-family:${/\s/.test(p.font) ? `'${p.font}'` : p.font}`);
  if (p.highlight) s.push(`background-color:${p.highlight}`);
  return s.join(";");
}

function buildHtml(
  text: string,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  paraSpans: Span<ParaProps>[],
  refMap: Map<number, NoteRef> = new Map(),
  cmtRefMap: Map<number, string> = new Map(),
  imageAt: (offset: number) => string | null = () => null,
  rmAuthors: string[] = [],
  sectionBreaks: Map<number, string> = new Map(),
  textboxes: string[] = [],
  equation: { mathml: string | null; present: boolean } = { mathml: null, present: false },
  floats: DocFloat[] = [],
): string {
  const blocks: { tag: string; attr: string; inner: string }[] = [];
  // Floating drawings by their anchor CP; the first image already shown inline (a shape that is also
  // an inline/textbox picture) is skipped so it does not appear twice.
  const floatsByCp = new Map<number, DocFloat[]>();
  for (const f of floats) (floatsByCp.get(f.cp) ?? floatsByCp.set(f.cp, []).get(f.cp)!).push(f);
  const inlineImgs = new Set<string>(); // image srcs already shown inline (e.g. in a textbox)
  for (const tb of textboxes) for (const m of tb.matchAll(/<img [^>]*src="([^"]+)"/g)) inlineImgs.add(m[1]!);
  let paraHasFloat = false; // the current paragraph anchors a float (mark it position:relative)
  let paraFloatReserve = 0; // px of vertical space a wrapping float reserves in this paragraph
  let textboxIdx = 0; // next textbox story to place (they follow the order of 0x08 anchors)
  let eqUsed = false; // the parsed equation has been placed (only one is recoverable)
  const emitEquation = (): string => {
    if (equation.mathml && !eqUsed) { eqUsed = true; return `<span class="docx-eq" data-rdoc-eq contenteditable="false">${equation.mathml}</span>`; }
    return `<span class="docx-eq-raw" contenteditable="false" title="Imported equation (not editable)">⟨equation⟩</span>`;
  };
  const pendingTextboxes: string[] = []; // boxes anchored in the current paragraph, emitted after it
  let runHtml = ""; // accumulated inline HTML of the current paragraph
  let curStyle: string | null = null;
  let curText = "";
  let inInstr = false; // between a field's begin (0x13) and separator (0x14)
  let instr = "";
  let anchorOpen = false; // an <a> from a HYPERLINK field is open
  let fieldClose = ""; // HTML to append at the field end (e.g. "</a>" or "</span>")
  let suppressResult = false; // skip the field result text (a live field span replaces it)
  let pendingToc = false; // inside a TOC field's result (emit an empty toc div; engine rebuilds it)
  let pendingEmbed = false; // inside an EMBED field wrapping an equation OLE object
  let swallowNextPara = false; // drop the para mark that terminates the last TOC entry
  let pendingSecBreak = ""; // SecGeom JSON to attach to the next paragraph (a section boundary)
  let curRev: { tag: "ins" | "del"; key: string } | null = null; // an open tracked-change wrapper
  const revAt = (cp: number): { tag: "ins" | "del"; key: string; open: string } | null => {
    const p = lookup(charSpans, cpToFc(cp));
    if (!p?.rev) return null;
    const author = rmAuthors[p.rmIbst ?? -1] ?? "";
    const date = p.rmDate ?? "";
    const title = author ? `${author}${date ? " – " + date : ""}` : "";
    return { tag: p.rev, key: `${p.rev}|${author}|${date}`, open: `<${p.rev} class="docx-${p.rev}" data-author="${esc(author)}" data-date="${esc(date)}" title="${esc(title)}">` };
  };
  const closeRev = (): void => { if (curRev) { flushRun(); runHtml += `</${curRev.tag}>`; curRev = null; } };
  const headingAt = (cp: number): number => lookup(paraSpans, cpToFc(cp))?.headingLevel ?? 0;
  let curHeading = headingAt(0);
  let tableRows: string[][] = [];
  let rowCells: string[] = [];
  const flushTable = (): void => {
    if (!tableRows.length) return;
    const td = 'style="border:1px solid #999;padding:2px 6px"';
    const rows = tableRows
      .map((cells) => `<tr>${cells.map((c) => `<td ${td}>${c || "<br>"}</td>`).join("")}</tr>`)
      .join("");
    blocks.push({ tag: "table", attr: ' style="border-collapse:collapse"', inner: rows });
    tableRows = [];
  };

  const flushRun = (): void => {
    if (!curText) return;
    const body = esc(curText).replace(/\n/g, "<br>");
    runHtml += curStyle ? `<span style="${curStyle}">${body}</span>` : body;
    curText = "";
  };
  const flushPara = (cp: number): void => {
    flushRun();
    closeRev();
    const pp = lookup(paraSpans, cpToFc(cp));
    const styles: string[] = [];
    const al = pp?.align;
    if (al === 1) styles.push("text-align:center");
    else if (al === 2) styles.push("text-align:right");
    else if (al === 3) styles.push("text-align:justify");
    if (pp?.indentTwips && pp.indentTwips > 0) styles.push(`margin-left:${Math.round(pp.indentTwips / 15)}px`);
    if (pp?.spaceBeforeTw) styles.push(`margin-top:${Math.round(pp.spaceBeforeTw / 15)}px`);
    if (pp?.spaceAfterTw) styles.push(`margin-bottom:${Math.round(pp.spaceAfterTw / 15)}px`);
    if (pp?.inTable) {
      if (pp.ttp) {
        // The row terminator: Word/LibreOffice put the last cell's content on the ttp
        // paragraph itself (no separate row mark), so add it before closing the row. Our
        // own writer emits an empty ttp paragraph, which contributes no extra cell.
        if (runHtml) rowCells.push(runHtml);
        tableRows.push(rowCells);
        rowCells = [];
      } else {
        rowCells.push(runHtml);
      }
      runHtml = "";
      curStyle = null;
      curHeading = headingAt(cp + 1);
      return;
    }
    flushTable();
    // A list-formatted paragraph (sprmPIlfo) renders as a bullet item: prepend the marker the
    // list grouping recognises (a real Word list can be a number too, but bullet is the common
    // run-in list and we do not resolve the LST/LFO tables). Deeper levels add an indent.
    if (pp?.ilfo && pp.ilfo > 0 && curHeading < 1 && !pp.inTable) {
      runHtml = `•\t${runHtml}`;
      if (pp.ilvl) styles.push(`margin-left:${pp.ilvl * 24}px`);
    }
    // A paragraph flagged "page break before" starts a new page: an inline manual-break marker the
    // paginator honours (it breaks before a block that contains one).
    if (pp?.pageBreakBefore) runHtml = `<span class="docx-pagebreak" contenteditable="false" data-docx-pagebreak="manual"></span>${runHtml}`;
    const secAttr = pendingSecBreak ? ` data-rdoc-secbreak="${esc(pendingSecBreak).replace(/"/g, "&quot;")}"` : "";
    pendingSecBreak = "";
    if (paraFloatReserve > 0) styles.push(`min-height:${paraFloatReserve}px`);
    paraFloatReserve = 0;
    const cls = paraHasFloat ? ' class="docx-float-anchor"' : "";
    paraHasFloat = false;
    const attr = cls + (styles.length ? ` style="${styles.join(";")}"` : "") + secAttr;
    const tag = curHeading >= 1 && curHeading <= 6 ? `h${curHeading}` : "p";
    blocks.push({ tag, attr, inner: runHtml || "<br>" });
    // A text box anchored in this paragraph renders as a bordered block right after it (its
    // float position is not modelled; on save it degrades to a plain paragraph).
    for (const box of pendingTextboxes)
      blocks.push({ tag: "div", attr: ' class="docx-textbox"', inner: box });
    pendingTextboxes.length = 0;
    runHtml = "";
    curStyle = null;
    curHeading = headingAt(cp + 1);
  };

  for (let cp = 0; cp < text.length; cp++) {
    const c = text.charCodeAt(cp);
    if (c === FIELD_BEGIN) {
      flushRun();
      inInstr = true;
      instr = "";
      fieldClose = "";
      suppressResult = false;
      continue;
    }
    if (c === FIELD_SEP) {
      inInstr = false;
      const m = instr.match(/HYPERLINK\s+"([^"]+)"|HYPERLINK\s+(\S+)/i);
      const url = m && (m[1] || m[2]);
      const kind = fieldKind(instr);
      flushRun();
      if (url) {
        runHtml += `<a href="${esc(url).replace(/"/g, "&quot;")}">`;
        anchorOpen = true;
      } else if (kind === "PAGE" || kind === "NUMPAGES") {
        // A live field the engine recomputes; drop the stale cached result.
        runHtml += `<span class="docx-field" data-field="${kind}" contenteditable="false">`;
        fieldClose = "</span>";
        suppressResult = true;
      } else if (kind === "TOC") {
        // The engine's decorateFields rebuilds a TOC's rows live from the document's headings,
        // so we emit an empty toc div and drop the cached entries (and their para marks).
        suppressResult = true;
        pendingToc = true;
      } else if (equation.present && /^\s*EMBED\b/i.test(instr)) {
        // An embedded OLE object; when the file carries an Equation Native stream we treat it as
        // the equation and replace the 0x01 placeholder char (suppressed) with recovered MathML.
        suppressResult = true;
        pendingEmbed = true;
      }
      continue;
    }
    if (c === FIELD_END) {
      if (pendingEmbed) {
        flushRun();
        runHtml += emitEquation();
        pendingEmbed = false;
        suppressResult = false;
        inInstr = false;
        continue;
      }
      if (pendingToc) {
        // Discard the (empty) first-TOC-paragraph accumulation and emit the toc div as its own
        // block; the last entry's own paragraph mark is swallowed so it adds no empty <p>.
        curText = ""; runHtml = ""; curStyle = null;
        blocks.push({ tag: "div", attr: ' class="docx-field-toc"', inner: "" });
        pendingToc = false;
        suppressResult = false;
        inInstr = false;
        swallowNextPara = true;
        continue;
      }
      flushRun();
      if (anchorOpen) { runHtml += "</a>"; anchorOpen = false; }
      else if (inInstr) {
        // A field with no separator (e.g. ruby, stored as an EQ field with no result).
        const ruby = rubyFromEq(instr);
        if (ruby) runHtml += `<ruby>${esc(ruby.base)}<rt>${esc(ruby.reading)}</rt></ruby>`;
      }
      runHtml += fieldClose;
      fieldClose = "";
      suppressResult = false;
      inInstr = false;
      continue;
    }
    if (inInstr) {
      instr += text[cp];
      continue;
    }
    if (suppressResult) continue;
    if (c === PARA || c === CELL) {
      if (swallowNextPara) { swallowNextPara = false; curHeading = headingAt(cp + 1); continue; }
      flushPara(cp);
      continue;
    }
    if (c === SECTION) {
      const secGeom = sectionBreaks.get(cp);
      if (secGeom) {
        // A section break: it also ends the paragraph, which carries the section's geometry.
        pendingSecBreak = secGeom;
        flushPara(cp);
      } else {
        // A plain manual page break: an inline marker richdoc renders in the pageless view.
        flushRun();
        runHtml += '<span class="docx-pagebreak" contenteditable="false" data-docx-pagebreak="manual"></span>';
      }
      continue;
    }
    if (c === LINEBREAK) {
      curText += "\n";
      continue;
    }
    if (c === 0x1e) {
      curText += "‑";
      continue;
    }
    if (c === 0x02) {
      // A footnote / endnote reference mark: emit the engine's inline reference (the visible
      // number is filled at render time). Non-note 0x02 chars are dropped.
      const r = refMap.get(cp);
      if (r) {
        flushRun();
        runHtml += `<sup class="docx-fnref" data-fn-id="${r.id}" data-fn-kind="${r.kind}" contenteditable="false"></sup>`;
      }
      continue;
    }
    if (c === 0x05) {
      // A comment (annotation) reference mark: emit the engine's inline comment reference.
      const cid = cmtRefMap.get(cp);
      if (cid) {
        flushRun();
        runHtml += `<span class="docx-comment-ref" data-comment-id="${cid}" contenteditable="false">\u{1F4AC}</span>`;
      }
      continue;
    }
    if (c === 0x01) {
      // A picture placeholder: its CHPX carries the Data-stream offset of the image.
      const off = lookup(charSpans, cpToFc(cp))?.picOffset;
      if (off !== undefined) {
        const img = imageAt(off);
        if (img) { flushRun(); runHtml += img; }
      }
      continue;
    }
    if (c === 0x08) {
      // A drawn-object anchor. A floating picture anchored here (logo/banner/watermark) is placed
      // absolutely inside this paragraph, offset from it and drawn behind the text; the paragraph is
      // marked position:relative. A shape that is also shown inline (a textbox picture) is skipped.
      let injected = false;
      for (const f of floatsByCp.get(cp) ?? []) {
        if (inlineImgs.has(f.img)) continue; // already shown inline (its textbox) - don't double it
        flushRun();
        runHtml += `<img class="docx-float" src="${f.img}" alt="" contenteditable="false" style="left:${f.dx}px;top:${f.dy}px;width:${f.w}px;height:${f.h}px">`;
        paraHasFloat = true;
        injected = true;
        // A wrapping float (top-and-bottom / square) reserves its vertical extent: make the anchor
        // paragraph at least tall enough to contain it so following content flows below, not over it.
        if (f.reserve) paraFloatReserve = Math.max(paraFloatReserve, f.dy + f.h);
      }
      // No floating image placed here: pair the anchor with the next textbox story (placed after
      // this paragraph). A deduped float still falls through to its textbox, which shows the image.
      if (!injected && textboxIdx < textboxes.length) pendingTextboxes.push(textboxes[textboxIdx++]);
      continue;
    }
    if (c === 0x1f || c === 0x00 || c === 0x03) continue;
    if (c === 0x09) {
      curText += "\t";
      continue;
    }
    // Open / close the tracked-change wrapper as the revision under the caret changes.
    const rev = revAt(cp);
    if ((rev?.key ?? "") !== (curRev?.key ?? "")) {
      flushRun();
      if (curRev) runHtml += `</${curRev.tag}>`;
      if (rev) { runHtml += rev.open; curRev = { tag: rev.tag, key: rev.key }; } else curRev = null;
    }
    // Merge the paragraph's named-style character formatting (base) under the run's own (override),
    // so a run inherits e.g. a bold Title style unless it sets those properties itself.
    const runChp = lookup(charSpans, cpToFc(cp));
    const styleChp = lookup(paraSpans, cpToFc(cp))?.styleChp;
    const style = runStyle(styleChp ? { ...styleChp, ...runChp } : runChp, curHeading > 0) || null;
    if (style !== curStyle) {
      flushRun();
      curStyle = style;
    }
    curText += text[cp];
  }
  if (curText || runHtml || pendingTextboxes.length) flushPara(text.length);
  flushTable();
  while (blocks.length > 1 && blocks[blocks.length - 1].inner === "<br>" && blocks[blocks.length - 1].tag === "p")
    blocks.pop();
  // Any textbox whose anchor we could not place (e.g. anchored outside the main story) is
  // appended so its content is never silently dropped.
  for (; textboxIdx < textboxes.length; textboxIdx++)
    blocks.push({ tag: "div", attr: ' class="docx-textbox"', inner: textboxes[textboxIdx] });
  return blocksToHtml(blocks) || "<p><br></p>";
}

// Detect the leading list marker textutil/our writer emits ("\t•\t" or "\tN.\t"),
// returning the list kind and the content with the marker stripped.
const BULLET_RE = /^\t?[•▪◦‣·]\t/;
const NUMBER_RE = /^\t?\d+[.)]\t/;
function listMarker(inner: string): { kind: "ul" | "ol"; rest: string } | null {
  // The marker is plain text at the very start (an unstyled run), so it precedes any span.
  const plain = inner.replace(/^(\t?)([•▪◦‣·]|\d+[.)])(\t)/, "");
  if (BULLET_RE.test(inner)) return { kind: "ul", rest: plain };
  if (NUMBER_RE.test(inner)) return { kind: "ol", rest: plain };
  return null;
}

function blocksToHtml(blocks: { tag: string; attr: string; inner: string }[]): string {
  const out: string[] = [];
  let i = 0;
  while (i < blocks.length) {
    const m = blocks[i].tag === "p" ? listMarker(blocks[i].inner) : null;
    if (m) {
      const kind = m.kind;
      const items: string[] = [];
      while (i < blocks.length && blocks[i].tag === "p") {
        const mm = listMarker(blocks[i].inner);
        if (!mm || mm.kind !== kind) break;
        items.push(`<li>${mm.rest}</li>`);
        i++;
      }
      out.push(`<${kind}>${items.join("")}</${kind}>`);
    } else {
      const b = blocks[i];
      if (b.tag === "table") out.push(`<table${b.attr}>${b.inner}</table>`);
      else out.push(`<${b.tag}${b.attr}>${b.inner}</${b.tag}>`);
      i++;
    }
  }
  return out.join("");
}
