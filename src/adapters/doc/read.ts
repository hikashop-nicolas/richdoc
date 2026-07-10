import { readCfb } from "./cfb";
import { parseFib, parsePieceTable, readPieceText, FC, type Piece } from "./fib";
import { bytesToBase64 } from "../../core/util";
import type { CommentThread, Note, PageGeometry } from "../../core/types";

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
  inTable?: boolean; // sprmPFInTable
  ttp?: boolean; // sprmPFTtp (table-terminating / row-end paragraph)
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

// Parse the first section's properties (SEPX) into richdoc page geometry.
function parseSection(wd: Uint8Array, table: Uint8Array, fcSed: number, lcbSed: number): PageGeometry | undefined {
  if (lcbSed < 16) return undefined;
  const nSed = (lcbSed - 4) / 16; // PLCF: 4*(n+1) CPs + 12*n Sed = 16n+4
  const sedBase = 4 * (nSed + 1);
  const sdv = new DataView(table.buffer, table.byteOffset + fcSed, lcbSed);
  const fcSepx = sdv.getUint32(sedBase + 2, true); // Sed: fn(2) then fcSepx(4)
  if (fcSepx === 0xffffffff || fcSepx + 2 > wd.length) return undefined;
  const cb = wd[fcSepx] | (wd[fcSepx + 1] << 8);
  const g = wd.subarray(fcSepx + 2, fcSepx + 2 + cb);
  const gdv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  const px = (tw: number) => Math.round(tw / 15);
  const page: PageGeometry = { widthPx: 816, heightPx: 1056, margin: { top: 96, right: 96, bottom: 96, left: 96 } };
  let any = false;
  let i = 0;
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
    else if (op === 0xb021) { page.margin.left = px(u16()); any = true; }
    else if (op === 0xb022) page.margin.right = px(u16());
    else if (op === 0x9023) page.margin.top = px(gdv.getInt16(start, true));
    else if (op === 0x9024) page.margin.bottom = px(gdv.getInt16(start, true));
    else if (op === 0x500b) { page.columns = u16() + 1; any = true; }
    else if (op === 0x900c) page.columnGapPx = px(u16());
    else if (op === 0x5453) { if (u16() !== 0) { page.vertical = true; any = true; } }
    else if (op === 0x5228) { if (u16() !== 0) { page.rtl = true; any = true; } }
  }
  return any ? page : undefined;
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
  for (let cp = start; cp < end && cp < full.length; cp++) {
    const c = full.charCodeAt(cp);
    if (!leadStripped) {
      if (c === 0x02) continue;
      if (c === 0x09) { leadStripped = true; continue; }
      leadStripped = true;
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
    notes.push({ id, kind, html: renderNoteBody(full, subStart + txtCps[i], subStart + txtCps[i + 1], cpToFc, charSpans) });
  }
  return notes;
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
): { header: string; footer: string } {
  const empty = { header: "", footer: "" };
  if (fib.ccpHdd <= 0) return empty;
  const hdd = fib.fc(FC.plcfHdd);
  if (hdd.lcb < 44) return empty; // need at least the 6 separators + section-0 header/footer
  const dv = new DataView(table.buffer, table.byteOffset + hdd.fc, hdd.lcb);
  const n = hdd.lcb / 4;
  const cp = (k: number) => (k < n ? dv.getUint32(k * 4, true) : dv.getUint32((n - 1) * 4, true));
  const base = fib.ccpText + fib.ccpFtn;
  const story = (i: number) => (cp(i + 1) > cp(i) ? renderNoteBody(full, base + cp(i), base + cp(i + 1), cpToFc, charSpans) : "");
  const clean = (h: string) => (h === "<p><br></p>" ? "" : h);
  return { header: clean(story(7)), footer: clean(story(9)) };
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
  for (const sp of paraSpans) if (sp.props.istd != null) sp.props.headingLevel = headings.get(sp.props.istd);
  const cpToFc = makeCpToFc(pieces);

  const sedFc = fib.fc(FC.plcfSed);
  const page = parseSection(wd, table, sedFc.fc, sedFc.lcb);

  // Footnote / endnote subdocuments follow the main text in CP order: main, footnote,
  // header, comment, endnote. Their references are 0x02 chars in the main text.
  const refMap = new Map<number, NoteRef>();
  const cmtRefMap = new Map<number, string>();
  const ftnStart = ccp;
  const atnStart = ccp + fib.ccpFtn + fib.ccpHdd;
  const ednStart = ccp + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn;
  const notes = [
    ...parseNotes(full, fib, table, "footnote", ftnStart, fib.ccpFtn, FC.plcffndRef, FC.plcffndTxt, cpToFc, charSpans, refMap),
    ...parseNotes(full, fib, table, "endnote", ednStart, fib.ccpEdn, FC.plcfendRef, FC.plcfendTxt, cpToFc, charSpans, refMap),
  ];
  const comments = parseComments(full, fib, table, atnStart, cpToFc, charSpans, cmtRefMap);
  const { header, footer } = parseHeaderFooter(full, fib, table, cpToFc, charSpans);
  const dataStream = cfb.get("Data");
  const imageAt = (offset: number) => extractImageHtml(dataStream, offset);
  return {
    body: buildHtml(text, cpToFc, charSpans, paraSpans, refMap, cmtRefMap, imageAt),
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
): string {
  const blocks: { tag: string; attr: string; inner: string }[] = [];
  let runHtml = ""; // accumulated inline HTML of the current paragraph
  let curStyle: string | null = null;
  let curText = "";
  let inInstr = false; // between a field's begin (0x13) and separator (0x14)
  let instr = "";
  let anchorOpen = false; // an <a> from a HYPERLINK field is open
  let fieldClose = ""; // HTML to append at the field end (e.g. "</a>" or "</span>")
  let suppressResult = false; // skip the field result text (a live field span replaces it)
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
    const pp = lookup(paraSpans, cpToFc(cp));
    const styles: string[] = [];
    const al = pp?.align;
    if (al === 1) styles.push("text-align:center");
    else if (al === 2) styles.push("text-align:right");
    else if (al === 3) styles.push("text-align:justify");
    if (pp?.indentTwips && pp.indentTwips > 0) styles.push(`margin-left:${Math.round(pp.indentTwips / 15)}px`);
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
    const attr = styles.length ? ` style="${styles.join(";")}"` : "";
    const tag = curHeading >= 1 && curHeading <= 6 ? `h${curHeading}` : "p";
    blocks.push({ tag, attr, inner: runHtml || "<br>" });
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
      }
      continue;
    }
    if (c === FIELD_END) {
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
      flushPara(cp);
      continue;
    }
    if (c === SECTION) {
      // manual page break: an inline marker richdoc renders in the pageless view
      flushRun();
      runHtml +=
        '<span class="docx-pagebreak" contenteditable="false" data-docx-pagebreak="manual"></span>';
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
    if (c === 0x1f || c === 0x00 || c === 0x03 || c === 0x08) continue;
    if (c === 0x09) {
      curText += "\t";
      continue;
    }
    const style = runStyle(lookup(charSpans, cpToFc(cp)), curHeading > 0) || null;
    if (style !== curStyle) {
      flushRun();
      curStyle = style;
    }
    curText += text[cp];
  }
  if (curText || runHtml) flushPara(text.length);
  flushTable();
  while (blocks.length > 1 && blocks[blocks.length - 1].inner === "<br>" && blocks[blocks.length - 1].tag === "p")
    blocks.pop();
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
