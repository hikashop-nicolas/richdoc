import { readCfb } from "./cfb";
import { parseFib, parsePieceTable, readPieceText, FC, type Piece } from "./fib";

// Read half of the .doc adapter: bytes -> HTML in richdoc's vocabulary. It extracts the
// text (piece table) and the character/paragraph formatting (CHPX/PAPX formatted-disk
// pages), then emits paragraphs with inline runs the engine edits like docx/odt.

export interface DocParts {
  body: string;
}

interface CharProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  sizeHalf?: number;
  color?: string; // #rrggbb
}
interface ParaProps {
  align?: number; // 0..3
  indentTwips?: number;
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Decode a CHPX grpprl (a run of sprms) into character properties.
function decodeCharSprms(g: Uint8Array): CharProps {
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
      case 0x6870: // sprmCCv: RR GG BB shade
        p.color = "#" + [v[0], v[1], v[2]].map((x) => x.toString(16).padStart(2, "0")).join("");
        break;
    }
  }
  return p;
}

function decodeParaSprms(g: Uint8Array): ParaProps {
  const dv = new DataView(g.buffer, g.byteOffset, g.byteLength);
  const p: ParaProps = {};
  let i = 0;
  while (i + 2 <= g.length) {
    const op = dv.getUint16(i, true);
    i += 2;
    const spra = (op >> 13) & 7;
    const len = spra === 6 ? g[i++] : [1, 1, 2, 4, 2, 2, 0, 3][spra];
    const v = g.subarray(i, i + len);
    i += len;
    if (op === 0x2403 || op === 0x2461) p.align = v[0];
    else if (op === 0x840f) p.indentTwips = dv.getInt16(i - len, true);
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
          props = decode(fkp.subarray(start + 2, start + grpprlLen)); // skip istd
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

  const chpxPlc = fib.fc(FC.plcfBteChpx);
  const papxPlc = fib.fc(FC.plcfBtePapx);
  const charSpans = readFkpSpans(wd, table, chpxPlc.fc, chpxPlc.lcb, false, decodeCharSprms);
  const paraSpans = readFkpSpans(wd, table, papxPlc.fc, papxPlc.lcb, true, decodeParaSprms);
  const cpToFc = makeCpToFc(pieces);

  return { body: buildHtml(text, cpToFc, charSpans, paraSpans) };
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

function runStyle(p: CharProps | undefined): string {
  if (!p) return "";
  const s: string[] = [];
  if (p.b) s.push("font-weight:bold");
  if (p.i) s.push("font-style:italic");
  const deco = [p.u ? "underline" : "", p.strike ? "line-through" : ""].filter(Boolean).join(" ");
  if (deco) s.push(`text-decoration:${deco}`);
  if (p.sizeHalf) s.push(`font-size:${p.sizeHalf / 2}pt`);
  if (p.color && p.color !== "#000000") s.push(`color:${p.color}`);
  return s.join(";");
}

function buildHtml(
  text: string,
  cpToFc: (cp: number) => number,
  charSpans: Span<CharProps>[],
  paraSpans: Span<ParaProps>[],
): string {
  const out: string[] = [];
  let runHtml = ""; // accumulated inline HTML of the current paragraph
  let curStyle: string | null = null;
  let curText = "";
  let fieldDepth = 0;

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
    const attr = styles.length ? ` style="${styles.join(";")}"` : "";
    out.push(runHtml ? `<p${attr}>${runHtml}</p>` : "<p><br></p>");
    runHtml = "";
    curStyle = null;
  };

  for (let cp = 0; cp < text.length; cp++) {
    const c = text.charCodeAt(cp);
    if (c === FIELD_BEGIN) {
      flushRun();
      fieldDepth++;
      continue;
    }
    if (c === FIELD_SEP) {
      if (fieldDepth > 0) fieldDepth--;
      continue;
    }
    if (c === FIELD_END) continue;
    if (fieldDepth > 0) continue;
    if (c === PARA || c === CELL || c === SECTION) {
      flushPara(cp);
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
    if (c === 0x1f || c === 0x00 || c === 0x01 || c === 0x02 || c === 0x03 || c === 0x08) continue;
    if (c === 0x09) {
      curText += "\t";
      continue;
    }
    const style = runStyle(lookup(charSpans, cpToFc(cp))) || null;
    if (style !== curStyle) {
      flushRun();
      curStyle = style;
    }
    curText += text[cp];
  }
  if (curText || runHtml) flushPara(text.length);
  while (out.length > 1 && out[out.length - 1] === "<p><br></p>") out.pop();
  return out.join("") || "<p><br></p>";
}
