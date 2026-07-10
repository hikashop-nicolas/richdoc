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
  font?: string; // font family name
  highlight?: string; // background colour #rrggbb
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
): string {
  const blocks: { tag: string; attr: string; inner: string }[] = [];
  let runHtml = ""; // accumulated inline HTML of the current paragraph
  let curStyle: string | null = null;
  let curText = "";
  let inInstr = false; // between a field's begin (0x13) and separator (0x14)
  let instr = "";
  let anchorOpen = false; // an <a> from a HYPERLINK field is open
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
      continue;
    }
    if (c === FIELD_SEP) {
      inInstr = false;
      const m = instr.match(/HYPERLINK\s+"([^"]+)"|HYPERLINK\s+(\S+)/i);
      const url = m && (m[1] || m[2]);
      if (url) {
        flushRun();
        runHtml += `<a href="${esc(url).replace(/"/g, "&quot;")}">`;
        anchorOpen = true;
      }
      continue;
    }
    if (c === FIELD_END) {
      if (anchorOpen) {
        flushRun();
        runHtml += "</a>";
        anchorOpen = false;
      }
      continue;
    }
    if (inInstr) {
      instr += text[cp];
      continue;
    }
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
    if (c === 0x1f || c === 0x00 || c === 0x01 || c === 0x02 || c === 0x03 || c === 0x08) continue;
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
