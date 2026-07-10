// The Word File Information Block (FIB) at the start of the WordDocument stream, and the
// piece table (Clx) that maps logical character positions to byte offsets. We parse only
// the fields the reader needs; the writer builds a matching FIB in write.ts.

// Indices into FibRgFcLcb97 (each entry is an fc/lcb pair of two uint32s).
export const FC = {
  stshf: 1,
  plcffndRef: 2, // footnote reference CPs (main doc)
  plcffndTxt: 3, // footnote text spans (footnote subdocument)
  plcfandRef: 4, // comment (annotation) reference CPs + ATRD (main doc)
  plcfandTxt: 5, // comment text spans (annotation subdocument)
  plcfSed: 6,
  plcfHdd: 11,
  plcfBteChpx: 12,
  plcfBtePapx: 13,
  sttbfffn: 15,
  plcffldMom: 16, // field boundaries in the main document (CPs + FLD per field char)
  grpXstAtnOwners: 36, // comment author names (group of Xst), indexed by ATRD.ibst
  clx: 33,
  plcfendRef: 46, // endnote reference CPs (main doc)
  plcfendTxt: 47, // endnote text spans (endnote subdocument)
} as const;

export interface Fib {
  /** "0Table" or "1Table", per fWhichTblStm. */
  tableStream: "0Table" | "1Table";
  ccpText: number;
  /** Char counts of the appended subdocuments, in CP order after the main text. */
  ccpFtn: number; // footnote subdocument
  ccpHdd: number; // header/footer subdocument
  ccpAtn: number; // comment (annotation) subdocument
  ccpEdn: number; // endnote subdocument
  /** Offset of the fc/lcb blob within the WordDocument stream. */
  blobOffset: number;
  wd: Uint8Array;
  fc(index: number): { fc: number; lcb: number };
}

export function parseFib(wd: Uint8Array): Fib {
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  const flags1 = dv.getUint16(10, true);
  const tableStream = ((flags1 >> 9) & 1) === 1 ? "1Table" : "0Table";
  const csw = dv.getUint16(32, true);
  const rgLwOff = 34 + csw * 2 + 2; // fibRgW (csw words) then cslw (2 bytes)
  const cslw = dv.getUint16(34 + csw * 2, true);
  const lw = (i: number) => dv.getInt32(rgLwOff + i * 4, true);
  const ccpText = lw(3); // fibRgLw index 3
  const cbRgFcLcbOff = rgLwOff + cslw * 4;
  const blobOffset = cbRgFcLcbOff + 2; // after the cbRgFcLcb count
  return {
    tableStream,
    ccpText,
    ccpFtn: lw(4),
    ccpHdd: lw(5),
    ccpAtn: lw(7),
    ccpEdn: lw(8),
    blobOffset,
    wd,
    fc(index: number) {
      return {
        fc: dv.getUint32(blobOffset + index * 8, true),
        lcb: dv.getUint32(blobOffset + index * 8 + 4, true),
      };
    },
  };
}

/** One piece of the piece table: a run of characters at a byte offset in WordDocument. */
export interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number; // byte offset into WordDocument of this piece's text
  unicode: boolean; // true = 2-byte UTF-16LE, false = 1-byte Windows-1252
}

/** Parse the Clx (piece table) from the table stream at fcClx/lcbClx. */
export function parsePieceTable(table: Uint8Array, fcClx: number, lcbClx: number): Piece[] {
  const clx = table.subarray(fcClx, fcClx + lcbClx);
  const dv = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);
  let i = 0;
  // Skip any Prc entries (each: 0x01, cbGrpprl uint16, grpprl).
  while (i < clx.length && clx[i] === 0x01) {
    const cb = dv.getUint16(i + 1, true);
    i += 3 + cb;
  }
  // Some real-world files (e.g. a doc that is mostly an embedded OLE object with a couple of
  // text chars) carry a Clx we can't recognise a Pcdt in. Rather than fail the whole open,
  // return no pieces: the reader then treats the body as empty instead of throwing.
  if (clx[i] !== 0x02) return [];
  const lcbPlcPcd = dv.getUint32(i + 1, true);
  i += 5;
  const plc = clx.subarray(i, i + lcbPlcPcd);
  const pdv = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
  const n = (lcbPlcPcd - 4) / 12; // (n+1) CPs of 4 bytes + n PCDs of 8 bytes
  const cps: number[] = [];
  for (let k = 0; k <= n; k++) cps.push(pdv.getUint32(k * 4, true));
  const pcdBase = (n + 1) * 4;
  const pieces: Piece[] = [];
  for (let k = 0; k < n; k++) {
    const fcVal = pdv.getUint32(pcdBase + k * 8 + 2, true);
    const compressed = ((fcVal >>> 30) & 1) === 1;
    const fc = compressed ? (fcVal & 0x3fffffff) >>> 1 : fcVal & 0x3fffffff;
    pieces.push({ cpStart: cps[k], cpEnd: cps[k + 1], fc, unicode: !compressed });
  }
  return pieces;
}

/** Reconstruct the document's full text (one char per CP) from the pieces. */
export function readPieceText(wd: Uint8Array, pieces: Piece[]): string {
  let out = "";
  for (const p of pieces) {
    const count = p.cpEnd - p.cpStart;
    if (p.unicode) {
      for (let k = 0; k < count; k++) {
        const off = p.fc + k * 2;
        out += String.fromCharCode(wd[off] | (wd[off + 1] << 8));
      }
    } else {
      for (let k = 0; k < count; k++) out += cp1252(wd[p.fc + k]);
    }
  }
  return out;
}

// Windows-1252 high range (0x80-0x9F) differs from Latin-1; map the common ones.
const CP1252_HIGH: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020,
  0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152,
  0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022,
  0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a,
  0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};
function cp1252(b: number): string {
  return String.fromCharCode(b >= 0x80 && b <= 0x9f ? (CP1252_HIGH[b] ?? b) : b);
}
