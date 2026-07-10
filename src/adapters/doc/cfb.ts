// Minimal, self-contained MS-CFB (Compound File Binary / OLE2) reader and writer, enough
// for Word 97-2003 `.doc` files. No dependency (richdoc keeps only fflate + temml).
//
// Reader: parses any well-formed compound file (regular + mini streams) into a map of
// stream name -> bytes. Writer: builds a compound file from named streams, keeping every
// stream in regular 512-byte sectors (padded to >= the 4096 mini cutoff) so no mini FAT is
// needed, which mirrors what macOS `textutil` emits and what Word/LibreOffice accept.

const SECTOR = 512;
const MINI_CUTOFF = 4096;
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** True if the bytes start with the CFB / OLE2 magic (so it may be a legacy .doc). */
export function isCfb(bytes: Uint8Array): boolean {
  return SIGNATURE.every((b, i) => bytes[i] === b);
}

// ---------------------------------------------------------------------------
// Reader
// ---------------------------------------------------------------------------

export function readCfb(bytes: Uint8Array): Map<string, Uint8Array> {
  if (!isCfb(bytes)) throw new Error("not a compound file");
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 1 << dv.getUint16(30, true);
  const miniSize = 1 << dv.getUint16(32, true);
  const numFatSec = dv.getUint32(44, true);
  const firstDirSec = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true);
  const firstMiniFat = dv.getUint32(60, true);
  const firstDifat = dv.getUint32(68, true);
  const off = (s: number) => (s + 1) * sectorSize;

  // DIFAT: 109 entries in the header, then any DIFAT-sector chain.
  const difat: number[] = [];
  for (let i = 0; i < 109; i++) {
    const s = dv.getUint32(76 + i * 4, true);
    if (s !== FREESECT) difat.push(s);
  }
  let ds = firstDifat;
  const perDifat = sectorSize / 4 - 1;
  while (ds < FATSECT && off(ds) < bytes.length) {
    for (let i = 0; i < perDifat; i++) {
      const s = dv.getUint32(off(ds) + i * 4, true);
      if (s !== FREESECT) difat.push(s);
    }
    ds = dv.getUint32(off(ds) + perDifat * 4, true);
  }

  // FAT.
  const fat: number[] = [];
  for (const s of difat.slice(0, numFatSec)) {
    for (let i = 0; i < sectorSize / 4; i++) fat.push(dv.getUint32(off(s) + i * 4, true));
  }
  const chain = (start: number): number[] => {
    const out: number[] = [];
    let s = start;
    while (s < FATSECT && s < fat.length) {
      out.push(s);
      s = fat[s];
      if (out.length > 1 << 22) break;
    }
    return out;
  };
  const readRegular = (start: number, size: number): Uint8Array => {
    const secs = chain(start);
    const out = new Uint8Array(secs.length * sectorSize);
    let o = 0;
    for (const s of secs) {
      out.set(bytes.subarray(off(s), off(s) + sectorSize), o);
      o += sectorSize;
    }
    return out.subarray(0, size);
  };

  // Directory entries.
  const dir = readRegular(firstDirSec, chain(firstDirSec).length * sectorSize);
  const ddv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  interface Entry {
    name: string;
    type: number;
    start: number;
    size: number;
  }
  const entries: Entry[] = [];
  let rootStart = 0;
  let rootSize = 0;
  for (let i = 0; i + 128 <= dir.length; i += 128) {
    const type = dir[i + 66];
    if (type === 0) continue;
    const nlen = ddv.getUint16(i + 64, true);
    let name = "";
    for (let j = 0; j < Math.max(0, nlen - 2); j += 2) name += String.fromCharCode(dir[i + j] | (dir[i + j + 1] << 8));
    const start = ddv.getUint32(i + 116, true);
    const size = ddv.getUint32(i + 120, true);
    if (type === 5) {
      rootStart = start;
      rootSize = size;
    } else if (type === 2) {
      entries.push({ name, type, start, size });
    }
  }

  // Mini stream (owned by the root entry) + mini FAT.
  const miniStream = readRegular(rootStart, rootSize);
  const miniFat: number[] = [];
  {
    let s = firstMiniFat;
    while (s < FATSECT && s < fat.length) {
      for (let i = 0; i < sectorSize / 4; i++) miniFat.push(dv.getUint32(off(s) + i * 4, true));
      s = fat[s];
    }
  }
  const readMini = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(Math.ceil(size / miniSize) * miniSize || miniSize);
    let o = 0;
    let s = start;
    while (s < FATSECT && s * miniSize < miniStream.length && o < out.length) {
      out.set(miniStream.subarray(s * miniSize, (s + 1) * miniSize), o);
      o += miniSize;
      s = miniFat[s];
    }
    return out.subarray(0, size);
  };

  const result = new Map<string, Uint8Array>();
  for (const e of entries) {
    result.set(e.name, e.size < miniCutoff ? readMini(e.start, e.size) : readRegular(e.start, e.size));
  }
  return result;
}

// ---------------------------------------------------------------------------
// Writer
// ---------------------------------------------------------------------------

// CFB directory ordering: by UTF-16 name length, then by uppercased code units.
function nameLess(a: string, b: string): boolean {
  if (a.length !== b.length) return a.length < b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? true : ua > ub ? false : a < b;
}

export interface CfbInput {
  name: string;
  data: Uint8Array;
}

/** Build a compound file from named streams (all placed in regular sectors). */
export function writeCfb(streams: CfbInput[]): Uint8Array {
  // Pad each stream to a whole number of sectors, and to at least the mini cutoff so it
  // never needs the mini stream. Assign sector chains in order starting at sector 0.
  interface Placed {
    name: string;
    size: number; // real (unpadded) size, written to the directory entry
    startSector: number;
    sectorCount: number;
    padded: Uint8Array;
  }
  const placed: Placed[] = [];
  let sector = 0;
  for (const s of streams) {
    const paddedLen = Math.max(MINI_CUTOFF, Math.ceil(s.data.length / SECTOR) * SECTOR) || MINI_CUTOFF;
    const padded = new Uint8Array(paddedLen);
    padded.set(s.data);
    const count = paddedLen / SECTOR;
    placed.push({ name: s.name, size: paddedLen, startSector: sector, sectorCount: count, padded });
    sector += count;
  }
  const dataSectors = sector;

  // Directory: Root Entry + one per stream. Count directory sectors (4 entries per sector).
  const dirEntryCount = placed.length + 1;
  const dirSectorCount = Math.ceil(dirEntryCount / (SECTOR / 128));
  const dirStartSector = dataSectors;

  // One FAT sector covers 128 sector entries (512/4). Iterate until it stabilises.
  let fatSectorCount = 1;
  for (;;) {
    const total = dataSectors + dirSectorCount + fatSectorCount;
    const need = Math.ceil(total / (SECTOR / 4));
    if (need <= fatSectorCount) break;
    fatSectorCount = need;
  }
  const fatStartSector = dirStartSector + dirSectorCount;
  const totalSectors = dataSectors + dirSectorCount + fatSectorCount;

  // FAT array.
  const fat = new Uint32Array(fatSectorCount * (SECTOR / 4)).fill(FREESECT);
  const linkChain = (start: number, count: number): void => {
    for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? ENDOFCHAIN : start + i + 1;
  };
  for (const p of placed) linkChain(p.startSector, p.sectorCount);
  linkChain(dirStartSector, dirSectorCount);
  for (let i = 0; i < fatSectorCount; i++) fat[fatStartSector + i] = FATSECT;

  // Directory entries. Build a balanced BST honouring CFB name ordering; entry 0 is Root.
  const dirBuf = new Uint8Array(dirSectorCount * SECTOR);
  const ddv = new DataView(dirBuf.buffer);
  const NOSTREAM = 0xffffffff;
  // Stream entry indices are 1..n, ordered as given; sort those indices by CFB name order.
  const order = placed.map((_, i) => i + 1).sort((a, b) => (nameLess(placed[a - 1].name, placed[b - 1].name) ? -1 : 1));
  const left = new Array(dirEntryCount).fill(NOSTREAM);
  const right = new Array(dirEntryCount).fill(NOSTREAM);
  const buildTree = (lo: number, hi: number): number => {
    if (lo > hi) return NOSTREAM;
    const mid = (lo + hi) >> 1;
    const idx = order[mid];
    left[idx] = buildTree(lo, mid - 1);
    right[idx] = buildTree(mid + 1, hi);
    return idx;
  };
  const rootChild = buildTree(0, order.length - 1);

  const writeEntry = (i: number, name: string, type: number, start: number, size: number, l: number, r: number, child: number): void => {
    const base = i * 128;
    for (let j = 0; j < name.length && j < 31; j++) {
      dirBuf[base + j * 2] = name.charCodeAt(j) & 0xff;
      dirBuf[base + j * 2 + 1] = (name.charCodeAt(j) >> 8) & 0xff;
    }
    ddv.setUint16(base + 64, (Math.min(name.length, 31) + 1) * 2, true); // name length incl. null
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = 1; // colour: black
    ddv.setUint32(base + 68, l, true);
    ddv.setUint32(base + 72, r, true);
    ddv.setUint32(base + 76, child, true);
    ddv.setUint32(base + 116, start, true);
    ddv.setUint32(base + 120, size, true);
    // size high dword (offset 124) stays 0.
  };
  // Root Entry: type 5, no mini stream (start ENDOFCHAIN, size 0), child = tree root.
  writeEntry(0, "Root Entry", 5, ENDOFCHAIN, 0, NOSTREAM, NOSTREAM, rootChild);
  for (let k = 0; k < placed.length; k++) {
    const p = placed[k];
    writeEntry(k + 1, p.name, 2, p.startSector, p.size, left[k + 1], right[k + 1], NOSTREAM);
  }
  // Unused directory entries: type 0, all pointers NOSTREAM.
  for (let i = dirEntryCount; i < dirSectorCount * (SECTOR / 128); i++) {
    ddv.setUint32(i * 128 + 68, NOSTREAM, true);
    ddv.setUint32(i * 128 + 72, NOSTREAM, true);
    ddv.setUint32(i * 128 + 76, NOSTREAM, true);
  }

  // Assemble: header + all sectors.
  const out = new Uint8Array((totalSectors + 1) * SECTOR);
  const odv = new DataView(out.buffer);
  out.set(SIGNATURE, 0);
  odv.setUint16(24, 0x003e, true); // minor version
  odv.setUint16(26, 0x0003, true); // major version 3 (512-byte sectors)
  odv.setUint16(28, 0xfffe, true); // byte order
  odv.setUint16(30, 9, true); // sector shift
  odv.setUint16(32, 6, true); // mini sector shift
  odv.setUint32(44, fatSectorCount, true);
  odv.setUint32(48, dirStartSector, true);
  odv.setUint32(56, MINI_CUTOFF, true);
  odv.setUint32(60, ENDOFCHAIN, true); // first mini FAT sector (none)
  odv.setUint32(64, 0, true); // number of mini FAT sectors
  odv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector (none)
  odv.setUint32(72, 0, true); // number of DIFAT sectors
  for (let i = 0; i < 109; i++) odv.setUint32(76 + i * 4, i < fatSectorCount ? fatStartSector + i : FREESECT, true);

  const sectorAt = (s: number) => SECTOR + s * SECTOR;
  for (const p of placed) out.set(p.padded, sectorAt(p.startSector));
  out.set(dirBuf, sectorAt(dirStartSector));
  // FAT sectors.
  const fatBytes = new Uint8Array(fat.buffer, 0, fatSectorCount * SECTOR);
  out.set(fatBytes, sectorAt(fatStartSector));
  return out;
}
