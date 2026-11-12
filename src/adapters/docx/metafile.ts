// The bitmap inside an EMF / WMF picture. Browsers cannot draw metafiles, but many (pasted
// screenshots, scans, charts saved as pictures) only wrap one device-independent bitmap: pull
// out the largest one as a .bmp (or the PNG / JPEG a DIB may carry) to show in its place.

export interface MetaBitmap {
  mime: string;
  bytes: Uint8Array;
}

/** A DIB (BITMAPINFO followed by its pixel bits) as an image the browser can show. */
function dibToImage(bmi: Uint8Array, bits: Uint8Array): MetaBitmap | undefined {
  if (bmi.length < 12) return undefined;
  const v = new DataView(bmi.buffer, bmi.byteOffset, bmi.byteLength);
  const hdr = v.getUint32(0, true);
  if (hdr === 12) {
    // BITMAPCOREHEADER: 3-byte colour entries.
    const bitCount = v.getUint16(10, true);
    const colors = bitCount <= 8 ? 1 << bitCount : 0;
    return bmp(bmi.subarray(0, 12 + colors * 3), bits);
  }
  if (hdr < 40 || bmi.length < 40) return undefined;
  const bitCount = v.getUint16(14, true);
  const compression = v.getUint32(16, true);
  if (compression === 4) return { mime: "image/jpeg", bytes: bits }; // BI_JPEG
  if (compression === 5) return { mime: "image/png", bytes: bits }; // BI_PNG
  const clrUsed = v.getUint32(32, true);
  const colors = clrUsed || (bitCount <= 8 ? 1 << bitCount : 0);
  const masks = hdr === 40 && (compression === 3 || compression === 6) ? (compression === 6 ? 16 : 12) : 0;
  return bmp(bmi.subarray(0, Math.min(bmi.length, hdr + masks + colors * 4)), bits);
}

function bmp(info: Uint8Array, bits: Uint8Array): MetaBitmap {
  const out = new Uint8Array(14 + info.length + bits.length);
  const v = new DataView(out.buffer);
  out[0] = 0x42; // "BM"
  out[1] = 0x4d;
  v.setUint32(2, out.length, true);
  v.setUint32(10, 14 + info.length, true);
  out.set(info, 14);
  out.set(bits, 14 + info.length);
  return { mime: "image/bmp", bytes: out };
}

// EMF records whose source bitmap is given by offBmiSrc/cbBmiSrc/offBitsSrc/cbBitsSrc, keyed by
// record type, valued by where offBmiSrc sits in the record.
const EMF_DIB_AT: Record<number, number> = { 76: 84, 77: 84, 80: 48, 81: 48 }; // BITBLT, STRETCHBLT, SETDIBITSTODEVICE, STRETCHDIBITS

function emfBitmap(b: Uint8Array): MetaBitmap | undefined {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (b.length < 88 || v.getUint32(0, true) !== 1 || v.getUint32(40, true) !== 0x464d4520) return undefined; // " EMF"
  let best: { bmi: Uint8Array; bits: Uint8Array } | undefined;
  for (let off = 0; off + 8 <= b.length; ) {
    const type = v.getUint32(off, true);
    const size = v.getUint32(off + 4, true);
    if (size < 8 || off + size > b.length) break;
    const at = EMF_DIB_AT[type];
    if (at !== undefined && size >= at + 16) {
      const offBmi = v.getUint32(off + at, true), cbBmi = v.getUint32(off + at + 4, true);
      const offBits = v.getUint32(off + at + 8, true), cbBits = v.getUint32(off + at + 12, true);
      if (cbBmi && cbBits && offBmi + cbBmi <= size && offBits + cbBits <= size && (!best || cbBits > best.bits.length)) {
        best = { bmi: b.subarray(off + offBmi, off + offBmi + cbBmi), bits: b.subarray(off + offBits, off + offBits + cbBits) };
      }
    }
    if (type === 14) break; // EMR_EOF
    off += size;
  }
  return best && dibToImage(best.bmi, best.bits);
}

// WMF records carrying a DIB, valued by where the DIB starts in the record.
const WMF_DIB_AT: Record<number, number> = { 0x0f43: 28, 0x0b41: 26, 0x0940: 22, 0x0d33: 24 }; // STRETCHDIB, DIBSTRETCHBLT, DIBBITBLT, SETDIBTODEV

function wmfBitmap(b: Uint8Array): MetaBitmap | undefined {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let off = b.length >= 22 && v.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0; // placeable header
  if (off + 18 > b.length) return undefined;
  const hdrWords = v.getUint16(off + 2, true);
  if (hdrWords !== 9) return undefined;
  off += 18;
  let best: Uint8Array | undefined;
  while (off + 6 <= b.length) {
    const size = v.getUint32(off, true) * 2;
    const fn = v.getUint16(off + 4, true);
    if (size < 6 || off + size > b.length) break;
    const at = WMF_DIB_AT[fn];
    if (at !== undefined && size > at + 40 && (!best || size - at > best.length)) best = b.subarray(off + at, off + size);
    if (fn === 0) break; // META_EOF
    off += size;
  }
  if (!best) return undefined;
  // In a WMF the DIB is contiguous: its header (and colour table) then the bits.
  const dv = new DataView(best.buffer, best.byteOffset, best.byteLength);
  const hdr = dv.getUint32(0, true);
  if (hdr !== 12 && hdr < 40) return undefined;
  let infoLen: number;
  if (hdr === 12) {
    const bc = dv.getUint16(10, true);
    infoLen = 12 + (bc <= 8 ? 1 << bc : 0) * 3;
  } else {
    const bc = dv.getUint16(14, true), comp = dv.getUint32(16, true), used = dv.getUint32(32, true);
    infoLen = hdr + (hdr === 40 && comp === 3 ? 12 : hdr === 40 && comp === 6 ? 16 : 0) + (used || (bc <= 8 ? 1 << bc : 0)) * 4;
  }
  if (infoLen >= best.length) return undefined;
  return dibToImage(best.subarray(0, infoLen), best.subarray(infoLen));
}

/** The bitmap a metafile wraps, or undefined for a vector-only one (or a malformed file). */
export function metafileBitmap(bytes: Uint8Array, ext: string): MetaBitmap | undefined {
  try {
    return ext === "emf" ? emfBitmap(bytes) : ext === "wmf" ? wmfBitmap(bytes) : undefined;
  } catch {
    return undefined;
  }
}
