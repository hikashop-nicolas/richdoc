import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { metafileBitmap } from "./metafile";
import { docxToParts, htmlToDocx } from "./index";

// A 2x1 24-bit DIB: BITMAPINFOHEADER (40 bytes) + one 8-byte row (6 bytes of pixels, padded).
function dib(): { bmi: Uint8Array; bits: Uint8Array } {
  const bmi = new Uint8Array(40);
  const v = new DataView(bmi.buffer);
  v.setUint32(0, 40, true);
  v.setInt32(4, 2, true);
  v.setInt32(8, 1, true);
  v.setUint16(12, 1, true);
  v.setUint16(14, 24, true);
  return { bmi, bits: new Uint8Array([0, 0, 255, 0, 255, 0, 0, 0]) };
}

function emf(withBitmap = true): Uint8Array {
  const recs: Uint8Array[] = [];
  const header = new Uint8Array(88);
  const hv = new DataView(header.buffer);
  hv.setUint32(0, 1, true);
  hv.setUint32(4, 88, true);
  hv.setUint32(40, 0x464d4520, true);
  recs.push(header);
  if (withBitmap) {
    const { bmi, bits } = dib();
    const rec = new Uint8Array(80 + bmi.length + bits.length);
    const rv = new DataView(rec.buffer);
    rv.setUint32(0, 81, true); // EMR_STRETCHDIBITS
    rv.setUint32(4, rec.length, true);
    rv.setUint32(48, 80, true);
    rv.setUint32(52, bmi.length, true);
    rv.setUint32(56, 80 + bmi.length, true);
    rv.setUint32(60, bits.length, true);
    rec.set(bmi, 80);
    rec.set(bits, 80 + bmi.length);
    recs.push(rec);
  }
  const eof = new Uint8Array(20);
  new DataView(eof.buffer).setUint32(0, 14, true);
  new DataView(eof.buffer).setUint32(4, 20, true);
  recs.push(eof);
  const out = new Uint8Array(recs.reduce((a, r) => a + r.length, 0));
  let o = 0;
  for (const r of recs) { out.set(r, o); o += r.length; }
  return out;
}

function wmf(): Uint8Array {
  const { bmi, bits } = dib();
  const rec = new Uint8Array(28 + bmi.length + bits.length);
  const rv = new DataView(rec.buffer);
  rv.setUint32(0, rec.length / 2, true);
  rv.setUint16(4, 0x0f43, true); // META_STRETCHDIB
  rec.set(bmi, 28);
  rec.set(bits, 28 + bmi.length);
  const head = new Uint8Array(22 + 18);
  const hv = new DataView(head.buffer);
  hv.setUint32(0, 0x9ac6cdd7, true);
  hv.setUint16(22, 1, true);
  hv.setUint16(24, 9, true);
  const eof = new Uint8Array([3, 0, 0, 0, 0, 0]);
  const out = new Uint8Array(head.length + rec.length + eof.length);
  out.set(head);
  out.set(rec, head.length);
  out.set(eof, head.length + rec.length);
  return out;
}

const bmpHeaderOk = (b: Uint8Array) => {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  return b[0] === 0x42 && b[1] === 0x4d && v.getUint32(2, true) === b.length && v.getUint32(10, true) === 54 && b.length === 54 + 8;
};

describe("metafile bitmaps", () => {
  it("pulls the DIB out of an EMF as a .bmp", () => {
    const r = metafileBitmap(emf(), "emf")!;
    expect(r.mime).toBe("image/bmp");
    expect(bmpHeaderOk(r.bytes)).toBe(true);
  });
  it("pulls the DIB out of a placeable WMF", () => {
    const r = metafileBitmap(wmf(), "wmf")!;
    expect(r.mime).toBe("image/bmp");
    expect(bmpHeaderOk(r.bytes)).toBe(true);
  });
  it("finds nothing in a vector-only or malformed file", () => {
    expect(metafileBitmap(emf(false), "emf")).toBeUndefined();
    expect(metafileBitmap(new Uint8Array([1, 2, 3]), "wmf")).toBeUndefined();
  });

  const docxWith = (media: Uint8Array, ext: string) => zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.${ext}"/></Relationships>`),
    [`word/media/image1.${ext}`]: media,
  });

  it("shows the bitmap at the picture's size and saves the metafile untouched", () => {
    const d = docxWith(emf(), "emf");
    const { body } = docxToParts(d);
    expect(body).toMatch(/<span class="docx-img-ph docx-img-meta"[^>]*><img src="data:image\/bmp;base64,[^"]+" alt="" style="width:100px;height:50px"><\/span>/);
    const out = unzipSync(htmlToDocx(body, d));
    expect(Array.from(out["word/media/image1.emf"]!)).toEqual(Array.from(emf()));
    expect(strFromU8(out["word/document.xml"]!)).toContain('r:embed="rId9"');
    expect(Object.keys(out).filter((k) => k.startsWith("word/media/"))).toEqual(["word/media/image1.emf"]);
  });

  it("shows a sized box for a vector-only metafile", () => {
    const { body } = docxToParts(docxWith(emf(false), "emf"));
    expect(body).toMatch(/<span class="docx-img-ph docx-img-vector"[^>]*style="width:100px;height:50px">Vector picture/);
  });
});
