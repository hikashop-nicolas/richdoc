import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { encodePng, metafileToSvgUrl, parseMetafile, pictureToSvg } from "./metafile";
import { docxToParts, htmlToDocx } from "./index";

// A little-endian record writer.
class W {
  bytes: number[] = [];
  u8(v: number) { this.bytes.push(v & 0xff); return this; }
  u16(v: number) { return this.u8(v).u8(v >> 8); }
  i16(v: number) { return this.u16(v & 0xffff); }
  u32(v: number) { return this.u16(v & 0xffff).u16((v >>> 16) & 0xffff); }
  i32(v: number) { return this.u32(v >>> 0); }
  raw(b: ArrayLike<number>) { for (let i = 0; i < b.length; i++) this.bytes.push(b[i]!); return this; }
  pad(n: number) { while (this.bytes.length < n) this.bytes.push(0); return this; }
}

// ---- EMF ----
const emfRec = (type: number, body: (w: W) => void): number[] => {
  const w = new W();
  body(w);
  while (w.bytes.length % 4) w.u8(0);
  return new W().u32(type).u32(w.bytes.length + 8).raw(w.bytes).bytes;
};
function emf(records: number[][]): Uint8Array {
  // 10 device px per mm; the frame is 10 x 5 mm, so the drawing spans device units 0..100 x 0..50.
  const h = new W().u32(1).u32(88).i32(0).i32(0).i32(100).i32(50).i32(0).i32(0).i32(1000).i32(500).u32(0x464d4520).pad(72).i32(1000).i32(1000).i32(100).i32(100);
  const recs = records.flat();
  const eof = new W().u32(14).u32(20).u32(0).u32(16).u32(20).bytes;
  return new Uint8Array([...h.bytes, ...recs, ...eof]);
}
const dib2x1 = () => {
  const info = new W().u32(40).i32(2).i32(1).u16(1).u16(24).pad(40).bytes;
  return { info, bits: [0, 0, 255, 0, 255, 0, 0, 0] }; // red, green
};
const stretchDibits = (x: number, y: number, w: number, h: number) => {
  const { info, bits } = dib2x1();
  return emfRec(81, (r) => {
    r.i32(0).i32(0).i32(0).i32(0) // bounds
      .i32(x).i32(y).i32(0).i32(0).i32(2).i32(1) // dest x/y, src x/y/w/h
      .u32(80).u32(info.length).u32(80 + info.length).u32(bits.length) // bmi / bits (offsets from the record start)
      .u32(0).u32(0x00cc0020).i32(w).i32(h)
      .raw(info).raw(bits);
  });
};
const text = (s: string, x: number, y: number) => {
  const units = [...s].map((c) => c.charCodeAt(0));
  return emfRec(84, (r) => {
    r.i32(0).i32(0).i32(0).i32(0).u32(1).u32(0).u32(0) // bounds, graphics mode, scales
      .i32(x).i32(y).u32(units.length).u32(76).u32(0).pad(64).u32(0) // reference, count, string at 76, options, rect, no dx
      .raw(new W().raw(units.flatMap((u) => [u & 0xff, u >> 8])).bytes);
  });
};
const VECTOR_EMF = () => emf([
  emfRec(39, (r) => r.u32(1).u32(0).u32(0x0000ff).u32(0)), // brush 1: red
  emfRec(37, (r) => r.u32(1)),
  emfRec(43, (r) => r.i32(10).i32(10).i32(90).i32(40)), // rectangle
  emfRec(38, (r) => r.u32(2).u32(0).i32(4).i32(0).u32(0xff0000)), // pen 2: blue, 4 wide
  emfRec(37, (r) => r.u32(2)),
  emfRec(87, (r) => r.i32(0).i32(0).i32(0).i32(0).u32(2).i16(0).i16(45).i16(100).i16(45)), // polyline16
  emfRec(82, (r) => { r.u32(3).i32(-10).i32(0).i32(0).i32(0).i32(700).u8(1).pad(32); for (const c of "Arial") r.u16(c.charCodeAt(0)); r.pad(32 + 64 + 4); }),
  emfRec(37, (r) => r.u32(3)),
  emfRec(24, (r) => r.u32(0x00ff00)), // text colour green
  text("Hi", 20, 20),
]);

describe("metafile vector rendering (ported from wordcraft)", () => {
  it("plays EMF shapes, pens and text into the picture frame", () => {
    const pic = parseMetafile(VECTOR_EMF())!;
    expect(pic.width).toBeCloseTo(28.35, 1); // 10 mm in points
    expect(pic.height).toBeCloseTo(14.17, 1);
    const kinds = pic.items.map((i) => i.k);
    expect(kinds).toEqual(["path", "path", "text"]);
    const svg = pictureToSvg(pic);
    expect(svg).toMatch(/<path d="M2\.835 2\.835L25\.512 2\.835L25\.512 11\.339L2\.835 11\.339Z" fill="#ff0000" fill-rule="evenodd" stroke="#000000" stroke-width="1" vector-effect="non-scaling-stroke"/);
    expect(svg).toMatch(/<path d="M0 12\.756L28\.346 12\.756" fill="none" stroke="#0000ff" stroke-width="1\.134"/);
    expect(svg).toMatch(/<text [^>]*font-family="Arial, Arial, sans-serif"[^>]* font-weight="bold" font-style="italic" fill="#00ff00"[^>]*>Hi<\/text>/);
  });

  it("decodes a DIB into a PNG image placed in the frame", () => {
    const svg = pictureToSvg(parseMetafile(emf([stretchDibits(0, 0, 100, 50)]))!);
    expect(svg).toMatch(/<image x="0" y="0" width="28\.346" height="14\.173" preserveAspectRatio="none" href="data:image\/png;base64,/);
  });

  it("encodes a valid PNG", () => {
    const png = encodePng({ w: 2, h: 1, px: new Uint8Array([255, 0, 0, 255, 0, 255, 0, 255]) });
    expect(Array.from(png.slice(0, 8))).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(strFromU8(png.slice(12, 16))).toBe("IHDR");
    expect(strFromU8(png.slice(-8, -4))).toBe("IEND");
  });

  it("plays a placeable WMF with a brush, a rectangle and Shift-JIS text", () => {
    const rec = (fn: number, body: (w: W) => void) => { const w = new W(); body(w); if (w.bytes.length % 2) w.u8(0); return new W().u32((w.bytes.length + 6) / 2).u16(fn).raw(w.bytes).bytes; };
    const sjis = [0x93, 0xfa, 0x96, 0x7b]; // 日本
    const records = [
      rec(0x02fc, (r) => r.u16(0).u32(0x00ff00).u16(0)), // brush (handle 0): green
      rec(0x012d, (r) => r.u16(0)),
      rec(0x041b, (r) => r.i16(40).i16(90).i16(10).i16(10)), // rectangle: bottom, right, top, left
      rec(0x02fb, (r) => { r.i16(-20).i16(0).i16(0).i16(0).i16(400).u8(0).u8(0).u8(0).u8(128).pad(18); for (const c of "Gothic") r.u8(c.charCodeAt(0)); r.pad(18 + 32); }),
      rec(0x012d, (r) => r.u16(1)),
      rec(0x0521, (r) => r.u16(sjis.length).raw(sjis).i16(30).i16(20)), // TEXTOUT: string, y, x
    ].flat();
    const head = new W().u32(0x9ac6cdd7).u16(0).i16(0).i16(0).i16(100).i16(50).u16(1440).u32(0).u16(0) // placeable
      .u16(1).u16(9).u16(0x300).u32(0).u16(0).u32(0).u16(0).bytes;
    const pic = parseMetafile(new Uint8Array([...head, ...records, 3, 0, 0, 0, 0, 0]))!;
    expect(pic.width).toBeCloseTo(5, 3); // 100 twips-per-inch units at 1440/in -> 5 pt
    const svg = pictureToSvg(pic);
    expect(svg).toContain('fill="#00ff00"');
    expect(svg).toMatch(/<text [^>]*font-family="Gothic, Arial, sans-serif"[^>]*>日本<\/text>/);
  });

  it("rejects what is not a metafile, and never throws on a truncated one", () => {
    expect(parseMetafile(new Uint8Array([1, 2, 3]))).toBeUndefined();
    const full = VECTOR_EMF();
    for (let n = 0; n < full.length; n += 7) expect(() => parseMetafile(full.subarray(0, n))).not.toThrow();
  });

  const docxWith = (media: Uint8Array, ext: string) => zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.${ext}"/></Relationships>`),
    [`word/media/image1.${ext}`]: media,
  });

  it("shows a vector EMF in a docx at the picture's size and saves it untouched", () => {
    const media = VECTOR_EMF();
    const d = docxWith(media, "emf");
    const { body } = docxToParts(d);
    expect(body).toMatch(/<span class="docx-img-ph docx-img-meta"[^>]*><img src="data:image\/svg\+xml;charset=utf-8,%3Csvg[^"]+" alt="" style="width:100px;height:50px"><\/span>/);
    const out = unzipSync(htmlToDocx(body, d));
    expect(Array.from(out["word/media/image1.emf"]!)).toEqual(Array.from(media));
    expect(Object.keys(out).filter((k) => k.startsWith("word/media/"))).toEqual(["word/media/image1.emf"]);
  });

  it("shows a sized box for a metafile that draws nothing", () => {
    expect(metafileToSvgUrl(emf([]))).toBeUndefined();
    const { body } = docxToParts(docxWith(emf([emfRec(33, () => {})]), "emf"));
    expect(body).toMatch(/<span class="docx-img-ph docx-img-vector"[^>]*style="width:100px;height:50px">Vector picture/);
  });
});
