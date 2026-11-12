import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { docxToParts, htmlToDocx } from "./index";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DSP = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="chevron"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1400"/><a:t>Step one</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`;
const inline = (gd: string) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="d"/><a:graphic>${gd}</a:graphic></wp:inline></w:drawing></w:r></w:p>`;
function docx(body: string, extra: Record<string, Uint8Array> = {}, rels = ""): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`),
    ...extra,
  });
}

describe("SmartArt and WordArt in a document", () => {
  it("draws a SmartArt diagram from its saved drawing and saves it untouched", () => {
    const gd = '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId4" r:lo="rId5" r:qs="rId6" r:cs="rId7"/></a:graphicData>';
    const rel = (id: string, t: string) => `<Relationship Id="${id}" Type="http://x/${t}" Target="diagrams/${t}1.xml"/>`;
    const d = docx(inline(gd), {
      "word/diagrams/data1.xml": strToU8('<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>'),
      "word/diagrams/drawing1.xml": strToU8(DSP),
    }, rel("rId4", "data") + rel("rId8", "drawing"));
    const { body } = docxToParts(d);
    expect(body).toMatch(/<span class="docx-img-ph docx-img-meta"[^>]*><img src="data:image\/svg\+xml;charset=utf-8,%3Csvg[^"]*Step[^"]*one[^"]*" alt="" style="width:200px;height:100px">/);
    const out = unzipSync(htmlToDocx(body, d));
    expect(strFromU8(out["word/document.xml"]!)).toContain('<dgm:relIds r:dm="rId4"');
  });

  it("draws WordArt from its warp and keeps the shape on save", () => {
    const wsp = '<wps:wsp><wps:spPr><a:prstGeom prst="rect"/><a:noFill/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="72"/></w:rPr><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></wps:bodyPr></wps:wsp>';
    const d = docx(inline(`<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp}</a:graphicData>`));
    const { body } = docxToParts(d);
    expect(body).toMatch(/class="docx-shape docx-shape-art"[^>]*><span class="docx-shape-box" style="width:200px;height:100px"><img src="data:image\/svg\+xml/);
    expect(body).not.toContain("docx-shape-body");
    expect(strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!)).toContain('prst="textArchUp"');
  });
});
