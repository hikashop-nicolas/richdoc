import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { deobfuscateFont, docxToHtml, docxToParts, htmlToDocx } from "./index";

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Titre</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">Bonjour </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>monde</w:t></w:r><w:r><w:t> ici.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
 </w:body>
</w:document>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

function makeDocx(document = DOCUMENT): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(document),
    "word/_rels/document.xml.rels": strToU8(RELS),
    "word/styles.xml": strToU8("<styles/>"),
    "extra.bin": new Uint8Array([9, 8, 7]),
  });
}

describe("docx <-> html", () => {
  it("reads the body to HTML, mapping headings and bold", () => {
    const html = docxToHtml(makeDocx());
    expect(html).toContain("<h1>Titre</h1>");
    expect(html).toContain("Bonjour");
    expect(html).toContain("<strong>monde</strong>");
    expect(html).toContain("ici.");
  });

  it("writes edited HTML back, preserving other parts and the section properties", () => {
    const docx = makeDocx();
    const out = htmlToDocx("<h1>Titre</h1><p>Bonjour <strong>planete</strong> la.</p>", docx);

    const files = unzipSync(out);
    expect(strFromU8(files["word/styles.xml"])).toBe("<styles/>");
    expect(strFromU8(files["[Content_Types].xml"])).toBe("<Types/>");
    expect(Array.from(files["extra.bin"])).toEqual([9, 8, 7]);
    // section properties (page setup) are preserved
    const docXml = strFromU8(files["word/document.xml"]);
    expect(docXml).toContain("w:sectPr");
    expect(docXml).toContain("pgSz");

    const html2 = docxToHtml(out);
    expect(html2).toContain("<h1>Titre</h1>");
    expect(html2).toContain("<strong>planete</strong>");
    expect(html2).toContain("la.");
    expect(html2).not.toContain("monde");
  });

  it("round-trips bold/italic/underline", () => {
    const out = htmlToDocx("<p><strong>b</strong> <em>i</em> <u>u</u></p>", makeDocx());
    const html = docxToHtml(out);
    expect(html).toContain("<strong>b</strong>");
    expect(html).toContain("<em>i</em>");
    expect(html).toContain("<u>u</u>");
  });

  it("renders paragraph alignment, run colour, highlight and font size", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>justified</w:t></w:r></w:p>
 <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="28"/></w:rPr><w:t>big red</w:t></w:r></w:p>
 <w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>marked</w:t></w:r></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('text-align:justify');
    expect(html).toContain('text-align:center');
    expect(html).toContain('color:#FF0000');
    expect(html).toContain('font-size:14pt'); // sz 28 half-points
    expect(html).toContain('background-color:#ffff00');
  });

  it("round-trips alignment, colour, highlight and size from edited HTML", () => {
    const html =
      '<p style="text-align:justify">j</p>' +
      '<p style="text-align:center"><span style="color:#1188ff;font-size:18pt">c</span></p>' +
      '<p><span style="background-color:rgb(255,255,0)">h</span></p>';
    const out = htmlToDocx(html, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain('w:jc w:val="both"');
    expect(xml).toContain('w:jc w:val="center"');
    expect(xml).toContain('w:color w:val="1188FF"');
    expect(xml).toContain('w:sz w:val="36"'); // 18pt -> 36 half-points
    expect(xml).toContain('w:highlight w:val="yellow"');
    // and it reads back
    const html2 = docxToHtml(out);
    expect(html2).toContain('text-align:justify');
    expect(html2).toContain('color:#1188FF');
    expect(html2).toContain('background-color:#ffff00');
  });

  it("renders a table and preserves it byte-exact through a save", () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
      "</w:tbl>";
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>before</w:t></w:r></w:p>${tbl}<w:p><w:r><w:t>after</w:t></w:r></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain("<table");
    expect(html).toContain("A1");
    expect(html).toContain("data-docx-xml");
    // editing the surrounding text and saving keeps the table intact
    const out = htmlToDocx(html.replace("before", "BEFORE"), makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("<w:tbl");
    expect(xml).toContain("A1");
    expect(xml).toContain("B1");
    expect(xml).toContain("BEFORE");
  });

  it("writes per-cell borders, resized column widths and a row height when a table is edited", () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>x</w:t></w:r></w:p>${tbl}
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    // Simulate an edit: drop the skeleton (forces build-from-DOM), add a colgroup with widths,
    // a row height, and top+left borders on the first cell.
    const edited = html
      .replace(/ data-docx-xml="[^"]*"/, "")
      .replace('<table class="docx-table"', '<table class="docx-table" style="margin-left: 48px"')
      .replace(/(<table class="docx-table"[^>]*>)/, '$1<colgroup><col style="width: 120px"><col style="width: 90px"></colgroup>')
      .replace(/<tr>/, '<tr style="height: 40px">')
      .replace("<td", '<td class="rdoc-bordered" data-rdoc-bt="2px dashed #ff0000" data-rdoc-bl="1px solid #000000" ');
    const out = htmlToDocx(edited, makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("<w:tcBorders");
    expect(xml).toContain('<w:top w:val="dashed"'); // chosen style
    expect(xml).toContain('w:sz="12"'); // 2px -> 1/8 pt
    expect(xml).toContain('w:color="ff0000"'); // chosen colour
    expect(xml).toContain('<w:bottom w:val="nil"');
    expect(xml).toContain('w:w="1800"'); // 120px -> twips
    expect(xml).toContain('w:w="1350"'); // 90px -> twips
    expect(xml).toContain("<w:trHeight");
    expect(xml).toContain('w:val="600"'); // 40px -> twips
    expect(xml).toContain('<w:tblInd w:w="720"'); // 48px indent -> twips
  });

  it("resolves table and cell borders into the editor's per-side model on read", () => {
    const tbl =
      "<w:tbl><w:tblPr><w:tblBorders>" +
      '<w:top w:val="single" w:sz="12" w:color="FF0000"/><w:left w:val="single" w:sz="12" w:color="FF0000"/>' +
      '<w:bottom w:val="single" w:sz="12" w:color="FF0000"/><w:right w:val="single" w:sz="12" w:color="FF0000"/>' +
      '<w:insideH w:val="dashed" w:sz="6" w:color="0000FF"/><w:insideV w:val="dashed" w:sz="6" w:color="0000FF"/>' +
      "</w:tblBorders></w:tblPr>" +
      "<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>" +
      "<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${tbl}</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    // First cell: red outer top/left (12 eighths -> 2px), blue dashed inner bottom/right (6 -> 1px).
    expect(html).toContain('data-rdoc-bt="2px solid #FF0000"');
    expect(html).toContain('data-rdoc-bl="2px solid #FF0000"');
    expect(html).toContain('data-rdoc-bb="1px dashed #0000FF"');
    expect(html).toContain('data-rdoc-br="1px dashed #0000FF"');
  });

  it("reads vertical (tategaki) text direction from the section", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>x</w:t></w:r></w:p>
 <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:textDirection w:val="tbRl"/></w:sectPr>
</w:body></w:document>`;
    expect(docxToParts(makeDocx(doc)).page?.vertical).toBe(true);
  });

  it("reads a right-to-left (bidi) section", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>x</w:t></w:r></w:p>
 <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:bidi/></w:sectPr>
</w:body></w:document>`;
    const page = docxToParts(makeDocx(doc)).page;
    expect(page?.rtl).toBe(true);
    expect(page?.vertical).toBe(false);
  });

  it("writes inserted fields (page number/count) and a table of contents", () => {
    const body =
      '<h1>Alpha</h1>' +
      '<div class="docx-field-toc"><div class="docx-field-toc-title">Contents</div>' +
      '<div class="docx-field-toc-row toc-h1"><span class="docx-field-toc-text">Alpha</span><span class="docx-field-toc-page">1</span></div></div>' +
      '<p>Page <span class="docx-field" data-field="PAGE">2</span> / <span class="docx-field" data-field="NUMPAGES">5</span></p>';
    const xml = strFromU8(unzipSync(htmlToDocx(body, makeDocx()))["word/document.xml"]);
    expect(xml).toContain('<w:fldSimple w:instr=" PAGE "');
    expect(xml).toContain('<w:fldSimple w:instr=" NUMPAGES "');
    expect(xml).toContain('w:fldCharType="begin"'); // TOC complex field
    expect(xml).toMatch(/TOC \\o/); // TOC instruction
    expect(xml).toContain('w:fldCharType="end"');
  });

  it("renders an inline image and preserves the drawing through a save", () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
    const drawing =
      '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="1143000" cy="571500"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="x">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
      '<a:blip r:embed="rId100"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
 <w:p><w:r><w:t>keep me</w:t></w:r>${drawing}</w:p>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="x" Target="media/image1.png"/></Relationships>`,
      ),
      "word/media/image1.png": png,
    };
    const docx = zipSync(files);
    const html = docxToHtml(docx);
    expect(html).toContain("data:image/png;base64,");
    expect(html).toContain("data-docx-xml");
    const out = htmlToDocx(html.replace("keep me", "KEEP ME"), docx);
    const files2 = unzipSync(out);
    const xml = strFromU8(files2["word/document.xml"]);
    expect(xml).toContain("w:drawing");
    expect(xml).toContain("rId100");
    expect(xml).toContain("KEEP ME");
    expect(Array.from(files2["word/media/image1.png"])).toEqual(Array.from(png));
  });

  // A floating image: a w:drawing whose container is wp:anchor (with a wrap + position),
  // rendered out of line and editable via the image toolbar (data-rdoc-* attributes).
  const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
  const anchorImg = (wrapEl: string, posH: string, behind = "0") =>
    '<w:r><w:drawing><wp:anchor behindDoc="' + behind + '" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    posH +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>476250</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="1143000" cy="571500"/>' + wrapEl +
    '<wp:docPr id="1" name="Image 1" descr="a cat"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="x">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
    '<a:blip r:embed="rId100"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const imgDocx = (drawing: string) =>
    zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
 <w:p><w:r><w:t>text</w:t></w:r>${drawing}</w:p>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="x" Target="media/image1.png"/></Relationships>`,
      ),
      "word/media/image1.png": PNG,
    });

  it("reads a floating image's wrap mode, alignment and alt text", () => {
    const html = docxToHtml(imgDocx(anchorImg('<wp:wrapSquare wrapText="bothSides"/>', '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>')));
    expect(html).toContain('data-rdoc-wrap="square"');
    expect(html).toContain('data-rdoc-align="right"');
    expect(html).toContain('alt="a cat"');
  });

  it("preserves a square image's exact offset position (posOffset, not alignment)", () => {
    const posH = '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>'; // 96px
    const docx = imgDocx(anchorImg('<wp:wrapSquare wrapText="bothSides"/>', posH)); // fixture's posV is posOffset 476250 (50px)
    const html = docxToHtml(docx);
    expect(html).toContain('data-rdoc-absx="1"');
    expect(html).toContain('data-rdoc-absy="1"');
    const xml = strFromU8(unzipSync(htmlToDocx(html.replace("text", "TEXT"), docx))["word/document.xml"]!);
    expect(xml).toContain("wp:wrapSquare");
    expect(xml).not.toContain("<wp:align>"); // positioned by offset on both axes, no alignment
    expect(xml).toMatch(/<wp:positionH[^>]*><wp:posOffset>914400<\/wp:posOffset>/);
    expect(xml).toMatch(/<wp:positionV[^>]*><wp:posOffset>476250<\/wp:posOffset>/);
  });

  it("keeps the V offset of a square image aligned in H", () => {
    // H = align right, V = posOffset 476250 (50px): the alignment AND the vertical offset both survive.
    const docx = imgDocx(anchorImg('<wp:wrapSquare wrapText="bothSides"/>', '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>'));
    const xml = strFromU8(unzipSync(htmlToDocx(docxToHtml(docx), docx))["word/document.xml"]!);
    expect(xml).toContain("<wp:align>right</wp:align>");
    expect(xml).toMatch(/<wp:positionV[^>]*><wp:posOffset>476250<\/wp:posOffset>/);
  });

  it("round-trips a floating image's wrap through an edit", () => {
    const docx = imgDocx(anchorImg('<wp:wrapSquare wrapText="bothSides"/>', '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>'));
    const html = docxToHtml(docx);
    const xml = strFromU8(unzipSync(htmlToDocx(html.replace("text", "TEXT"), docx))["word/document.xml"]);
    expect(xml).toContain("wp:anchor");
    expect(xml).toContain("wp:wrapSquare");
    expect(xml).toContain("<wp:align>right</wp:align>");
    expect(xml).toContain("rId100"); // the original picture relationship is preserved
  });

  it("round-trips a floating image's wrap distances (text padding)", () => {
    const drawing = anchorImg('<wp:wrapSquare wrapText="bothSides"/>', '<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>')
      .replace('behindDoc="0"', 'behindDoc="0" distT="0" distB="0" distL="228600" distR="228600"');
    const docx = imgDocx(drawing);
    const html = docxToHtml(docx);
    expect(html).toContain('data-rdoc-wrapdist="0,24,0,24"'); // 228600 EMU = 24px (t,r,b,l)
    const xml = strFromU8(unzipSync(htmlToDocx(html, docx))["word/document.xml"]);
    expect(xml).toMatch(/distL="228600"/);
    expect(xml).toMatch(/distR="228600"/);
  });

  it("converts an inline image to wrap text when the toolbar sets a wrap mode", () => {
    const inline =
      '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="1143000" cy="571500"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="x">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
      '<a:blip r:embed="rId100"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const docx = imgDocx(inline);
    const html = docxToHtml(docx);
    expect(html).not.toContain("data-rdoc-wrap"); // starts inline
    const edited = html.replace("<img ", '<img data-rdoc-wrap="square" data-rdoc-align="left" ');
    const xml = strFromU8(unzipSync(htmlToDocx(edited, docx))["word/document.xml"]);
    expect(xml).toContain("wp:anchor");
    expect(xml).toContain("wp:wrapSquare");
  });

  it("embeds a new floating image as a wp:anchor with a wrap", () => {
    const out = htmlToDocx('<p><img src="data:image/png;base64,iVBORw0KGgo=" width="120" height="80" data-rdoc-wrap="topbottom" data-rdoc-align="center"></p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("wp:anchor");
    expect(xml).toContain("wp:wrapTopAndBottom");
  });

  it("reads a behind-text image's offset and writes behindDoc back", () => {
    const docx = imgDocx(anchorImg("<wp:wrapNone/>", '<wp:positionH relativeFrom="column"><wp:posOffset>952500</wp:posOffset></wp:positionH>', "1"));
    const html = docxToHtml(docx);
    expect(html).toContain('data-rdoc-wrap="behind"');
    expect(html).toContain('data-rdoc-x="100"'); // 952500 EMU = 100px
    const xml = strFromU8(unzipSync(htmlToDocx(html, docx))["word/document.xml"]);
    expect(xml).toContain('behindDoc="1"');
    expect(xml).toContain("wp:wrapNone");
  });

  it("round-trips a tab character and the paragraph's custom tab stops", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2160"/><w:tab w:val="right" w:pos="6480" w:leader="dot"/></w:tabs></w:pPr><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-docx-tab="1"'); // the tab character is shown as an atomic span
    expect(html).toContain("data-rdoc-tabstops"); // the paragraph's stops are preserved
    const xml = strFromU8(unzipSync(htmlToDocx(html.replace(">a<", ">A<"), makeDocx(doc)))["word/document.xml"]);
    expect(xml).toContain("<w:tab/>"); // the tab character round-trips
    expect(xml).toMatch(/<w:tabs>/); // the custom stops round-trip
    expect(xml).toMatch(/w:pos="2160"/); // 2160 twips = 144px, back to 2160
    expect(xml).toMatch(/w:leader="dot"/);
    expect(xml).toContain("A");
  });

  it("preserves a mid-document section break through an edit (no flattening)", () => {
    // Section one ends at its paragraph (w:pPr/w:sectPr, next-page, portrait); the trailing
    // body sectPr governs section two (landscape).
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>section one</w:t></w:r></w:p>
 <w:p><w:r><w:t>section two</w:t></w:r></w:p>
 <w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain("data-docx-sectpr"); // the break is surfaced on the paragraph
    expect(html).toContain("docx-pagebreak-auto"); // shown as a page boundary
    const xml = strFromU8(unzipSync(htmlToDocx(html.replace("section one", "SECTION ONE"), makeDocx(doc)))["word/document.xml"]);
    expect(xml).toContain("SECTION ONE");
    expect(xml).toContain("section two");
    // Both sections survive: the mid-document sectPr (in a w:pPr) and the trailing body sectPr.
    expect(xml.match(/<w:sectPr/g)?.length).toBe(2);
    expect(xml).toContain("<w:pPr><w:sectPr"); // the mid-doc break stays inside its paragraph
    expect(xml).toContain('w:val="nextPage"'); // section-one type preserved
    expect(xml).toContain('w:orient="landscape"'); // section-two trailing sectPr preserved
  });

  it("writes <ol start> as a startOverride and restarts separate ordered lists", () => {
    // An explicit start.
    const started = strFromU8(unzipSync(htmlToDocx('<ol start="5"><li>a</li><li>b</li></ol>', makeDocx()))["word/numbering.xml"]);
    expect(started).toContain("w:startOverride");
    expect(started).toContain('w:val="5"');
    // Two separate ordered lists must use different numIds so each restarts at 1.
    const xml = strFromU8(unzipSync(htmlToDocx("<ol><li>a</li></ol><p>x</p><ol><li>b</li></ol>", makeDocx()))["word/document.xml"]);
    const numIds = [...xml.matchAll(/<w:numId w:val="(\d+)"/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(2);
    expect(numIds[0]).not.toBe(numIds[1]);
  });

  it("reads an ordered list's start number, and continuation across lists", () => {
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
    const li = (numId: string, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    // numId 1 starts at 3; numId 2 is used by two lists separated by a paragraph (continues).
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 ${li("1", "three")}
 <w:p><w:r><w:t>gap</w:t></w:r></w:p>
 ${li("2", "one")}${li("2", "two")}
 <w:p><w:r><w:t>gap</w:t></w:r></w:p>
 ${li("2", "cont")}
</w:body></w:document>`;
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(RELS),
      "word/numbering.xml": strToU8(numbering),
    };
    const html = docxToHtml(zipSync(files));
    expect(html).toContain('<ol start="3"><li>three</li></ol>'); // explicit start of 3
    expect(html).toContain("<ol><li>one</li><li>two</li></ol>"); // first list of numId 2 starts at 1
    expect(html).toContain('<ol start="3"><li>cont</li></ol>'); // third list continues numId 2 (1,2 -> 3)
  });

  it("shows manual page breaks (and Word's auto breaks) and round-trips the manual one", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>one</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>two</w:t></w:r></w:p>
 <w:p><w:r><w:lastRenderedPageBreak/><w:t>three</w:t></w:r></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-docx-pagebreak="manual"');
    expect(html).toContain('data-docx-pagebreak="auto"');
    const out = htmlToDocx(html, makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain('w:type="page"'); // manual break preserved
    expect(xml).not.toContain("lastRenderedPageBreak"); // auto markers dropped (Word regenerates)
    expect(xml).toContain("one");
    expect(xml).toContain("two");
  });

  it("renders a referenced header", () => {
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
 <w:p><w:r><w:t>body</w:t></w:r></w:p>
 <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="x" Target="header1.xml"/></Relationships>`,
      ),
      "word/header1.xml": strToU8(
        `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>My header</w:t></w:r></w:p></w:hdr>`,
      ),
    };
    const parts = docxToParts(zipSync(files));
    expect(parts.body).toContain("body");
    expect(parts.header).toContain("My header");
  });

  it("writes back an edited header part", () => {
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
 <w:p><w:r><w:t>body</w:t></w:r></w:p>
 <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="x" Target="header1.xml"/></Relationships>`,
      ),
      "word/header1.xml": strToU8(
        `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>old header</w:t></w:r></w:p></w:hdr>`,
      ),
    };
    const docx = zipSync(files);
    const parts = docxToParts(docx);
    expect(parts.headerPath).toBe("word/header1.xml");
    const out = htmlToDocx(parts.body, docx, [{ path: parts.headerPath!, html: "<p>new header</p>" }]);
    const files2 = unzipSync(out);
    const hdr = strFromU8(files2["word/header1.xml"]);
    expect(hdr).toContain("new header");
    expect(hdr).not.toContain("old header");
    expect(hdr).toContain("w:hdr"); // root element kept
  });

  it("round-trips arbitrary background as shading, named as highlight", () => {
    const out = htmlToDocx(
      '<p><span style="background-color:#123456">a</span><span style="background-color:rgb(255,255,0)">b</span></p>',
      makeDocx(),
    );
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain('w:fill="123456"'); // arbitrary -> w:shd
    expect(xml).toContain("w:shd");
    expect(xml).toContain('w:highlight w:val="yellow"'); // exact match -> highlight
  });

  it("embeds a newly inserted image (media + relationship + content type)", () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, 1, 2, 3]);
    const dataUrl = "data:image/png;base64," + Buffer.from(png).toString("base64");
    const out = htmlToDocx(`<p><img src="${dataUrl}" width="120" height="60"></p>`, makeDocx());
    const files = unzipSync(out);
    const xml = strFromU8(files["word/document.xml"]);
    expect(xml).toContain("w:drawing");
    expect(xml).toContain("r:embed");
    const media = Object.keys(files).find((k) => k.startsWith("word/media/"));
    expect(media).toBeTruthy();
    expect(Array.from(files[media!])).toEqual(Array.from(png));
    expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("/image");
    expect(strFromU8(files["[Content_Types].xml"])).toContain('Extension="png"');
  });

  it("displays an existing comment and preserves it on save", () => {
    const files = {
      "[Content_Types].xml": strToU8(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:commentRangeStart w:id="0"/><w:r><w:t>flagged</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t> rest</w:t></w:r></w:p>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`,
      ),
      "word/comments.xml": strToU8(
        `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Alice" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>please check</w:t></w:r></w:p></w:comment></w:comments>`,
      ),
    };
    const docx = zipSync(files);
    const html = docxToHtml(docx);
    expect(html).toContain("docx-comment-ref");
    expect(html).toContain("Alice");
    expect(html).toContain("please check");
    expect(html).toContain("flagged");
    const out = htmlToDocx(html, docx);
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain('w:commentRangeStart');
    expect(xml).toContain('w:commentRangeEnd');
    expect(xml).toContain("w:commentReference");
    expect(xml).toContain("flagged");
    // comments.xml untouched
    expect(strFromU8(unzipSync(out)["word/comments.xml"])).toContain("please check");
  });

  it("adds a new comment, creating comments.xml when absent", () => {
    const html =
      '<p>before<span class="docx-cmark" data-docx-xml="&lt;w:commentRangeStart xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot; w:id=&quot;0&quot;/&gt;"></span>' +
      '<span class="docx-comment">target</span>' +
      '<span class="docx-cmark" data-docx-xml="&lt;w:commentRangeEnd xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot; w:id=&quot;0&quot;/&gt;"></span>' +
      '<span class="docx-comment-ref" data-comment-id="0" data-comment-new="1" data-comment-author="Bob" data-comment-date="2026-03-04T00:00:00Z" data-comment-text="needs work" ' +
      'data-docx-xml="&lt;w:r xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot;&gt;&lt;w:commentReference w:id=&quot;0&quot;/&gt;&lt;/w:r&gt;">\u{1F4AC}</span>after</p>';
    const out = htmlToDocx(html, makeDocx());
    const files = unzipSync(out);
    const body = strFromU8(files["word/document.xml"]);
    expect(body).toContain("w:commentRangeStart");
    expect(body).toContain("w:commentReference");
    expect(body).toContain("target");
    const cmt = strFromU8(files["word/comments.xml"]);
    expect(cmt).toContain('w:author="Bob"');
    expect(cmt).toContain("needs work");
    expect(strFromU8(files["[Content_Types].xml"])).toContain("comments+xml");
    expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("/comments");
  });

  it("groups replies into threads and parses reactions out of the text", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    const cmt = `<?xml version="1.0"?><w:comments xmlns:w="${W}" xmlns:w14="${W14}">
 <w:comment w:id="0" w:author="Alice"><w:p w14:paraId="AA"><w:r><w:t>top comment</w:t></w:r></w:p>
   <w:p><w:r><w:t>Nicolas a réagi avec 👍 à 2026-06-12</w:t></w:r></w:p></w:comment>
 <w:comment w:id="1" w:author="Bob"><w:p w14:paraId="BB"><w:r><w:t>a reply</w:t></w:r></w:p></w:comment>
</w:comments>`;
    const cmtEx = `<?xml version="1.0"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
 <w15:commentEx w15:paraId="AA"/><w15:commentEx w15:paraId="BB" w15:paraIdParent="AA"/></w15:commentsEx>`;
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="${W}"><w:body>
 <w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
   <w:commentRangeStart w:id="1"/><w:r><w:t>y</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r></w:p>
</w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
      "word/comments.xml": strToU8(cmt),
      "word/commentsExtended.xml": strToU8(cmtEx),
    };
    const parts = docxToParts(zipSync(files));
    expect(parts.comments.length).toBe(1); // one thread
    expect(parts.comments[0]!.id).toBe("0");
    expect(parts.comments[0]!.text).toBe("top comment"); // reaction line stripped
    expect(parts.comments[0]!.reactions).toEqual([{ emoji: "👍", people: ["Nicolas"] }]);
    expect(parts.comments[0]!.replies.map((r) => r.id)).toEqual(["1"]);
  });

  it("appends an emoji reaction to an existing comment on save", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const files = {
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
      "word/comments.xml": strToU8(`<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`),
    };
    const out = htmlToDocx("<p>hi edited</p>", zipSync(files), [], {
      reactions: [{ commentId: "0", emoji: "👍", person: "Me", date: "2026-06-23" }],
    });
    const cmt = strFromU8(unzipSync(out)["word/comments.xml"]);
    expect(cmt).toContain("👍");
    expect(cmt).toContain("Me a réagi avec");
  });

  it("adds a reply (comments.xml + commentsExtended), resolves and deletes", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    const base = {
      "[Content_Types].xml": strToU8(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
      "word/comments.xml": strToU8(
        `<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="0" w:author="A"><w:p w14:paraId="AA"><w:r><w:t>note</w:t></w:r></w:p></w:comment><w:comment w:id="9" w:author="B"><w:p w14:paraId="ZZ"><w:r><w:t>kill me</w:t></w:r></w:p></w:comment></w:comments>`,
      ),
    };
    // reply to comment 0 (paraId AA)
    const out1 = htmlToDocx("<p>hi</p>", zipSync(base), [], {
      replies: [{ id: "5", paraId: "BB", parentParaId: "AA", author: "Me", date: "2026-06-23", text: "my reply" }],
    });
    const f1 = unzipSync(out1);
    expect(strFromU8(f1["word/comments.xml"])).toContain("my reply");
    expect(strFromU8(f1["word/commentsExtended.xml"])).toContain('w15:paraIdParent="AA"');
    // resolve thread AA
    const out2 = htmlToDocx("<p>hi</p>", zipSync(base), [], { done: new Map([["AA", true]]) });
    expect(strFromU8(unzipSync(out2)["word/commentsExtended.xml"])).toMatch(/w15:paraId="AA"[^>]*w15:done="1"/);
    // delete comment 9
    const out3 = htmlToDocx("<p>hi</p>", zipSync(base), [], { deletedComments: ["9"] });
    const cmt3 = strFromU8(unzipSync(out3)["word/comments.xml"]);
    expect(cmt3).not.toContain("kill me");
    expect(cmt3).toContain("note");
  });

  it("renders and round-trips tracked changes (w:ins / w:del)", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:r><w:t>keep </w:t></w:r><w:ins w:id="1" w:author="Al" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" w:author="Bo"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('<ins class="docx-ins"');
    expect(html).toContain("added");
    expect(html).toContain('<del class="docx-del"');
    expect(html).toContain("gone");
    // round-trip back to w:ins / w:del (delText for deletions)
    const out = htmlToDocx(html, makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("<w:del");
    expect(xml).toContain("w:delText");
    expect(xml).toContain("added");
    expect(xml).toContain("gone");
    // accepting an insertion = unwrap it; rejecting a deletion = unwrap it -> plain text
    const accepted = htmlToDocx("<p>keep added</p>", makeDocx(doc));
    const axml = strFromU8(unzipSync(accepted)["word/document.xml"]);
    expect(axml).not.toContain("<w:ins");
    expect(axml).toContain("added");
  });

  it("round-trips a paragraph-mark revision (split) and a formatting change", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Al" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>first half</w:t></w:r></w:p>
 <w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Bo"><w:rPr/></w:rPrChange></w:rPr><w:t>now bold</w:t></w:r></w:p>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-rev-para="ins"');
    expect(html).toContain('docx-rpr-change');
    const out = htmlToDocx(html, makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    // paragraph mark insertion preserved (w:ins inside pPr > rPr)
    expect(xml).toMatch(/<w:pPr><w:rPr><w:ins[^>]*\/><\/w:rPr><\/w:pPr>/);
    // formatting change preserved (new bold + rPrChange with old)
    expect(xml).toContain("w:rPrChange");
    expect(xml).toContain("now bold");
  });

  it("preserves bookmarks, fields and other unmodelled content through an edit", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:bookmarkStart w:id="0" w:name="mark1"/><w:r><w:t>before </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple><w:r><w:t> after</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
 <w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain("before");
    expect(html).toContain("data-docx-xml"); // bookmark/field/sdt carried as passthrough
    // edit the surrounding text and save: the unmodelled bits must survive
    const out = htmlToDocx(html.replace("before", "BEFORE"), makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("BEFORE");
    expect(xml).toContain("w:bookmarkStart");
    expect(xml).toContain("w:bookmarkEnd");
    expect(xml).toContain("w:fldSimple");
    expect(xml).toContain("w:sdt");
  });

  it("de-obfuscates an .odttf font and leaves plain .ttf untouched", () => {
    const hex = "0123456789ABCDEF0123456789ABCDEF";
    const guid = `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
    const key: number[] = [];
    for (let i = 0; i < 16; i++) key.push(parseInt(hex.substring(i * 2, i * 2 + 2), 16));
    key.reverse();
    const plain = Uint8Array.from({ length: 40 }, (_, i) => i & 0xff);
    const obf = plain.slice();
    for (let i = 0; i < 32; i++) obf[i] ^= key[i % 16]!;
    expect(Array.from(deobfuscateFont(obf, "fonts/x.odttf", guid))).toEqual(Array.from(plain));
    // plain .ttf with a zero key is returned unchanged
    expect(Array.from(deobfuscateFont(plain, "fonts/x.ttf", "{00000000-0000-0000-0000-000000000000}"))).toEqual(Array.from(plain));
  });

  it("adds a hyperlink relationship and reads it back", () => {
    const out = htmlToDocx('<p>see <a href="https://example.com/x">link</a></p>', makeDocx());
    const files = unzipSync(out);
    const rels = strFromU8(files["word/_rels/document.xml.rels"]);
    expect(rels).toContain("https://example.com/x");
    expect(rels).toContain("hyperlink");
    const html = docxToHtml(out);
    expect(html).toContain('href="https://example.com/x"');
    expect(html).toContain(">link</a>");
  });
});

describe("page geometry (w:sectPr)", () => {
  it("reads A4 page size from w:pgSz and defaults margins to 1in", () => {
    // the shared fixture declares <w:pgSz w:w="11906" w:h="16838"/> (A4) with no w:pgMar
    const page = docxToParts(makeDocx()).page;
    expect(page).toBeTruthy();
    expect(page!.widthPx).toBe(794); // 11906 twips / 15
    expect(page!.heightPx).toBe(1123); // 16838 twips / 15
    expect(page!.margin).toEqual({ top: 96, right: 96, bottom: 96, left: 96 });
  });

  it("reads explicit margins from w:pgMar (twips -> px)", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>
</w:body></w:document>`;
    const page = docxToParts(makeDocx(doc)).page;
    expect(page!.widthPx).toBe(816); // 12240 / 15 (US Letter)
    expect(page!.heightPx).toBe(1056); // 15840 / 15
    expect(page!.margin).toEqual({ top: 96, right: 120, bottom: 96, left: 120 });
  });

  it("reads section columns from w:cols (count + gap)", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2" w:space="720"/></w:sectPr>
</w:body></w:document>`;
    const page = docxToParts(makeDocx(doc)).page;
    expect(page!.columns).toBe(2);
    expect(page!.columnGapPx).toBe(48); // 720 twips = 48px
  });

  it("returns no geometry when there is no w:pgSz", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
</w:body></w:document>`;
    expect(docxToParts(makeDocx(doc)).page).toBeUndefined();
  });

  it("reads a page border from w:pgBorders (top side, sz eighths -> px)", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgBorders w:offsetFrom="page"><w:top w:val="double" w:sz="12" w:space="24" w:color="FF0000"/><w:left w:val="double" w:sz="12" w:space="24" w:color="FF0000"/><w:bottom w:val="double" w:sz="12" w:space="24" w:color="FF0000"/><w:right w:val="double" w:sz="12" w:space="24" w:color="FF0000"/></w:pgBorders></w:sectPr>
</w:body></w:document>`;
    const pb = docxToParts(makeDocx(doc)).page!.pageBorder!;
    expect(pb.style).toBe("double");
    expect(pb.widthPx).toBe(2); // sz 12 eighths = 1.5pt = 2px
    expect(pb.color).toBe("FF0000");
    expect(pb.spacePt).toBe(24);
  });
});

describe("page border write-back (w:pgBorders)", () => {
  const withBorder = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="6" w:space="24" w:color="000000"/></w:pgBorders></w:sectPr>
</w:body></w:document>`;

  it("writes a page border into w:pgBorders, after w:pgMar and before w:cols", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(), undefined, {
      pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 }, pageBorder: { style: "solid", widthPx: 2, color: "0000FF", spacePt: 20 } },
    });
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain("<w:pgBorders");
    const top = /<w:top\b[^>]*\/>/.exec(xml)?.[0] ?? "";
    expect(top).toContain('w:val="single"');
    expect(top).toContain('w:sz="12"'); // 2px -> 12 eighths of a point
    expect(top).toContain('w:space="20"');
    expect(top).toContain('w:color="0000FF"');
    expect(xml).toMatch(/<w:right\b/); // all four sides emitted
    expect(xml.indexOf("w:pgMar")).toBeLessThan(xml.indexOf("w:pgBorders")); // schema order
    expect(xml.indexOf("w:pgBorders")).toBeLessThan(xml.indexOf("w:cols"));
  });

  it("removes the page border when the geometry has none", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(withBorder), undefined, {
      pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 } },
    });
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).not.toContain("w:pgBorders");
  });
});

describe("page margin write-back (w:pgMar)", () => {
  it("writes edited margins into the section's w:pgMar (px -> twips)", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(), undefined, {
      pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 48, right: 48, bottom: 48, left: 48 } },
    });
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toMatch(/w:pgMar[^>]*w:top="720"/); // 48px * 15 = 720 twips
    expect(xml).toMatch(/w:pgMar[^>]*w:left="720"/);
  });
});

describe("section columns write-back (w:cols)", () => {
  const unequalDoc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2" w:equalWidth="0" w:sep="1"><w:col w:w="3000" w:space="200"/><w:col w:w="6000"/></w:cols></w:sectPr>
</w:body></w:document>`;

  it("preserves an unequal-width / separated column layout when the count is unchanged", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(unequalDoc), undefined, {
      pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 48, right: 48, bottom: 48, left: 48 }, columns: 2, columnGapPx: 36 },
    });
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toMatch(/w:pgMar[^>]*w:top="720"/); // the unrelated margin edit still applies
    expect(xml).toContain('w:w="3000"'); // the custom column widths survive
    expect(xml).toContain('w:w="6000"');
    expect(xml).toMatch(/w:cols[^>]*w:equalWidth="0"/); // not flattened to equal width
    expect(xml).not.toMatch(/w:equalWidth="1"/);
  });

  it("regenerates equal-width columns when the count actually changes", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(unequalDoc), undefined, {
      pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 }, columns: 3, columnGapPx: 48 },
    });
    const xml = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(xml).toMatch(/w:cols[^>]*w:num="3"/);
    expect(xml).toMatch(/w:cols[^>]*w:equalWidth="1"/);
    expect(xml).not.toContain('w:w="3000"'); // stale per-column widths are dropped
    expect(xml).toContain('w:sep="1"'); // the separator-line flag is count-independent, kept
  });
});

describe("create a header/footer from scratch", () => {
  it("adds a new header part, its relationship, a section reference, and a content type", () => {
    // The "header" path is the sentinel the engine sends for a band created in-editor.
    const out = htmlToDocx("<p>x</p>", makeDocx(), [{ path: "header", html: "<p>My header</p>" }]);
    const files = unzipSync(out);
    const part = strFromU8(files["word/header1.xml"]!);
    expect(part).toContain("<w:hdr");
    expect(part).toContain("My header");
    const rels = strFromU8(files["word/_rels/document.xml.rels"]!);
    expect(rels).toContain("header1.xml");
    expect(rels).toContain("relationships/header");
    expect(strFromU8(files["word/document.xml"]!)).toContain("w:headerReference");
    expect(strFromU8(files["[Content_Types].xml"]!)).toContain("/word/header1.xml");
  });

  it("creates a footer part with w:ftr when the footer sentinel is sent", () => {
    const out = htmlToDocx("<p>x</p>", makeDocx(), [{ path: "footer", html: "<p>My footer</p>" }]);
    const files = unzipSync(out);
    expect(strFromU8(files["word/footer1.xml"]!)).toContain("<w:ftr");
    expect(strFromU8(files["word/footer1.xml"]!)).toContain("My footer");
    expect(strFromU8(files["word/document.xml"]!)).toContain("w:footerReference");
  });
});

describe("editable tables (cell content round-trip)", () => {
  const TABLE_DOC = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>
  <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p><w:r><w:t>after</w:t></w:r></w:p>
</w:body></w:document>`;

  it("renders a table with editable cells", () => {
    const html = docxToHtml(makeDocx(TABLE_DOC));
    expect(html).toContain('class="docx-table"');
    expect(html).toContain('class="docx-cell"');
    expect(html).toContain("A1");
    expect(html).toContain("B2");
  });

  it("builds a fresh w:tbl from a table inserted in the editor (no skeleton)", () => {
    const html = '<p>x</p><table class="docx-table"><tr><td><div class="docx-cell">X</div></td><td><div class="docx-cell">Y</div></td></tr><tr><td><div class="docx-cell">Z</div></td><td><div class="docx-cell">W</div></td></tr></table>';
    const out = htmlToDocx(html, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain("<w:tbl");
    expect(xml).toContain("<w:tblGrid");
    expect((xml.match(/<w:gridCol\b/g) || []).length).toBe(2);
    expect((xml.match(/<w:tr\b/g) || []).length).toBe(2);
    expect((xml.match(/<w:tc\b/g) || []).length).toBe(4);
    expect(xml).toContain("X");
    expect(xml).toContain("W");
  });

  it("round-trips a vertical merge (w:vMerge <-> rowspan) through the grid", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>M</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl></w:body></w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('rowspan="2"'); // the restart cell spans both rows
    const fromDom = html.replace(/ data-docx-xml="[^"]*"/, ""); // simulate a structural edit
    const out = htmlToDocx(fromDom, makeDocx(doc));
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain('w:vMerge w:val="restart"');
    expect((xml.match(/<w:vMerge\b/g) || []).length).toBe(2); // restart + one continue cell
    expect((xml.match(/<w:tr\b/g) || []).length).toBe(2);
  });

  it("writes back edited cell content while keeping structure (tblGrid, rows, spans)", () => {
    const docx = makeDocx(TABLE_DOC);
    const html = docxToHtml(docx).replace(">A1<", ">EDITED<");
    const out = htmlToDocx(html, docx);
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain("EDITED");
    expect(xml).not.toContain(">A1<");
    expect(xml).toContain("B2"); // other cells intact
    expect(xml).toContain("w:tblGrid"); // grid preserved
    expect((xml.match(/<w:tr\b/g) || []).length).toBe(2); // still 2 rows
    expect((xml.match(/<w:tc\b/g) || []).length).toBe(4); // still 4 cells
  });
});

describe("paragraph indent + line spacing", () => {
  it("writes w:ind / w:spacing and reads them back", () => {
    const out = htmlToDocx('<p style="margin-left:48px;line-height:1.5">x</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:ind[^>]*w:left="720"/); // 48px * 15 = 720 twips
    expect(xml).toMatch(/w:spacing[^>]*w:line="360"[^>]*w:lineRule="auto"|w:spacing[^>]*w:lineRule="auto"[^>]*w:line="360"/); // 1.5 * 240
    const html = docxToHtml(out);
    expect(html).toMatch(/margin-left:\s*48px/);
    expect(html).toMatch(/line-height:\s*1\.5/);
  });

  it("writes paragraph space before/after (w:spacing @before/@after) and reads it back", () => {
    const out = htmlToDocx('<p style="margin-top:14px;margin-bottom:7px">x</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:spacing[^>]*w:before="210"/); // 14px * 15
    expect(xml).toMatch(/w:spacing[^>]*w:after="105"/); // 7px * 15
    const html = docxToHtml(out);
    expect(html).toMatch(/margin-top:\s*14px/);
    expect(html).toMatch(/margin-bottom:\s*7px/);
  });

  it("round-trips an explicit zero space-before (no-space paragraph)", () => {
    const out = htmlToDocx('<p style="margin-top:0px">x</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:spacing[^>]*w:before="0"/);
    expect(docxToHtml(out)).toMatch(/margin-top:\s*0px/);
  });
});

describe("run formatting: strike, superscript, subscript", () => {
  it("writes w:strike / w:vertAlign and reads them back", () => {
    const out = htmlToDocx("<p><s>a</s><sup>b</sup><sub>c</sub></p>", makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain("<w:strike");
    expect(xml).toMatch(/w:vertAlign[^>]*w:val="superscript"/);
    expect(xml).toMatch(/w:vertAlign[^>]*w:val="subscript"/);
    const html = docxToHtml(out);
    expect(html).toContain("<s>a</s>");
    expect(html).toContain("<sup>b</sup>");
    expect(html).toContain("<sub>c</sub>");
  });
});

describe("named paragraph styles", () => {
  const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="Hidden"/><w:semiHidden/></w:style>
</w:styles>`;
  const DOC = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body><w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>quoted</w:t></w:r></w:p>
 <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  const makeStyledDocx = () => zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(DOC),
    "word/_rels/document.xml.rels": strToU8(RELS),
    "word/styles.xml": strToU8(STYLES),
  });

  it("lists the document's paragraph styles (excluding headings, default, hidden)", () => {
    const parts = docxToParts(makeStyledDocx());
    const ids = (parts.paragraphStyles ?? []).map((s) => s.id);
    expect(ids).toContain("Quote");
    expect(ids).not.toContain("Normal"); // default
    expect(ids).not.toContain("ListBullet"); // semiHidden
    expect(parts.styleCss).toMatch(/\[data-rdoc-style="Quote"\]\{[^}]*font-style:italic/);
  });

  it("reads a styled paragraph as data-rdoc-style and writes it back as w:pStyle", () => {
    const html = docxToHtml(makeStyledDocx());
    expect(html).toMatch(/<p[^>]*data-rdoc-style="Quote"[^>]*>quoted<\/p>/);
    const out = htmlToDocx(html, makeStyledDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:pStyle[^>]*w:val="Quote"/);
  });
});

describe("authoring new styles", () => {
  const EMPTY_STYLES = `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`;
  const make = (doc: string) => zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(doc),
    "word/_rels/document.xml.rels": strToU8(RELS),
    "word/styles.xml": strToU8(EMPTY_STYLES),
  });
  const P = (body: string) => `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

  it("adds an authored paragraph style to styles.xml and reads it back", () => {
    const out = htmlToDocx('<p data-rdoc-style="MyHeading">Hi</p>', make(P("<w:p/>")), undefined, {
      newStyles: [{ id: "MyHeading", name: "My Heading", kind: "paragraph", css: { "text-align": "center", "font-weight": "bold", "margin-left": "24px" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(stylesXml).toMatch(/w:style[^>]*w:type="paragraph"[^>]*w:styleId="MyHeading"|w:style[^>]*w:styleId="MyHeading"[^>]*w:type="paragraph"/);
    expect(stylesXml).toMatch(/w:jc[^>]*w:val="center"/);
    expect(stylesXml).toContain("<w:b");
    expect(stylesXml).toMatch(/w:ind[^>]*w:left="360"/); // 24px * 15
    // the picker now lists it
    expect(docxToParts(out).paragraphStyles?.some((s) => s.id === "MyHeading")).toBe(true);
  });

  it("adds an authored character style to styles.xml", () => {
    const out = htmlToDocx('<p>a <span data-rdoc-cstyle="Em">b</span></p>', make(P("<w:p/>")), undefined, {
      newStyles: [{ id: "Em", name: "Emph", kind: "character", css: { "font-style": "italic", color: "#cc0000" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(stylesXml).toMatch(/w:style[^>]*w:type="character"[^>]*w:styleId="Em"|w:style[^>]*w:styleId="Em"[^>]*w:type="character"/);
    expect(stylesXml).toContain("<w:i");
    expect(stylesXml).toMatch(/w:color[^>]*w:val="cc0000"/);
    expect(docxToParts(out).characterStyles?.some((s) => s.id === "Em")).toBe(true);
  });

  it("writes background colour as shading (paragraph shading / run shading) and reads it back", () => {
    const out = htmlToDocx('<p data-rdoc-style="Box">x</p><p>y <span data-rdoc-cstyle="Hi">z</span></p>', make(P("<w:p/>")), undefined, {
      newStyles: [
        { id: "Box", name: "Box", kind: "paragraph", css: { "background-color": "#ffeecc" } },
        { id: "Hi", name: "Hi", kind: "character", css: { "background-color": "#ffff00" } },
      ],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(stylesXml).toMatch(/w:shd[^>]*w:fill="ffeecc"/);
    expect(stylesXml).toMatch(/w:shd[^>]*w:fill="ffff00"/);
    const defs = docxToParts(out).styleDefs ?? [];
    expect(defs.find((d) => d.id === "Box")?.css["background-color"]).toBe("#ffeecc");
    expect(defs.find((d) => d.id === "Hi")?.css["background-color"]).toBe("#ffff00");
  });

  it("writes a paragraph style's border as a w:pBdr and reads it back", () => {
    const out = htmlToDocx('<p data-rdoc-style="Framed">x</p>', make(P("<w:p/>")), undefined, {
      newStyles: [{ id: "Framed", name: "Framed", kind: "paragraph", css: { "border-top": "1px solid #ff0000", "border-bottom": "1px solid #ff0000", padding: "2px 6px" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(stylesXml).toMatch(/<w:pBdr>/);
    expect(stylesXml).toMatch(/<w:top[^>]*w:val="single"/);
    expect(stylesXml).toMatch(/<w:top[^>]*w:color="FF0000"/i);
    const def = (docxToParts(out).styleDefs ?? []).find((d) => d.id === "Framed");
    expect(def?.css["border-top"]?.toLowerCase()).toBe("1px solid #ff0000"); // round-trips into the dialog's model
  });

  it("writes a paragraph style's tab stops as w:tabs and reads them back", () => {
    const out = htmlToDocx('<p data-rdoc-style="Tabbed">x</p>', make(P("<w:p/>")), undefined, {
      newStyles: [{ id: "Tabbed", name: "Tabbed", kind: "paragraph", css: { "--rdoc-tabstops": JSON.stringify([{ pos: 96, val: "right", leader: "dot" }]) } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(stylesXml).toMatch(/<w:tabs>/);
    expect(stylesXml).toMatch(/<w:tab[^>]*w:val="right"/);
    expect(stylesXml).toMatch(/<w:tab[^>]*w:pos="1440"/); // 96px * 15 = 1440 twips
    const def = (docxToParts(out).styleDefs ?? []).find((d) => d.id === "Tabbed");
    expect(JSON.parse(def?.css["--rdoc-tabstops"] ?? "[]")[0]?.pos).toBe(96); // round-trips into the style def
  });

  it("editing an existing style replaces its definition in place (no duplicate)", () => {
    const STYLES = `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(P('<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>q</w:t></w:r></w:p>')),
      "word/_rels/document.xml.rels": strToU8(RELS),
      "word/styles.xml": strToU8(STYLES),
    });
    const out = htmlToDocx('<p data-rdoc-style="Quote">q</p>', docx, undefined, {
      newStyles: [{ id: "Quote", name: "Quote", kind: "paragraph", css: { "font-weight": "bold", "text-align": "center" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect((stylesXml.match(/w:styleId="Quote"/g) ?? []).length).toBe(1); // not duplicated
    expect(stylesXml).toContain("<w:b"); // new prop present
    expect(stylesXml).toMatch(/w:jc[^>]*w:val="center"/);
    expect(stylesXml).not.toContain("<w:i"); // old prop replaced
  });

  it("editing a style keeps its unmodelled properties, basedOn, and does not flatten inherited props", () => {
    const STYLES = `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="44"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:smallCaps/><w:b/></w:rPr></w:style></w:styles>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(P('<w:p><w:pPr><w:pStyle w:val="Fancy"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>')),
      "word/_rels/document.xml.rels": strToU8(RELS),
      "word/styles.xml": strToU8(STYLES),
    });
    // The edit dialog sees Fancy's OWN props only (bold), not the inherited size.
    const def = (docxToParts(docx).styleDefs ?? []).find((d) => d.id === "Fancy");
    expect(def?.css["font-weight"]).toBe("bold");
    expect(def?.css["font-size"]).toBeUndefined(); // inherited from Normal, not flattened in
    // Edit Fancy: add italic, keep bold.
    const out = htmlToDocx('<p data-rdoc-style="Fancy">x</p>', docx, undefined, {
      newStyles: [{ id: "Fancy", name: "Fancy", kind: "paragraph", css: { "font-weight": "bold", "font-style": "italic" } }],
    });
    const xml = strFromU8(unzipSync(out)["word/styles.xml"]!);
    expect(xml).toContain("<w:keepNext"); // unmodelled paragraph prop preserved
    expect(xml).toContain("<w:smallCaps"); // unmodelled run prop preserved
    expect(xml).toMatch(/w:basedOn[^>]*w:val="Normal"/); // inheritance preserved
    expect(xml).toContain("<w:b");
    expect(xml).toContain("<w:i");
    const fancyEl = xml.slice(xml.indexOf('w:styleId="Fancy"')).split("</w:style>")[0]!;
    expect(fancyEl).not.toContain("w:sz"); // the inherited size is not copied into Fancy
  });
});

describe("named character styles", () => {
  const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;
  const DOC = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body><w:p><w:r><w:t xml:space="preserve">a </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>bold</w:t></w:r></w:p>
 <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  const makeStyledDocx = () => zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(DOC),
    "word/_rels/document.xml.rels": strToU8(RELS),
    "word/styles.xml": strToU8(STYLES),
  });

  it("lists character styles and emits their CSS", () => {
    const parts = docxToParts(makeStyledDocx());
    expect((parts.characterStyles ?? []).map((s) => s.id)).toContain("Strong");
    expect(parts.styleCss).toMatch(/\[data-rdoc-cstyle="Strong"\]\{[^}]*font-weight:bold/);
  });

  it("reads a styled run as data-rdoc-cstyle and writes it back as w:rStyle", () => {
    const html = docxToHtml(makeStyledDocx());
    expect(html).toMatch(/data-rdoc-cstyle="Strong"/);
    const out = htmlToDocx(html, makeStyledDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:rStyle[^>]*w:val="Strong"/);
  });
});

describe("list fidelity: nesting and ordered/bullet", () => {
  it("writes numbering.xml with a bullet and an ordered list and registers the part", () => {
    const out = htmlToDocx("<ul><li>a</li></ul><ol><li>b</li></ol>", makeDocx());
    const files = unzipSync(out);
    const num = strFromU8(files["word/numbering.xml"]!);
    expect(num).toMatch(/w:numFmt[^>]*w:val="bullet"/);
    expect(num).toMatch(/w:numFmt[^>]*w:val="decimal"/);
    // the part is declared and related so Word opens it
    expect(strFromU8(files["[Content_Types].xml"]!)).toContain("/word/numbering.xml");
    expect(strFromU8(files["word/_rels/document.xml.rels"]!)).toContain("numbering.xml");
    // both list types reference a numId via numPr
    const doc = strFromU8(files["word/document.xml"]!);
    expect(doc).toMatch(/w:numId/);
  });

  it("round-trips a nested list, preserving levels and ordered/bullet kind", () => {
    const html = "<ul><li>top<ol><li>sub1</li><li>sub2</li></ol></li><li>top2</li></ul>";
    const out = htmlToDocx(html, makeDocx());
    const doc = strFromU8(unzipSync(out)["word/document.xml"]!);
    // a level-1 paragraph exists (nested item)
    expect(doc).toMatch(/w:ilvl[^>]*w:val="1"/);
    const back = docxToHtml(out);
    // nested ordered list survives inside the outer unordered list
    expect(back).toMatch(/<ul>[\s\S]*<ol>[\s\S]*sub1[\s\S]*<\/ol>[\s\S]*<\/ul>/);
    expect(back).toContain("top2");
  });

  it("reuses an existing numbering.xml rather than duplicating numIds", () => {
    const existingNum = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
 <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
 <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
 <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
    const base = unzipSync(makeDocx());
    base["word/numbering.xml"] = strToU8(existingNum);
    const out = htmlToDocx("<ul><li>a</li></ul><ol><li>b</li></ol>", zipSync(base));
    const num = strFromU8(unzipSync(out)["word/numbering.xml"]!);
    // no new w:num beyond the two that already existed
    expect((num.match(/<w:num /g) ?? []).length).toBe(2);
    const doc = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(doc).toMatch(/w:numId[^>]*w:val="1"/); // bullet reused
    expect(doc).toMatch(/w:numId[^>]*w:val="2"/); // ordered reused
  });

  it("reads first-page and even/odd header & footer variants", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
      `<w:sectPr><w:headerReference w:type="default" r:id="rH1"/><w:headerReference w:type="first" r:id="rH2"/><w:headerReference w:type="even" r:id="rH3"/><w:titlePg/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
    const hdr = (t: string) => `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rH3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/></Relationships>`;
    const zip = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/settings.xml": strToU8('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>'),
      "word/header1.xml": strToU8(hdr("DEF")),
      "word/header2.xml": strToU8(hdr("FIRST")),
      "word/header3.xml": strToU8(hdr("EVEN")),
      "word/_rels/document.xml.rels": strToU8(rels),
    });
    const parts = docxToParts(zip);
    expect(parts.headerFirst?.html).toContain("FIRST");
    expect(parts.headerEven?.html).toContain("EVEN");
    expect(parts.page?.titlePage).toBe(true);
    expect(parts.page?.evenOdd).toBe(true);
  });

  it("writes header/footer variant parts + their flags from the model", () => {
    const out = htmlToDocx("<p>Body</p>", makeDocx(), [
      { path: "header:first", html: "<p>FIRSTH</p>" },
      { path: "footer:even", html: "<p>EVENF</p>" },
    ], { pageGeometry: { widthPx: 794, heightPx: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 }, titlePage: true, evenOdd: true } });
    const files = unzipSync(out);
    const docXml = strFromU8(files["word/document.xml"]!);
    expect(docXml).toContain("w:titlePg");
    expect(docXml).toMatch(/headerReference[^>]*w:type="first"/);
    expect(docXml).toMatch(/footerReference[^>]*w:type="even"/);
    expect(strFromU8(files["word/settings.xml"]!)).toContain("w:evenAndOddHeaders");
    const partTexts = Object.entries(files).filter(([k]) => /word\/(header|footer)\d+\.xml/.test(k)).map(([, v]) => strFromU8(v)).join("\n");
    expect(partTexts).toContain("FIRSTH");
    expect(partTexts).toContain("EVENF");
  });

  it("reads a bookmark + a REF/PAGEREF cross-reference, dropping Word's _GoBack", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/><w:bookmarkStart w:id="1" w:name="intro"/><w:r><w:t>Intro</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
  <w:p><w:fldSimple w:instr=" REF intro \\h "><w:r><w:t>Intro</w:t></w:r></w:fldSimple> and <w:fldSimple w:instr=" PAGEREF intro "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
 </w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-rdoc-bm="intro"');
    expect(html).toContain("docx-bookmark-end");
    expect(html).not.toContain("_GoBack"); // transient bookmark dropped
    expect(html).toContain('data-rdoc-xref="intro"');
    expect(html).toContain('data-rdoc-xref-fmt="text"');
    expect(html).toContain('data-rdoc-xref-fmt="page"');
  });

  it("writes a bookmark range and REF/PAGEREF cross-references back", () => {
    const body =
      '<p><a class="docx-bookmark" data-rdoc-bm="intro" data-rdoc-bm-id="1" contenteditable="false"></a>Intro' +
      '<a class="docx-bookmark-end" data-rdoc-bm-id="1" data-rdoc-bm-end="intro" contenteditable="false"></a></p>' +
      '<p><a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="text" contenteditable="false">Intro</a> ' +
      '<a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="page" contenteditable="false">1</a></p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:bookmarkStart[^>]*w:name="intro"/);
    expect(xml).toContain("w:bookmarkEnd");
    expect(xml).toMatch(/w:instr=" REF intro/);
    expect(xml).toMatch(/w:instr=" PAGEREF intro/);
    // and it survives a re-read
    const html = docxToHtml(out);
    expect(html).toContain('data-rdoc-bm="intro"');
    expect(html).toContain('data-rdoc-xref="intro"');
  });

  it("reads an OMML equation as a MathML span, keeping the original for a verbatim rewrite", () => {
    const m = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><m:oMath ${m}><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></w:p></w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('class="docx-eq"');
    expect(html).toContain("<mfrac");
    expect(html).toContain("data-docx-xml"); // original OMML kept
    // an un-edited equation rewrites its original OMML verbatim
    const out = htmlToDocx(html, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain("m:oMath");
    expect(xml).toMatch(/<m:f>/);
  });

  it("writes an authored MathML equation as OMML", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq data-latex="x^2" contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><msup><mi>x</mi><mn>2</mn></msup></math></span></p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain("m:oMath");
    expect(xml).toMatch(/<m:sSup>/);
    expect(xml).toMatch(/<m:sup>/);
    // and it re-reads as an equation
    expect(docxToHtml(out)).toContain('class="docx-eq"');
  });

  it("writes an authored MathML matrix as an OMML m:m", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mtable>' +
      "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>" +
      "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr>" +
      "</mtable></math></span></p>";
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<m:m>/);
    expect((xml.match(/<m:mr>/g) ?? []).length).toBe(2);
    expect((xml.match(/<m:e>/g) ?? []).length).toBe(4); // 2x2 cells
    expect(xml).toMatch(/<m:count m:val="2"/);
  });

  it("reads an OMML matrix back as a MathML mtable", () => {
    const m = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><m:oMath ${m}><m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m></m:oMath></w:p></w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain("<mtable>");
    expect(html).toContain("<mtr>");
    expect(html).toContain("<mtd>");
  });

  it("writes a delimited matrix (pmatrix) as an OMML m:d around m:m", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>' +
      '<mo fence="true" stretchy="true">(</mo>' +
      "<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr></mtable>" +
      '<mo fence="true" stretchy="true">)</mo>' +
      "</mrow></math></span></p>";
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<m:d>/);
    expect(xml).toMatch(/<m:begChr m:val="\("/);
    expect(xml).toMatch(/<m:m>/); // the matrix sits inside the delimiter
    // re-reads as a bracketed matrix
    const html = docxToHtml(out);
    expect(html).toContain("<mtable>");
    expect(html).toMatch(/<mo[^>]*>\(<\/mo>/);
  });

  it("round-trips an accent (hat) as an OMML m:acc", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mover accent="true"><mi>x</mi><mo>^</mo></mover></math></span></p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<m:acc>/);
    expect(xml).toMatch(/<m:chr m:val="̂"/); // combining circumflex
    // and it re-reads as an accent over the base
    expect(docxToHtml(out)).toContain("<mover");
  });

  it("round-trips a labeled overbrace as an OMML m:groupChr inside m:limUpp", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mover><mover>' +
      "<mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy=\"true\">⏞</mo></mover><mi>n</mi></mover></math></span></p>";
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<m:limUpp>/);
    expect(xml).toMatch(/<m:groupChr>/);
    expect(xml).toMatch(/<m:pos m:val="top"/);
    // re-reads as a brace with a label above
    const html = docxToHtml(out);
    expect(html).toContain("<mover>");
    expect(html).toContain("⏞");
  });

  it("round-trips an overline (menclose) as an OMML m:bar", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><menclose notation="top"><mrow><mi>A</mi><mi>B</mi></mrow></menclose></math></span></p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<m:bar>/);
    expect(xml).toMatch(/<m:pos m:val="top"/);
    expect(docxToHtml(out)).toContain("<menclose");
  });

  it("keeps function names and \\mathrm upright via an OMML run style", () => {
    const body = '<p><span class="docx-eq" data-rdoc-eq contenteditable="false">' +
      '<math xmlns="http://www.w3.org/1998/Math/MathML">' +
      "<mi>sin</mi><mi>x</mi><mi mathvariant=\"normal\">d</mi><mi mathvariant=\"bold\">v</mi>" +
      "</math></span></p>";
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect((xml.match(/<m:sty m:val="p"\/>/g) ?? []).length).toBe(2); // sin (multi-letter) + \mathrm d
    expect(xml).toMatch(/<m:sty m:val="b"\/>/); // bold v
    // the single italic identifier x carries no explicit style
    expect((xml.match(/<m:sty/g) ?? []).length).toBe(3);
    // and an upright run re-reads as a normal-variant identifier
    expect(docxToHtml(out)).toMatch(/<mi mathvariant="normal">d<\/mi>/);
  });

  it("renders a Symbol-font w:sym as Unicode and rewrites it verbatim", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><w:r><w:sym w:font="Symbol" w:char="F061"/></w:r></w:p></w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('class="docx-sym"');
    expect(html).toContain("α"); // Symbol F061 -> Greek alpha
    expect(html).toContain("data-docx-xml"); // the original run is stashed
    // an untouched symbol rewrites its w:sym verbatim
    const out = htmlToDocx(html, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:sym[^>]*w:char="F061"/);
    expect(xml).toMatch(/w:font="Symbol"/);
  });

  it("renders a Wingdings w:sym via the bundled MaterialDings font, rewriting it verbatim", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r></w:p></w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('class="docx-sym docx-dings"');
    expect(html).toContain("font-family:'MaterialDings'");
    expect(html).toContain("J"); // F04A -> low byte 0x4A, MaterialDings' classic codepoint
    const out = htmlToDocx(html, makeDocx());
    expect(strFromU8(unzipSync(out)["word/document.xml"]!)).toMatch(/<w:sym[^>]*w:font="Wingdings"/);
  });

  it("maps Webdings and Wingdings 2/3 to Unicode and rewrites them verbatim", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p>
  <w:r><w:sym w:font="Webdings" w:char="F061"/></w:r>
  <w:r><w:sym w:font="Wingdings 2" w:char="F041"/></w:r>
  <w:r><w:sym w:font="Wingdings 3" w:char="F041"/></w:r>
 </w:p></w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain("✔"); // Webdings F061 -> U+2714
    expect(html).toContain("🖛"); // Wingdings 2 F041 -> U+1F59B
    expect(html).toContain("⮑"); // Wingdings 3 F041 -> U+2B91
    expect(html).not.toContain("docx-dings"); // these use Unicode, not MaterialDings
    expect(html).not.toContain("font-family:'Webdings'"); // mapped, not the named font
    // all three round-trip verbatim
    const out = htmlToDocx(html, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:font="Webdings"/);
    expect(xml).toMatch(/w:font="Wingdings 2"/);
    expect(xml).toMatch(/w:font="Wingdings 3"/);
  });

  it("writes inserted info fields (date / author / file name) as w:fldSimple", () => {
    const body = "<p>" +
      '<span class="docx-field" data-field="DATE" contenteditable="false">6/30/2026</span> ' +
      '<span class="docx-field" data-field="AUTHOR" contenteditable="false">Jane</span> ' +
      '<span class="docx-field" data-field="FILENAME" contenteditable="false">report.docx</span>' +
      "</p>";
    const xml = strFromU8(unzipSync(htmlToDocx(body, makeDocx()))["word/document.xml"]!);
    expect(xml).toMatch(/<w:fldSimple[^>]*w:instr=" DATE "/);
    expect(xml).toMatch(/<w:fldSimple[^>]*w:instr=" AUTHOR "/);
    expect(xml).toMatch(/<w:fldSimple[^>]*w:instr=" FILENAME "/);
    expect(xml).toContain("Jane"); // the cached snapshot survives
  });

  it("round-trips paragraph shading as a pPr w:shd", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="ffeecc"/></w:pPr><w:r><w:t>hi</w:t></w:r></w:p></w:body>
</w:document>`;
    expect(docxToHtml(makeDocx(doc))).toContain("background-color:#ffeecc"); // read -> rendered

    const out = htmlToDocx('<p style="background-color:#ffeecc">hi</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:shd[^>]*w:fill="ffeecc"/i); // authored -> w:shd
    expect(docxToHtml(out)).toContain("background-color:#ffeecc"); // and round-trips
  });

  it("round-trips paragraph borders as a pPr w:pBdr", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:color="000000"/></w:pBdr></w:pPr><w:r><w:t>hi</w:t></w:r></w:p></w:body>
</w:document>`;
    expect(docxToHtml(makeDocx(doc))).toContain("border-top:1px solid #000000"); // read -> rendered

    const out = htmlToDocx('<p style="border-top:1px solid #000000;border-bottom:1px solid #000000;padding:2px 6px">hi</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:pBdr>/); // authored -> w:pBdr
    expect(xml).toMatch(/<w:top[^>]*w:val="single"/);
    expect(xml).toMatch(/<w:bottom[^>]*w:val="single"/);
    expect(docxToHtml(out)).toContain("border-bottom:1px solid #000000"); // and round-trips
  });

  it("round-trips a single-side (left only) paragraph border", () => {
    const out = htmlToDocx('<p style="border-left:2px solid #008000;padding:2px 6px">x</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:pBdr>/);
    expect(xml).toMatch(/<w:left[^>]*w:val="single"/);
    expect(xml).not.toMatch(/<w:top[^>]*w:val=/); // only the left side
    expect(docxToHtml(out)).toMatch(/border-left:2px solid #008000/i);
    expect(docxToHtml(out)).not.toMatch(/border-top:/);
  });

  it("round-trips a paragraph border's width, line style and colour", () => {
    const out = htmlToDocx('<p style="border-top:3px dashed #ff0000;padding:2px 6px">x</p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:top[^>]*w:val="dashed"/);
    expect(xml).toMatch(/<w:top[^>]*w:sz="18"/); // 3px * 6 = 18 eighths of a point
    expect(xml).toMatch(/<w:top[^>]*w:color="FF0000"/i);
    expect(docxToHtml(out)).toMatch(/border-top:3px dashed #ff0000/i); // and round-trips
  });

  it("round-trips an internal hyperlink as a w:hyperlink w:anchor", () => {
    const out = htmlToDocx('<p>See <a href="#intro">the intro</a></p>', makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/<w:hyperlink[^>]*w:anchor="intro"/);
    expect(xml).not.toMatch(/r:id="rId/); // no external relationship was minted
    expect(docxToHtml(out)).toContain('<a href="#intro">');
  });

  it("reads a w:hyperlink w:anchor as an internal link", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body><w:p><w:hyperlink w:anchor="intro"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p></w:body>
</w:document>`;
    expect(docxToHtml(makeDocx(doc))).toContain('<a href="#intro">go</a>');
  });

  it("round-trips an above/below cross-reference via the \\p switch", () => {
    const body = '<p><a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="direction" contenteditable="false">below</a></p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:instr=" REF intro \\p \\h "/);
    expect(docxToHtml(out)).toContain('data-rdoc-xref-fmt="direction"');
  });

  it("reads a SEQ caption paragraph as a caption with a seq field", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Figure </w:t></w:r><w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t>: A chart</w:t></w:r></w:p>
 </w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-rdoc-caption="figure"');
    expect(html).toContain('data-field="seq"');
    expect(html).toContain('data-seq="Figure"');
  });

  it("reads complex (fldChar) REF/PAGEREF and SEQ fields, not just fldSimple", () => {
    const r = (xml: string) => `<w:r>${xml}</w:r>`;
    const ref = (instr: string, cached: string) =>
      r('<w:fldChar w:fldCharType="begin"/>') +
      r(`<w:instrText xml:space="preserve">${instr}</w:instrText>`) +
      r('<w:fldChar w:fldCharType="separate"/>') +
      r(`<w:t>${cached}</w:t>`) +
      r('<w:fldChar w:fldCharType="end"/>');
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p>${ref(" REF intro \\h ", "Introduction")} on page ${ref(" PAGEREF intro ", "3")}</w:p>
  <w:p>Figure ${ref(" SEQ Figure \\* ARABIC ", "2")}: chart</w:p>
  <w:p>Built ${ref(" DATE ", "2026-01-01")}</w:p>
 </w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-rdoc-xref="intro"');
    expect(html).toContain('data-rdoc-xref-fmt="text"');
    expect(html).toContain('data-rdoc-xref-fmt="page"');
    expect(html).toContain(">Introduction</a>"); // the cached result becomes the xref text
    expect(html).toContain('data-rdoc-caption="figure"');
    expect(html).toContain('data-seq="Figure"');
    expect(html).not.toMatch(/REF intro/); // the instruction text is no longer shown
    expect(html).toContain("docx-pass"); // the unmodelled DATE field stays passthrough
  });

  it("reads and writes an equation caption (SEQ Equation)", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Equation </w:t></w:r><w:fldSimple w:instr=" SEQ Equation \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
 </w:body>
</w:document>`;
    const html = docxToHtml(makeDocx(doc));
    expect(html).toContain('data-rdoc-caption="equation"');
    expect(html).toContain('data-seq="Equation"');
    // and authoring one writes a SEQ Equation field back
    const out = htmlToDocx('<p data-rdoc-caption="equation">Equation <span class="docx-field docx-field-seq" data-field="seq" data-seq="Equation" contenteditable="false">1</span></p>', makeDocx());
    expect(strFromU8(unzipSync(out)["word/document.xml"]!)).toMatch(/w:instr=" SEQ Equation/);
  });

  it("writes a caption paragraph back as a Caption-styled SEQ field", () => {
    const body =
      '<p data-rdoc-caption="table">Table <span class="docx-field docx-field-seq" data-field="seq" data-seq="Table" contenteditable="false">1</span>: Results</p>';
    const out = htmlToDocx(body, makeDocx());
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toMatch(/w:instr=" SEQ Table/);
    expect(xml).toMatch(/<w:pStyle w:val="Caption"/);
    const html = docxToHtml(out);
    expect(html).toContain('data-rdoc-caption="table"');
    expect(html).toContain('data-seq="Table"');
  });
});
