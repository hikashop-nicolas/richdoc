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

  it("returns no geometry when there is no w:pgSz", () => {
    const doc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
</w:body></w:document>`;
    expect(docxToParts(makeDocx(doc)).page).toBeUndefined();
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
});
