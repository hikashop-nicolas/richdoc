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
