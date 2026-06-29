import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { htmlToOdt, odtToHtml, odtToParts } from "./index";

const CONTENT = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">
 <office:automatic-styles>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
 </office:automatic-styles>
 <office:body>
  <office:text>
   <text:h text:outline-level="1">Titre</text:h>
   <text:p>Bonjour <text:span text:style-name="T1">monde</text:span> ici.</text:p>
  </office:text>
 </office:body>
</office:document-content>`;

function makeOdt(content = CONTENT): Uint8Array {
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "content.xml": strToU8(content),
    "styles.xml": strToU8("<x/>"),
    "META-INF/manifest.xml": strToU8("<m/>"),
    "extra.bin": new Uint8Array([1, 2, 3, 4]),
  });
}

describe("odt <-> html", () => {
  it("reads the body to HTML, mapping headings and bold", () => {
    const html = odtToHtml(makeOdt());
    expect(html).toContain("<h1>Titre</h1>");
    expect(html).toContain("Bonjour");
    expect(html).toContain("<strong>monde</strong>");
    expect(html).toContain("ici.");
  });

  it("writes edited HTML back to a valid .odt, preserving other parts", () => {
    const odt = makeOdt();
    const edited = "<h1>Titre</h1><p>Bonjour <strong>planete</strong> la.</p>";
    const out = htmlToOdt(edited, odt);

    const files = unzipSync(out);
    // other archive parts are preserved untouched
    expect(strFromU8(files["mimetype"])).toBe("application/vnd.oasis.opendocument.text");
    expect(strFromU8(files["styles.xml"])).toBe("<x/>");
    expect(Array.from(files["extra.bin"])).toEqual([1, 2, 3, 4]);

    // the edit round-trips back through the reader
    const html2 = odtToHtml(out);
    expect(html2).toContain("<h1>Titre</h1>");
    expect(html2).toContain("<strong>planete</strong>");
    expect(html2).toContain("la.");
    expect(html2).not.toContain("monde");
  });

  it("preserves bold/italic/underline added in the editor", () => {
    const out = htmlToOdt("<p><strong>b</strong> <em>i</em> <u>u</u></p>", makeOdt());
    const html = odtToHtml(out);
    expect(html).toContain("<strong>b</strong>");
    expect(html).toContain("<em>i</em>");
    expect(html).toContain("<u>u</u>");
  });

  it("preserves tables, comments and tracked changes through an edit", () => {
    const content = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:body><office:text>
 <text:tracked-changes><text:changed-region text:id="c1"><text:deletion><office:change-info><dc:creator>Al</dc:creator></office:change-info><text:p>gone</text:p></text:deletion></text:changed-region></text:tracked-changes>
 <text:p>before<text:change text:change-id="c1"/><office:annotation><dc:creator>Al</dc:creator><text:p>my note</text:p></office:annotation> after</text:p>
 <table:table table:name="T"><table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell></table:table-row></table:table>
</office:text></office:body></office:document-content>`;
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain("before");
    expect(html).toContain("data-odt-xml");
    const out = htmlToOdt(html.replace("before", "BEFORE"), makeOdt(content));
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain("BEFORE");
    expect(xml).toContain("table:table");
    expect(xml).toContain("A1");
    expect(xml).toContain("office:annotation");
    expect(xml).toContain("text:tracked-changes");
  });

  it("writes per-cell borders, resized column widths and a row height when a table is edited", () => {
    const content = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:body><office:text>
 <text:p>x</text:p>
 <table:table table:name="T"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row></table:table>
</office:text></office:body></office:document-content>`;
    const html = odtToHtml(makeOdt(content));
    const edited = html
      .replace(/ data-odt-xml="[^"]*"/, "")
      .replace('<table class="docx-table"', '<table class="docx-table" style="margin-left: 48px"')
      .replace(/(<table class="docx-table"[^>]*>)/, '$1<colgroup><col style="width: 120px"><col style="width: 90px"></colgroup>')
      .replace(/<tr>/, '<tr style="height: 40px">')
      .replace(/<td>(<div class="docx-cell")/, '<td class="rdoc-bordered" data-rdoc-bt="2px dashed #ff0000" data-rdoc-bl="1px solid #000000">$1');
    const out = htmlToOdt(edited, makeOdt(content));
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain('fo:border-top="0.053cm dashed #ff0000"'); // chosen width/style/colour
    expect(xml).toContain('fo:border-left="0.026cm solid #000000"');
    expect(xml).toContain('fo:border-bottom="none"');
    expect(xml).toContain('style:column-width="3.175cm"'); // 120px
    expect(xml).toContain('style:column-width="2.381cm"'); // 90px
    expect(xml).toContain('style:min-row-height="1.058cm"'); // 40px
    expect(xml).toContain('fo:margin-left="1.27cm"'); // 48px table indent
  });

  it("resolves a cell style's borders into the editor's per-side model on read", () => {
    const content = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>
 <style:style style:name="Tc1" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt dashed #00ff00"/></style:style>
</office:automatic-styles><office:body><office:text>
 <table:table table:name="T"><table:table-column table:number-columns-repeated="1"/><table:table-row><table:table-cell table:style-name="Tc1"><text:p>A1</text:p></table:table-cell></table:table-row></table:table>
</office:text></office:body></office:document-content>`;
    const html = odtToHtml(makeOdt(content));
    // 0.5pt -> 1px, fo:border shorthand applies to all four sides.
    expect(html).toContain('data-rdoc-bt="1px dashed #00ff00"');
    expect(html).toContain('data-rdoc-bl="1px dashed #00ff00"');
    expect(html).toContain('data-rdoc-br="1px dashed #00ff00"');
  });

  it("reads vertical (tategaki) writing-mode from the page layout", () => {
    const styles = `<?xml version="1.0"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>
 <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" style:writing-mode="tb-rl"/></style:page-layout>
</office:automatic-styles></office:document-styles>`;
    const odt = zipSync({ mimetype: strToU8("application/vnd.oasis.opendocument.text"), "content.xml": strToU8(CONTENT), "styles.xml": strToU8(styles) });
    expect(odtToParts(odt).page?.vertical).toBe(true);
  });

  it("reads a right-to-left (rl-tb) page layout", () => {
    const styles = `<?xml version="1.0"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>
 <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" style:writing-mode="rl-tb"/></style:page-layout>
</office:automatic-styles></office:document-styles>`;
    const odt = zipSync({ mimetype: strToU8("application/vnd.oasis.opendocument.text"), "content.xml": strToU8(CONTENT), "styles.xml": strToU8(styles) });
    const page = odtToParts(odt).page;
    expect(page?.rtl).toBe(true);
    expect(page?.vertical).toBe(false);
  });

  it("writes inserted fields (page number/count) and a table of contents", () => {
    const body =
      '<h1>Alpha</h1>' +
      '<div class="docx-field-toc"><div class="docx-field-toc-title">Contents</div>' +
      '<div class="docx-field-toc-row toc-h1"><span class="docx-field-toc-text">Alpha</span><span class="docx-field-toc-page">1</span></div></div>' +
      '<p>Page <span class="docx-field" data-field="PAGE">2</span> / <span class="docx-field" data-field="NUMPAGES">5</span></p>';
    const xml = strFromU8(unzipSync(htmlToOdt(body, makeOdt()))["content.xml"]);
    expect(xml).toContain("<text:page-number");
    expect(xml).toContain("<text:page-count");
    expect(xml).toContain("<text:table-of-content");
    expect(xml).toContain("<text:index-body");
  });

  it("keeps mimetype as the first, uncompressed entry", () => {
    const out = htmlToOdt("<p>x</p>", makeOdt());
    // local file header: signature(4) + ... + compression method at offset 8 (0 = stored)
    // and the first entry's name follows at offset 30.
    expect(out[0]).toBe(0x50); // 'P'
    expect(out[1]).toBe(0x4b); // 'K'
    expect(out[8]).toBe(0); // first entry stored (mimetype)
    const name = strFromU8(out.slice(30, 38));
    expect(name).toBe("mimetype");
  });
});

describe("odt formatting (colour, font, size, alignment)", () => {
  it("writes run colour/font/size into a synthesized text style and round-trips", () => {
    const html = '<p>x<span style="color:#FF0000;background-color:#FFFF00;font-family:\'Arial\';font-size:18pt">y</span></p>';
    const out = htmlToOdt(html, makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain('fo:color="#FF0000"');
    expect(xml).toContain('fo:background-color="#FFFF00"');
    expect(xml).toContain('fo:font-family="Arial"');
    expect(xml).toContain('fo:font-size="18pt"');
    // re-read: the inline styles come back
    const back = odtToHtml(out);
    expect(back).toContain("color:#FF0000");
    expect(back).toContain("background-color:#FFFF00");
    expect(back).toContain("font-size:18pt");
    expect(back).toMatch(/font-family:'?Arial/);
  });

  it("writes paragraph alignment into a synthesized paragraph style and round-trips", () => {
    const out = htmlToOdt('<p style="text-align:center">mid</p><p style="text-align:right">end</p>', makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain('fo:text-align="center"');
    expect(xml).toContain('fo:text-align="end"');
    const back = odtToHtml(out);
    expect(back).toContain("text-align:center");
    expect(back).toContain("text-align:right");
  });

  it("does not emit a style for default (left) alignment", () => {
    const out = htmlToOdt('<p style="text-align:left">x</p>', makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).not.toContain("fo:text-align");
  });

  it("reads existing run colour from an automatic text style", () => {
    const content = CONTENT.replace(
      '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>',
      '<style:style style:name="T1" style:family="text"><style:text-properties fo:color="#0000FF" fo:font-size="14pt"/></style:style>',
    );
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain("color:#0000FF");
    expect(html).toContain("font-size:14pt");
  });
});

describe("odt images (draw:frame)", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
    'xmlns:xlink="http://www.w3.org/1999/xlink"';
  const MANIFEST =
    '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">' +
    '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>';
  const PIC = new Uint8Array([1, 2, 3, 4, 5]);

  function makeImgOdt(content: string): Uint8Array {
    return zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8(MANIFEST),
      "Pictures/p.png": PIC,
    });
  }

  it("renders a draw:frame image as an <img> with a data URL and px size", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}><office:body><office:text>` +
      '<text:p><draw:frame svg:width="2.54cm" svg:height="1.27cm"><draw:image xlink:href="Pictures/p.png"/></draw:frame></text:p>' +
      "</office:text></office:body></office:document-content>";
    const html = odtToHtml(makeImgOdt(content));
    expect(html).toContain('<img src="data:image/png;base64,');
    expect(html).toContain('width="96"'); // 2.54cm = 1in = 96px
    expect(html).toContain('height="48"'); // 1.27cm = 48px
  });

  it("embeds an <img> data URL as a draw:frame + Pictures file + manifest entry", () => {
    const base = `<?xml version="1.0"?><office:document-content ${ROOT}><office:body><office:text><text:p/></office:text></office:body></office:document-content>`;
    const out = htmlToOdt('<p><img src="data:image/png;base64,AQIDBAU=" width="100" height="50"></p>', makeImgOdt(base));
    const files = unzipSync(out);
    const xml = strFromU8(files["content.xml"]);
    expect(xml).toContain("draw:frame");
    expect(xml).toContain('xlink:href="Pictures/ot_img0.png"');
    expect(xml).toContain('svg:width="2.646cm"'); // 100px -> cm
    expect(Array.from(files["Pictures/ot_img0.png"])).toEqual([1, 2, 3, 4, 5]);
    expect(strFromU8(files["META-INF/manifest.xml"])).toContain('manifest:full-path="Pictures/ot_img0.png"');
  });

  // A floating frame: anchored to the paragraph with a graphic style carrying style:wrap.
  const FLOAT_ROOT = ROOT + ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"';
  it("reads a floating draw:frame's wrap mode from its graphic style", () => {
    const content = `<?xml version="1.0"?><office:document-content ${FLOAT_ROOT}>` +
      "<office:automatic-styles>" +
      '<style:style style:name="fr1" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="right"/></style:style>' +
      "</office:automatic-styles><office:body><office:text>" +
      '<text:p>hi<draw:frame text:anchor-type="paragraph" draw:style-name="fr1" svg:width="2.54cm" svg:height="1.27cm"><draw:image xlink:href="Pictures/p.png"/><svg:desc>a cat</svg:desc></draw:frame></text:p>' +
      "</office:text></office:body></office:document-content>";
    const html = odtToHtml(makeImgOdt(content));
    expect(html).toContain('data-rdoc-wrap="square"');
    expect(html).toContain('data-rdoc-align="right"');
    expect(html).toContain('alt="a cat"');
  });

  it("round-trips a floating frame's wrap distances", () => {
    const content = `<?xml version="1.0"?><office:document-content ${FLOAT_ROOT} xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
      "<office:automatic-styles>" +
      '<style:style style:name="fr1" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="left" fo:margin-left="0.635cm" fo:margin-right="0.635cm" fo:margin-top="0cm" fo:margin-bottom="0cm"/></style:style>' +
      "</office:automatic-styles><office:body><office:text>" +
      '<text:p>hi<draw:frame text:anchor-type="paragraph" draw:style-name="fr1" svg:width="2.54cm" svg:height="1.27cm"><draw:image xlink:href="Pictures/p.png"/></draw:frame></text:p>' +
      "</office:text></office:body></office:document-content>";
    const odt = makeImgOdt(content);
    const html = odtToHtml(odt);
    expect(html).toContain('data-rdoc-wrapdist="0,24,0,24"'); // 0.635cm = 24px
    const xml = strFromU8(unzipSync(htmlToOdt(html.replace(">hi<", ">HI<"), odt))["content.xml"]);
    expect(xml).toMatch(/fo:margin-left="0.635cm"/);
    expect(xml).toMatch(/fo:margin-right="0.635cm"/);
  });

  it("writes a wrapped image as an anchored frame + graphic style", () => {
    const base = `<?xml version="1.0"?><office:document-content ${FLOAT_ROOT}><office:body><office:text><text:p/></office:text></office:body></office:document-content>`;
    const out = htmlToOdt('<p><img src="data:image/png;base64,AQIDBAU=" width="100" height="50" data-rdoc-wrap="square" data-rdoc-align="left"></p>', makeImgOdt(base));
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain('text:anchor-type="paragraph"');
    expect(xml).toContain('style:family="graphic"');
    expect(xml).toContain('style:wrap="parallel"');
  });

  it("round-trips a floating frame's wrap through an edit", () => {
    const content = `<?xml version="1.0"?><office:document-content ${FLOAT_ROOT}>` +
      "<office:automatic-styles>" +
      '<style:style style:name="fr1" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="center"/></style:style>' +
      "</office:automatic-styles><office:body><office:text>" +
      '<text:p>hi<draw:frame text:anchor-type="paragraph" draw:style-name="fr1" svg:width="2.54cm" svg:height="1.27cm"><draw:image xlink:href="Pictures/p.png"/></draw:frame></text:p>' +
      "</office:text></office:body></office:document-content>";
    const odt = makeImgOdt(content);
    const html = odtToHtml(odt);
    expect(html).toContain('data-rdoc-wrap="topbottom"');
    const xml = strFromU8(unzipSync(htmlToOdt(html.replace(">hi<", ">HI<"), odt))["content.xml"]);
    expect(xml).toContain('xlink:href="Pictures/p.png"'); // original picture reused
    expect(xml).toContain('style:wrap="none"');
  });
});

describe("odt sections (master page / page break)", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

  it("preserves a master-page section break through an edit", () => {
    // An automatic paragraph style carrying a new page master + a page break (LibreOffice's
    // way of starting a new section), referenced by the second paragraph.
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}>` +
      "<office:automatic-styles>" +
      '<style:style style:name="P2" style:family="paragraph" style:master-page-name="Landscape"><style:paragraph-properties fo:break-before="page"/></style:style>' +
      "</office:automatic-styles><office:body><office:text>" +
      "<text:p>one</text:p><text:p text:style-name=\"P2\">two</text:p>" +
      "</office:text></office:body></office:document-content>";
    const odt = makeOdt(content);
    const html = odtToHtml(odt);
    expect(html).toContain('data-odt-masterpage="Landscape"'); // section surfaced on the paragraph
    expect(html).toContain('data-odt-break-before="page"');
    expect(html).toContain("docx-pagebreak"); // shown as a page boundary
    const xml = strFromU8(unzipSync(htmlToOdt(html.replace(">two<", ">TWO<"), odt))["content.xml"]);
    expect(xml).toContain("TWO");
    expect(xml).toContain('style:master-page-name="Landscape"'); // the section survives the edit
    expect(xml).toContain('fo:break-before="page"');
  });

  it("preserves a plain page break (fo:break-before) through an edit", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}>` +
      "<office:automatic-styles>" +
      '<style:style style:name="P3" style:family="paragraph"><style:paragraph-properties fo:break-before="page"/></style:style>' +
      "</office:automatic-styles><office:body><office:text>" +
      "<text:p>a</text:p><text:p text:style-name=\"P3\">b</text:p>" +
      "</office:text></office:body></office:document-content>";
    const odt = makeOdt(content);
    const html = odtToHtml(odt);
    expect(html).toContain('data-odt-break-before="page"');
    const xml = strFromU8(unzipSync(htmlToOdt(html.replace(">b<", ">B<"), odt))["content.xml"]);
    expect(xml).toContain('fo:break-before="page"');
    expect(xml).toContain("B");
  });
});

describe("odt ordered-list numbering", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"';
  const numStyle = (name: string, start?: number) =>
    `<text:list-style style:name="${name}"><text:list-level-style-number text:level="1" style:num-format="1"${start ? ` text:start-value="${start}"` : ""}/></text:list-style>`;
  const item = (text: string) => `<text:list-item><text:p>${text}</text:p></text:list-item>`;

  it("reads a list's start-value as <ol start>", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}>` +
      `<office:automatic-styles>${numStyle("L1", 4)}</office:automatic-styles><office:body><office:text>` +
      `<text:list text:style-name="L1">${item("a")}${item("b")}</text:list>` +
      "</office:text></office:body></office:document-content>";
    expect(odtToHtml(makeOdt(content))).toContain('<ol start="4"><li>a</li><li>b</li></ol>');
  });

  it("reads continue-numbering as a continued <ol start>", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}>` +
      `<office:automatic-styles>${numStyle("L1")}</office:automatic-styles><office:body><office:text>` +
      `<text:list text:style-name="L1">${item("a")}${item("b")}</text:list>` +
      "<text:p>gap</text:p>" +
      `<text:list text:style-name="L1" text:continue-numbering="true">${item("c")}</text:list>` +
      "</office:text></office:body></office:document-content>";
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain("<ol><li>a</li><li>b</li></ol>"); // first list starts at 1
    expect(html).toContain('<ol start="3"><li>c</li></ol>'); // continues after two items
  });

  it("writes <ol start> as a list style start-value", () => {
    const xml = strFromU8(unzipSync(htmlToOdt('<ol start="4"><li>a</li></ol>', makeOdt()))["content.xml"]);
    expect(xml).toContain('text:start-value="4"');
  });
});

describe("odt tabs (text:tab + tab stops)", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';

  it("round-trips a tab character and the paragraph's tab stops", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ROOT}>` +
      "<office:automatic-styles>" +
      '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties><style:tab-stops>' +
      '<style:tab-stop style:position="3.81cm"/>' +
      '<style:tab-stop style:position="11.43cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>' +
      "</style:tab-stops></style:paragraph-properties></style:style>" +
      "</office:automatic-styles><office:body><office:text>" +
      '<text:p text:style-name="P1">a<text:tab/>b</text:p>' +
      "</office:text></office:body></office:document-content>";
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain('data-docx-tab="1"');
    expect(html).toContain("data-rdoc-tabstops");
    const xml = strFromU8(unzipSync(htmlToOdt(html.replace(">a<", ">A<"), makeOdt(content)))["content.xml"]);
    expect(xml).toContain("<text:tab"); // tab character round-trips
    expect(xml).toContain("style:tab-stops");
    expect(xml).toMatch(/style:position="3.81cm"/); // 3.81cm = 144px, back to 3.81cm
    expect(xml).toMatch(/style:type="right"/);
    expect(xml).toMatch(/style:leader-style="dotted"/);
  });
});

describe("odt comments (office:annotation)", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
    'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';
  const make = (inner: string) =>
    makeOdt(`<?xml version="1.0"?><office:document-content ${ROOT}><office:body><office:text>${inner}</office:text></office:body></office:document-content>`);

  it("reads a ranged annotation into a thread + highlight + ref", () => {
    const odt = make(
      '<text:p>Hi <office:annotation office:name="A1"><dc:creator>Alice</dc:creator><dc:date>2026-06-24T10:00:00</dc:date><text:p>Nice</text:p></office:annotation>there<office:annotation-end office:name="A1"/>!</text:p>',
    );
    const { body, comments } = odtToParts(odt);
    expect(comments).toHaveLength(1);
    expect(comments[0]).toMatchObject({ id: "A1", author: "Alice", text: "Nice", resolved: false });
    expect(body).toContain('class="docx-comment" data-comment-id="A1"');
    expect(body).toContain('class="docx-comment-ref"');
    expect(body).toContain(">there<"); // the commented text stays in the body
    expect(body).toContain('data-comment-text="Nice"'); // comment text is metadata, not body text
    expect(body).toContain(">\u{1F4AC}</span>"); // the comment text is not rendered inline
  });

  it("writes a commented range back to an office:annotation + annotation-end", () => {
    const html =
      '<p>Hi <span class="docx-comment" data-comment-id="A1">there</span>' +
      '<span class="docx-comment-ref" data-comment-id="A1" data-comment-paraid="A1" data-comment-author="Bob" data-comment-date="2026-06-24T00:00:00" data-comment-text="check">\u{1F4AC}</span>!</p>';
    const out = htmlToOdt(html, make("<text:p/>"));
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain('office:annotation office:name="A1"');
    expect(xml).toContain("<dc:creator>Bob</dc:creator>");
    expect(xml).toContain("check");
    expect(xml).toContain('office:annotation-end office:name="A1"');
    expect(xml).toContain("there"); // body text preserved
  });

  it("marks an annotation resolved from the done edit map", () => {
    const html =
      '<p><span class="docx-comment" data-comment-id="A1">x</span>' +
      '<span class="docx-comment-ref" data-comment-id="A1" data-comment-paraid="P1" data-comment-author="A" data-comment-text="t">\u{1F4AC}</span></p>';
    const out = htmlToOdt(html, make("<text:p/>"), { done: new Map([["P1", true]]) });
    expect(strFromU8(unzipSync(out)["content.xml"])).toContain('loext:resolved="true"');
  });

  it("round-trips a comment through read -> write -> read", () => {
    const odt = make(
      '<text:p>a <office:annotation office:name="A1"><dc:creator>Z</dc:creator><text:p>note</text:p></office:annotation>b<office:annotation-end office:name="A1"/> c</text:p>',
    );
    const { body } = odtToParts(odt);
    const out = htmlToOdt(body, odt);
    const again = odtToParts(out);
    expect(again.comments).toHaveLength(1);
    expect(again.comments[0]).toMatchObject({ id: "A1", author: "Z", text: "note" });
  });
});

describe("odt header/footer (styles.xml master page)", () => {
  const STYLES = (header: string, footer: string) =>
    '<?xml version="1.0"?><office:document-styles ' +
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
    "<office:master-styles><style:master-page style:name=\"Standard\">" +
    (header ? `<style:header>${header}</style:header>` : "") +
    (footer ? `<style:footer>${footer}</style:footer>` : "") +
    "</style:master-page></office:master-styles></office:document-styles>";

  function makeHfOdt(header: string, footer: string): Uint8Array {
    return zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(STYLES(header, footer)),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
  }

  it("reads the master-page header and footer into HTML", () => {
    const parts = odtToParts(makeHfOdt("<text:p>My header</text:p>", "<text:p>The footer</text:p>"));
    expect(parts.header).toContain("My header");
    expect(parts.footer).toContain("The footer");
  });

  it("returns empty header/footer when the master page has none", () => {
    const parts = odtToParts(makeHfOdt("", ""));
    expect(parts.header).toBe("");
    expect(parts.footer).toBe("");
  });

  it("writes edited header/footer back into styles.xml", () => {
    const out = htmlToOdt("<p>body</p>", makeHfOdt("<text:p>old h</text:p>", "<text:p>old f</text:p>"), {
      parts: [
        { path: "header", html: "<p>new header</p>" },
        { path: "footer", html: "<p>new footer</p>" },
      ],
    });
    const styles = strFromU8(unzipSync(out)["styles.xml"]);
    expect(styles).toContain("<style:header>");
    expect(styles).toContain("new header");
    expect(styles).toContain("new footer");
    expect(styles).not.toContain("old h");
  });

  it("round-trips header/footer through read -> write -> read", () => {
    const odt = makeHfOdt("<text:p>H1</text:p>", "<text:p>F1</text:p>");
    const parts = odtToParts(odt);
    const out = htmlToOdt(parts.body, odt, { parts: [{ path: "header", html: parts.header }, { path: "footer", html: parts.footer }] });
    const again = odtToParts(out);
    expect(again.header).toContain("H1");
    expect(again.footer).toContain("F1");
  });
});

describe("odt track changes (text:tracked-changes)", () => {
  const ROOT =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/"';
  const make = (inner: string) =>
    makeOdt(`<?xml version="1.0"?><office:document-content ${ROOT}><office:body><office:text>${inner}</office:text></office:body></office:document-content>`);

  it("reads an insertion range into <ins>", () => {
    const odt = make(
      '<text:tracked-changes><text:changed-region text:id="c1"><text:insertion><office:change-info><dc:creator>Ann</dc:creator><dc:date>2026-06-24T00:00:00</dc:date></office:change-info></text:insertion></text:changed-region></text:tracked-changes>' +
        '<text:p>keep <text:change-start text:change-id="c1"/>added<text:change-end text:change-id="c1"/> end</text:p>',
    );
    const html = odtToHtml(odt);
    expect(html).toContain('<ins class="docx-ins"');
    expect(html).toContain('data-author="Ann"');
    expect(html).toContain(">added</ins>");
  });

  it("reads a deletion into <del> with the removed text", () => {
    const odt = make(
      '<text:tracked-changes><text:changed-region text:id="d1"><text:deletion><office:change-info><dc:creator>Bo</dc:creator></office:change-info><text:p>removed words</text:p></text:deletion></text:changed-region></text:tracked-changes>' +
        '<text:p>a <text:change text:change-id="d1"/>b</text:p>',
    );
    const html = odtToHtml(odt);
    expect(html).toContain('<del class="docx-del"');
    expect(html).toContain("removed words</del>");
  });

  it("writes <ins>/<del> back to tracked-changes regions + body markers", () => {
    const html =
      '<p>x<ins class="docx-ins" data-author="Ann" data-date="2026-06-24T00:00:00">new</ins>y' +
      '<del class="docx-del" data-author="Bo">old</del>z</p>';
    const out = htmlToOdt(html, make("<text:p/>"));
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain("text:tracked-changes");
    expect(xml).toContain("text:insertion");
    expect(xml).toContain("<dc:creator>Ann</dc:creator>");
    expect(xml).toContain("text:change-start");
    expect(xml).toContain("text:change-end");
    expect(xml).toContain("text:deletion");
    expect(xml).toContain("<text:p>old</text:p>"); // deleted text lives in the region
    expect(xml).toContain("text:change "); // the deletion point marker in the body
    expect(xml).toContain(">new<"); // inserted text stays in the body
    // the body paragraph holds only the marker, not the deleted text
    expect(xml).toMatch(/<text:p>x.*?new.*?y<text:change [^>]*\/>z<\/text:p>/);
  });

  it("round-trips an insertion and a deletion", () => {
    const odt = make(
      '<text:tracked-changes>' +
        '<text:changed-region text:id="c1"><text:insertion><office:change-info><dc:creator>A</dc:creator></office:change-info></text:insertion></text:changed-region>' +
        '<text:changed-region text:id="d1"><text:deletion><office:change-info><dc:creator>B</dc:creator></office:change-info><text:p>dead</text:p></text:deletion></text:changed-region>' +
        '</text:tracked-changes>' +
        '<text:p>k<text:change-start text:change-id="c1"/>ins<text:change-end text:change-id="c1"/> <text:change text:change-id="d1"/>m</text:p>',
    );
    const out = htmlToOdt(odtToHtml(odt), odt);
    const html2 = odtToHtml(out);
    expect(html2).toContain(">ins</ins>");
    expect(html2).toContain("dead</del>");
  });
})

describe("odt page margins (page-layout)", () => {
  const STYLES =
    '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
    '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:margin-top="2cm"/></style:page-layout></office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
  function makeMarginOdt(): Uint8Array {
    return zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(STYLES),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
  }
  it("writes edited margins into the page-layout (px -> cm)", () => {
    const out = htmlToOdt("<p>x</p>", makeMarginOdt(), {
      page: { widthPx: 794, heightPx: 1123, margin: { top: 96, right: 96, bottom: 96, left: 96 } },
    });
    const styles = strFromU8(unzipSync(out)["styles.xml"]);
    expect(styles).toContain('fo:margin-left="2.54cm"'); // 96px = 1in = 2.54cm
    expect(styles).toContain('fo:margin-top="2.54cm"');
  });
});

describe("odt page geometry (page-layout)", () => {
  const stylesWith = (props: string) =>
    '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
    `<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties ${props}/></style:page-layout></office:automatic-styles>` +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
  const odtWith = (props: string) =>
    zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(stylesWith(props)),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });

  it("reads page size + margins (cm -> px)", () => {
    const page = odtToParts(odtWith('fo:page-width="18.2cm" fo:page-height="25.7cm" fo:margin-top="2cm" fo:margin-left="2cm"')).page;
    expect(page?.widthPx).toBe(688); // 18.2cm
    expect(page?.heightPx).toBe(971); // 25.7cm (JIS B5)
    expect(page?.margin.top).toBe(76); // 2cm
  });

  it("reads landscape (swapped page-width/height) as a wide page", () => {
    const page = odtToParts(odtWith('fo:page-width="42cm" fo:page-height="29.7cm" style:print-orientation="landscape"')).page;
    expect(page!.widthPx).toBeGreaterThan(page!.heightPx); // A3 landscape
  });

  it("returns no geometry when the page-layout has no size", () => {
    expect(odtToParts(odtWith('fo:margin-top="2cm"')).page).toBeUndefined();
  });

  it("reads section columns from style:columns (count + gap)", () => {
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
      '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm">' +
      '<style:columns fo:column-count="2"><style:column-sep style:width="0.5cm"/></style:columns>' +
      '</style:page-layout-properties></style:page-layout></office:automatic-styles>' +
      '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const page = odtToParts(odt).page;
    expect(page?.columns).toBe(2);
    expect(page?.columnGapPx).toBe(19); // 0.5cm = ~19px
  });
});

describe("odt editable tables (cell content round-trip)", () => {
  const TABLE = makeOdt(
    '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">' +
      "<office:body><office:text>" +
      '<table:table table:name="T1"><table:table-column table:number-columns-repeated="2"/>' +
      "<table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row>" +
      "<table:table-row><table:table-cell><text:p>A2</text:p></table:table-cell><table:table-cell><text:p>B2</text:p></table:table-cell></table:table-row>" +
      "</table:table><text:p>after</text:p></office:text></office:body></office:document-content>",
  );

  it("renders a table:table with editable cells", () => {
    const html = odtToHtml(TABLE);
    expect(html).toContain('class="docx-table"');
    expect(html).toContain('class="docx-cell"');
    expect(html).toContain("A1");
    expect(html).toContain("B2");
  });

  it("builds a fresh table:table from a table inserted in the editor (no skeleton)", () => {
    const html = '<p>x</p><table class="docx-table"><tr><td><div class="docx-cell">X</div></td><td><div class="docx-cell">Y</div></td></tr></table>';
    const out = htmlToOdt(html, makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain("<table:table ");
    expect(xml).toContain("table:number-columns-repeated=\"2\"");
    expect((xml.match(/<table:table-cell\b/g) || []).length).toBe(2);
    expect(xml).toContain("X");
    expect(xml).toContain("Y");
  });

  it("writes back edited cell content, keeping rows/cells", () => {
    const html = odtToHtml(TABLE).replace(">A1<", ">EDITED<");
    const out = htmlToOdt(html, TABLE);
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toContain("EDITED");
    expect(xml).not.toContain("<text:p>A1</text:p>");
    expect(xml).toContain("B2");
    expect((xml.match(/<table:table-row\b/g) || []).length).toBe(2);
    expect((xml.match(/<table:table-cell\b/g) || []).length).toBe(4);
  });
});

describe("odt paragraph indent + line spacing", () => {
  it("writes fo:margin-left / fo:line-height and reads them back", () => {
    const out = htmlToOdt('<p style="margin-left:48px;line-height:1.5">x</p>', makeOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toContain('fo:margin-left="1.27cm"'); // 48px ~= 1.27cm
    expect(content).toContain('fo:line-height="150%"');
    const html = odtToHtml(out);
    expect(html).toMatch(/margin-left:\s*48px/);
    expect(html).toMatch(/line-height:\s*1\.5/);
  });

  it("writes paragraph space before/after (fo:margin-top/bottom) and reads it back", () => {
    const out = htmlToOdt('<p style="margin-top:19px;margin-bottom:10px">x</p>', makeOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toMatch(/fo:margin-top="0\.50\d*cm"/); // 19px ~= 0.503cm
    expect(content).toMatch(/fo:margin-bottom="0\.2\d*cm"/); // 10px ~= 0.265cm
    const html = odtToHtml(out);
    expect(html).toMatch(/margin-top:\s*19px/);
    expect(html).toMatch(/margin-bottom:\s*10px/);
  });

  it("round-trips an explicit zero space-before", () => {
    const out = htmlToOdt('<p style="margin-top:0px">x</p>', makeOdt());
    expect(strFromU8(unzipSync(out)["content.xml"])).toContain('fo:margin-top="0cm"');
    expect(odtToHtml(out)).toMatch(/margin-top:\s*0px/);
  });
});

describe("odt run formatting: strike, superscript, subscript", () => {
  it("writes the ODF text properties and reads them back", () => {
    const out = htmlToOdt("<p><s>a</s><sup>b</sup><sub>c</sub></p>", makeOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toContain("style:text-line-through-style");
    expect(content).toContain('style:text-position="super 58%"');
    expect(content).toContain('style:text-position="sub 58%"');
    const html = odtToHtml(out);
    expect(html).toContain("<s>a</s>");
    expect(html).toContain("<sup>b</sup>");
    expect(html).toContain("<sub>c</sub>");
  });
});

describe("odt list fidelity: nesting and ordered/bullet", () => {
  it("writes a number list style for <ol> and a bullet style for <ul>", () => {
    const out = htmlToOdt("<ul><li>a</li></ul><ol><li>b</li></ol>", makeOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toContain("text:list-level-style-bullet");
    expect(content).toContain("text:list-level-style-number");
    // each list element carries a style-name
    expect(content).toMatch(/<text:list [^>]*text:style-name="OT_L[BO]"/);
  });

  it("round-trips a nested list, preserving ordered/bullet kind at each level", () => {
    const html = "<ol><li>top<ul><li>sub</li></ul></li></ol>";
    const out = htmlToOdt(html, makeOdt());
    const back = odtToHtml(out);
    // outer ordered, inner unordered, no flattening or duplication
    expect(back).toMatch(/<ol>[\s\S]*<ul>[\s\S]*sub[\s\S]*<\/ul>[\s\S]*<\/ol>/);
    // "sub" appears exactly once (not duplicated inline + as a nested item)
    expect((back.match(/sub/g) ?? []).length).toBe(1);
  });
});

describe("odt named paragraph styles", () => {
  const STYLES = `<?xml version="1.0"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
 <office:styles>
  <style:style style:name="Standard" style:family="paragraph"/>
  <style:style style:name="Quote" style:display-name="Quote" style:family="paragraph">
   <style:paragraph-properties fo:margin-left="1cm"/><style:text-properties fo:font-style="italic"/>
  </style:style>
 </office:styles>
</office:document-styles>`;
  const CONTENT = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
 <office:body><office:text><text:p text:style-name="Quote">quoted</text:p></office:text></office:body></office:document-content>`;
  const makeStyledOdt = () => zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "content.xml": strToU8(CONTENT),
    "styles.xml": strToU8(STYLES),
    "META-INF/manifest.xml": strToU8("<m/>"),
  });

  it("lists named styles (excluding Standard) and emits their CSS", () => {
    const parts = odtToParts(makeStyledOdt());
    const ids = (parts.paragraphStyles ?? []).map((s) => s.id);
    expect(ids).toContain("Quote");
    expect(ids).not.toContain("Standard");
    expect(parts.styleCss).toMatch(/\[data-rdoc-style="Quote"\]\{[^}]*font-style:italic/);
  });

  it("reads a styled paragraph as data-rdoc-style and writes the style-name back", () => {
    const html = odtToHtml(makeStyledOdt());
    expect(html).toMatch(/<p[^>]*data-rdoc-style="Quote"[^>]*>quoted<\/p>/);
    const out = htmlToOdt(html, makeStyledOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toMatch(/<text:p[^>]*text:style-name="Quote"[^>]*>quoted<\/text:p>/);
  });
});

describe("odt authoring new styles", () => {
  const STYLES = `<?xml version="1.0"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:styles/></office:document-styles>`;
  const make = () => zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "content.xml": strToU8(`<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:body><office:text><text:p>x</text:p></office:text></office:body></office:document-content>`),
    "styles.xml": strToU8(STYLES),
    "META-INF/manifest.xml": strToU8("<m/>"),
  });

  it("adds an authored paragraph style to styles.xml office:styles", () => {
    const out = htmlToOdt('<p data-rdoc-style="MyHeading">Hi</p>', make(), {
      newStyles: [{ id: "MyHeading", name: "My Heading", kind: "paragraph", css: { "text-align": "center", "font-weight": "bold" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["styles.xml"]);
    expect(stylesXml).toMatch(/<style:style[^>]*style:name="MyHeading"[^>]*style:family="paragraph"/);
    expect(stylesXml).toMatch(/fo:text-align="center"/);
    expect(stylesXml).toMatch(/fo:font-weight="bold"/);
    expect(odtToParts(out).paragraphStyles?.some((s) => s.id === "MyHeading")).toBe(true);
  });

  it("adds an authored character style to styles.xml", () => {
    const out = htmlToOdt('<p>a <span data-rdoc-cstyle="Em">b</span></p>', make(), {
      newStyles: [{ id: "Em", name: "Emph", kind: "character", css: { "font-style": "italic" } }],
    });
    const stylesXml = strFromU8(unzipSync(out)["styles.xml"]);
    expect(stylesXml).toMatch(/<style:style[^>]*style:name="Em"[^>]*style:family="text"/);
    expect(stylesXml).toMatch(/fo:font-style="italic"/);
    expect(odtToParts(out).characterStyles?.some((s) => s.id === "Em")).toBe(true);
  });

  it("editing a style keeps its unmodelled attributes, parent, and does not flatten inherited props", () => {
    const STYLES = `<?xml version="1.0"?>` +
      `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:styles>` +
      `<style:style style:name="Normal" style:family="paragraph"><style:text-properties fo:font-size="22pt"/></style:style>` +
      `<style:style style:name="Fancy" style:family="paragraph" style:parent-style-name="Normal"><style:paragraph-properties fo:keep-with-next="always"/><style:text-properties fo:font-weight="bold"/></style:style>` +
      `</office:styles></office:document-styles>`;
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(`<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p text:style-name="Fancy">x</text:p></office:text></office:body></office:document-content>`),
      "styles.xml": strToU8(STYLES),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    // The edit dialog sees Fancy's OWN props only (bold), not the inherited font size.
    const def = (odtToParts(odt).styleDefs ?? []).find((d) => d.id === "Fancy");
    expect(def?.css["font-weight"]).toBe("bold");
    expect(def?.css["font-size"]).toBeUndefined();
    const out = htmlToOdt('<p data-rdoc-style="Fancy">x</p>', odt, {
      newStyles: [{ id: "Fancy", name: "Fancy", kind: "paragraph", css: { "font-weight": "bold", "font-style": "italic" } }],
    });
    const xml = strFromU8(unzipSync(out)["styles.xml"]);
    const fancyEl = xml.slice(xml.indexOf('style:name="Fancy"')).split("</style:style>")[0]!;
    expect(fancyEl).toContain('fo:keep-with-next="always"'); // unmodelled attr preserved
    expect(fancyEl).toMatch(/style:parent-style-name="Normal"/); // inheritance preserved
    expect(fancyEl).toContain('fo:font-weight="bold"');
    expect(fancyEl).toContain('fo:font-style="italic"');
    expect(fancyEl).not.toContain("fo:font-size"); // inherited size not flattened in
  });
});

describe("odt named character styles", () => {
  const STYLES = `<?xml version="1.0"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
 <office:styles>
  <style:style style:name="Emphasis" style:display-name="Emphasis" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
 </office:styles>
</office:document-styles>`;
  const CONTENT = `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
 <office:body><office:text><text:p>a <text:span text:style-name="Emphasis">em</text:span></text:p></office:text></office:body></office:document-content>`;
  const makeStyledOdt = () => zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "content.xml": strToU8(CONTENT),
    "styles.xml": strToU8(STYLES),
    "META-INF/manifest.xml": strToU8("<m/>"),
  });

  it("lists character styles and emits their CSS", () => {
    const parts = odtToParts(makeStyledOdt());
    expect((parts.characterStyles ?? []).map((s) => s.id)).toContain("Emphasis");
    expect(parts.styleCss).toMatch(/\[data-rdoc-cstyle="Emphasis"\]\{[^}]*font-style:italic/);
  });

  it("reads a styled span as data-rdoc-cstyle and writes the style-name back", () => {
    const html = odtToHtml(makeStyledOdt());
    expect(html).toMatch(/data-rdoc-cstyle="Emphasis"/);
    const out = htmlToOdt(html, makeStyledOdt());
    const content = strFromU8(unzipSync(out)["content.xml"]);
    expect(content).toMatch(/<text:span[^>]*text:style-name="Emphasis"[^>]*>em<\/text:span>/);
  });

  it("reads and writes an even-page (left) header variant (odt)", () => {
    const styles = `<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
      `<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout></office:automatic-styles>` +
      `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>ODD</text:p></style:header><style:header-left><text:p>EVEN</text:p></style:header-left></style:master-page></office:master-styles></office:document-styles>`;
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const parts = odtToParts(odt);
    expect(parts.header).toContain("ODD");
    expect(parts.headerEven?.html).toContain("EVEN");
    expect(parts.page?.evenOdd).toBe(true);
    // write the even variant back (edited content) to style:header-left
    const out = htmlToOdt("<p>body</p>", odt, { parts: [{ path: "header-left@Standard", html: "<p>EVEN2</p>" }] });
    const sx = strFromU8(unzipSync(out)["styles.xml"]);
    expect(sx).toMatch(/<style:header-left>[\s\S]*EVEN2[\s\S]*<\/style:header-left>/);
  });

  it("reads and writes a first-page header variant (odt header-first)", () => {
    const styles = `<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
      `<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout></office:automatic-styles>` +
      `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>MAIN</text:p></style:header><style:header-first><text:p>TITLE</text:p></style:header-first></style:master-page></office:master-styles></office:document-styles>`;
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(CONTENT),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const parts = odtToParts(odt);
    expect(parts.headerFirst?.html).toContain("TITLE");
    expect(parts.page?.titlePage).toBe(true);
    // a freshly toggled first-page variant (sentinel path) writes into style:header-first
    const out = htmlToOdt("<p>body</p>", odt, { parts: [{ path: "header:first", html: "<p>TITLE2</p>" }] });
    const sx = strFromU8(unzipSync(out)["styles.xml"]);
    expect(sx).toMatch(/<style:header-first>[\s\S]*TITLE2[\s\S]*<\/style:header-first>/);
    // an enabled-but-empty first header still writes the element, so a blank first page round-trips
    const blank = htmlToOdt("<p>body</p>", odt, { parts: [{ path: "header:first", html: "<p><br></p>" }] });
    expect(strFromU8(unzipSync(blank)["styles.xml"])).toContain("style:header-first");
  });
});

describe("odt bookmarks and cross-references", () => {
  const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;

  it("reads point + range bookmarks and a bookmark-ref cross-reference", () => {
    const content = `<?xml version="1.0"?>
<office:document-content ${NS}><office:body><office:text>
 <text:p><text:bookmark text:name="here"/>A <text:bookmark-start text:name="intro"/>Intro<text:bookmark-end text:name="intro"/></text:p>
 <text:p>See <text:bookmark-ref text:reference-format="text" text:ref-name="intro">Intro</text:bookmark-ref> on page <text:bookmark-ref text:reference-format="page" text:ref-name="intro">1</text:bookmark-ref></text:p>
</office:text></office:body></office:document-content>`;
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain('data-rdoc-bm="here"'); // point bookmark -> start/end pair
    expect(html).toContain('data-rdoc-bm="intro"');
    expect(html).toContain("docx-bookmark-end");
    expect(html).toContain('data-rdoc-xref="intro"');
    expect(html).toContain('data-rdoc-xref-fmt="text"');
    expect(html).toContain('data-rdoc-xref-fmt="page"');
  });

  it("writes bookmark markers and cross-references back to ODF", () => {
    const body =
      '<p><a class="docx-bookmark" data-rdoc-bm="intro" data-rdoc-bm-id="intro" contenteditable="false"></a>Intro' +
      '<a class="docx-bookmark-end" data-rdoc-bm-id="intro" data-rdoc-bm-end="intro" contenteditable="false"></a></p>' +
      '<p><a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="text" contenteditable="false">Intro</a> ' +
      '<a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="page" contenteditable="false">1</a></p>';
    const out = htmlToOdt(body, makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toMatch(/<text:bookmark-start[^>]*text:name="intro"/);
    expect(xml).toMatch(/<text:bookmark-end[^>]*text:name="intro"/);
    expect(xml).toMatch(/<text:bookmark-ref[^>]*text:reference-format="text"[^>]*text:ref-name="intro"/);
    expect(xml).toMatch(/text:reference-format="page"/);
    // and survives a re-read
    const html = odtToHtml(out);
    expect(html).toContain('data-rdoc-bm="intro"');
    expect(html).toContain('data-rdoc-xref="intro"');
  });

  it("round-trips an above/below cross-reference via reference-format=direction", () => {
    const body = '<p><a class="docx-xref" data-rdoc-xref="intro" data-rdoc-xref-fmt="direction" contenteditable="false">below</a></p>';
    const out = htmlToOdt(body, makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toMatch(/text:reference-format="direction"/);
    expect(odtToHtml(out)).toContain('data-rdoc-xref-fmt="direction"');
  });

  it("reads a text:sequence caption and writes it back", () => {
    const content = `<?xml version="1.0"?>
<office:document-content ${NS} xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:body><office:text>
 <text:p>Table <text:sequence text:name="Table" text:formula="ooow:Table+1" style:num-format="1">1</text:sequence>: Results</text:p>
</office:text></office:body></office:document-content>`;
    const html = odtToHtml(makeOdt(content));
    expect(html).toContain('data-rdoc-caption="table"');
    expect(html).toContain('data-field="seq"');
    expect(html).toContain('data-seq="Table"');

    const body = '<p data-rdoc-caption="figure">Figure <span class="docx-field docx-field-seq" data-field="seq" data-seq="Figure" contenteditable="false">1</span>: A chart</p>';
    const out = htmlToOdt(body, makeOdt());
    const xml = strFromU8(unzipSync(out)["content.xml"]);
    expect(xml).toMatch(/<text:sequence[^>]*text:name="Figure"/);
    expect(odtToHtml(out)).toContain('data-rdoc-caption="figure"');
  });
});
