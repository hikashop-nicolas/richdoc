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
