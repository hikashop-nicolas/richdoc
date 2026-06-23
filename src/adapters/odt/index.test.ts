import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { htmlToOdt, odtToHtml } from "./index";

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
 <text:tracked-changes><text:changed-region text:id="c1"><text:deletion><office:change-info><dc:creator>Al</dc:creator></office:change-info></text:deletion></text:changed-region></text:tracked-changes>
 <text:p>before<office:annotation><dc:creator>Al</dc:creator><text:p>my note</text:p></office:annotation> after</text:p>
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
