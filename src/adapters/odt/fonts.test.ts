import { afterEach, describe, expect, it, vi } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createOdtAdapter, htmlToOdt, odtToParts } from "./index";
import { leadingScript, splitScripts } from "./fonts";

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

// A synthetic CJK document: a Western default font, an East Asian default at a smaller size,
// an embedded East Asian brush font, and a paragraph style that sets both scripts' size.
const FACES = `<office:font-face-decls>
  <style:font-face style:name="Serif1" svg:font-family="'Body Serif'" style:font-family-generic="roman"/>
  <style:font-face style:name="Brush" svg:font-family="Brush Kai">
   <svg:font-face-src><svg:font-face-uri xlink:href="Fonts/brush.ttf" xlink:type="simple" loext:font-style="normal" loext:font-weight="normal"><svg:font-face-format svg:string="truetype"/></svg:font-face-uri></svg:font-face-src>
  </style:font-face>
  <style:font-face style:name="Mincho" svg:font-family="Doc Mincho"/>
 </office:font-face-decls>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS} office:version="1.3">
 ${FACES}
 <office:styles>
  <style:default-style style:family="paragraph">
   <style:text-properties style:font-name="Serif1" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="10.5pt"/>
  </style:default-style>
  <style:style style:name="Standard" style:family="paragraph"/>
 </office:styles>
</office:document-styles>`;

const content = (body: string, auto = "") => `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.3">
 ${FACES}
 <office:automatic-styles>
  <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard">
   <style:text-properties fo:font-size="28pt" style:font-name-asian="Brush" style:font-size-asian="26pt"/>
  </style:style>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-size="12pt" style:font-size-asian="10pt" style:font-name-asian="Brush"/></style:style>
  <style:style style:name="T2" style:family="text"><style:text-properties style:font-name="Serif1"/></style:style>
  ${auto}
 </office:automatic-styles>
 <office:body><office:text>${body}</office:text></office:body>
</office:document-content>`;

const FONT_BYTES = new Uint8Array([0, 1, 0, 0, 9, 9, 9, 9]);
function makeOdt(body: string, auto = ""): Uint8Array {
  return zipSync({
    mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
    "content.xml": strToU8(content(body, auto)),
    "styles.xml": strToU8(STYLES),
    "Fonts/brush.ttf": FONT_BYTES,
    "META-INF/manifest.xml": strToU8("<m/>"),
  });
}

describe("odt script detection", () => {
  it("lets weak characters follow the script before them", () => {
    const state = { script: leadingScript("一、千里") };
    expect(state.script).toBe("a");
    expect(splitScripts("一、□□ (abc) 12", state)).toEqual([
      { ea: true, text: "一、□□ (" },
      { ea: false, text: "abc) 12" },
    ]);
    // A weak lead-in takes the first strong character's script.
    expect(leadingScript("☆△ さん")).toBe("a");
    expect(leadingScript("12 abc")).toBe("l");
  });
});

describe("odt fonts", () => {
  afterEach(() => vi.restoreAllMocks());

  it("registers embedded fonts under their declared family", () => {
    const created: Blob[] = [];
    (URL as unknown as { createObjectURL: (b: Blob) => string }).createObjectURL = (b: Blob) => (created.push(b), `blob:font${created.length}`);
    const doc = createOdtAdapter(makeOdt("<text:p>x</text:p>")).read();
    expect(doc.fontCss).toContain('font-family:"Brush Kai"');
    expect(doc.fontCss).toContain("src:url(blob:font1)");
    expect(doc.fontUrls).toEqual(["blob:font1"]);
    expect(created).toHaveLength(1); // declared in no other part, loaded once
  });

  it("reads the document defaults per script and resolves font-face names", () => {
    const p = odtToParts(makeOdt("<text:p>x</text:p>"));
    expect(p.defaultFont).toBe("Body Serif"); // style:font-name="Serif1" -> its svg:font-family
    expect(p.defaultFontEastAsia).toBe("Doc Mincho");
    expect(p.defaultSizeHalfPt).toBe(24);
    expect(p.styleCss).toContain("--rdoc-ea-size:10.5pt");
    expect(p.styleCss).toContain("--rdoc-ea-font:'Doc Mincho'");
  });

  it("renders East Asian text in the East Asian font and size, Western text in the Western ones", () => {
    const body = odtToParts(makeOdt('<text:p text:style-name="P1">一、千里 <text:span text:style-name="T2">(abc)</text:span><text:span text:style-name="T1">□□</text:span></text:p>')).body;
    // The paragraph style's own run properties, both scripts.
    expect(body).toMatch(/<p style="[^"]*font-size:28pt[^"]*--rdoc-ea-font:'Brush Kai'[^"]*--rdoc-ea-size:26pt/);
    expect(body).toContain('<span data-rdoc-script="ea">一、千里 </span>');
    // A style:font-name reference is resolved through the font-face declarations.
    expect(body).toContain("font-family:'Body Serif'");
    // "(" follows the East Asian text before it; "abc)" is Western.
    expect(body).toContain('<span data-rdoc-script="ea">(</span>abc)');
    // A run's own East Asian size rides as a custom property; weak squares stay East Asian.
    expect(body).toMatch(/font-size:12pt;--rdoc-ea-font:'Brush Kai'[^"]*;--rdoc-ea-size:10pt">/);
  });

  it("keeps a run's East Asian size apart from its Western size on save", () => {
    const odt = makeOdt('<text:p>本文 <text:span text:style-name="T1">漢字</text:span> end</text:p>');
    const html = odtToParts(odt).body.replace("end", "edited");
    const out = strFromU8(unzipSync(htmlToOdt(html, odt))["content.xml"]);
    const name = /<text:span text:style-name="([^"]+)">(?:<[^>]+>)*漢字/.exec(out)?.[1];
    const st = new RegExp(`<style:style[^>]*style:name="${name}"[^>]*>[\\s\\S]*?</style:style>`).exec(out)?.[0] ?? "";
    expect(st).toContain('fo:font-size="12pt"');
    expect(st).toContain('style:font-size-asian="10pt"');
    expect(out).toContain("edited");
  });
});
