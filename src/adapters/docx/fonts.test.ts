import { beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocxAdapter, createDocxEditor, docxToParts, htmlToDocx } from "./index";

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
  const zeroRect = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  if (!Range.prototype.getBoundingClientRect) Range.prototype.getBoundingClientRect = zeroRect;
  if (!Range.prototype.getClientRects) Range.prototype.getClientRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
});

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

// What Word writes in a stock document: every font is a theme reference.
const THEME = `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:fontScheme name="Office">` +
  `<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Yu Mincho"/><a:cs typeface=""/></a:minorFont>` +
  `</a:fontScheme></a:themeElements></a:theme>`;
const STYLES = `<?xml version="1.0"?><w:styles ${NS}>` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style>` +
  `</w:styles>`;

function wordDocx(body: string, extra: Record<string, Uint8Array> = {}): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(
      `<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    ),
    "word/theme/theme1.xml": strToU8(THEME),
    "word/styles.xml": strToU8(STYLES),
    ...extra,
  });
}

// What Omnitext's New dialog creates: no styles.xml, no theme.
const BLANK = zipSync({
  "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
  "_rels/.rels": strToU8("<Relationships/>"),
  "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body><w:p/><w:sectPr/></w:body></w:document>`),
  "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="${RELS_NS}"></Relationships>`),
});

describe("docx fonts", () => {
  it("resolves the default font and size through docDefaults, the Normal style and the theme", () => {
    const parts = createDocxAdapter(wordDocx("<w:p><w:r><w:t>body</w:t></w:r></w:p>")).read();
    expect(parts.defaultFont).toBe("Calibri");
    expect(parts.defaultFontEastAsia).toBe("Yu Mincho");
    expect(parts.defaultSizeHalfPt).toBe(22); // Normal (11pt) beats docDefaults (12pt)
  });

  it("renders a run's theme font and gives headings their style's look", () => {
    const parts = docxToParts(wordDocx(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr><w:t>themed</w:t></w:r></w:p>`,
    ));
    expect(parts.body).toMatch(/font-family:'Calibri Light'[^"]*">themed/);
    expect(parts.styleCss).toMatch(/\.docxedit-doc h1\{[^}]*font-family:'Calibri Light'/);
    expect(parts.styleCss).toMatch(/\.docxedit-doc h1\{[^}]*font-size:16pt/);
    expect(parts.styleCss).toMatch(/\.docxedit-doc h1\{[^}]*color:#2F5496/);
  });

  it("drops the theme reference when the user picks another font for a themed run", () => {
    const docx = wordDocx(`<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsia="MS Mincho"/></w:rPr><w:t>x</w:t></w:r></w:p>`);
    const html = docxToParts(docx).body.replace(/font-family:'Calibri Light'/, "font-family:'Georgia'");
    const xml = strFromU8(unzipSync(htmlToDocx(html, docx))["word/document.xml"]!);
    expect(xml).toMatch(/<w:rFonts[^>]*w:ascii="Georgia"/);
    expect(xml).not.toContain("asciiTheme");
    expect(xml).toContain('w:eastAsia="MS Mincho"'); // the East Asian font is untouched
    // An untouched themed run keeps its theme reference.
    const same = strFromU8(unzipSync(htmlToDocx(docxToParts(docx).body, docx))["word/document.xml"]!);
    expect(same).toContain('w:asciiTheme="majorHAnsi"');
  });

  it("round-trips explicit run fonts and sizes", () => {
    const docx = wordDocx(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="28"/></w:rPr><w:t>g</w:t></w:r></w:p>`);
    const out = htmlToDocx(docxToParts(docx).body, docx);
    const body = docxToParts(out).body;
    expect(body).toContain("font-size:14pt");
    expect(body).toContain("font-family:'Georgia'");
  });

  it("stores the default font of a new document so it reopens in it", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, BLANK);
    const docEl = host.querySelector(".docxedit-doc") as HTMLElement;
    expect(docEl.style.fontFamily).toContain("Arial"); // the new-document default the toolbar shows
    docEl.innerHTML = "<p>Hello</p>";
    docEl.dispatchEvent(new Event("input", { bubbles: true }));
    const out = await ed.getBytes();
    ed.destroy();
    host.remove();
    const files = unzipSync(out);
    const styles = strFromU8(files["word/styles.xml"]!);
    expect(styles).toMatch(/<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts[^>]*w:ascii="Arial"/);
    expect(styles).toContain('<w:sz w:val="22"/>');
    expect(strFromU8(files["word/_rels/document.xml.rels"]!)).toContain('Target="styles.xml"');
    expect(strFromU8(files["[Content_Types].xml"]!)).toContain('PartName="/word/styles.xml"');
    const parts = createDocxAdapter(out).read();
    expect(parts.defaultFont).toBe("Arial");
    expect(parts.defaultSizeHalfPt).toBe(22);
  });

  it("leaves styles.xml alone when the page font is the file's own default", () => {
    const docx = wordDocx("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    const out = htmlToDocx("<p>y</p>", docx, [], { docDefaults: { font: "Calibri", sizeHalfPt: 22 } });
    expect(strFromU8(unzipSync(out)["word/styles.xml"]!)).toBe(STYLES);
  });

  it("changes a default the Normal style overrides in the Normal style", () => {
    const out = htmlToDocx("<p>y</p>", wordDocx("<w:p/>"), [], { docDefaults: { font: "Georgia", sizeHalfPt: 24 } });
    const parts = createDocxAdapter(out).read();
    expect(parts.defaultFont).toBe("Georgia");
    expect(parts.defaultSizeHalfPt).toBe(24);
  });
});
