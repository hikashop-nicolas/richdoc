import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { docxToParts, htmlToDocx } from "./index";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
function docx(body: string, numbering?: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body>${body}</w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    ...(numbering ? { "word/numbering.xml": strToU8(`<?xml version="1.0"?><w:numbering ${NS}>${numbering}</w:numbering>`) } : {}),
  });
}
const li = (lvl: number, text: string, numId = "1") => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const lvl = (i: number, fmt: string, text: string) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/></w:lvl>`;

describe("docx list level formats", () => {
  it("renders each level in its own counter style and kind", () => {
    const num = `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}${lvl(1, "lowerLetter", "%2)")}${lvl(2, "bullet", "")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const parts = docxToParts(docx(li(0, "one") + li(1, "sub") + li(2, "dot"), num));
    expect(parts.body).toMatch(/^<ol><li[^>]*>one<ol style="list-style-type:rdoc-num-0-1"><li[^>]*>sub<ul style="list-style-type:square">/);
    expect(parts.styleCss).toContain('@counter-style rdoc-num-0-1{system:extends lower-alpha;prefix:"";suffix:") ";}');
  });

  it("shows a bulleted level under a numbered one as bullets, and roman numerals as such", () => {
    const num = `<w:abstractNum w:abstractNumId="0">${lvl(0, "upperRoman", "%1.")}${lvl(1, "bullet", "o")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const { body } = docxToParts(docx(li(0, "I") + li(1, "b"), num));
    expect(body).toContain('<ol style="list-style-type:upper-roman">');
    expect(body).toContain('<ul style="list-style-type:circle">');
  });

  it("keeps the list's numbering untouched on save", () => {
    const num = `<w:abstractNum w:abstractNumId="0">${lvl(0, "lowerLetter", "(%1)")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const d = docx(li(0, "a"), num);
    const out = unzipSync(htmlToDocx(docxToParts(d).body, d));
    expect(strFromU8(out["word/numbering.xml"]!)).toContain('w:val="(%1)"');
    expect(strFromU8(out["word/document.xml"]!)).toContain('<w:numId w:val="1"/>');
  });
});

describe("docx run properties shown from the preserved rPr", () => {
  const run = (rPr: string, t = "x") => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${t}</w:t></w:r></w:p>`;
  it("shows caps, small caps, double strike, hidden text, spacing and underline styles", () => {
    const { body } = docxToParts(docx(run('<w:caps/><w:spacing w:val="40"/>') + run("<w:smallCaps/><w:dstrike/>") + run("<w:vanish/>") + run('<w:u w:val="double"/>') + run('<w:caps w:val="0"/>')));
    expect(body).toMatch(/class="docx-caps" style="letter-spacing:2pt"/);
    expect(body).toContain('class="docx-smallcaps docx-dstrike"');
    expect(body).toContain('class="docx-hidden"');
    expect(body).toMatch(/class="docx-u-double"><u>x<\/u>/);
    expect(body).not.toMatch(/docx-caps"[^>]*><\/span>|class="docx-caps"><span|class="docx-caps">x/); // caps off: no class
  });

  it("does not let the display styles leak into the saved run", () => {
    const d = docx(run('<w:smallCaps/><w:spacing w:val="40"/><w:position w:val="6"/><w:u w:val="wave"/>'));
    const xml = strFromU8(unzipSync(htmlToDocx(docxToParts(d).body, d))["word/document.xml"]!);
    expect(xml).toContain("<w:smallCaps/>");
    expect(xml).toContain('<w:spacing w:val="40"/>');
    expect(xml).toContain('<w:position w:val="6"/>');
    expect(xml).toContain('<w:u w:val="wave"/>');
    expect(xml).not.toContain("w:vertAlign");
  });
});
