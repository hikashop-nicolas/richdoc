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

  it("labels nested levels 1.1. the way Word's legal-style lists do", () => {
    const num = `<w:abstractNum w:abstractNumId="3">${lvl(0, "decimal", "%1.")}${lvl(1, "decimal", "%1.%2.")}${lvl(2, "decimal", "%1.%2.%3")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="3"/></w:num>`;
    const parts = docxToParts(docx(li(0, "a") + li(1, "b") + li(2, "c"), num));
    expect(parts.body).toContain('<ol class="rdoc-num-3-1">');
    expect(parts.body).toContain('<ol class="rdoc-num-3-2">');
    expect(parts.styleCss).toContain('.docxedit-doc ol.rdoc-num-3-1>li::marker{content:counters(list-item,".") ". "}');
    expect(parts.styleCss).toContain('.docxedit-doc ol.rdoc-num-3-2>li::marker{content:counters(list-item,".") " "}');
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

describe("docx table display", () => {
  const STYLES = `<?xml version="1.0"?><w:styles ${NS}>` +
    `<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style>` +
    `<w:style w:type="table" w:styleId="Grid4"><w:name w:val="Grid Table 4"/><w:basedOn w:val="TableGrid"/>` +
    `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr></w:style>` +
    `</w:styles>`;
  const withStyles = (body: string) => {
    const files = unzipSync(docx(body));
    files["word/styles.xml"] = strToU8(STYLES);
    return zipSync(files);
  };
  const row = (...cells: string[]) => `<w:tr>${cells.map((c) => `<w:tc><w:p><w:r><w:t>${c}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`;
  const table = (style: string, look: string, rows: string) =>
    `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="5000" w:type="pct"/>${look}</w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;

  it("shows grid column widths and the table width", () => {
    const { body } = docxToParts(withStyles(table("TableGrid", "", row("a", "b"))));
    expect(body).toContain('style="width:100%"');
    expect(body).toMatch(/<td[^>]*width:100px[^>]*>.*<td[^>]*width:200px/);
  });

  it("draws a table style's borders and spacing", () => {
    const { body, styleCss } = docxToParts(withStyles(table("TableGrid", "", row("a", "b"))));
    expect(body).toContain('data-rdoc-bt="1px solid #000000"');
    expect(body).toContain('data-rdoc-tstyle="TableGrid"');
    expect(styleCss).toContain('.docxedit-doc table[data-rdoc-tstyle="TableGrid"] p{margin-bottom:0px;line-height:1}');
    expect(styleCss).toContain(".docxedit-doc p{margin-bottom:11px;line-height:1.08}");
  });

  it("shades the header row and the bands, honouring w:tblLook", () => {
    const rows = row("H1", "H2") + row("r1", "x") + row("r2", "y") + row("r3", "z");
    const { body } = docxToParts(withStyles(table("Grid4", '<w:tblLook w:firstRow="1" w:noHBand="0" w:noVBand="1"/>', rows)));
    expect(body).toMatch(/font-weight:bold;color:#FFFFFF;background-color:#4472C4[^>]*><div[^>]*>.*H1/);
    expect(body).toMatch(/background-color:#D9E2F3[^>]*><div[^>]*><p>r1/);
    expect(body).not.toMatch(/background-color:#D9E2F3[^>]*><div[^>]*><p>r2/);
    expect(body).toMatch(/background-color:#D9E2F3[^>]*><div[^>]*><p>r3/);
    const off = docxToParts(withStyles(table("Grid4", '<w:tblLook w:firstRow="0" w:noHBand="1"/>', rows))).body;
    expect(off).not.toContain("4472C4");
    expect(off).not.toContain("D9E2F3");
  });

  it("leaves the saved table untouched by its display styles", () => {
    const d = withStyles(table("Grid4", "", row("H", "x") + row("a", "b")));
    const xml = strFromU8(unzipSync(htmlToDocx(docxToParts(d).body, d))["word/document.xml"]!);
    expect(xml).not.toContain("w:tcBorders");
    expect(xml).not.toContain("4472C4");
    expect(xml).toContain('<w:gridCol w:w="1500"/>');
  });
});

describe("docx keep with next", () => {
  it("marks a w:keepNext paragraph for pagination and keeps the property on save", () => {
    const d = docx('<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>caption</w:t></w:r></w:p><w:p><w:pPr><w:keepNext w:val="0"/></w:pPr><w:r><w:t>off</w:t></w:r></w:p>');
    const { body } = docxToParts(d);
    expect(body).toMatch(/<p[^>]*data-rdoc-keepnext[^>]*>caption/);
    expect(body).not.toMatch(/<p[^>]*data-rdoc-keepnext[^>]*>off/);
    expect(strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!)).toContain("<w:keepNext/>");
  });
});

describe("docx numbering from a paragraph style", () => {
  const STYLES = `<?xml version="1.0"?><w:styles ${NS}>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="MyList"><w:name w:val="My List"/><w:basedOn w:val="ListNumber"/></w:style>` +
    `</w:styles>`;
  const NUM = `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}${lvl(1, "lowerLetter", "%2.")}</w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num>`;
  const make = (body: string) => {
    const files = unzipSync(docx(body, NUM));
    files["word/styles.xml"] = strToU8(STYLES);
    return zipSync(files);
  };
  const sp = (style: string, t: string) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;

  it("renders style-numbered paragraphs as a list, inheriting along basedOn", () => {
    const { body } = docxToParts(make(sp("ListNumber", "one") + sp("MyList", "two")));
    expect(body).toMatch(/^<ol><li[^>]*data-rdoc-style="ListNumber"[^>]*>one<\/li><li[^>]*data-rdoc-style="MyList"[^>]*>two<\/li><\/ol>$/);
  });

  it("starts a new list for a top-level item of another list", () => {
    const { body } = docxToParts(docx(li(0, "one", "1") + li(0, "dot", "2"), `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl(0, "bullet", "\uF0B7")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`));
    expect(body).toMatch(/^<ol><li[^>]*>one<\/li><\/ol><ul><li[^>]*>dot<\/li><\/ul>$/);
  });

  it("saves them with their style only, adding a direct numPr once the level changes", () => {
    const d = make(sp("ListNumber", "one") + sp("ListNumber", "two"));
    const html = docxToParts(d).body;
    const same = strFromU8(unzipSync(htmlToDocx(html, d))["word/document.xml"]!);
    expect(same).not.toContain("w:numPr");
    expect(same.match(/<w:pStyle w:val="ListNumber"\/>/g)).toHaveLength(2);
    // Indent the second item: it now needs its own level.
    const nested = html.replace(/<\/li><li([^>]*)>two<\/li>/, "<ol><li$1>two</li></ol></li>");
    const xml = strFromU8(unzipSync(htmlToDocx(nested, d))["word/document.xml"]!);
    expect(xml).toMatch(/<w:ilvl w:val="1"\/><w:numId w:val="5"\/>/);
  });
});

describe("docx heading numbering from the Heading styles", () => {
  const STYLES = `<?xml version="1.0"?><w:styles ${NS}>` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style>` +
    `</w:styles>`;
  const NUM = `<w:abstractNum w:abstractNumId="2">${lvl(0, "decimal", "%1")}${lvl(1, "decimal", "%1.%2")}</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="2"/></w:num>`;
  const h = (n: number, t: string) => `<w:p><w:pPr><w:pStyle w:val="Heading${n}"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const make = (body: string) => {
    const files = unzipSync(docx(body, NUM));
    files["word/styles.xml"] = strToU8(STYLES);
    return zipSync(files);
  };

  it("keeps headings as headings and numbers them with CSS counters", () => {
    const d = make(h(1, "Intro") + h(2, "Scope") + h(1, "Method"));
    const { body, styleCss } = docxToParts(d);
    expect(body).toMatch(/^<h1[^>]*>Intro<\/h1><h2[^>]*>Scope<\/h2><h1[^>]*>Method<\/h1>$/);
    expect(styleCss).toContain(".docxedit-doc{counter-reset:rdoc-hn-7-0 0 rdoc-hn-7-1 0}");
    expect(styleCss).toContain(".docxedit-doc h1{counter-increment:rdoc-hn-7-0;counter-set:rdoc-hn-7-1 0}");
    expect(styleCss).toContain('.docxedit-doc h2::before{content:counter(rdoc-hn-7-0,decimal) "." counter(rdoc-hn-7-1,decimal) " "}');
    // Nothing of the label is saved: the headings keep only their style.
    const xml = strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!);
    expect(xml).not.toContain("w:numPr");
    expect(xml).toContain("<w:t xml:space=\"preserve\">Intro</w:t>");
  });
});

describe("docx drop caps", () => {
  it("floats a drop-cap paragraph and keeps its frame on save", () => {
    const d = docx('<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr><w:r><w:rPr><w:sz w:val="96"/></w:rPr><w:t>O</w:t></w:r></w:p><w:p><w:r><w:t>nce upon a time</w:t></w:r></w:p>');
    const { body } = docxToParts(d);
    expect(body).toMatch(/<p[^>]*data-rdoc-dropcap="drop"[^>]*>.*O/);
    const xml = strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!);
    expect(xml).toContain('w:dropCap="drop"');
    expect(xml).toContain('w:lines="3"');
  });
});

describe("docx text frames and list overrides", () => {
  it("floats a framed paragraph at its width and keeps the frame on save", () => {
    const d = docx('<w:p><w:pPr><w:framePr w:w="2880" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/></w:pPr><w:r><w:t>side note</w:t></w:r></w:p><w:p><w:r><w:t>body</w:t></w:r></w:p>');
    const { body } = docxToParts(d);
    expect(body).toMatch(/<p style="width:192px"[^>]*data-rdoc-frame="right"[^>]*>side note/);
    const xml = strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!);
    expect(xml).toContain('w:xAlign="right"');
    expect(xml).not.toContain("w:ind");
  });

  it("applies a list's own level override", () => {
    const num = `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride></w:num>`;
    const parts = docxToParts(docx(li(0, "plain", "1") + "<w:p/>" + li(0, "over", "2"), num));
    expect(parts.body).toContain('<ol start="3" style="list-style-type:rdoc-num-o2-0">');
    expect(parts.body).toMatch(/^<ol><li/);
    expect(parts.styleCss).toContain("@counter-style rdoc-num-o2-0{system:extends upper-alpha;");
  });
});
