import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { docxToParts, htmlToDocx } from "./index";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2></a:clrScheme></a:themeElements></a:theme>`;
function docx(body: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8("<Types/>"),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
    "word/theme/theme1.xml": strToU8(THEME),
  });
}
const txbx = (t: string) => `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${t}</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>second</w:t></w:r></w:p></w:txbxContent></wps:txbx>`;
const wsp = (spPr: string, inner = "", style = "") => `<wps:wsp><wps:cNvSpPr/><wps:spPr>${spPr}</wps:spPr>${style}${inner}<wps:bodyPr anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>`;
const drawing = (shape: string, anchor = false, uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape") =>
  `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${anchor ? '<wp:anchor behindDoc="0"><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1905000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/>' : '<wp:inline><wp:extent cx="1905000" cy="952500"/>'}<wp:docPr id="1" name="s"/><a:graphic><a:graphicData uri="${uri}">${shape}</a:graphicData></a:graphic>${anchor ? "</wp:anchor>" : "</wp:inline>"}</w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="width:150pt;height:75pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>fallback</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`;

describe("docx shapes and text boxes", () => {
  it("draws a text box with its fill, outline and text, inside the paragraph", () => {
    const d = docx(drawing(wsp('<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>', txbx("Note"))));
    const { body } = docxToParts(d);
    expect(body).toMatch(/^<p>Before <span class="docx-shape" contenteditable="false" data-docx-xml="[^"]+"><span class="docx-shape-box" style="width:200px;height:100px;background-color:#FFF2CC;border:2px solid #C00000">/);
    expect(body).toContain('<span class="docx-shape-body" contenteditable="true" style="padding:0px 0px 0px 0px;justify-content:center"><span class="docx-tb-p">Note</span><span class="docx-tb-p"><strong>second</strong></span></span>');
    expect(body.replace(/data-docx-xml="[^"]*"/g, "")).not.toContain("fallback"); // the DrawingML choice wins over the VML fallback
    expect(body).toMatch(/<\/span> after<\/p>$/); // the paragraph stays whole
    const xml = strFromU8(unzipSync(htmlToDocx(body, d))["word/document.xml"]!);
    expect(xml).toContain("<wps:txbx>");
    expect(xml).toContain("mc:Fallback");
    expect(xml).toContain('<w:t>Note</w:t>');
  });

  it("takes the shape's colours from its style and the theme, and its geometry", () => {
    const style = '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>';
    const { body } = docxToParts(docx(drawing(wsp('<a:xfrm rot="5400000"/><a:prstGeom prst="ellipse"/>', txbx("Hi"), style), true)));
    expect(body).toMatch(/class="docx-shape"[^>]*data-shape-wrap="square" data-shape-align="right"/);
    expect(body).toContain("background-color:#4472C4;border:1px solid #223962;border-radius:50%;color:#FFFFFF;transform:rotate(90deg)");
  });

  it("places the members of a group", () => {
    const member = (x: number, fill: string) => `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></wps:spPr>${txbx("m")}<wps:bodyPr/></wps:wsp>`;
    const grp = `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>${member(0, "FF0000")}${member(952500, "00FF00")}</wpg:wgp>`;
    const { body } = docxToParts(docx(drawing(grp, false, "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")));
    expect(body).toContain("position:absolute;left:0px;top:0px;width:100px;height:100px;background-color:#FF0000");
    expect(body).toContain("position:absolute;left:100px;top:0px;width:100px;height:100px;background-color:#00FF00");
  });

  it("draws an old VML text box", () => {
    const { body } = docxToParts(docx('<w:p><w:r><w:pict><v:shape style="width:150pt;height:30pt" fillcolor="#ddeeff" strokecolor="navy" strokeweight="1.5pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>vml text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'));
    expect(body).toMatch(/<span class="docx-shape-box" style="width:200px;height:40px;background-color:#DDEEFF;border:2px solid #000000">.*vml text/);
  });

  it("colours a run whose colour is only a theme reference", () => {
    const { body } = docxToParts(docx('<w:p><w:r><w:rPr><w:color w:val="auto" w:themeColor="accent2" w:themeShade="BF"/></w:rPr><w:t>t</w:t></w:r></w:p>'));
    expect(body).toContain("color:#B25E25");
  });
});

describe("editing shapes and text boxes", () => {
  const TB = drawing(wsp('<a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>', txbx("Note")));
  const edit = (html: string, fn: (doc: Document) => void) => {
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
    fn(doc);
    return doc.body.innerHTML;
  };
  const xmlOf = (d: Uint8Array, html: string) => strFromU8(unzipSync(htmlToDocx(html, d))["word/document.xml"]!);

  it("saves an untouched shape byte for byte, fallback included", () => {
    const d = docx(TB);
    const xml = xmlOf(d, docxToParts(d).body);
    expect(xml).toContain("mc:Fallback");
    expect(xml).toContain('<w:t>Note</w:t>');
  });

  it("writes edited text box text back into w:txbxContent, keeping paragraph formatting", () => {
    const d = docx(TB);
    const html = edit(docxToParts(d).body, (doc) => {
      const shape = doc.querySelector(".docx-shape")!;
      shape.setAttribute("data-shape-dirty", "text");
      const ps = shape.querySelectorAll(".docx-tb-p");
      ps[0]!.textContent = "Edited note";
      const extra = doc.createElement("span");
      extra.className = "docx-tb-p";
      extra.textContent = "third line";
      ps[1]!.after(extra);
    });
    const xml = xmlOf(d, html);
    const txbx = /<w:txbxContent>(.*?)<\/w:txbxContent>/.exec(xml)![1]!;
    expect(txbx).toContain("Edited note");
    expect(txbx).toMatch(/<w:b\/>.*second/); // the bold run of paragraph 2 survives
    expect(txbx).toContain("third line");
    expect((txbx.match(/<w:p>|<w:p /g) ?? []).length).toBe(3);
    expect(xml).not.toContain("mc:Fallback"); // the stale VML copy is dropped
    expect(docxToParts(unzipSync(htmlToDocx(html, d)) && htmlToDocx(html, d)).body).toContain("Edited note");
  });

  it("writes an edited size, fill and outline", () => {
    const d = docx(TB);
    const html = edit(docxToParts(d).body, (doc) => {
      const shape = doc.querySelector(".docx-shape")!;
      shape.setAttribute("data-shape-dirty", "look");
      const box = shape.querySelector<HTMLElement>(".docx-shape-box")!;
      box.style.width = "300px";
      box.style.height = "150px";
      box.style.backgroundColor = "#00ff00";
      box.style.border = "3px solid #0000ff";
    });
    const xml = xmlOf(d, html);
    expect(xml).toContain('<wp:extent cx="2857500" cy="1428750"/>');
    expect(xml).toContain('<a:ext cx="2857500" cy="1428750"/>');
    expect(xml).toMatch(/<a:solidFill><a:srgbClr val="00FF00"\/><\/a:solidFill><a:ln w="28575"><a:solidFill><a:srgbClr val="0000FF"\/><\/a:solidFill><\/a:ln>/);
    expect(xml).not.toContain("FFF2CC");
  });

  it("removes the fill and outline when cleared", () => {
    const d = docx(TB);
    const html = edit(docxToParts(d).body, (doc) => {
      const shape = doc.querySelector(".docx-shape")!;
      shape.setAttribute("data-shape-dirty", "look");
      const box = shape.querySelector<HTMLElement>(".docx-shape-box")!;
      box.style.backgroundColor = "";
      box.style.border = "none";
    });
    expect(xmlOf(d, html)).toMatch(/<a:noFill\/><a:ln><a:noFill\/><\/a:ln>/);
  });

  it("creates a new text box that reads back", () => {
    const d = docx("<w:p/>");
    const html = '<p>Hi <span class="docx-shape" contenteditable="false" data-shape-new="textbox"><span class="docx-shape-box" style="width: 192px; height: 96px; background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0);"><span class="docx-shape-body" contenteditable="true"><span class="docx-tb-p">Fresh <strong>box</strong></span></span></span></span></p>';
    const out = htmlToDocx(html, d);
    const xml = strFromU8(unzipSync(out)["word/document.xml"]!);
    expect(xml).toContain('<wps:cNvSpPr txBox="1"/>');
    expect(xml).toContain('<wp:extent cx="1828800" cy="914400"/>');
    expect(xml).toMatch(/<w:txbxContent><w:p>.*Fresh.*<w:b\/>.*box.*<\/w:txbxContent>/);
    const back = docxToParts(out).body;
    expect(back).toMatch(/class="docx-shape"[^>]*data-docx-xml/);
    expect(back).toContain("background-color:#FFFFFF;border:1px solid #000000");
    expect(back).toContain("<strong>box</strong>");
  });
});
