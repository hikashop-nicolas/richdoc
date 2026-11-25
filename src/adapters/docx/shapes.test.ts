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
    expect(body).toMatch(/^<p><span class="docx-shape"[^]*<\/span>Before {2}after<\/p>$/); // a wrapped shape leads its paragraph
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

  describe("drawn outlines", () => {
    const fill = '<a:solidFill><a:srgbClr val="DBE9F6"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>';
    const red = (ends: string) => `<a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>${ends}</a:ln>`;

    it("draws a preset from its guides, keeping the look on the box for the shape bar", () => {
      const { body } = docxToParts(docx(drawing(wsp(`<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>${fill}`))));
      expect(body).toContain('class="docx-shape-box" data-geom style="width:200px;height:100px;background-color:#DBE9F6;border:1.33px solid #1F4E79"');
      // The default arrow: a shaft half the height, a head as long as half the height.
      expect(body).toContain('<path d="M0,25 L150,25 L150,0 L200,50 L150,100 L150,75 L0,75 Z" data-fill="norm" data-stroke="1" fill="#DBE9F6" stroke="#1F4E79" stroke-width="1.33"/>');
    });

    it("takes adjust values from the file", () => {
      const { body } = docxToParts(docx(drawing(wsp(`<a:prstGeom prst="rightArrow"><a:avLst><a:gd name="adj1" fmla="val 20000"/><a:gd name="adj2" fmla="val 20000"/></a:avLst></a:prstGeom>${fill}`))));
      expect(body).toContain('d="M0,40 L180,40 L180,0 L200,50 L180,100 L180,60 L0,60 Z"');
    });

    it("shades 3-D faces and mirrors a flipped shape", () => {
      const { body } = docxToParts(docx(drawing(wsp(`<a:xfrm flipH="1"/><a:prstGeom prst="cube"><a:avLst/></a:prstGeom>${fill}`))));
      expect(body).toContain('data-fill="darkenLess" data-stroke="0" fill="#afbac5"');
      expect(body).toContain('transform="translate(200,0) scale(-1,1)"');
    });

    it("puts the line ends on a line, sized from its width", () => {
      const { body } = docxToParts(docx(drawing(wsp(`<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>${red('<a:headEnd type="oval"/><a:tailEnd type="triangle"/>')}`))));
      expect(body).toContain('<path d="M0,0 L200,100"');
      expect(body).toMatch(/data-head="solid" fill="#C00000" stroke="none" d="M200,100 L/);
      expect((body.match(/data-head=/g) ?? []).length).toBe(2);
    });

    it("keeps a line's arrowheads when its colour is edited", () => {
      const d = docx(drawing(wsp(`<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>${red('<a:prstDash val="dash"/><a:tailEnd type="triangle"/>')}`)));
      const html = docxToParts(d).body.replace("<span class=\"docx-shape\"", '<span class="docx-shape" data-shape-dirty="look"').replace("border:2px solid #C00000", "border:2px solid #0000FF");
      const xml = strFromU8(unzipSync(htmlToDocx(html, d))["word/document.xml"]!);
      expect(xml).toMatch(/<a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"\/><\/a:solidFill><a:prstDash val="dash"\/><a:tailEnd type="triangle"\/><\/a:ln>/);
    });

    it("saves a dragged floating shape's new offsets", () => {
      const anchored = `<w:p><w:r><w:drawing><wp:anchor behindDoc="0"><wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionV><wp:extent cx="1905000" cy="952500"/><wp:wrapNone/><wp:docPr id="1" name="s"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp(`<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>${fill}`)}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
      const d = docx(anchored);
      const body = docxToParts(d).body;
      // Placed from the page: centred across, 100 px down.
      expect(body).toMatch(/class="docx-shape"[^>]*data-shape-wrap="front"[^>]*data-shape-relh="page" data-shape-relv="page" data-shape-alignh="center" data-shape-y="100"/);
      // The editor keeps a drag as offsets from the same references.
      const html = body.replace('data-shape-alignh="center" data-shape-y="100"', 'data-shape-x="50" data-shape-y="120" data-shape-dirty="pos"');
      const xml = strFromU8(unzipSync(htmlToDocx(html, d))["word/document.xml"]!);
      expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>476250</wp:posOffset></wp:positionH>');
      expect(xml).toContain('<wp:positionV relativeFrom="page"><wp:posOffset>1143000</wp:posOffset></wp:positionV>');
    });

    it("saves a shape taken out of the text flow, wrapped, and put back in it", () => {
      const d = docx(drawing(wsp('<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>', txbx("Note"))));
      const body = docxToParts(d).body;
      const save = (attrs: string) => strFromU8(unzipSync(htmlToDocx(body.replace('<span class="docx-shape"', `<span class="docx-shape" data-shape-dirty="wrap" ${attrs}`), d))["word/document.xml"]!);
      const front = save('data-shape-wrap="front" data-shape-align="left" data-shape-relh="margin" data-shape-relv="para" data-shape-x="20" data-shape-y="-10" data-shape-z="251660000"');
      expect(front).toMatch(/<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251660000" behindDoc="0"[^>]*><wp:simplePos x="0" y="0"\/><wp:positionH relativeFrom="column"><wp:posOffset>190500<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-95250<\/wp:posOffset><\/wp:positionV><wp:extent cx="1905000" cy="952500"\/><wp:effectExtent l="0" t="0" r="0" b="0"\/><wp:wrapNone\/><wp:docPr id="1" name="s"\/><a:graphic>/);
      expect(front).toContain(">Note</w:t>"); // the drawing itself is kept
      const behind = save('data-shape-wrap="behind" data-shape-relh="page" data-shape-relv="page" data-shape-x="0" data-shape-y="0"');
      expect(behind).toMatch(/behindDoc="1"[^>]*><wp:simplePos x="0" y="0"\/><wp:positionH relativeFrom="page">/);
      const square = save('data-shape-wrap="square" data-shape-align="center"');
      expect(square).toMatch(/<wp:positionH relativeFrom="column"><wp:align>center<\/wp:align><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0<\/wp:posOffset><\/wp:positionV>[^]*<wp:wrapSquare wrapText="bothSides"\/>/);
      // And back in line from a floating one.
      const floating = docx(drawing(wsp('<a:prstGeom prst="rect"/>', txbx("Note")), true));
      const fb = docxToParts(floating).body.replace(/data-shape-wrap="square" data-shape-align="right"/, 'data-shape-dirty="wrap"');
      const inline = strFromU8(unzipSync(htmlToDocx(fb, floating))["word/document.xml"]!);
      expect(inline).toMatch(/<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1905000" cy="952500"\/><wp:docPr id="1" name="s"\/><a:graphic>/);
      expect(inline).not.toContain("wp:anchor");
    });

    it("saves a new stacking order", () => {
      const anchored = `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="5"><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1905000" cy="952500"/><wp:wrapNone/><wp:docPr id="1" name="s"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wsp('<a:prstGeom prst="rect"/>')}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
      const d = docx(anchored);
      const body = docxToParts(d).body;
      expect(body).toContain('data-shape-z="5"');
      const xml = strFromU8(unzipSync(htmlToDocx(body.replace('data-shape-z="5"', 'data-shape-z="9" data-shape-dirty="z"'), d))["word/document.xml"]!);
      expect(xml).toContain('relativeHeight="9"');
    });

    it("saves an inserted shape floating in front of the text", () => {
      const d = docx("<w:p><w:r><w:t>Text</w:t></w:r></w:p>");
      const html = '<p>Text<span class="docx-shape" data-shape-new="rect" data-shape-prst="star5" data-shape-wrap="front" data-shape-align="left" data-shape-relh="margin" data-shape-relv="para" data-shape-x="30" data-shape-y="0" data-shape-z="1"><span class="docx-shape-box" data-geom="" style="width:144px;height:96px;background-color:#4472c4;border:1px solid #2f528f"><span class="docx-shape-body"><span class="docx-tb-p"><br></span></span></span></span></p>';
      const xml = strFromU8(unzipSync(htmlToDocx(html, d))["word/document.xml"]!);
      expect(xml).toMatch(/<wp:anchor [^>]*relativeHeight="1" behindDoc="0"[^>]*>[^]*<wp:posOffset>285750<\/wp:posOffset>[^]*<a:prstGeom prst="star5">/);
      expect(xml).not.toContain("<wp:inline");
    });

    it("draws a custom geometry scaled from its path box", () => {
      const cust = '<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:arcTo wR="25" hR="25" stAng="10800000" swAng="10800000"/><a:close/></a:path></a:pathLst></a:custGeom>';
      const { body } = docxToParts(docx(drawing(wsp(cust + fill))));
      expect(body).toMatch(/d="M0,100 L100,0 C[^"]+ 200,0 Z"/);
    });

    it("keeps plain boxes, ellipses and default rounded rectangles as CSS boxes", () => {
      for (const prst of ["rect", "ellipse", "roundRect"]) {
        const { body } = docxToParts(docx(drawing(wsp(`<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}`))));
        expect(body).not.toContain("data-geom");
      }
    });
  });
});
