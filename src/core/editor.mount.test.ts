import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createDocxEditor } from "../adapters/docx/index";
import { createOdtEditor } from "../adapters/odt/index";

// jsdom lacks ResizeObserver (used to reposition comment cards) and Range.getBoundingClientRect /
// getClientRects (the ruler reads the caret rect); both exist in every real browser, so stub them
// with zero-rects so reflow can run under jsdom.
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

const DOCX = zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "_rels/.rels": strToU8("<Relationships/>"),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>`,
  ),
  "word/_rels/document.xml.rels": strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  ),
});

const ODT = zipSync({
  mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
  "content.xml": strToU8(
    `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Hi</text:p></office:text></office:body></office:document-content>`,
  ),
});

describe("shared engine mount", () => {
  it("mounts a docx editor with the full toolbar and renders the body", () => {
    const host = document.createElement("div");
    const ed = createDocxEditor(host, DOCX);
    expect(host.querySelector(".docxedit-toolbar")).toBeTruthy();
    expect(host.querySelector(".docxedit-doc")?.textContent).toContain("Hi");
    expect(ed.isDirty()).toBe(false);
    ed.destroy();
  });

  it("mounts an odt editor on the same engine, with unsupported controls gated off", () => {
    const docxHost = document.createElement("div");
    const docxEd = createDocxEditor(docxHost, DOCX);
    const docxBtns = docxHost.querySelector(".docxedit-toolbar")!.querySelectorAll("button").length;

    const odtHost = document.createElement("div");
    const odtEd = createOdtEditor(odtHost, ODT);
    const odtTb = odtHost.querySelector(".docxedit-toolbar")!;
    expect(odtTb).toBeTruthy();
    expect(odtHost.querySelector(".docxedit-doc")?.textContent).toContain("Hi");

    // odt supports colour/font/alignment, but still gates off comments, images, page
    // breaks and track changes, so its toolbar is smaller than docx's.
    expect(odtTb.querySelector('input[type="color"]')).not.toBeNull();
    expect(odtTb.querySelectorAll("button").length).toBeLessThan(docxBtns);

    docxEd.destroy();
    odtEd.destroy();
  });

  it("shows the image layout toolbar on select and applies a wrap mode", () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
    const anchor =
      '<w:r><w:drawing><wp:anchor behindDoc="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>' +
      '<wp:extent cx="1143000" cy="571500"/><wp:wrapSquare wrapText="bothSides"/>' +
      '<wp:docPr id="1" name="Image 1"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="x">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
      '<a:blip r:embed="rId100"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(
        `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>hi</w:t></w:r>${anchor}</w:p></w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="x" Target="media/image1.png"/></Relationships>`,
      ),
      "word/media/image1.png": png,
    });
    const host = document.createElement("div");
    const ed = createDocxEditor(host, docx);
    const img = host.querySelector("img") as HTMLImageElement;
    expect(img).toBeTruthy();
    expect(img.getAttribute("data-rdoc-wrap")).toBe("square"); // read as floating
    const bar = host.querySelector(".docxedit-imgbar") as HTMLElement;
    expect(bar.hidden).toBe(true); // hidden until an image is selected
    img.click(); // select it -> the layout toolbar appears
    expect(bar.hidden).toBe(false);
    // Clicking "Behind text" changes the wrap mode (found by title, not position).
    const behindBtn = [...bar.querySelectorAll(".docxedit-imgbar-btn")].find((b) => b.getAttribute("title") === "Behind text") as HTMLButtonElement;
    behindBtn.click();
    expect(img.getAttribute("data-rdoc-wrap")).toBe("behind");
    expect(ed.isDirty()).toBe(true);
    ed.destroy();
  });

  it("authors ordered-list numbering via the list-numbering menu", () => {
    // Two separate ordered lists (distinct numIds, each starting at 1).
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
    const li = (numId: string, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(
        `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${li("1", "a")}${li("1", "b")}<w:p><w:r><w:t>gap</w:t></w:r></w:p>${li("2", "c")}</w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
      "word/numbering.xml": strToU8(numbering),
    });
    const host = document.createElement("div");
    document.body.appendChild(host); // attached so the caret selection registers
    const ed = createDocxEditor(host, docx);
    const ols = host.querySelectorAll(".docxedit-doc ol");
    expect(ols).toHaveLength(2); // two independent lists, both starting at 1
    expect(ols[1]!.getAttribute("start")).toBeNull();
    // Put the caret in the second list and choose "Continue previous list".
    const li2 = ols[1]!.querySelector("li")!;
    const range = document.createRange();
    range.selectNodeContents(li2);
    range.collapse(true);
    const sel = getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
    const btn = [...host.querySelectorAll("button")].find((b) => b.title === "List numbering")!;
    btn.click();
    const item = [...host.querySelectorAll(".docxedit-menu-item")].find((b) => b.textContent === "Continue previous list")!;
    item.click();
    expect(ols[1]!.getAttribute("start")).toBe("3"); // first list has 2 items -> continue at 3
    expect(ed.isDirty()).toBe(true);
    ed.destroy();
    host.remove();
  });

  it("returns the original bytes unchanged when nothing was edited", async () => {
    const host = document.createElement("div");
    const ed = createOdtEditor(host, ODT);
    const out = await ed.getBytes();
    expect(out).toEqual(ODT);
    ed.destroy();
  });

  it("renders mixed per-section page sizes and strips the section boxes on save", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const a3l = '<w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>EndOne</w:t></w:r></w:p>` +
      "<w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      `<w:sectPr>${a3l}</w:sectPr></w:body></w:document>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const boxes = [...host.querySelectorAll(".docxedit-secpage")] as HTMLElement[];
    expect(boxes).toHaveLength(2); // one box per section
    expect(boxes[0]!.style.width).toBe("794px"); // A4 portrait
    expect(boxes[1]!.style.width).toBe("1587px"); // A3 landscape (wider)
    // Edit, then save: the section boxes are unwrapped and both sections round-trip.
    const p = host.querySelector(".docxedit-doc p")!;
    p.firstChild!.textContent = "One edited";
    (host.querySelector(".docxedit-doc") as HTMLElement).dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).not.toContain("docxedit-secpage");
    expect(xml).toContain("One edited");
    expect(xml).toContain("Two");
    expect(xml).toContain('w:w="11906"'); // section 1 A4 size preserved
    expect(xml).toContain('w:w="23811"'); // section 2 A3 landscape size preserved
    expect((xml.match(/<w:sectPr/g) ?? []).length).toBe(2);
    ed.destroy();
    host.remove();
  });

  it("renders mixed per-section page sizes for odt (master-page change)", () => {
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>' +
      '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="29.7cm" fo:page-height="21cm"/></style:page-layout>' +
      "</office:automatic-styles><office:master-styles>" +
      '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
      '<style:master-page style:name="Landscape" style:page-layout-name="pm2"/>' +
      "</office:master-styles></office:document-styles>";
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">' +
      '<office:automatic-styles><style:style style:name="P2" style:family="paragraph" style:master-page-name="Landscape"/></office:automatic-styles>' +
      '<office:body><office:text><text:p>One</text:p><text:p text:style-name="P2">Two</text:p></office:text></office:body></office:document-content>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    const boxes = [...host.querySelectorAll(".docxedit-secpage")] as HTMLElement[];
    expect(boxes).toHaveLength(2);
    expect(boxes[0]!.style.width).toBe("794px"); // A4 portrait (default master)
    expect(boxes[1]!.style.width).toBe("1123px"); // landscape master (wider)
    ed.destroy();
    host.remove();
  });

  it("paginates columns and strips the column wrappers on save", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p><w:p><w:r><w:t>Gamma</w:t></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:body></w:document>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    // The body is wrapped into per-page column boxes for display.
    expect(host.querySelector(".docxedit-colpage")).toBeTruthy();
    // An edit makes it dirty; saving unwraps the column boxes back to flat paragraphs.
    const p = host.querySelector(".docxedit-doc p")!;
    p.firstChild!.textContent = "Alpha edited";
    (host.querySelector(".docxedit-doc") as HTMLElement).dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).not.toContain("docxedit-colpage"); // wrappers stripped
    expect(xml).toContain("Alpha edited");
    expect(xml).toContain("Beta");
    expect(xml).toContain("Gamma");
    expect(xml).toContain("<w:cols"); // the column section is preserved
    ed.destroy();
    host.remove();
  });

  it("authors page setup: size + columns write back to the trailing sectPr (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    (host.querySelector(".docxedit-pagesetup-btn:not(.docxedit-sectionbreak-btn)") as HTMLElement).click();
    const sels = [...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[];
    sels[0]!.value = "a3"; // page size
    sels[3]!.value = "2"; // columns
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).toContain('w:w="16845"'); // A3 width in twips (1123px * 15)
    expect(xml).toContain('w:num="2"'); // two columns
    expect(xml).not.toContain('w:w="11906"'); // old A4 size replaced
    ed.destroy();
    host.remove();
  });

  it("renders the header/footer on every section page (docx)", () => {
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>EndOne</w:t></w:r></w:p>` +
      "<w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      `<w:sectPr><w:headerReference w:type="default" r:id="rId1"/>${a4}</w:sectPr></w:body></w:document>`;
    const hdr = '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>RUNNING HEADER</w:t></w:r></w:p></w:hdr>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/header1.xml": strToU8(hdr),
      "word/_rels/document.xml.rels": strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    expect(host.querySelectorAll(".docxedit-secpage").length).toBe(2); // two sections
    const headers = [...host.querySelectorAll(".docxedit-hf-clone")].filter((c) => c.textContent?.includes("RUNNING HEADER"));
    expect(headers.length).toBe(2); // the header renders on both section pages
    ed.destroy();
    host.remove();
  });

  it("renders + saves distinct per-section headers (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId1"/>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>EndOne</w:t></w:r></w:p>` +
      "<w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      `<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>${a4}</w:sectPr></w:body></w:document>`;
    const hdr = (txt: string) => `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${txt}</w:t></w:r></w:p></w:hdr>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/header1.xml": strToU8(hdr("HEADER ONE")),
      "word/header2.xml": strToU8(hdr("HEADER TWO")),
      "word/_rels/document.xml.rels": strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/></Relationships>'),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const headers = [...host.querySelectorAll(".docxedit-hf-clone")];
    expect(headers.some((c) => c.textContent === "HEADER ONE")).toBe(true); // section 1's own header
    expect(headers.some((c) => c.textContent === "HEADER TWO")).toBe(true); // section 2's (the default)
    // Edit section 1's header; only header1.xml should change, header2.xml stays.
    const h1 = headers.find((c) => c.textContent === "HEADER ONE") as HTMLElement;
    h1.textContent = "HEADER ONE EDITED";
    h1.dispatchEvent(new Event("input", { bubbles: true }));
    const out = unzipSync(await ed.getBytes());
    expect(strFromU8(out["word/header1.xml"]!)).toContain("HEADER ONE EDITED");
    expect(strFromU8(out["word/header2.xml"]!)).toContain("HEADER TWO");
    expect(strFromU8(out["word/header2.xml"]!)).not.toContain("EDITED");
    ed.destroy();
    host.remove();
  });

  it("breaks the header link on a section, minting a new part (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>EndOne</w:t></w:r></w:p>` + // section 1: no header ref -> inherits
      "<w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      `<w:sectPr><w:headerReference w:type="default" r:id="rId1"/>${a4}</w:sectPr></w:body></w:document>`;
    const docx = zipSync({
      "[Content_Types].xml": strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/header1.xml": strToU8('<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>DEFAULT HEADER</w:t></w:r></w:p></w:hdr>'),
      "word/_rels/document.xml.rels": strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    // Section 1 inherits the default header and shows a link chip; click it to make it independent.
    const chip = host.querySelector(".docxedit-hf-link") as HTMLElement;
    expect(chip).toBeTruthy();
    chip.click();
    // Edit section 1's now-own header (the first header clone in document order).
    const clone = [...host.querySelectorAll(".docxedit-hf-clone")][0] as HTMLElement;
    clone.textContent = "SECTION 1 HEADER";
    clone.dispatchEvent(new Event("input", { bubbles: true }));
    const out = unzipSync(await ed.getBytes());
    const headerParts = Object.keys(out).filter((k) => /^word\/header\d+\.xml$/.test(k));
    expect(headerParts.length).toBe(2); // the original + a newly minted one
    const texts = headerParts.map((k) => strFromU8(out[k]!));
    expect(texts.some((t) => t.includes("SECTION 1 HEADER"))).toBe(true);
    expect(texts.some((t) => t.includes("DEFAULT HEADER"))).toBe(true);
    expect((strFromU8(out["word/document.xml"]!).match(/headerReference/g) ?? []).length).toBe(2); // section 1 + body
    ed.destroy();
    host.remove();
  });

  it("inserts a section break and authors the new section (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    // Caret in the first paragraph, then insert a section break after it.
    const p1 = host.querySelector(".docxedit-doc p")!;
    const r = document.createRange(); r.setStart(p1.firstChild!, 0); r.collapse(true);
    const sel = getSelection()!; sel.removeAllRanges(); sel.addRange(r);
    (host.querySelector(".docxedit-sectionbreak-btn") as HTMLElement).click();
    // The first paragraph now ends a section; customise it to A3 via Page setup.
    const p1b = [...host.querySelectorAll(".docxedit-doc p")].find((p) => p.textContent === "First")!;
    const r2 = document.createRange(); r2.setStart(p1b.firstChild!, 0); r2.collapse(true);
    sel.removeAllRanges(); sel.addRange(r2);
    (host.querySelector(".docxedit-pagesetup-btn:not(.docxedit-sectionbreak-btn)") as HTMLElement).click();
    ([...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[])[0]!.value = "a3";
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect((xml.match(/<w:sectPr/g) ?? []).length).toBe(2); // the inserted break + the body section
    expect(xml).toContain('w:w="16845"'); // the new section is A3 (1123px * 15)
    expect(xml).toContain('w:w="11906"'); // the trailing body section keeps its A4 size
    ed.destroy();
    host.remove();
  });

  it("authors page setup: columns write back to the page-layout (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout>' +
      '</office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Hello</text:p></office:text></office:body></office:document-content>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    (host.querySelector(".docxedit-pagesetup-btn:not(.docxedit-sectionbreak-btn)") as HTMLElement).click();
    const sels = [...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[];
    sels[3]!.value = "2"; // columns
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["styles.xml"]!);
    expect(xml).toContain('fo:column-count="2"');
    ed.destroy();
    host.remove();
  });

  it("renders + saves distinct per-section headers (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>' +
      '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>' +
      "</office:automatic-styles><office:master-styles>" +
      '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>STD HEADER</text:p></style:header></style:master-page>' +
      '<style:master-page style:name="Custom" style:page-layout-name="pm2"><style:header><text:p>CUSTOM HEADER</text:p></style:header></style:master-page>' +
      "</office:master-styles></office:document-styles>";
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">' +
      '<office:automatic-styles><style:style style:name="P2" style:family="paragraph" style:master-page-name="Custom"/></office:automatic-styles>' +
      '<office:body><office:text><text:p>One</text:p><text:p text:style-name="P2">Two</text:p></office:text></office:body></office:document-content>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    const headers = [...host.querySelectorAll(".docxedit-hf-clone")];
    expect(headers.some((c) => c.textContent === "STD HEADER")).toBe(true); // section 1 (default master)
    expect(headers.some((c) => c.textContent === "CUSTOM HEADER")).toBe(true); // section 2 (custom master)
    // Edit the custom section's header; its master changes, the default master is untouched.
    const cust = headers.find((c) => c.textContent === "CUSTOM HEADER") as HTMLElement;
    cust.textContent = "CUSTOM HEADER EDITED";
    cust.dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["styles.xml"]!);
    expect(xml).toContain("CUSTOM HEADER EDITED");
    expect(xml).toContain("STD HEADER");
    ed.destroy();
    host.remove();
  });

  it("breaks the header link on a section, writing its master (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>' +
      '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>' +
      "</office:automatic-styles><office:master-styles>" +
      '<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>STD HEADER</text:p></style:header></style:master-page>' +
      '<style:master-page style:name="Custom" style:page-layout-name="pm2"/>' + // section 2's master has no header -> inherits
      "</office:master-styles></office:document-styles>";
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">' +
      '<office:automatic-styles><style:style style:name="P2" style:family="paragraph" style:master-page-name="Custom"/></office:automatic-styles>' +
      '<office:body><office:text><text:p>One</text:p><text:p text:style-name="P2">Two</text:p></office:text></office:body></office:document-content>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    const chip = host.querySelector(".docxedit-hf-link") as HTMLElement; // section 2's header chip
    expect(chip).toBeTruthy();
    chip.click();
    const clone = [...host.querySelectorAll(".docxedit-hf-clone")][1] as HTMLElement; // section 2's header
    clone.textContent = "SECTION 2 HEADER";
    clone.dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["styles.xml"]!);
    expect((xml.match(/<style:header/g) ?? []).length).toBe(2); // Standard + Custom now each have one
    expect(xml).toContain("SECTION 2 HEADER");
    expect(xml).toContain("STD HEADER");
    ed.destroy();
    host.remove();
  });

  it("inserts a section break and authors the new section (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout>' +
      '</office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>First</text:p><text:p>Second</text:p></office:text></office:body></office:document-content>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    // Caret in the first paragraph, insert a break (the second paragraph starts the new section).
    const p1 = host.querySelector(".docxedit-doc p")!;
    const r = document.createRange(); r.setStart(p1.firstChild!, 0); r.collapse(true);
    const sel = getSelection()!; sel.removeAllRanges(); sel.addRange(r);
    (host.querySelector(".docxedit-sectionbreak-btn") as HTMLElement).click();
    // Caret in the second paragraph (the new section), customise it to A3.
    const p2 = [...host.querySelectorAll(".docxedit-doc p")].find((p) => p.textContent === "Second")!;
    const r2 = document.createRange(); r2.setStart(p2.firstChild!, 0); r2.collapse(true);
    sel.removeAllRanges(); sel.addRange(r2);
    (host.querySelector(".docxedit-pagesetup-btn:not(.docxedit-sectionbreak-btn)") as HTMLElement).click();
    ([...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[])[0]!.value = "a3";
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["styles.xml"]!);
    expect((xml.match(/<style:master-page/g) ?? []).length).toBe(2); // Standard + the new section master
    expect(xml).toContain("rdoc-sec-1"); // the new master / page-layout
    expect(xml).toContain('fo:page-width="29.713cm"'); // the new section is A3 (1123px)
    ed.destroy();
    host.remove();
  });

  it("lays vertical multi-column text into stacked bands (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2" w:space="720"/><w:textDirection w:val="tbRl"/></w:sectPr></w:body></w:document>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const band = host.querySelector<HTMLElement>(".docxedit-vband");
    expect(band).toBeTruthy(); // vertical columns lay out as band wrappers
    expect(band!.style.writingMode).toBe("vertical-rl");
    // Save: bands are stripped; the vertical multi-column section round-trips.
    const p = host.querySelector(".docxedit-vband p")!;
    p.firstChild!.textContent = "One edited";
    (host.querySelector(".docxedit-doc") as HTMLElement).dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).not.toContain("docxedit-vband");
    expect(xml).toContain('w:num="2"');
    expect(xml).toContain("tbRl");
    ed.destroy();
    host.remove();
  });

  it("renders mixed vertical + horizontal sections (docx)", () => {
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>EndOne</w:t></w:r></w:p>` + // section 1: horizontal
      "<w:p><w:r><w:t>Two</w:t></w:r></w:p>" +
      `<w:sectPr>${a4}<w:textDirection w:val="tbRl"/></w:sectPr></w:body></w:document>`; // section 2: vertical
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    expect(host.querySelector(".docxedit-page")!.classList.contains("is-vertical")).toBe(false); // per-section, not whole-page
    const boxes = [...host.querySelectorAll<HTMLElement>(".docxedit-secpage")];
    expect(boxes.some((b) => b.style.writingMode === "vertical-rl")).toBe(true); // the vertical section
    expect(boxes.some((b) => b.style.writingMode !== "vertical-rl")).toBe(true); // the horizontal section
    ed.destroy();
    host.remove();
  });

  it("renders a vertical multi-column section as bands without leaking columns to a plain section (docx)", () => {
    const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Intro one</w:t></w:r></w:p>" +
      `<w:p><w:pPr><w:sectPr>${a4}<w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>Intro two</w:t></w:r></w:p>` + // section 1: plain, NO columns
      "<w:p><w:r><w:t>本文</w:t></w:r></w:p>" +
      `<w:sectPr>${a4}<w:cols w:num="2" w:space="720"/><w:textDirection w:val="tbRl"/></w:sectPr></w:body></w:document>`; // section 2: vertical + 2 columns
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const boxes = [...host.querySelectorAll<HTMLElement>(".docxedit-secpage")];
    const plain = boxes.find((b) => !b.querySelector(".docxedit-vband"))!;
    const vcol = boxes.find((b) => b.querySelector(".docxedit-vband"))!;
    expect(vcol).toBeTruthy(); // the vertical multi-column section lays out as bands
    expect(plain).toBeTruthy();
    expect(plain.style.columnCount).toBe(""); // the plain section does NOT inherit the document's 2 columns
    ed.destroy();
    host.remove();
  });

  it("authors writing direction: toggling vertical writes back w:textDirection (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    (host.querySelector(".docxedit-pagesetup-btn:not(.docxedit-sectionbreak-btn)") as HTMLElement).click();
    const sels = [...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[];
    sels[sels.length - 1]!.value = "vertical"; // the direction select
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).toContain('w:textDirection');
    expect(xml).toContain('tbRl');
    ed.destroy();
    host.remove();
  });

  it("renders, edits and round-trips footnotes (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>Original note</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const ref = host.querySelector(".docx-fnref");
    expect(ref?.textContent).toBe("1"); // numbered by order
    const note = host.querySelector(".docxedit-note") as HTMLElement;
    expect(note?.textContent).toContain("Original note");
    note.innerHTML = "<p>Edited note</p>";
    note.dispatchEvent(new Event("input", { bubbles: true }));
    const out = unzipSync(await ed.getBytes());
    expect(strFromU8(out["word/footnotes.xml"]!)).toContain("Edited note");
    expect(strFromU8(out["word/footnotes.xml"]!)).toContain('w:type="separator"'); // separator kept
    expect(strFromU8(out["word/document.xml"]!)).toContain("footnoteReference"); // reference preserved
    ed.destroy();
    host.remove();
  });

  it("renders, edits and round-trips footnotes (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text>' +
      '<text:p>Body text<text:note text:note-class="footnote" text:id="ftn1"><text:note-citation>1</text:note-citation>' +
      "<text:note-body><text:p>Original note</text:p></text:note-body></text:note></text:p>" +
      "</office:text></office:body></office:document-content>";
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    expect(host.querySelector(".docx-fnref")?.textContent).toBe("1");
    const note = host.querySelector(".docxedit-note") as HTMLElement;
    expect(note?.textContent).toContain("Original note");
    note.innerHTML = "<p>Edited note</p>";
    note.dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["content.xml"]!);
    expect(xml).toContain("text:note");
    expect(xml).toContain("Edited note");
    ed.destroy();
    host.remove();
  });

  it("formats footnote body text and round-trips the formatting (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>Original note</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const note = host.querySelector(".docxedit-note") as HTMLElement;
    // A footnote body is a real editing host: formatting applied to it (here, bold) is saved.
    note.innerHTML = "<p><b>Bold note</b></p>";
    note.dispatchEvent(new Event("input", { bubbles: true }));
    const fx = strFromU8(unzipSync(await ed.getBytes())["word/footnotes.xml"]!);
    expect(fx).toContain("Bold note");
    expect(fx).toContain("<w:b"); // the bold run property round-trips
    ed.destroy();
    host.remove();
  });

  it("renders footnotes at the document's footnote-text style size (docx)", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const styles = '<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style></w:styles>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/styles.xml": strToU8(styles),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    // sz 20 half-points = 10pt: the note area inherits the document's footnote-text size.
    expect((host.querySelector(".docxedit-noteslayer") as HTMLElement).style.fontSize).toBe("10pt");
    ed.destroy();
    host.remove();
  });

  it("renders footnotes at the document's Footnote style size (odt)", () => {
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text>' +
      '<text:p>Body text<text:note text:note-class="footnote" text:id="ftn1"><text:note-citation>1</text:note-citation>' +
      "<text:note-body><text:p>Note</text:p></text:note-body></text:note></text:p>" +
      "</office:text></office:body></office:document-content>";
    const styles =
      '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
      '<office:styles><style:style style:name="Footnote" style:family="paragraph" style:class="extra">' +
      '<style:text-properties fo:font-size="10pt"/></style:style></office:styles></office:document-styles>';
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "styles.xml": strToU8(styles),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    expect((host.querySelector(".docxedit-noteslayer") as HTMLElement).style.fontSize).toBe("10pt");
    ed.destroy();
    host.remove();
  });

  it("places footnotes in a per-page vertical band for tategaki documents (docx)", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:textDirection w:val="tbRl"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>Vnote</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const band = host.querySelector(".docxedit-fnarea.is-vertical");
    expect(band).toBeTruthy(); // the footnote renders in a vertical band, not the doc-end area
    expect(band?.textContent).toContain("Vnote");
    expect((host.querySelector(".docxedit-noteslayer") as HTMLElement).hidden).toBe(true); // not duplicated at doc end
    ed.destroy();
    host.remove();
  });

  it("places footnotes in a per-page area below the columns (docx multi-column)", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>Colnote</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const area = host.querySelector(".docxedit-fnarea");
    expect(area).toBeTruthy();
    expect(area?.classList.contains("is-vertical")).toBe(false); // horizontal: a bottom strip
    expect(area?.textContent).toContain("Colnote");
    expect((host.querySelector(".docxedit-noteslayer") as HTMLElement).hidden).toBe(true);
    ed.destroy();
    host.remove();
  });

  it("inserts a note via the popup with the chosen kind and text (docx)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, DOCX);
    const docEl = host.querySelector(".docxedit-doc") as HTMLElement;
    const p = docEl.querySelector("p, h1, h2, h3") ?? docEl;
    docEl.focus();
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(false);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(r);
    // open the insert-note popup (default locale is en)
    const btn = host.querySelector('[title="Insert note"]') as HTMLElement;
    expect(btn).toBeTruthy();
    btn.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    btn.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    const overlay = host.querySelector(".docxedit-noteinsert")!.closest(".docxedit-dialog-overlay") as HTMLElement;
    expect(overlay.hidden).toBe(false);
    (host.querySelector(".docxedit-noteinsert .docxedit-dialog-textarea") as HTMLTextAreaElement).value = "Popup endnote";
    const enRadio = [...host.querySelectorAll<HTMLInputElement>(".docxedit-noteinsert-opt input")].find((i) => i.value === "endnote")!;
    enRadio.checked = true;
    (host.querySelector(".docxedit-noteinsert .docxedit-dialog-primary") as HTMLElement).click();
    expect(overlay.hidden).toBe(true); // closes after insert
    const newRef = [...host.querySelectorAll(".docx-fnref")].find((x) => x.getAttribute("data-fn-id")?.startsWith("rdoc-note-new"));
    expect(newRef?.getAttribute("data-fn-kind")).toBe("endnote"); // the radio choice is honoured
    expect([...host.querySelectorAll(".docxedit-note")].some((n) => n.textContent?.includes("Popup endnote"))).toBe(true);
    ed.destroy();
    host.remove();
  });

  it("removes a footnote when its reference is deleted (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>One</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
      "<w:r><w:t> two</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>FirstNote</w:t></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:t>SecondNote</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    expect(host.querySelectorAll(".docx-fnref").length).toBe(2);
    const firstRef = host.querySelector(".docx-fnref") as HTMLElement;
    const docEl = host.querySelector(".docxedit-doc") as HTMLElement;
    docEl.focus();
    const range = document.createRange();
    range.setStartAfter(firstRef);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
    // Backspace right after the reference removes it (and its note).
    docEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", bubbles: true, cancelable: true }));
    expect(host.querySelectorAll(".docx-fnref").length).toBe(1);
    const fx = strFromU8(unzipSync(await ed.getBytes())["word/footnotes.xml"]!);
    expect(fx).not.toContain("FirstNote"); // the removed note is dropped on save
    expect(fx).toContain("SecondNote"); // the kept note survives
    ed.destroy();
    host.remove();
  });

  it("places footnotes per section box (docx sections)", () => {
    const breakP = "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:pPr>" +
      "<w:r><w:t>End one</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Sec one body</w:t></w:r></w:p>" + breakP +
      "<w:p><w:r><w:t>Sec two body</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>SecOneNote</w:t></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:t>SecTwoNote</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const areas = [...host.querySelectorAll(".docxedit-fnarea")];
    expect(areas.length).toBe(2); // one per section box
    const text = areas.map((a) => a.textContent).join(" ");
    expect(text).toContain("SecOneNote");
    expect(text).toContain("SecTwoNote");
    expect((host.querySelector(".docxedit-noteslayer") as HTMLElement).hidden).toBe(true);
    ed.destroy();
    host.remove();
  });

  it("places footnotes in a vertical band for tategaki multi-column (docx)", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2"/><w:textDirection w:val="tbRl"/></w:sectPr></w:body></w:document>';
    const footnotes = '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:t>VColNote</w:t></w:r></w:p></w:footnote></w:footnotes>';
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/footnotes.xml": strToU8(footnotes),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const area = host.querySelector(".docxedit-fnarea.is-vertical");
    expect(area).toBeTruthy();
    expect(area?.textContent).toContain("VColNote");
    ed.destroy();
    host.remove();
  });

  it("renders and round-trips furigana / ruby (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:ruby><w:rubyPr><w:rubyAlign w:val=\"center\"/></w:rubyPr><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby></w:p>" +
      "<w:p><w:r><w:t>tail</w:t></w:r></w:p></w:body></w:document>";
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const ruby = host.querySelector("ruby");
    expect(ruby).toBeTruthy();
    expect(ruby!.querySelector("rt")?.textContent).toBe("かんじ"); // reading
    expect(ruby!.textContent).toContain("漢字"); // base
    // Edit (to mark dirty) then save: the ruby round-trips as a w:ruby.
    const tail = [...host.querySelectorAll(".docxedit-doc p")].pop()!;
    tail.firstChild!.textContent = "tail edited";
    (host.querySelector(".docxedit-doc") as HTMLElement).dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).toContain("<w:ruby");
    expect(xml).toContain("<w:rubyBase");
    expect(xml).toContain("かんじ");
    expect(xml).toContain("漢字");
    ed.destroy();
    host.remove();
  });

  it("keyboard shortcuts: Ctrl+Shift+F adds furigana, Ctrl+Shift+Enter inserts a section break (docx)", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>漢字テスト</w:t></w:r></w:p></w:body></w:document>";
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    const region = host.querySelector(".docxedit-doc") as HTMLElement;
    // Ctrl+Shift+F over a selection -> ruby.
    const tn = region.querySelector("p")!.firstChild!;
    const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, 2);
    const sel = getSelection()!; sel.removeAllRanges(); sel.addRange(r);
    const orig = window.prompt; window.prompt = () => "かんじ";
    try { region.dispatchEvent(new KeyboardEvent("keydown", { key: "F", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true })); } finally { window.prompt = orig; }
    expect(host.querySelector("ruby")).toBeTruthy();
    // Ctrl+Shift+Enter -> a section break (the doc lays out as sections).
    const p = region.querySelector("p")!;
    const r2 = document.createRange(); r2.setStart(p.firstChild!, 0); r2.collapse(true); sel.removeAllRanges(); sel.addRange(r2);
    region.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    expect(host.querySelector(".docxedit-secpage")).toBeTruthy();
    ed.destroy();
    host.remove();
  });

  it("authors furigana: wraps a selection in ruby (docx)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      "<w:p><w:r><w:t>漢字テスト</w:t></w:r></w:p></w:body></w:document>";
    const docx = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, docx);
    // Select the first two characters (漢字) and apply furigana.
    const tn = host.querySelector(".docxedit-doc p")!.firstChild!;
    const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, 2);
    const sel = getSelection()!; sel.removeAllRanges(); sel.addRange(r);
    const orig = window.prompt;
    window.prompt = () => "かんじ";
    try {
      ([...host.querySelectorAll("button")].find((b) => /Furigana/.test(b.title)) as HTMLElement).click();
    } finally { window.prompt = orig; }
    const ruby = host.querySelector(".docxedit-doc ruby");
    expect(ruby).toBeTruthy();
    expect(ruby!.querySelector("rt")?.textContent).toBe("かんじ");
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).toContain("<w:ruby");
    expect(xml).toContain("かんじ");
    ed.destroy();
    host.remove();
  });

  it("renders and round-trips furigana / ruby (odt)", async () => {
    const { strFromU8, unzipSync } = await import("fflate");
    const content =
      '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text>' +
      "<text:p><text:ruby><text:ruby-base>漢字</text:ruby-base><text:ruby-text>かんじ</text:ruby-text></text:ruby></text:p>" +
      "<text:p>tail</text:p></office:text></office:body></office:document-content>";
    const odt = zipSync({
      mimetype: [strToU8("application/vnd.oasis.opendocument.text"), { level: 0 }],
      "content.xml": strToU8(content),
      "META-INF/manifest.xml": strToU8("<m/>"),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createOdtEditor(host, odt);
    const ruby = host.querySelector("ruby");
    expect(ruby).toBeTruthy();
    expect(ruby!.querySelector("rt")?.textContent).toBe("かんじ");
    expect(ruby!.textContent).toContain("漢字");
    const tail = [...host.querySelectorAll(".docxedit-doc p")].pop()!;
    tail.firstChild!.textContent = "tail edited";
    (host.querySelector(".docxedit-doc") as HTMLElement).dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["content.xml"]!);
    expect(xml).toContain("text:ruby");
    expect(xml).toContain("text:ruby-base");
    expect(xml).toContain("text:ruby-text");
    expect(xml).toContain("かんじ");
    ed.destroy();
    host.remove();
  });

  it("lists the document's headings in the outline pane and updates on edit", () => {
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Detail</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>body</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const bytes = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, bytes);
    const pane = host.querySelector(".docxedit-outline") as HTMLElement;
    expect(pane.hidden).toBe(true); // collapsed by default
    (host.querySelector(".docxedit-bottombar-btn") as HTMLButtonElement).click(); // open the outline
    expect(pane.hidden).toBe(false);
    const rows = [...host.querySelectorAll(".docxedit-outline-row")].map((r) => r.textContent);
    expect(rows).toEqual(["Intro", "Detail"]);
    // clicking a row scrolls its heading into view
    const h2 = host.querySelector(".docxedit-doc h2") as HTMLElement;
    let scrolled: Element | null = null;
    h2.scrollIntoView = function () { scrolled = this; };
    ([...host.querySelectorAll(".docxedit-outline-row")].find((r) => r.textContent === "Detail") as HTMLButtonElement).click();
    expect(scrolled).toBe(h2);
    ed.destroy();
    host.remove();
  });

  it("resolves a cross-reference to a bookmark that spans a block boundary, separating the blocks", () => {
    const bm = (id: string, end = false) =>
      `<w:bookmark${end ? "End" : "Start"} w:id="${id}"${end ? "" : ` w:name="${id}"`}/>`;
    const doc = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>${bm("span1")}<w:r><w:t>beta gamma</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>delta epsilon</w:t></w:r>${bm("span1", true)}<w:r><w:t xml:space="preserve"> zeta</w:t></w:r></w:p>` +
      `<w:p><w:fldSimple w:instr=" REF span1 \\h "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>` +
      `</w:body></w:document>`;
    const bytes = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "_rels/.rels": strToU8("<Relationships/>"),
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
    });
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, bytes);
    const xref = host.querySelector(".docx-xref");
    expect(xref?.textContent).toBe("beta gamma delta epsilon"); // a space at the block boundary, not "gammadelta"
    ed.destroy();
    host.remove();
  });
});
