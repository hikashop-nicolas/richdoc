import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createDocxEditor } from "../adapters/docx/index";
import { createOdtEditor } from "../adapters/odt/index";

// jsdom lacks ResizeObserver, which the engine uses to reposition comment cards.
beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
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
    (host.querySelector(".docxedit-pagesetup-btn") as HTMLElement).click();
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
    (host.querySelector(".docxedit-pagesetup-btn") as HTMLElement).click();
    const sels = [...host.querySelectorAll(".docxedit-pagesetup select")] as HTMLSelectElement[];
    sels[3]!.value = "2"; // columns
    (host.querySelector(".docxedit-pagesetup .docxedit-dialog-primary") as HTMLElement).click();
    const xml = strFromU8(unzipSync(await ed.getBytes())["styles.xml"]!);
    expect(xml).toContain('fo:column-count="2"');
    ed.destroy();
    host.remove();
  });
});
