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
    // The "behind text" button is the 4th wrap button; clicking it changes the wrap.
    const behindBtn = bar.querySelectorAll(".docxedit-imgbar-btn")[3] as HTMLButtonElement;
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
});
