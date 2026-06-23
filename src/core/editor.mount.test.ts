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

  it("returns the original bytes unchanged when nothing was edited", async () => {
    const host = document.createElement("div");
    const ed = createOdtEditor(host, ODT);
    const out = await ed.getBytes();
    expect(out).toEqual(ODT);
    ed.destroy();
  });
});
