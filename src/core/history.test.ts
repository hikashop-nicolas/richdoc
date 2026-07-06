import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { createDocxEditor } from "../adapters/docx/index";

// Same jsdom stubs as editor.mount.test.ts (reflow needs them to run).
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

function mount() {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const ed = createDocxEditor(host, DOCX);
  const docEl = host.querySelector(".docxedit-doc") as HTMLElement;
  return { host, ed, docEl };
}

/** Simulate a typing burst: mutate the block and fire the input event the engine listens to. */
function typeInto(docEl: HTMLElement, html: string) {
  const p = docEl.querySelector("p") as HTMLElement;
  p.innerHTML = html;
  docEl.dispatchEvent(new Event("input", { bubbles: true }));
}

describe("undo/redo history", () => {
  it("starts with nothing to undo or redo", () => {
    const { host, ed } = mount();
    expect(ed.canUndo()).toBe(false);
    expect(ed.canRedo()).toBe(false);
    ed.undo(); // no-op, must not throw
    ed.destroy();
    host.remove();
  });

  it("undoes and redoes a typing run as one step", () => {
    const { host, ed, docEl } = mount();
    typeInto(docEl, "Hie");
    typeInto(docEl, "Hiedit"); // same run: coalesces
    expect(ed.canUndo()).toBe(true);
    ed.undo();
    expect(docEl.textContent).toContain("Hi");
    expect(docEl.textContent).not.toContain("Hiedit");
    expect(ed.canRedo()).toBe(true);
    ed.redo();
    expect(docEl.textContent).toContain("Hiedit");
    ed.destroy();
    host.remove();
  });

  it("treats discrete operations (no open typing run) as separate steps", () => {
    const { host, ed, docEl } = mount();
    const p = docEl.querySelector("p") as HTMLElement;
    p.focus();
    const sel = window.getSelection()!;
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
    const open = [...host.querySelectorAll("button")].find((b) => b.title === "Insert special character") as HTMLButtonElement;
    open.click();
    ([...host.querySelectorAll(".docxedit-symbol")].find((b) => b.textContent === "→") as HTMLButtonElement).click();
    open.click();
    ([...host.querySelectorAll(".docxedit-symbol")].find((b) => b.textContent === "†") as HTMLButtonElement).click();
    expect(docEl.textContent).toContain("→");
    expect(docEl.textContent).toContain("†");
    ed.undo();
    expect(docEl.textContent).toContain("→");
    expect(docEl.textContent).not.toContain("†");
    ed.undo();
    expect(docEl.textContent).not.toContain("→");
    ed.destroy();
    host.remove();
  });

  it("handles Ctrl+Z / Ctrl+Y keydown inside the body and suppresses the native default", () => {
    const { host, ed, docEl } = mount();
    typeInto(docEl, "HiKEY");
    const p = docEl.querySelector("p") as HTMLElement;
    const zDown = new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true, cancelable: true });
    p.dispatchEvent(zDown);
    expect(zDown.defaultPrevented).toBe(true);
    expect(docEl.textContent).not.toContain("HiKEY");
    const yDown = new KeyboardEvent("keydown", { key: "y", ctrlKey: true, bubbles: true, cancelable: true });
    // The restore swapped the body's nodes; dispatch from the live block.
    (docEl.querySelector("p") as HTMLElement).dispatchEvent(yDown);
    expect(docEl.textContent).toContain("HiKEY");
    ed.destroy();
    host.remove();
  });

  it("drives undo/redo from the toolbar buttons", () => {
    const { host, ed, docEl } = mount();
    typeInto(docEl, "HiTB");
    const undoBtn = [...host.querySelectorAll("button")].find((b) => b.title.startsWith("Undo")) as HTMLButtonElement;
    const redoBtn = [...host.querySelectorAll("button")].find((b) => b.title.startsWith("Redo")) as HTMLButtonElement;
    expect(undoBtn).toBeTruthy();
    expect(redoBtn).toBeTruthy();
    undoBtn.click();
    expect(docEl.textContent).not.toContain("HiTB");
    redoBtn.click();
    expect(docEl.textContent).toContain("HiTB");
    ed.destroy();
    host.remove();
  });

  it("restores interned data-URL images intact across undo/redo", () => {
    const { host, ed, docEl } = mount();
    const dataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUg";
    typeInto(docEl, `Hi<img src="${dataUrl}">`);
    ed.undo();
    expect(docEl.querySelector("img")).toBeNull();
    ed.redo();
    const img = docEl.querySelector("img") as HTMLImageElement;
    expect(img).toBeTruthy();
    expect(img.getAttribute("src")).toBe(dataUrl);
    ed.destroy();
    host.remove();
  });

  it("keeps band typing out of the body history", () => {
    const { host, ed, docEl } = mount();
    typeInto(docEl, "HiBODY");
    ed.undo();
    expect(docEl.textContent).toContain("Hi");
    expect(ed.canUndo()).toBe(false); // only the one body step existed
    ed.destroy();
    host.remove();
  });

  it("round-trips the undone state through getBytes", async () => {
    const { host, ed, docEl } = mount();
    typeInto(docEl, "HiSAVED");
    ed.undo();
    const bytes = await ed.getBytes();
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain("HiSAVED");
    ed.destroy();
    host.remove();
  });
});
