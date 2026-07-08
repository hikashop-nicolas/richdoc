import { beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocxEditor } from "../../adapters/docx/index";

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
  const g = globalThis as unknown as { CSS?: { escape?: (s: string) => string } };
  if (!g.CSS?.escape) g.CSS = { ...g.CSS, escape: (s: string) => s.replace(/[^a-zA-Z0-9_-]/g, (c) => `\\${c}`) };
  const zeroRect = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  if (!Range.prototype.getBoundingClientRect) Range.prototype.getBoundingClientRect = zeroRect;
  if (!Range.prototype.getClientRects) Range.prototype.getClientRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
});

const DOCX_WITH_COMMENT = () =>
  zipSync({
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
 <w:p><w:commentRangeStart w:id="0"/><w:r><w:t>flagged</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t> rest</w:t></w:r></w:p>
</w:body></w:document>`),
    "word/_rels/document.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`,
    ),
    "word/comments.xml": strToU8(
      `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Alice" w:date="2026-01-02T00:00:00Z"><w:p w14:paraId="AAAA1111" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>please check</w:t></w:r></w:p></w:comment></w:comments>`,
    ),
  });

describe("comment editing", () => {
  it("edits a comment's text from the panel and saves it into comments.xml", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, DOCX_WITH_COMMENT());
    const card = host.querySelector(".docxedit-cmt-card") as HTMLElement;
    expect(card).toBeTruthy();
    expect(card.textContent).toContain("please check");

    (card.querySelector(".docxedit-cmt-editbtn") as HTMLButtonElement).click();
    const ta = card.querySelector(".docxedit-cmt-editinput") as HTMLTextAreaElement;
    expect(ta.value).toBe("please check");
    ta.value = "checked and approved";
    (card.querySelector(".docxedit-cmt-send") as HTMLButtonElement).click();
    expect(card.textContent).toContain("checked and approved");

    const out = unzipSync(await ed.getBytes());
    const comments = strFromU8(out["word/comments.xml"]!);
    expect(comments).toContain("checked and approved");
    expect(comments).not.toContain("please check");
    expect(comments).toContain('w:author="Alice"'); // metadata untouched
    expect(comments).toContain('w14:paraId="AAAA1111"'); // threading anchor preserved
    ed.destroy();
    host.remove();
  });

  it("shows a live word and character count", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ed = createDocxEditor(host, DOCX_WITH_COMMENT());
    const count = host.querySelector(".docxedit-count") as HTMLElement;
    expect(count.textContent).toContain("2"); // "flagged rest" = 2 words
    ed.destroy();
    host.remove();
  });
});
