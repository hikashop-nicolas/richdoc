import { beforeAll, describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocxEditor } from "../../adapters/docx/index";
import { isRichdocHtml, normalizeBlocks, normalizeClipboardHtml } from "./paste";

// jsdom stubs so the editor engine can mount (same as editor.mount.test.ts).
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

const htmlOf = (nodes: Node[]): string => {
  const div = document.createElement("div");
  div.append(...nodes.map((n) => n.cloneNode(true)));
  return div.innerHTML;
};

describe("isRichdocHtml (internal copy detection)", () => {
  it("recognizes editor vocabulary markers", () => {
    expect(isRichdocHtml('<p data-docx-ppr="&lt;w:pPr/&gt;">x</p>')).toBe(true);
    expect(isRichdocHtml('<table class="docx-table"><tr><td><div class="docx-cell">x</div></td></tr></table>')).toBe(true);
    expect(isRichdocHtml('<ins class="docx-ins" data-author="A">x</ins>')).toBe(true);
  });
  it("treats plain web HTML as external", () => {
    expect(isRichdocHtml("<p style=\"margin:0\"><b>Hello</b></p>")).toBe(false);
  });
});

describe("normalizeClipboardHtml: structure", () => {
  it("keeps paragraphs and headings, drops scripts/styles/classes", () => {
    const r = normalizeClipboardHtml(
      '<style>p{color:red}</style><h2 class="title" id="t">Title</h2><script>alert(1)</script><p class="lead" data-x="1">Body</p>',
    );
    expect(r.inline).toBe(false);
    expect(htmlOf(r.nodes)).toBe("<h2>Title</h2><p>Body</p>");
  });

  it("a single paragraph becomes inline content (merges at the caret)", () => {
    const r = normalizeClipboardHtml("<p>just <b>one</b> para</p>");
    expect(r.inline).toBe(true);
    expect(htmlOf(r.nodes)).toBe("just <b>one</b> para");
  });

  it("unwraps the Google Docs bold wrapper without bolding everything", () => {
    const r = normalizeClipboardHtml(
      '<b style="font-weight:normal" id="docs-internal-guid-1"><p><span style="font-weight:700">Bold</span> plain</p><p>Second</p></b>',
    );
    expect(r.inline).toBe(false);
    expect(htmlOf(r.nodes)).toBe('<p><span style="font-weight: 700;">Bold</span> plain</p><p>Second</p>');
  });

  it("divs with inline content become paragraphs; nested containers are flattened", () => {
    const r = normalizeClipboardHtml('<div><div style="text-align:center">line one</div><div>line two</div></div>');
    expect(htmlOf(r.nodes)).toBe('<p style="text-align: center;">line one</p><p>line two</p>');
  });

  it("keeps lists with nesting inside the li (writer vocabulary)", () => {
    const r = normalizeClipboardHtml("<ul><li>a<ul><li>a1</li></ul></li><li>b</li></ul>");
    expect(htmlOf(r.nodes)).toBe("<ul><li>a<ul><li>a1</li></ul></li><li>b</li></ul>");
  });

  it("maps semantic tags onto writer tags", () => {
    const r = normalizeClipboardHtml("<p><em>i</em> <strong>b</strong> <del>s</del> <ins>u</ins> <code>c</code></p>");
    expect(htmlOf(r.nodes)).toBe(
      '<i>i</i> <b>b</b> <s>s</s> <u>u</u> <span style="font-family: &quot;Courier New&quot;, monospace;">c</span>',
    );
  });

  it("keeps only safe hrefs", () => {
    const r = normalizeClipboardHtml('<p><a href="https://x.test/a">ok</a> <a href="javascript:alert(1)">bad</a></p>');
    expect(htmlOf(r.nodes)).toBe('<a href="https://x.test/a">ok</a> bad');
  });

  it("converts pre blocks to monospace paragraphs, keeping blank lines", () => {
    const r = normalizeClipboardHtml("<pre>line1\n\nline3</pre>");
    const html = htmlOf(r.nodes);
    expect(html).toContain(">line1</p>");
    expect(html).toContain("><br></p>");
    expect(html).toContain(">line3</p>");
    expect((r.nodes[0] as HTMLElement).style.fontFamily).toContain("Courier New");
  });

  it("indents blockquote content", () => {
    const r = normalizeClipboardHtml("<blockquote><p>quoted</p></blockquote>");
    expect((r.nodes[0] as HTMLElement).style.marginLeft).toBe("40px");
  });

  it("drops empty paragraphs and collapses source whitespace", () => {
    const r = normalizeClipboardHtml("<p></p>\n  <p>a\n   b</p>\n<p>&nbsp;</p><p>c</p>");
    expect(htmlOf(r.nodes)).toBe("<p>a b</p><p>c</p>");
  });
});

describe("normalizeClipboardHtml: styles", () => {
  it("whitelists modeled run styles and drops the rest", () => {
    const r = normalizeClipboardHtml(
      '<p><span style="color:#ff0000;cursor:pointer;position:absolute;font-size:14px">x</span></p>',
    );
    const span = (r.nodes[0] as HTMLElement);
    expect(span.tagName).toBe("SPAN");
    expect(span.style.color).toBeTruthy();
    expect(span.style.fontSize).toBe("14px");
    expect(span.getAttribute("style")).not.toContain("cursor");
    expect(span.getAttribute("style")).not.toContain("position");
  });

  it("unwraps spans with no surviving styles", () => {
    const r = normalizeClipboardHtml('<p><span class="x" style="cursor:pointer">plain</span></p>');
    expect(htmlOf(r.nodes)).toBe("plain");
  });

  it("drops transparent backgrounds (would become black shading on save)", () => {
    const r = normalizeClipboardHtml('<p><span style="background-color:transparent">x</span> <span style="background-color:rgba(0, 0, 0, 0)">y</span></p>');
    expect(htmlOf(r.nodes)).toBe("x y");
  });

  it("drops font sizes the writers cannot parse", () => {
    const r = normalizeClipboardHtml('<p><span style="font-size:1.2em">x</span></p>');
    expect(htmlOf(r.nodes)).toBe("x");
  });

  it("keeps paragraph alignment and indent", () => {
    const r = normalizeClipboardHtml('<p style="text-align:right;margin-left:80px;padding:12px">x</p><p>y</p>');
    const p = r.nodes[0] as HTMLElement;
    expect(p.style.textAlign).toBe("right");
    expect(p.style.marginLeft).toBe("80px");
    expect(p.getAttribute("style")).not.toContain("padding");
  });
});

describe("normalizeClipboardHtml: tables and images", () => {
  it("converts external tables to the editor table vocabulary", () => {
    const r = normalizeClipboardHtml(
      "<table><thead><tr><th>H1</th><th>H2</th></tr></thead><tbody><tr><td colspan=\"2\">joined</td></tr></tbody></table>",
    );
    const table = r.nodes[0] as HTMLTableElement;
    expect(table.className).toBe("docx-table");
    expect(table.getAttribute("contenteditable")).toBe("false");
    expect(table.rows.length).toBe(2);
    const th = table.rows[0]!.cells[0]!;
    expect(th.querySelector(".docx-cell")).toBeTruthy();
    expect(th.querySelector(".docx-cell")!.getAttribute("contenteditable")).toBe("true");
    expect(th.querySelector("b")!.textContent).toBe("H1");
    expect(table.rows[1]!.cells[0]!.colSpan).toBe(2);
  });

  it("collects external images for inlining and keeps data URLs as-is", () => {
    const r = normalizeClipboardHtml(
      '<p>pic <img src="https://x.test/a.png" alt="A" width="900" height="450"> and <img src="data:image/gif;base64,R0lGOD=="></p>',
    );
    expect(r.externalImages.length).toBe(1);
    const ext = r.externalImages[0]!;
    expect(ext.getAttribute("alt")).toBe("A");
    expect(ext.getAttribute("width")).toBe("600"); // capped, ratio kept
    expect(ext.getAttribute("height")).toBe("300");
  });

  it("flags a sole image so the clipboard file is preferred", () => {
    expect(normalizeClipboardHtml('<img src="https://x.test/a.png">').soleImage).toBe(true);
    expect(normalizeClipboardHtml('<p><img src="https://x.test/a.png">caption text</p>').soleImage).toBe(false);
  });

  it("drops javascript: image sources", () => {
    const r = normalizeClipboardHtml('<p><img src="javascript:alert(1)">x</p>');
    expect(htmlOf(r.nodes)).toBe("x");
  });
});

describe("normalizeBlocks: loose inline content", () => {
  it("wraps stray inline runs into one paragraph, preserving inter-run spaces", () => {
    const div = document.createElement("div");
    div.innerHTML = "<b>a</b> <i>b</i>";
    const blocks = normalizeBlocks(div);
    expect(blocks.length).toBe(1);
    expect(blocks[0]!.innerHTML).toBe("<b>a</b> <i>b</i>");
  });
});

describe("end to end: normalized paste output serializes through the docx writer", () => {
  const DOCX = zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    ),
    "_rels/.rels": strToU8("<Relationships/>"),
    "word/document.xml": strToU8(
      '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
    ),
    "word/_rels/document.xml.rels": strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    ),
  });

  it("writes headings, formatted runs, links, lists and tables", async () => {
    const host = document.createElement("div");
    const ed = createDocxEditor(host, DOCX);
    const docEl = host.querySelector(".docxedit-doc") as HTMLElement;
    const r = normalizeClipboardHtml(
      '<h1>Title</h1><p style="text-align:center"><b>bold</b> and <a href="https://x.test/">link</a></p>' +
        "<ul><li>item</li></ul><table><tr><td>cell</td></tr></table>",
    );
    expect(r.inline).toBe(false);
    docEl.append(...(r.nodes as HTMLElement[]));
    docEl.dispatchEvent(new Event("input", { bubbles: true }));
    const xml = strFromU8(unzipSync(await ed.getBytes())["word/document.xml"]!);
    expect(xml).toContain("Heading1");
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("w:hyperlink");
    expect(xml).toContain('w:val="center"');
    expect(xml).toContain("w:numPr");
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain("cell");
    ed.destroy();
  });
});
