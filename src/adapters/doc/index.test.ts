import { describe, expect, it } from "vitest";
import type { Note } from "../../core/types";
import { isCfb, readCfb, writeCfb } from "./cfb";
import { docToHtml, docToParts } from "./read";
import { htmlToDoc } from "./write";

describe("cfb", () => {
  it("round-trips stream content (padded to regular sectors)", () => {
    // The writer keeps streams in regular sectors (padded to >= the 4096 mini cutoff), so
    // content is preserved as a prefix; readers use the FIB's fc/lcb, not the padded size.
    const streams = [
      { name: "WordDocument", data: new Uint8Array([1, 2, 3, 4, 5]) },
      { name: "1Table", data: new Uint8Array(300).fill(7) },
    ];
    const built = writeCfb(streams);
    expect(isCfb(built)).toBe(true);
    const back = readCfb(built);
    expect(Array.from(back.get("WordDocument")!.subarray(0, 5))).toEqual([1, 2, 3, 4, 5]);
    expect(back.get("1Table")!.subarray(0, 300).every((b) => b === 7)).toBe(true);
  });
});

describe("doc write -> read round trip", () => {
  it("preserves text and paragraphs", () => {
    const doc = htmlToDoc("<p>First para.</p><p>Second para.</p>");
    expect(isCfb(doc)).toBe(true);
    const html = docToHtml(doc);
    expect(html).toContain("First para.");
    expect(html).toContain("Second para.");
    expect((html.match(/<p/g) || []).length).toBe(2);
  });

  it("preserves bold / italic / underline", () => {
    const html = docToHtml(htmlToDoc("<p>a <b>bold</b> <i>it</i> <u>un</u> z</p>"));
    expect(html).toMatch(/font-weight:bold[^<]*>bold/);
    expect(html).toMatch(/font-style:italic[^<]*>it/);
    expect(html).toMatch(/text-decoration:underline[^<]*>un/);
  });

  it("preserves alignment, size and colour", () => {
    const html = docToHtml(
      htmlToDoc('<p style="text-align:center">mid</p><p><span style="font-size:20pt;color:#cc0000">big red</span></p>'),
    );
    expect(html).toContain("text-align:center");
    expect(html).toContain("font-size:20pt");
    expect(html).toContain("color:#cc0000");
  });

  it("preserves font family and highlight", () => {
    const html = docToHtml(
      htmlToDoc('<p><span style="font-family:Arial">a</span> <span style="background-color:#ffff00">b</span></p>'),
    );
    expect(html).toMatch(/font-family:Arial/);
    expect(html).toContain("background-color:#ffff00");
  });

  it("preserves hyperlinks as HYPERLINK fields", () => {
    const html = docToHtml(htmlToDoc('<p>see <a href="https://ex.com/p">here</a> ok</p>'));
    expect(html).toMatch(/<a href="https:\/\/ex\.com\/p">/);
    expect(html).toContain("here");
    expect(html).toContain("ok");
  });

  it("preserves headings as h1-h6", () => {
    const html = docToHtml(htmlToDoc("<h1>Title</h1><h2>Sub</h2><p>Body.</p>"));
    expect(html).toMatch(/<h1>Title<\/h1>/);
    expect(html).toMatch(/<h2>Sub<\/h2>/);
    expect(html).toContain("<p>Body.</p>");
  });

  it("preserves bullet and numbered lists", () => {
    const html = docToHtml(htmlToDoc("<ul><li>a</li><li>b</li></ul><ol><li>x</li><li>y</li></ol>"));
    expect(html).toMatch(/<ul><li>a<\/li><li>b<\/li><\/ul>/);
    expect(html).toMatch(/<ol><li>x<\/li><li>y<\/li><\/ol>/);
  });

  it("preserves tables", () => {
    const html = docToHtml(htmlToDoc("<p>x</p><table><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></table><p>y</p>"));
    expect(html).toMatch(/<table[^>]*><tr><td[^>]*>A1<\/td><td[^>]*>B1<\/td><\/tr>/);
    expect(html).toContain("A2");
    expect(html).toContain("B2");
  });

  it("preserves manual page breaks", () => {
    const html = docToHtml(htmlToDoc('<p>a<span data-docx-pagebreak="manual"></span>b</p>'));
    expect(html).toContain('data-docx-pagebreak="manual"');
  });

  it("preserves page geometry, columns and vertical (tategaki)", () => {
    const page = { widthPx: 900, heightPx: 1200, margin: { top: 80, right: 70, bottom: 80, left: 70 }, vertical: true, columns: 2, columnGapPx: 40 };
    const bytes = htmlToDoc("<p>x</p>", page as unknown as Parameters<typeof htmlToDoc>[1]);
    const pg = docToParts(bytes).page!;
    expect(pg.widthPx).toBe(900);
    expect(pg.heightPx).toBe(1200);
    expect(pg.margin.left).toBe(70);
    expect(pg.columns).toBe(2);
    expect(pg.vertical).toBe(true);
  });

  it("preserves footnotes and endnotes as note subdocuments", () => {
    const body =
      '<p>Intro<sup class="docx-fnref" data-fn-id="fn1" data-fn-kind="footnote"></sup> and' +
      '<sup class="docx-fnref" data-fn-id="fn2" data-fn-kind="footnote"></sup> then an' +
      '<sup class="docx-fnref" data-fn-id="en1" data-fn-kind="endnote"></sup> end.</p>';
    const notes: Note[] = [
      { id: "fn1", kind: "footnote", html: "<p>First note.</p>" },
      { id: "fn2", kind: "footnote", html: "<p>Second note.</p>" },
      { id: "en1", kind: "endnote", html: "<p>End note.</p>" },
    ];
    const parts = docToParts(htmlToDoc(body, undefined, notes));
    // Three inline references survive, tagged by kind.
    expect((parts.body.match(/docx-fnref/g) || []).length).toBe(3);
    expect(parts.body).toContain('data-fn-kind="endnote"');
    // Note bodies come back in reference order, footnotes before endnotes.
    expect(parts.notes?.map((n) => n.kind)).toEqual(["footnote", "footnote", "endnote"]);
    expect(parts.notes?.map((n) => n.html.replace(/<[^>]+>/g, ""))).toEqual(["First note.", "Second note.", "End note."]);
  });

  it("keeps run formatting inside a footnote body", () => {
    const body = '<p>x<sup class="docx-fnref" data-fn-id="fn1" data-fn-kind="footnote"></sup></p>';
    const notes: Note[] = [{ id: "fn1", kind: "footnote", html: "<p>plain <b>bold</b> tail</p>" }];
    const parts = docToParts(htmlToDoc(body, undefined, notes));
    expect(parts.notes?.[0]?.html).toMatch(/font-weight:bold[^<]*>bold/);
  });

  it("preserves comments (author + body) as an annotation subdocument", () => {
    const body =
      '<p>Alpha <span class="docx-comment-ref" data-comment-id="dc1"></span>beta ' +
      '<span class="docx-comment-ref" data-comment-id="dc2"></span>gamma.</p>';
    const comments = [
      { id: "dc1", author: "Alice Smith", text: "First remark." },
      { id: "dc2", author: "Bob Jones", text: "Second remark." },
    ];
    const parts = docToParts(htmlToDoc(body, undefined, undefined, comments));
    expect((parts.body.match(/docx-comment-ref/g) || []).length).toBe(2);
    expect(parts.comments?.map((c) => c.author)).toEqual(["Alice Smith", "Bob Jones"]);
    expect(parts.comments?.map((c) => c.text)).toEqual(["First remark.", "Second remark."]);
  });

  it("keeps footnotes and comments together in the right subdocument order", () => {
    const body =
      '<p>x<sup class="docx-fnref" data-fn-id="fn1" data-fn-kind="footnote"></sup>' +
      '<span class="docx-comment-ref" data-comment-id="dc1"></span>' +
      'y<sup class="docx-fnref" data-fn-id="en1" data-fn-kind="endnote"></sup>z</p>';
    const notes: Note[] = [
      { id: "fn1", kind: "footnote", html: "<p>a note</p>" },
      { id: "en1", kind: "endnote", html: "<p>an endnote</p>" },
    ];
    const comments = [{ id: "dc1", author: "Al", text: "a comment" }];
    const parts = docToParts(htmlToDoc(body, undefined, notes, comments));
    expect(parts.notes?.map((n) => `${n.kind}:${n.html.replace(/<[^>]+>/g, "")}`)).toEqual(["footnote:a note", "endnote:an endnote"]);
    expect(parts.comments?.[0]?.text).toBe("a comment");
    expect((parts.body.match(/docx-fnref/g) || []).length).toBe(2);
    expect((parts.body.match(/docx-comment-ref/g) || []).length).toBe(1);
  });

  it("preserves a header and footer (header/footer subdocument)", () => {
    const bytes = htmlToDoc("<p>Body.</p>", undefined, undefined, undefined, { header: "<p>My header</p>", footer: "<p>My footer</p>" });
    const parts = docToParts(bytes);
    expect(parts.header).toBe("<p>My header</p>");
    expect(parts.footer).toBe("<p>My footer</p>");
    expect(parts.body).toContain("Body.");
  });

  it("supports a footer without a header", () => {
    const parts = docToParts(htmlToDoc("<p>x</p>", undefined, undefined, undefined, { footer: "<p>just a footer</p>" }));
    expect(parts.header).toBeUndefined();
    expect(parts.footer).toBe("<p>just a footer</p>");
  });

  it("does not emit spurious images for a doc without pictures", () => {
    // Picture chars (0x01) only become <img> when a CHPX points at a Data-stream blip; a plain
    // doc has neither, so no image should ever appear.
    expect(docToHtml(htmlToDoc("<p>no image here</p>"))).not.toContain("<img");
  });

  it("is idempotent across a second round trip", () => {
    const once = docToHtml(htmlToDoc('<p><b>x</b> y <i>z</i> <a href="http://a.b/c">L</a></p>'));
    const twice = docToHtml(htmlToDoc(once));
    expect(twice).toBe(once);
  });
});
