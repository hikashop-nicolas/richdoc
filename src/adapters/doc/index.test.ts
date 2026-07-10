import { describe, expect, it } from "vitest";
import { isCfb, readCfb, writeCfb } from "./cfb";
import { docToHtml } from "./read";
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

  it("is idempotent across a second round trip", () => {
    const once = docToHtml(htmlToDoc('<p><b>x</b> y <i>z</i> <a href="http://a.b/c">L</a></p>'));
    const twice = docToHtml(htmlToDoc(once));
    expect(twice).toBe(once);
  });
});
