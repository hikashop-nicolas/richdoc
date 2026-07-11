import { gunzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Note } from "../../core/types";
import { isCfb, readCfb, writeCfb } from "./cfb";
import { docToHtml, docToParts, parseSepx } from "./read";
import { EQUATION_DOC_GZ_B64, TEXTBOX_DOC_GZ_B64 } from "./textbox.fixture";
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

  it("prefers a top-level stream over a same-named one nested in a storage", () => {
    // Real .doc files (e.g. one with an embedded OLE object under an ObjectPool) carry a second
    // "WordDocument" inside the nested object. A flat directory scan can pick the wrong one; the
    // reader must resolve names via the directory tree so the top-level stream wins.
    const SEC = 512;
    const ENDOFCHAIN = 0xfffffffe, FREESECT = 0xffffffff, FATSECT = 0xfffffffd, NOSTREAM = 0xffffffff;
    // Sectors: 0=FAT, 1=directory, 2..11=top WordDocument (>=4096 -> regular), 12..21=nested.
    const buf = new Uint8Array(SEC * 23);
    const dv = new DataView(buf.buffer);
    buf.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0); // CFB magic
    dv.setUint16(30, 9, true); // sector shift -> 512
    dv.setUint16(32, 6, true); // mini shift -> 64
    dv.setUint32(44, 1, true); // one FAT sector
    dv.setUint32(48, 1, true); // first directory sector
    dv.setUint32(56, 4096, true); // mini cutoff
    dv.setUint32(60, ENDOFCHAIN, true); // first mini FAT
    dv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
    for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, FREESECT, true);
    dv.setUint32(76, 0, true); // DIFAT[0] -> FAT is in sector 0
    const secOff = (s: number) => (s + 1) * SEC;
    const fat = (s: number, v: number) => dv.setUint32(secOff(0) + s * 4, v, true);
    for (let s = 0; s < 128; s++) fat(s, FREESECT);
    fat(0, FATSECT); fat(1, ENDOFCHAIN);
    for (let s = 2; s <= 10; s++) fat(s, s + 1); fat(11, ENDOFCHAIN); // top chain 2..11
    for (let s = 12; s <= 20; s++) fat(s, s + 1); fat(21, ENDOFCHAIN); // nested chain 12..21
    const entry = (i: number, name: string, type: number, left: number, right: number, child: number, start: number, size: number) => {
      const b = secOff(1) + i * 128;
      for (let j = 0; j < name.length; j++) dv.setUint16(b + j * 2, name.charCodeAt(j), true);
      dv.setUint16(b + 64, (name.length + 1) * 2, true); // name length incl. terminator
      buf[b + 66] = type;
      dv.setUint32(b + 68, left, true);
      dv.setUint32(b + 72, right, true);
      dv.setUint32(b + 76, child, true);
      dv.setUint32(b + 116, start, true);
      dv.setUint32(b + 120, size, true);
    };
    // Root -> child "OP" (storage). OP's right sibling is the top-level WordDocument; OP's child
    // is the nested WordDocument. So both are reachable, at different depths.
    entry(0, "Root Entry", 5, NOSTREAM, NOSTREAM, 1, ENDOFCHAIN, 0);
    entry(1, "OP", 1, NOSTREAM, 2, 3, 0, 0);
    entry(2, "WordDocument", 2, NOSTREAM, NOSTREAM, NOSTREAM, 2, 5000); // top-level
    entry(3, "WordDocument", 2, NOSTREAM, NOSTREAM, NOSTREAM, 12, 5000); // nested in OP
    buf.set([0x54, 0x4f, 0x50, 0x21], secOff(2)); // "TOP!" at start of top data
    buf.set([0x4e, 0x45, 0x53, 0x54], secOff(12)); // "NEST" at start of nested data
    const streams = readCfb(buf);
    expect(Array.from(streams.get("WordDocument")!.subarray(0, 4))).toEqual([0x54, 0x4f, 0x50, 0x21]);
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

  it("preserves table cell shading (sprmTDefTableShd)", () => {
    const html = docToHtml(htmlToDoc('<table><tr><td style="background-color:#800000">H</td><td style="background-color:#800000">I</td></tr><tr><td>a</td><td>b</td></tr></table>'));
    // header cells keep the fill; the data row stays unshaded
    expect(html).toMatch(/<td[^>]*background-color:#800000[^>]*>H<\/td>/);
    expect(html).toMatch(/<td[^>]*background-color:#800000[^>]*>I<\/td>/);
    expect(html).not.toMatch(/<td[^>]*background-color[^>]*>a<\/td>/);
  });

  it("preserves paragraph spacing (space before / after)", () => {
    const html = docToHtml(htmlToDoc('<p style="margin-top:24px;margin-bottom:12px">s</p>'));
    expect(html).toContain("margin-top:24px");
    expect(html).toContain("margin-bottom:12px");
  });

  it("preserves a floating picture (FSPA anchor + OfficeArt drawing)", () => {
    const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAFElEQVR42mP8z8BQz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC";
    const html = docToHtml(htmlToDoc(`<p>C<img class="docx-float" src="${png}" data-reserve="1" style="left:40px;top:45px;width:120px;height:100px"></p>`));
    const m = html.match(/<img class="docx-float"[^>]*>/);
    expect(m).not.toBeNull();
    expect(m![0]).toContain("data:image/png"); // the blip round-trips through the delay stream
    expect(m![0]).toMatch(/width:120px/);
    expect(m![0]).toMatch(/data-reserve="1"/); // top-and-bottom wrap survives via the FSPA flags
  });

  it("writes a large document across multiple FKP pages", () => {
    // A single 512-byte FKP page holds only a few dozen runs; this many paragraphs forces the
    // writer to split CHPX/PAPX across pages and the PlcfBte to map each FC range to its page.
    const n = 400;
    const src = Array.from({ length: n }, (_, i) => `<p>Paragraph number ${i} with some <b>bold</b> text.</p>`).join("");
    const html = docToHtml(htmlToDoc(src));
    const paras = html.match(/<p>/g) || [];
    expect(paras.length).toBe(n);
    expect(html).toContain("Paragraph number 0 ");
    expect(html).toContain(`Paragraph number ${n - 1} `);
    expect(html).toMatch(/font-weight:bold">bold</); // formatting survives on a late page
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

  it("round-trips an embedded PNG image (Data stream blip)", () => {
    // A 1x1 PNG as a data URL, written into the Data stream and read back as an <img>.
    const png =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC";
    const body = `<p>x</p><p><img src="data:image/png;base64,${png}" style="width:16px;height:16px"></p>`;
    const html = docToHtml(htmlToDoc(body));
    expect(html).toMatch(/<img src="data:image\/png;base64,iVBORw0KGgo/);
  });

  it("round-trips PAGE / NUMPAGES fields", () => {
    const body = '<p>Page <span class="docx-field" data-field="PAGE">1</span> of <span class="docx-field" data-field="NUMPAGES">1</span></p>';
    const html = docToHtml(htmlToDoc(body));
    expect(html).toMatch(/data-field="PAGE"/);
    expect(html).toMatch(/data-field="NUMPAGES"/);
  });

  it("round-trips DATE / TIME / AUTHOR / FILENAME info fields, keeping the cached value", () => {
    const body =
      '<p>Built <span class="docx-field" data-field="DATE" contenteditable="false">2026-01-02</span>' +
      ' at <span class="docx-field" data-field="TIME" contenteditable="false">09:30</span>' +
      ' by <span class="docx-field" data-field="AUTHOR" contenteditable="false">Jane</span>' +
      ' in <span class="docx-field" data-field="FILENAME" contenteditable="false">report.doc</span></p>';
    const html = docToHtml(htmlToDoc(body));
    expect(html).toContain('data-field="DATE"');
    expect(html).toContain('data-field="TIME"');
    expect(html).toContain('data-field="AUTHOR"');
    expect(html).toContain('data-field="FILENAME"');
    // The cached snapshot is preserved inside each span (info fields are not recomputed).
    expect(html).toMatch(/data-field="DATE"[^>]*>2026-01-02</);
    expect(html).toMatch(/data-field="AUTHOR"[^>]*>Jane</);
    expect(html).toMatch(/data-field="FILENAME"[^>]*>report\.doc</);
  });

  it("round-trips ruby (furigana) as an EQ field", () => {
    const html = docToHtml(htmlToDoc("<p>x<ruby>漢字<rt>かんじ</rt></ruby>y</p>"));
    expect(html).toMatch(/<ruby>漢字<rt>かんじ<\/rt><\/ruby>/);
  });

  it("round-trips tracked insertions and deletions", () => {
    const body = '<p>a <ins class="docx-ins" data-author="Al" data-date="2026-07-11">new</ins> b <del class="docx-del" data-author="Bo" data-date="2026-06-01">old</del> c</p>';
    const html = docToHtml(htmlToDoc(body));
    expect(html).toMatch(/<ins class="docx-ins"[^>]*data-author="Al"[^>]*>new<\/ins>/);
    // Deletion author must survive independently of the insertion (distinct rmark sprms).
    expect(html).toMatch(/<del class="docx-del"[^>]*data-author="Bo"[^>]*>old<\/del>/);
  });

  it("round-trips a multi-section document (per-section geometry)", () => {
    const sec1 = JSON.stringify({ w: 794, h: 1123, mt: 76, mr: 76, mb: 76, ml: 76 }).replace(/"/g, "&quot;");
    const body = `<p data-rdoc-secbreak="${sec1}">portrait</p><p>landscape</p>`;
    const page = { widthPx: 1123, heightPx: 794, margin: { top: 76, right: 76, bottom: 76, left: 76 } };
    const parts = docToParts(htmlToDoc(body, page as unknown as Parameters<typeof htmlToDoc>[1]));
    expect(parts.page?.widthPx).toBe(1123); // last section = landscape
    expect(parts.body).toMatch(/data-rdoc-secbreak="[^"]*&quot;w&quot;:794/); // first section = portrait
  });

  it("reads text boxes and places them after their anchor paragraph", () => {
    const bytes = gunzipSync(Uint8Array.from(atob(TEXTBOX_DOC_GZ_B64), (c) => c.charCodeAt(0)));
    const html = docToParts(bytes).body;
    // Each box renders as a bordered div in document order, right after its anchor paragraph.
    // Styling is via the .docx-textbox class in the stylesheet, not inline.
    expect(html).toMatch(/Anchor one\.<\/p><div class="docx-textbox"><p>Box ALPHA content\.<\/p><\/div>/);
    expect(html).toMatch(/Anchor two\.<\/p><div class="docx-textbox"><p>Box BETA content\.<\/p><\/div>/);
    // Body text around the boxes is preserved in order.
    expect(html).toMatch(/First body line\..*Middle body line\..*Last body line\./);
  });

  it("round-trips a table of contents as a TOC field", () => {
    const body =
      '<h1>Alpha</h1><p>a</p><h2>Beta</h2><p>b</p>' +
      '<div class="docx-field-toc"><div class="docx-field-toc-title">Contents</div>' +
      '<div class="docx-field-toc-row toc-h1"><span class="docx-field-toc-text">Alpha</span><span class="docx-field-toc-page">1</span></div>' +
      '<div class="docx-field-toc-row toc-h2"><span class="docx-field-toc-text">Beta</span><span class="docx-field-toc-page">1</span></div></div>' +
      '<p>After.</p>';
    const html = docToParts(htmlToDoc(body)).body;
    // The TOC becomes an empty div the engine repopulates from the headings; entries/para marks
    // of the cached field result are dropped, and the surrounding content is preserved in order.
    expect(html).toContain('<div class="docx-field-toc"></div>');
    expect(html).not.toContain("docx-field-toc-row");
    expect(html).toMatch(/<h1>Alpha<\/h1>.*<h2>Beta<\/h2>.*<div class="docx-field-toc"><\/div><p>After\.<\/p>/);
  });

  it("reads an MS Equation object as a MathML equation span", () => {
    const bytes = gunzipSync(Uint8Array.from(atob(EQUATION_DOC_GZ_B64), (c) => c.charCodeAt(0)));
    const html = docToParts(bytes).body;
    expect(html).toContain('<span class="docx-eq" data-rdoc-eq');
    expect(html).toContain("<msup><mi>x</mi><mn>2</mn></msup>"); // x squared
    expect(html).toMatch(/The formula is .*<\/span> here\./); // placed inline at its field
    // On save the equation degrades to the math's text content (no OLE synthesis).
    expect(docToHtml(htmlToDoc(html))).toContain("x2+1");
  });

  it("resolves section margins: negative bottom -> footer distance, gutter added to left", () => {
    // A SEPX matching a real Word doc (MULTINAT): left 2592tw, bottom -1440tw (negative -> use the
    // 720tw footer distance), plus a 245tw binding gutter added to the left. Twips / 15 = px.
    const sprm = (op: number, val: number) => [op & 0xff, (op >> 8) & 0xff, val & 0xff, (val >> 8) & 0xff];
    const body = [
      ...sprm(0xb01f, 12240), // page width 8.5in
      ...sprm(0xb020, 15840), // page height 11in
      ...sprm(0xb021, 2592), // left margin (before gutter)
      ...sprm(0xb022, 1080), // right margin
      ...sprm(0x9023, 1440), // top margin (positive -> used directly)
      ...sprm(0x9024, -1440 & 0xffff), // bottom margin negative -> footer distance
      ...sprm(0xb025, 245), // gutter
      ...sprm(0xb018, 720), // footer distance from bottom
    ];
    const wd = Uint8Array.from([body.length & 0xff, (body.length >> 8) & 0xff, ...body]);
    const geom = parseSepx(wd, 0)!;
    expect(geom.margin.left).toBe(Math.round((2592 + 245) / 15)); // 189
    expect(geom.margin.right).toBe(72);
    expect(geom.margin.top).toBe(96);
    expect(geom.margin.bottom).toBe(48); // footer distance, not the raw -96
  });

  it("is idempotent across a second round trip", () => {
    const once = docToHtml(htmlToDoc('<p><b>x</b> y <i>z</i> <a href="http://a.b/c">L</a></p>'));
    const twice = docToHtml(htmlToDoc(once));
    expect(twice).toBe(once);
  });
});
