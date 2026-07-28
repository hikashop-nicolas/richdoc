#!/usr/bin/env node
// Generate the test corpus: synthetic .docx and .odt fixtures for the verification checks.
//
// Why generated rather than collected: demo/samples holds real files, whose provenance is
// not established and which therefore stay out of a public repository. Everything here is
// authored below in plain sight, so the corpus is provably synthetic and reviewable as
// text even though what it produces is binary.
//
// Each fixture is written as flat ODF (a single XML file, exactly what LibreOffice saves
// as "ODF Text Document (Flat XML)") and converted by LibreOffice to both .docx and .odt.
// Going through LibreOffice matters: it produces the same kind of real-world markup the
// editor meets in the wild, rather than the tidy markup we would write by hand.
//
// The OUTPUT is committed alongside this script. CI must not need LibreOffice merely to
// have fixtures, and the preservation check needs byte-stable inputs: regenerating them
// on every run would compare a file against a different file and prove nothing.
//
// Usage: node scripts/gen-corpus.mjs [--soffice <path>] [name ...]

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const ROOT = new URL("..", import.meta.url).pathname;
const OUT = join(ROOT, "test-corpus");

const SOFFICE_CANDIDATES = [
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
  "/usr/bin/soffice",
  "/usr/bin/libreoffice",
  "soffice",
];

// A 2x2 transparent PNG, so the image fixtures carry a real (tiny) payload.
const PNG_2X2 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP8z4AATAxUZoxq" +
  "AAAAAP//AwAWkgHhTdSFAAAAAElFTkSuQmCC";

const NS = [
  ['office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'],
  ['style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'],
  ['text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'],
  ['table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'],
  ['draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'],
  ['fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'],
  ['xlink', 'http://www.w3.org/1999/xlink'],
  ['dc', 'http://purl.org/dc/elements/1.1/'],
  ['meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0'],
  ['svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'],
  ['number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0'],
  ['loext', 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0'],
].map(([p, u]) => `xmlns:${p}="${u}"`).join(" ");

/** Wrap a fixture's styles and body in the flat-ODF envelope. */
const fodt = ({ styles = "", masterStyles = "", body }) =>
  `<?xml version="1.0" encoding="UTF-8"?>
<office:document ${NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:automatic-styles>${styles}</office:automatic-styles>
 <office:master-styles>${masterStyles || '<style:master-page style:name="Standard" style:page-layout-name="PL"/>'}</office:master-styles>
 <office:body><office:text>${body}</office:text></office:body>
</office:document>`;

const PAGE_LAYOUT =
  '<style:page-layout style:name="PL"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin="2cm"/></style:page-layout>';

// ---------------------------------------------------------------------------
// The fixtures. One per feature area the editor claims to preserve.
// ---------------------------------------------------------------------------

const FIXTURES = {
  // Runs and paragraph properties: the baseline every other fixture builds on.
  paragraphs: {
    styles:
      PAGE_LAYOUT +
      '<style:style style:name="Tb" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>' +
      '<style:style style:name="Ti" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>' +
      '<style:style style:name="Tu" style:family="text"><style:text-properties style:text-underline-style="solid"/></style:style>' +
      '<style:style style:name="Tc" style:family="text"><style:text-properties fo:color="#c00000" fo:font-size="14pt"/></style:style>' +
      '<style:style style:name="Pc" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>' +
      '<style:style style:name="Pj" style:family="paragraph"><style:paragraph-properties fo:text-align="justify" fo:margin-left="1cm" fo:text-indent="0.5cm"/></style:style>',
    body:
      '<text:h text:outline-level="1">First heading</text:h>' +
      '<text:p>A plain paragraph with <text:span text:style-name="Tb">bold</text:span>, ' +
      '<text:span text:style-name="Ti">italic</text:span>, ' +
      '<text:span text:style-name="Tu">underlined</text:span> and ' +
      '<text:span text:style-name="Tc">coloured</text:span> runs.</text:p>' +
      '<text:h text:outline-level="2">Second level heading</text:h>' +
      '<text:p text:style-name="Pc">A centred paragraph.</text:p>' +
      '<text:p text:style-name="Pj">A justified paragraph with a left margin and a first-line ' +
      'indent, long enough to wrap onto a second line so the justification is visible.</text:p>' +
      '<text:p>Trailing paragraph with a  double space and a\ttab.</text:p>',
  },

  // Bullets and numbering, including a nested level.
  lists: {
    styles:
      PAGE_LAYOUT +
      '<text:list-style style:name="LB">' +
      '<text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:space-before="0.5cm" text:min-label-width="0.5cm"/></text:list-level-style-bullet>' +
      '<text:list-level-style-bullet text:level="2" text:bullet-char="◦"><style:list-level-properties text:space-before="1cm" text:min-label-width="0.5cm"/></text:list-level-style-bullet>' +
      '</text:list-style>' +
      '<text:list-style style:name="LN">' +
      '<text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties text:space-before="0.5cm" text:min-label-width="0.5cm"/></text:list-level-style-number>' +
      '<text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"><style:list-level-properties text:space-before="1cm" text:min-label-width="0.5cm"/></text:list-level-style-number>' +
      '</text:list-style>',
    body:
      '<text:p>A bulleted list:</text:p>' +
      '<text:list text:style-name="LB">' +
      '<text:list-item><text:p>First bullet</text:p></text:list-item>' +
      '<text:list-item><text:p>Second bullet</text:p>' +
      '<text:list><text:list-item><text:p>Nested bullet</text:p></text:list-item></text:list>' +
      '</text:list-item>' +
      '<text:list-item><text:p>Third bullet</text:p></text:list-item>' +
      '</text:list>' +
      '<text:p>A numbered list:</text:p>' +
      '<text:list text:style-name="LN">' +
      '<text:list-item><text:p>First item</text:p></text:list-item>' +
      '<text:list-item><text:p>Second item</text:p>' +
      '<text:list><text:list-item><text:p>Nested letter</text:p></text:list-item></text:list>' +
      '</text:list-item>' +
      '</text:list>',
  },

  // A table with a header row, spans and cell borders.
  tables: {
    styles:
      PAGE_LAYOUT +
      '<style:style style:name="Tbl" style:family="table"><style:table-properties style:width="15cm" table:align="margins"/></style:style>' +
      '<style:style style:name="Col" style:family="table-column"><style:table-column-properties style:column-width="5cm"/></style:style>' +
      '<style:style style:name="CellH" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000" fo:background-color="#dddddd" fo:padding="0.1cm"/></style:style>' +
      '<style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000" fo:padding="0.1cm"/></style:style>',
    body:
      '<text:p>A table follows.</text:p>' +
      '<table:table table:name="T1" table:style-name="Tbl">' +
      '<table:table-column table:style-name="Col" table:number-columns-repeated="3"/>' +
      '<table:table-header-rows><table:table-row>' +
      '<table:table-cell table:style-name="CellH" office:value-type="string"><text:p>Header A</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="CellH" office:value-type="string"><text:p>Header B</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="CellH" office:value-type="string"><text:p>Header C</text:p></table:table-cell>' +
      '</table:table-row></table:table-header-rows>' +
      '<table:table-row>' +
      '<table:table-cell table:style-name="Cell" table:number-columns-spanned="2" office:value-type="string"><text:p>Spans two columns</text:p></table:table-cell>' +
      '<table:covered-table-cell/>' +
      '<table:table-cell table:style-name="Cell" office:value-type="string"><text:p>Third</text:p></table:table-cell>' +
      '</table:table-row>' +
      '<table:table-row>' +
      '<table:table-cell table:style-name="Cell" office:value-type="string"><text:p>One</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="Cell" office:value-type="string"><text:p>Two</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="Cell" office:value-type="string"><text:p>Three</text:p></table:table-cell>' +
      '</table:table-row>' +
      '</table:table>' +
      '<text:p>Text after the table.</text:p>',
  },

  // Footnotes and endnotes, whose parts live outside the body in docx.
  notes: {
    styles: PAGE_LAYOUT,
    body:
      '<text:p>A sentence with a footnote' +
      '<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation>' +
      '<text:note-body><text:p>The first footnote body.</text:p></text:note-body></text:note>' +
      ' and then more text.</text:p>' +
      '<text:p>A second sentence with an endnote' +
      '<text:note text:id="edn1" text:note-class="endnote"><text:note-citation>i</text:note-citation>' +
      '<text:note-body><text:p>The endnote body.</text:p></text:note-body></text:note>.</text:p>' +
      '<text:p>A third sentence with another footnote' +
      '<text:note text:id="ftn2" text:note-class="footnote"><text:note-citation>2</text:note-citation>' +
      '<text:note-body><text:p>The second footnote body.</text:p></text:note-body></text:note>.</text:p>',
  },

  // Comments: a point comment and one anchored over a range.
  comments: {
    styles: PAGE_LAYOUT,
    body:
      '<text:p>A paragraph carrying a point comment.' +
      '<office:annotation office:name="c1"><dc:creator>Reviewer One</dc:creator>' +
      '<dc:date>2026-01-01T09:00:00</dc:date><text:p>A first remark.</text:p></office:annotation>' +
      '</text:p>' +
      '<text:p>A paragraph where ' +
      '<office:annotation office:name="c2"><dc:creator>Reviewer Two</dc:creator>' +
      '<dc:date>2026-01-02T10:30:00</dc:date><text:p>A remark on a range.</text:p></office:annotation>' +
      'this whole phrase is commented' +
      '<office:annotation-end office:name="c2"/>' +
      ' and the rest is not.</text:p>',
  },

  // Tracked changes: an insertion and a deletion, both attributed.
  trackchanges: {
    styles: PAGE_LAYOUT,
    body:
      '<text:tracked-changes>' +
      '<text:changed-region xml:id="ct1" text:id="ct1"><text:insertion>' +
      '<office:change-info><dc:creator>Editor One</dc:creator><dc:date>2026-01-03T08:00:00</dc:date></office:change-info>' +
      '</text:insertion></text:changed-region>' +
      '<text:changed-region xml:id="ct2" text:id="ct2"><text:deletion>' +
      '<office:change-info><dc:creator>Editor Two</dc:creator><dc:date>2026-01-04T11:15:00</dc:date></office:change-info>' +
      '<text:p>a removed sentence.</text:p>' +
      '</text:deletion></text:changed-region>' +
      '</text:tracked-changes>' +
      '<text:p>This paragraph contains ' +
      '<text:change-start text:change-id="ct1"/>an inserted phrase<text:change-end text:change-id="ct1"/>' +
      ' and once contained <text:change text:change-id="ct2"/>the end.</text:p>' +
      '<text:p>An untouched paragraph after the changes.</text:p>',
  },

  // A multi-column section next to single-column text.
  sections: {
    styles:
      PAGE_LAYOUT +
      '<style:style style:name="Sec2" style:family="section"><style:section-properties>' +
      '<style:columns fo:column-count="2" fo:column-gap="0.6cm"><style:column style:rel-width="1*"/><style:column style:rel-width="1*"/></style:columns>' +
      '</style:section-properties></style:style>',
    body:
      '<text:p>Single-column text before the section.</text:p>' +
      '<text:section text:name="TwoCols" text:style-name="Sec2">' +
      '<text:p>The first paragraph inside a two-column section, written long enough that it ' +
      'flows from the first column into the second one and the layout is actually exercised.</text:p>' +
      '<text:p>A second paragraph inside the same section.</text:p>' +
      '</text:section>' +
      '<text:p>Single-column text after the section.</text:p>',
  },

  // Header and footer on the master page, with a page-number field in the footer.
  headerfooter: {
    styles: PAGE_LAYOUT,
    masterStyles:
      '<style:master-page style:name="Standard" style:page-layout-name="PL">' +
      '<style:header><text:p>The running header</text:p></style:header>' +
      '<style:footer><text:p>Page <text:page-number text:select-page="current">1</text:page-number></text:p></style:footer>' +
      '</style:master-page>',
    body:
      '<text:p>Body text on a page that has both a header and a footer.</text:p>' +
      '<text:p>A second body paragraph.</text:p>',
  },

  // Fields, which must survive as fields rather than collapsing to their cached text.
  fields: {
    styles: PAGE_LAYOUT,
    body:
      '<text:p>Page <text:page-number text:select-page="current">1</text:page-number> of ' +
      '<text:page-count>1</text:page-count>.</text:p>' +
      '<text:p>Written on <text:date text:date-value="2026-01-05">2026-01-05</text:date> at ' +
      '<text:time text:time-value="09:30:00">09:30:00</text:time>.</text:p>' +
      '<text:p>File name: <text:file-name text:display="name">corpus.odt</text:file-name>.</text:p>' +
      '<text:p>A <text:bookmark-start text:name="bm1"/>bookmarked phrase<text:bookmark-end text:name="bm1"/> ' +
      'and a <text:a xlink:href="https://example.test/page">hyperlink</text:a>.</text:p>',
  },

  // An inline image and a wrapped one, both with a real payload.
  images: {
    styles:
      PAGE_LAYOUT +
      '<style:style style:name="FrW" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="left" style:horizontal-rel="paragraph" fo:margin-right="0.3cm"/></style:style>',
    body:
      '<text:p>An inline image <draw:frame draw:name="Inline1" text:anchor-type="as-char" svg:width="1cm" svg:height="1cm">' +
      `<draw:image><office:binary-data>${PNG_2X2}</office:binary-data></draw:image></draw:frame> inside a sentence.</text:p>` +
      '<text:p><draw:frame draw:name="Wrapped1" draw:style-name="FrW" text:anchor-type="paragraph" svg:width="2cm" svg:height="2cm">' +
      `<draw:image><office:binary-data>${PNG_2X2}</office:binary-data></draw:image></draw:frame>` +
      'A paragraph with an image anchored to it and text wrapping around, long enough that the ' +
      'wrap has something to do on more than one line of the paragraph.</text:p>',
  },

  // Right-to-left text, which has its own paragraph and run properties.
  rtl: {
    styles:
      PAGE_LAYOUT +
      '<style:style style:name="Prtl" style:family="paragraph"><style:paragraph-properties style:writing-mode="rl-tb" fo:text-align="end"/></style:style>',
    body:
      '<text:p>A left-to-right paragraph before the right-to-left one.</text:p>' +
      '<text:p text:style-name="Prtl">مرحبا بالعالم</text:p>' +
      '<text:p text:style-name="Prtl">שלום עולם</text:p>' +
      '<text:p>A left-to-right paragraph after them.</text:p>',
  },

  // Vertical Japanese text, which richdoc renders and round-trips specially.
  vertical: {
    styles:
      PAGE_LAYOUT.replace(
        'fo:margin="2cm"/>',
        'fo:margin="2cm" style:writing-mode="tb-rl"/>',
      ),
    body:
      '<text:p>縦書きの段落です。</text:p>' +
      '<text:p>二つ目の段落も縦書きです。</text:p>',
  },
};

// ---------------------------------------------------------------------------

function findSoffice(explicit) {
  const list = explicit ? [explicit] : SOFFICE_CANDIDATES;
  for (const c of list) {
    try {
      execFileSync(c, ["--version"], { stdio: "ignore" });
      return c;
    } catch {
      /* try the next */
    }
  }
  return null;
}

const argv = process.argv.slice(2);
const sofficeArg = argv.includes("--soffice") ? argv[argv.indexOf("--soffice") + 1] : null;
const only = argv.filter((a) => FIXTURES[a]);

const soffice = findSoffice(sofficeArg);
if (!soffice) {
  console.error("LibreOffice not found. Install it, or pass --soffice <path>.");
  console.error("The committed fixtures under test-corpus/ mean this is only needed to CHANGE them.");
  process.exit(2);
}

const names = only.length ? only : Object.keys(FIXTURES);
const work = join(tmpdir(), `richdoc-corpus-${process.pid}`);
mkdirSync(work, { recursive: true });
mkdirSync(OUT, { recursive: true });

for (const name of names) {
  const src = join(work, `${name}.fodt`);
  writeFileSync(src, fodt(FIXTURES[name]));
  // .doc as well: the legacy writer rebuilds the whole file rather than preserving it, so
  // LibreOffice is the only judge it can have, and it needs fixtures to judge.
  for (const ext of ["docx", "odt", "doc"]) {
    const target = ext === "doc" ? "doc:MS Word 97" : ext;
    execFileSync(soffice, ["--headless", "--convert-to", target, "--outdir", work, src], { stdio: "ignore" });
    const made = join(work, `${name}.${ext}`);
    if (!existsSync(made)) {
      console.error(`${name}: LibreOffice produced no .${ext}`);
      process.exit(1);
    }
    renameSync(made, join(OUT, `${name}.${ext}`));
  }
  console.log(`${name}.docx, ${name}.odt`);
}

rmSync(work, { recursive: true, force: true });
console.log(`\n${readdirSync(OUT).length} fixtures in test-corpus/`);
