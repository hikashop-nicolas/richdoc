# richdoc - remaining gaps

The shipped feature set is in the [README](../README.md#features). This file tracks
only what is **not** done yet.

**The passthrough guarantee.** Any element the reader does not model is preserved
byte-for-byte: it is stashed in `data-docx-xml` / `data-odt-xml` (block / run / table /
image level) and re-emitted on save. So an unmodelled feature is *preserved untouched*,
not lost.

---

## Rendering fidelity

These round-trip correctly on save; only the in-editor preview is approximate.

- **Wrapped-image offset.** A square / tight wrapped image's exact offset (docx `posOffset`,
  odt `svg:x` / `svg:y`) round-trips per axis, but the editor renders the wrap by float at the
  nearest alignment, because arbitrary-offset text wrapping is not expressible in CSS. Behind /
  front images are positioned exactly (CSS-positioned and draggable).
- **Line numbers** render only in the single-column horizontal layout; other layouts (sections,
  multi-column, vertical) round-trip the setting without drawing the numbers.
- **Page vertical alignment** is previewed only for a single-page document (centre / bottom);
  multi-page and "both" (justified) round-trip without an in-editor preview.

---

## Preserved losslessly but not yet editable

These round-trip untouched today. Adding an insert/edit UI would make them authorable; they
are realistic to do, just not yet built.

- Less-common fields: the document title / subject, ASK / input fields, and similar. (The common
  fields - PAGE / NUMPAGES / TOC, the cross-reference / caption fields REF / PAGEREF / SEQ, and
  the date / time / author / file-name fields - are authored.)

---

## Out of scope

Left as the lossless passthrough they already are. Authoring these in a browser is not
realistic / not worth it; "complete" for them means surviving a save, which they do.

- Charts, SmartArt, text boxes, shapes, drawing groups (anything beyond inline images).
- Embedded OLE objects.
- Content controls / structured document tags (`w:sdt`) authoring.
- VML legacy markup (beyond image extraction).
- Equation arrays (`m:eqArr`) and boxes, and per-column matrix alignment, within math.

---

## `.doc` (Word 97-2003 binary)

The `.doc` adapter reads and writes MS-DOC and is wired into Omnitext. Read is a
normal binary parser and fairly complete; write is **from-scratch** (regenerates a
whole valid `.doc` on every save), so anything richdoc's HTML model does not
represent (macros, exotic/unmapped formatting) is dropped on save, the same
capability-gated trade-off as odt. Validated against the LibreOffice oracle. Beyond
the shipped set (text/runs, headings, lists, links, tables + cell shading, floating
and inline images, headers/footers, footnotes/endnotes, comments, TOC field,
tategaki/furigana, multi-page FKP for large docs), the known gaps are:

- **Single header/footer model**: one header + one footer for the document; no
  per-section header/footer variants and no first-page / even-odd variants.
- **Intra-row table splitting**: a table row taller than a whole page cannot be
  broken across the page boundary (would need to split a cell's content mid-row).
  Multi-row tables split cleanly at row boundaries.

PAGE / NUMPAGES / TOC and the information fields (DATE / TIME / AUTHOR / FILENAME)
all read as live docx-field spans and round-trip as real .doc fields, matching the
docx/odt adapters.

### Found by the LibreOffice oracle (2026-07-28)

`npm run check:lo` converts each fixture and richdoc's rewrite of it with LibreOffice and
compares the two readings. It is the only judge `.doc` can have, and the first time the
legacy writer has been checked automatically. It found six defects.

**Fixed:**

- **Note reference marks came out as `?`.** The FRD's `nAuto` was written as 0, which means
  the note carries a custom mark; 1 means automatically numbered.
- **A field code leaked into footer text**, rendering a page number as "Page PAGE 1". Each
  subdocument keeps its own field table and only the main document's was written.
- **A TIME field rendered as a date**, because the instruction was written bare and the
  format left to the reader. The document's own switches now travel with the field.
- **Table header rows were lost.** `sprmTDefTable`'s size field counts one byte more than
  the operand that follows, and the reader consumed the full count, so every table property
  after it was read from the wrong offset. (The writer had the mirror of the same bug.) With
  the walk aligned, the header-row flag reads and writes correctly.
- **Lists were not lists.** They were written as ordinary paragraphs with a literal bullet
  character, and the reader turned every list paragraph into a bullet without consulting the
  numbering tables, so a numbered list came back bulleted. The writer now emits real list
  definitions (PlfLst / PlfLfo) and marks each paragraph with its list and level, and the
  reader resolves the tables to tell a numbered list from a bulleted one.

**Also fixed:**

- **Bookmarks were dropped entirely.** Neither side handled them. The reader now parses the
  bookmark tables (`SttbfBkmk`, `PlcfBkf`, `PlcfBkl`) into the same markers the docx and odt
  adapters use, and the writer emits them back. The oracle gained a bookmark comparison
  first, since nothing had been guarding this.
- **A commented range collapsed to its mark**, so the comment attached at the end of the
  phrase instead of highlighting it. The range lives in the annotation bookmark tables and
  each comment's ATRD names its entry; both sides now carry it. A range crossing a paragraph
  break is skipped rather than emitted unbalanced.

**Still open:**

- **A comment's date is written as zeroes.** Word 97's ATRD has no date field. The date is a
  separate 18-byte record per comment, starting with a packed DTTM, that FIB entry 112
  (`fcAtrdExtra`) points at. Our FIB declares 108 entries, so reaching it means claiming a
  later Word format version for every document the writer produces: a decision of its own
  rather than a fix, and the reason this one is left alone.
- **Endnotes number 1, 2, 3 where the original used i, ii, iii.** Not a section property:
  both fixtures carry identical section sprms whether or not they have endnotes, and adding
  the pair that looked like the note formats changed nothing. It is in the document
  properties, whose field offset is not established, and the two files' property blocks
  differ wholesale so a byte diff does not isolate it.
- **Hyperlink text is written in the browser's default link blue** (#0000ee) rather than the
  colour the document used. The writer hardcodes it because the reader does not resolve the
  Hyperlink character style, so nothing in the editor's HTML carries a colour to preserve.
- **A list nested deeper than two levels** falls back to numbering for its lower levels. Each
  list is written as a single level definition. Written with nine, a reader finds the first
  list's levels but not the second's, which points at the per-level stride: one level parses
  correctly, so the size of a level as we write it must differ from what a reader expects
  somewhere after the first. The levels the fixtures actually use are correct.

Each open item is recorded as a known difference in `scripts/lo-check.py`, so fixing one
turns that entry into a failure until the entry is removed.

Two earlier readings of these results were wrong and are worth recording so they are not
repeated: headers and footers looked like they were dropped entirely, and odt footnote
bodies looked like they were emptied, in both cases because the check called the adapter
without handing back what the reader had returned. The editor passes every band and every
note on every save. When a check says the product loses data, suspect the check first.

## Notes

- The odt adapter mirrors docx throughout. A few page-setup properties have no odt home and are
  therefore docx-only: the page-number restart "start at N", the line-number "start at" /
  "restart each section", and page vertical alignment. A page border is authored as a uniform
  four-side box (an imported per-side border reads as its top side).
- Single-section documents keep their full section properties; per-section authoring is
  implemented, and untouched sections round-trip byte-for-byte. Editing a section's geometry
  preserves a custom column layout (unequal widths, separator line) when the column count is
  unchanged; only changing the count rebuilds equal-width columns.
