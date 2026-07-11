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
all read as live docx-field spans and round-trip as real .doc fields (field-type
ids validated against LibreOffice), matching the docx/odt adapters.

## Notes

- The odt adapter mirrors docx throughout. A few page-setup properties have no odt home and are
  therefore docx-only: the page-number restart "start at N", the line-number "start at" /
  "restart each section", and page vertical alignment. A page border is authored as a uniform
  four-side box (an imported per-side border reads as its top side).
- Single-section documents keep their full section properties; per-section authoring is
  implemented, and untouched sections round-trip byte-for-byte. Editing a section's geometry
  preserves a custom column layout (unequal widths, separator line) when the column count is
  unchanged; only changing the count rebuilds equal-width columns.
