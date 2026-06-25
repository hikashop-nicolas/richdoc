# richdoc - capability status and remaining gaps

This is the accurate, audited picture (docx-centric; odt mirrors it unless noted).

**The passthrough guarantee.** Any element the reader does not model is preserved
byte-for-byte: it is stashed in `data-docx-xml` (block / run / table / image
level) and re-emitted on save. So an unmodelled feature is *preserved untouched*,
not lost. The only things lost are in section C below, where the surrounding
context (a paragraph, or the document body) is regenerated from the edited HTML.

---

## A. Editable (modelled, round-trips)

- Paragraphs, headings (H1-H3), nested ordered/unordered lists, with ordered-list
  numbering preserved: each list restarts independently, an explicit start number
  round-trips, and a list that continues an earlier one keeps its running count
  (docx per-list `numId` + `startOverride`; odt list-style `start-value` /
  `continue-numbering`).
- Run formatting: bold, italic, underline, strike, super/subscript, text colour,
  highlight, shading, font family, size.
- Paragraph: alignment, left indent, line spacing, space before/after.
- Named **paragraph + character styles**: apply, author new, edit existing (name,
  weight/italic/underline/strike, alignment, size, text + background colour, font).
- Tables: editable cells, cell borders (colour/style/width), column/row/indent
  resize, merge, in-cell formatting.
- Page size, margins, orientation; vertical (tategaki) + RTL writing with
  paginated layout, header/footer, and graduated rulers + magnet.
- Header / footer (the default one), with live page-number / page-count / TOC
  fields.
- Inline images, hyperlinks.
- **Floating / anchored images + text wrap**: inline, wrap (square), break
  (top-and-bottom), behind text, in front of text; alignment; alt text; behind /
  front are draggable to position. An on-select image toolbar drives it; docx maps
  to `wp:anchor` (wrap + positionH/V), odt to a `draw:frame` + graphic style. The
  picture relationship is reused on save, so size/wrap edits never re-embed.
- Comments (+ replies, reactions, resolve) and tracked changes (insert / delete /
  formatting change).
- **Mid-document section breaks**: preserved through edits and shown as a page
  boundary, so editing a multi-section document no longer flattens it. docx keeps a
  `w:sectPr` inside a paragraph; odt keeps a paragraph style's `fo:break-before` /
  `fo:break-after` and `style:master-page-name` (a new page master, i.e. a section).
  (Authoring new breaks + per-section page setup is still pending, see C1.)

---

## C. Lost when you edit (the real gaps to close, priority order)

These are the only things that do not round-trip once the document is edited,
because their context is regenerated. This is the work for "feature complete".

1. **Authoring sections / per-section page setup.** Mid-document section breaks now
   round-trip on both formats (preserved, see bucket A). What remains is a UI to
   *insert* a section break and edit per-section page setup (size / orientation /
   margins / columns / headers) rather than only preserving what a file already has.
2. **Columns (`w:cols`).** Preserved for a single section (it rides the trailing
   sectPr) but not authorable, and authored sections (C1) would need a per-section
   column control.
3. **Tab stops (`w:tabs`).** Dropped from any paragraph that is edited. Needs a
   model + ruler tab markers (browsers do not render custom tab stops natively, so
   this is real work).
4. **List numbering authoring.** Start / restart / continue now round-trip (see
   bucket A), but there is no UI yet to *set* them: a control to restart an ordered
   list at 1, continue the previous one, or set an explicit start value.
5. **Style editing depth.** The style dialog covers the common properties;
   `w:basedOn` inheritance is flattened when a style is edited, and tab stops,
   borders and the long tail of style properties are not exposed.
6. **Image layout fine detail.** Wrap + alignment + behind/front offset round-trip,
   but the offset is approximate (mapped to a CSS-positioned element, not a Word
   layout engine), the toolbar exposes "wrap" as square (a file's `wrapTight` is
   preserved on read but authored as square), and distT/distB wrap padding uses a
   fixed default rather than the file's own values.

---

## B. Preserved losslessly but not editable (passthrough; add insert UI to "do" them)

These round-trip untouched today. Adding an insert/edit UI would make them
authorable; they are realistic to do, just not yet built.

- Footnotes / endnotes (`w:footnoteReference`; footnotes.xml/endnotes.xml parts
  preserved) - no insert/edit.
- Symbols / special characters (`w:sym`) - no insert picker.
- Bookmarks (`w:bookmarkStart/End`) and cross-references - no insert UI.
- Complex fields - only PAGE / NUMPAGES / TOC are authored; date, file name,
  author, etc. are preserved but not insertable.
- First-page / even-odd header & footer parts - preserved as files; only the
  default header/footer is editable.
- Page borders, page-number restart - preserved on the trailing section, not
  authorable.

---

## Out of scope (passthrough-only; not realistically authorable client-side)

Left as the lossless passthrough they already are. Authoring these in a browser
is not realistic / not worth it; "complete" for them means they survive a save,
which they do.

- Charts, SmartArt, text boxes, shapes, drawing groups (anything beyond inline
  images).
- Embedded OLE objects.
- Content controls / structured document tags (`w:sdt`) authoring.
- Equations / math (OMML) authoring - would require an embedded equation editor
  component (e.g. MathLive). Preserved as-is unless we decide to add one.
- VML legacy markup (beyond image extraction).

---

## Notes

- Single-section documents keep their full section properties (columns, borders,
  page-number restart, etc.); multi-section flattening (C2) is the real risk.
- First/even/odd and per-section header/footer *parts* are not regenerated, so
  they survive a save; only the default header/footer is shown for editing.
- The odt adapter mirrors docx for floating images (`draw:frame` anchor-type +
  a graphic style carrying `style:wrap` / `style:run-through` / `style:horizontal-pos`,
  and `svg:x`/`svg:y` for behind/front). The remaining gaps still shared with docx
  are the section model and tab stops (C1, C3).
