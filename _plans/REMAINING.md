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
  numbering preserved and authorable: each list restarts independently, an explicit
  start round-trips, and a list that continues an earlier one keeps its running
  count. A list-numbering menu sets restart-at-1 / continue-previous / start-at-N
  (docx per-list `numId` + `startOverride`; odt list-style `start-value` /
  `continue-numbering`).
- Run formatting: bold, italic, underline, strike, super/subscript, text colour,
  highlight, shading, font family, size.
- Paragraph: alignment, left indent, line spacing, space before/after.
- Named **paragraph + character styles**: apply, author new, edit existing (name,
  weight/italic/underline/strike, alignment, size, text + background colour, font).
  Editing reconciles only the dialog's own properties, so the style's other
  properties (tabs, keep-with-next, ...) and its `w:basedOn` / parent inheritance
  are preserved, and inherited props are not flattened into the edited style.
- Tables: editable cells, cell borders (colour/style/width), column/row/indent
  resize, merge, in-cell formatting.
- Page size, margins, orientation; vertical (tategaki) + RTL writing with
  paginated layout, header/footer, and graduated rulers + magnet.
- **Multi-column sections**: a section's `w:cols` (docx) / page-layout
  `style:columns` (odt) is read and rendered with **true per-page balanced columns**:
  the body is bucketed into one multi-column box per page (the browser's own column
  overflow decides where each page fills; `column-fill: balance` equalises the
  columns), so columns reset per page like Word. Editing preserves the caret across
  the reflow; the wrappers are stripped on save and `w:cols` round-trips.
- Header / footer (the default one), with live page-number / page-count / TOC
  fields.
- Inline images, hyperlinks.
- **Tabs**: a tab character round-trips (docx `w:tab`, odt `text:tab`) and renders at
  tab stops (the browser's `tab-size`, i.e. the default 0.5in grid); Tab inserts one,
  Shift+Tab removes it; copy/paste yields a real tab. A paragraph's custom tab stops
  are preserved through edits.
- **Floating / anchored images + text wrap**: inline, wrap (square), tight, break
  (top-and-bottom), behind text, in front of text; alignment; alt text; the wrap
  padding (distance kept clear of text) round-trips; behind / front are draggable to
  position. An on-select image toolbar drives it; docx maps to `wp:anchor` (wrap +
  positionH/V + dist), odt to a `draw:frame` + graphic style (wrap + fo:margins). The
  picture relationship is reused on save, so size/wrap edits never re-embed.
- Comments (+ replies, reactions, resolve) and tracked changes (insert / delete /
  formatting change).
- **Mid-document section breaks**: preserved through edits and shown as a page
  boundary, so editing a multi-section document no longer flattens it. docx keeps a
  `w:sectPr` inside a paragraph; odt keeps a paragraph style's `fo:break-before` /
  `fo:break-after` and `style:master-page-name` (a new page master, i.e. a section).
- **Mixed per-section page setup (docx)**: a document with section breaks now renders
  each section at its own page size / orientation / margins / columns (e.g. A4 portrait
  then A3 landscape), as centred, stacked per-section page boxes; editing preserves the
  caret and the boxes are stripped on save. odt parity + per-section headers/footers +
  authoring are pending (see C1). Design in `_plans/SECTIONS_PLAN.md`.

---

## C. Lost when you edit (the real gaps to close, priority order)

These are the only things that do not round-trip once the document is edited,
because their context is regenerated. This is the work for "feature complete".

1. **Sections: odt rendering + headers + authoring.** docx renders mixed per-section
   page setup (bucket A); what remains: (a) odt parity (resolve each section's
   master-page geometry and emit the same data-rdoc-secbreak), (b) per-section
   header/footer in the section page boxes, and (c) a UI to *insert* a break and edit a
   section's page setup (the ruler editing the section at the caret). See
   `_plans/SECTIONS_PLAN.md`.
2. **Columns authoring.** Columns now read, render with true per-page balancing, and
   round-trip (see bucket A). What remains is only a control to *set* the column count
   on a document (the rendering already updates live on a geometry change, since
   columns are paginated, not a separate mode).
3. **Tab-stop positioning + authoring.** Tabs and a paragraph's custom stops now
   round-trip (see bucket A), but custom stop *positions* render at the default 0.5in
   grid (preserved on save, not shown at their real x), and right / center / decimal
   alignment is approximated as left. A ruler to add/move stops and honour their
   alignment is the remaining work (browsers don't render arbitrary stops natively).
4. **Style authoring depth.** Editing now preserves a style's other properties and
   its inheritance (see bucket A), but the dialog still only *authors* the common
   properties; tab stops, borders and the long tail cannot yet be set from the UI
   (tab stops overlap with C3).
5. **Image layout fine detail.** Wrap mode (incl. tight), alignment, behind/front
   offset and wrap padding now round-trip and are authorable. What remains is minor:
   square/tight use alignment only (a file's exact `posOffset` for a wrapped image is
   not honored), and behind/front offsets map to a CSS-positioned element rather than
   a true layout engine.

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
