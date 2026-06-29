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
- **Multi-column vertical (tategaki) text**: a vertical document with `w:cols` lays out as N
  horizontal bands stacked top-to-bottom per page (CSS multicol can't do this for `vertical-rl`,
  so blocks are bucketed into band wrappers by block-axis overflow); pages advance right-to-left.
  The caret survives the reflow and `w:cols` round-trips.
- Page size, margins, orientation; vertical (tategaki) + RTL writing with
  paginated layout, header/footer, and a graduated ruler (horizontal + vertical)
  on the page holding the caret, sized to that page and following the caret between
  pages, with a drag-to-set-margin magnet. The margin handles are draggable on every
  page, including per-section pages: dragging a section page's handles rewrites that
  section's margins. A **Page setup dialog** (bottom bar) sets the page size,
  orientation, margin preset and column count of whichever section the caret is in; it
  re-renders live and writes back to the trailing `w:sectPr` (docx) / page-layout (odt)
  for the document, or the section's own geometry for a sub-section.
- **Multi-column sections**: a section's `w:cols` (docx) / page-layout
  `style:columns` (odt) is read and rendered with **true per-page balanced columns**:
  the body is bucketed into one multi-column box per page (the browser's own column
  overflow decides where each page fills; `column-fill: balance` equalises the
  columns), so columns reset per page like Word. Editing preserves the caret across
  the reflow; the wrappers are stripped on save and `w:cols` round-trips.
- Header / footer (the default one), with live page-number / page-count / TOC
  fields.
- **First-page and even/odd header & footer variants**: the document's first-page and even-page
  header/footer are read, rendered on the right pages (first on page 1, even on even pages, default
  otherwise; the reserve uses the tallest variant), and individually editable. Two "Different first
  page" / "Different odd & even pages" checkboxes in the Page setup dialog author them (turning one
  on adds the empty variant bands, off drops them). docx writes the typed `w:headerReference`
  variants, `w:titlePg`, and `w:evenAndOddHeaders` (minting settings.xml if absent); odt writes the
  even and first-page variants as `style:header-left` / `style:footer-left` and `style:header-first` /
  `style:footer-first` (ODF 1.3). An enabled-but-empty variant is still written (an empty element),
  so a blank first/even page round-trips in both formats. Design in `_plans/HEADERS_PLAN.md`.
- Inline images, hyperlinks.
- **Furigana / ruby**: `w:ruby` (docx) / `text:ruby` (odt) read, rendered as a native HTML
  `<ruby>base<rt>reading</rt></ruby>` (the browser places the reading above in horizontal text and
  to the right in vertical/tategaki, automatically), and written back. The `w:rubyPr` is preserved.
  A toolbar button authors it: wrap a selection in ruby with a prompted reading, edit the reading
  of an existing ruby, or remove it (empty reading) keeping the base text.
- **Tabs + custom tab stops**: a tab character round-trips (docx `w:tab`, odt `text:tab`); Tab
  inserts one, Shift+Tab removes it, copy/paste yields a real tab. A paragraph's custom tab stops
  round-trip and are **rendered at their real positions with alignment** (left / center / right /
  decimal, plus dot leaders): a per-paragraph layout pass sizes each tab span so the following
  segment aligns to its governing stop, resolving each tab against its own live x (so wrapped lines
  and successive tabs work) and falling back to the default 0.5in grid past the last stop. The
  horizontal ruler **authors** them: a corner type selector cycles the type for new stops, clicking
  the ruler adds a stop, dragging a marker moves it (snap magnet), clicking a marker cycles its type,
  and dragging it off the ruler removes it; edits apply to the selected blocks. Vertical (tategaki)
  text keeps the default grid (deferred).
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
- **Mixed per-section page setup + authoring**: a document with section breaks renders each
  section at its own page size / orientation / margins / columns (e.g. A4 portrait then A3
  landscape), as centred, stacked per-section page boxes; editing preserves the caret and the
  boxes are stripped on save. An **Insert section break** control splits the document at the
  caret (the new section inherits the current setup), and the **Page setup dialog** edits the
  section the caret is in. Both formats: docx regenerates that section's `w:sectPr` from its
  geometry (merging onto the preserved original); odt creates / updates a per-section
  page-layout + master-page. Untouched sections still round-trip byte-for-byte. Each section
  renders, edits and saves its own header/footer (docx multiple header/footer parts; odt
  per-master `style:header`/`style:footer`), falling back to the document default (Word's
  link-to-previous) when it has none. A corner chip on a non-main section's header/footer toggles
  that link: unlinking mints the section its own part (docx) / master header (odt), pre-filled
  with the inherited content; relinking drops it. Design in `_plans/SECTIONS_PLAN.md`.

---

## C. Lost when you edit (the real gaps to close, priority order)

These are the only things that do not round-trip once the document is edited,
because their context is regenerated. This is the work for "feature complete".

The layout / section model is now complete (see bucket A): mixed page sizes / orientation /
margins / columns per section, section-break insertion, per-section page setup, draggable
per-section ruler margins, distinct per-section header/footer with a link toggle, mixed
vertical + horizontal sections, a Page setup Direction control, vertical (tategaki) layout
including the ruler, the floating toolbar, and **multi-column vertical text (N stacked bands)**,
plus furigana. The only vertical edge left: a *mid-document section* that is BOTH vertical AND
multi-column (the whole-document vertical-columns path handles the common case; a vertical
section box with `w:cols` renders as a single vertical flow). See `_plans/SECTIONS_PLAN.md`.
2. **Tab-stop positioning + authoring** is now done (see bucket A): custom stops render at their
   real positions with left / center / right / decimal alignment and dot leaders, and the ruler
   authors them. Only vertical (tategaki) tab alignment is deferred (kept on the default grid).
3. **Style authoring depth.** Editing now preserves a style's other properties and
   its inheritance (see bucket A), but the dialog still only *authors* the common
   properties; borders and the long tail cannot yet be set from the UI (tab stops can now be set
   per-paragraph via the ruler, though not yet as part of a named style definition).
4. **Image layout fine detail.** Wrap mode (incl. tight), alignment, behind/front
   offset and wrap padding now round-trip and are authorable. What remains is minor:
   square/tight use alignment only (a file's exact `posOffset` for a wrapped image is
   not honored), and behind/front offsets map to a CSS-positioned element rather than
   a true layout engine.

---

## B. Preserved losslessly but not editable (passthrough; add insert UI to "do" them)

These round-trip untouched today. Adding an insert/edit UI would make them
authorable; they are realistic to do, just not yet built.

- Footnotes / endnotes: references render as numbered superscripts; footnotes render at the
  bottom of their page (above the footer, with a separator rule; the body reserves the space via
  the paginator), endnotes in a notes area at the document end. Bodies are full editing hosts:
  text in them is formattable with the toolbar, the floating bar and the keyboard shortcuts
  (bold/italic/colour/font), the toolbar reflects their formatting state, and it round-trips as
  run properties. The note area inherits the document's footnote-text style (docx "FootnoteText",
  odt "Footnote": font family / size / line-height / colour), falling back to a built-in size when
  the document defines none. An "Insert note" control opens a small popup to add one: the note text
  plus a footnote / endnote choice (footnote default), inserted at the caret (minting docx
  footnotes.xml / endnotes.xml + its content-type/relationship when the document has none); deleting
  its reference mark removes it
  (Backspace just after the mark or Delete just before it, since the mark is an atomic non-selectable
  superscript), dropping the body from the view and the save. Everything round-trips (docx
  footnotes.xml / endnotes.xml; odt inline `text:note`). Per-page placement covers every paginated
  layout, reserving the space via a shared two-pass measure-then-bucket pass: horizontal layouts put
  the area at the page bottom (full content width, below all columns in multi-column); vertical
  (tategaki) layouts put a band down the page's left edge (the end of the right-to-left flow), with
  the separator on the body side and the notes set vertical-rl; section documents place each note in
  its own section box, at that box's bottom (horizontal) or left band (vertical). Only the pageless
  view (no page boundaries) keeps the doc-end area. Endnotes always use the doc-end area. See
  `_plans/FOOTNOTES_PLAN.md`.
- Symbols / special characters (`w:sym`) - no insert picker.
- Bookmarks (`w:bookmarkStart/End`) and cross-references - no insert UI.
- Complex fields - only PAGE / NUMPAGES / TOC are authored; date, file name,
  author, etc. are preserved but not insertable.
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
  page-number restart, etc.); per-section authoring (C1) is the real risk.
- First-page and even/odd header/footer variants are read, rendered and authored in both formats
  (see bucket A). Per-section header/footer parts are preserved + editable.
- The odt adapter mirrors docx for floating images (`draw:frame` anchor-type +
  a graphic style carrying `style:wrap` / `style:run-through` / `style:horizontal-pos`,
  and `svg:x`/`svg:y` for behind/front). The section model and tab stops are shared with docx and
  now complete; the odt tab stops (`style:tab-stops`) render and author the same way.
