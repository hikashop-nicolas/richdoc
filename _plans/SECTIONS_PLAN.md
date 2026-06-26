# Per-section page setup (mixed page size / orientation / margins / columns)

Status: Phases 1-2 + 4 DONE (docx + odt). A document with section breaks renders each
section at its own size / orientation / margins / columns as centred, stacked, editable
page boxes, caret-preserving, stripped on save. docx reads each in-paragraph w:sectPr;
odt reads each master page's page-layout (a master-page change begins a section, emitted
as data-rdoc-secstart).

Phase 4 (authoring) DONE: an "Insert section break" control splits the document at the
caret (the new section inherits the current setup, written as data-rdoc-secbreak (docx) /
data-rdoc-secstart + a fresh master name (odt) + a data-rdoc-secedited marker), and the
Page setup dialog edits whichever section the caret is in. On save, an edited / inserted
section's geometry is regenerated: docx merges it onto the preserved w:sectPr; odt creates
or updates a per-section page-layout + master-page. Untouched sections pass through
byte-for-byte (gated on data-rdoc-secedited). Capability flag: caps.sections =
"trailing" (docx, geometry on the section's last paragraph) | "leading" (odt, on its
first paragraph) | false.

Phase 3 (per-section header/footer) DONE, including DISTINCT content per section. Each section
resolves its own header/footer source band: the reader emits opaque keys on the boundary
paragraph (data-rdoc-secheaderkey / data-rdoc-secfooterkey) and a sectionBands map (key ->
{html, path}); the engine builds one editable band per key and clones the right band over each
section box (page-numbered cumulatively, space reserved at top/bottom); getBytes emits each
band with its path. docx keys by the header/footer r:id, path = the real part (rebuildPart);
odt keys by oh:/of:<master>, path = header@/footer@<master> (applyHeaderFooter routes to that
master). A section without its own band falls back to the document default (Word's
link-to-previous). Untouched parts/sections stay byte-for-byte.

Still pending: an inserted section break inherits the default header/footer (no new part/master
is minted yet), and the per-section ruler is display-only.

Goal: render each section at its own page geometry, so e.g. an A4 portrait section
followed by an A3 landscape section displays correctly in the editor (not just
round-trips). Page size, orientation, margins, columns and headers/footers are all
per-section properties; a "section" is a run of content between section breaks.

## Today

`geometry` is ONE PageGeometry applied to the whole body. The paginator lays the body
out as one continuous single-column flow and pushes blocks to page tops with inert
spacers; one page-card size for all pages. Mid-document section breaks round-trip (the
sectPr / odt master-page change is preserved) but every page renders at the one size.

## Model

A document is a sequence of sections. In docx the section break lives in the LAST
paragraph of a section (`w:pPr/w:sectPr`, already stashed as data-docx-sectpr); the
final section is the body-level sectPr (today's `geometry`). In odt a section starts
at a paragraph whose style carries `fo:break-before` + `style:master-page-name`
(already surfaced as data-odt-*), and its geometry is that master page's page-layout.

The two formats put the break on opposite paragraphs, so the readers normalise to one
convention the engine understands: **the last paragraph of a non-final section carries
`data-rdoc-secbreak` = JSON of that section's geometry** `{w,h,mt,mr,mb,ml,cols,colGap}`.
The final section uses the document `geometry`. The engine never parses format XML.

## Rendering: section wrappers

Wrap each section's blocks in a `.docxedit-section` box sized to that section's page
width, centred horizontally (the page element is as wide as the widest section), with
the section's margins as padding. Paginate WITHIN each section wrapper:

- single-column section: reuse the existing spacer paginator (`paginate()`), using that
  section's page height; spacers go between the section's own blocks.
- multi-column section: the per-page column wrappers from the columns work, sized to
  that section.

Section wrappers stack vertically with page gaps; page cards are drawn per page at each
section's size and centred to match. Cumulative Y threads through all sections.

Why wrappers (not pure spacers): variable width cannot be expressed in one fixed-width
flow, so each section needs its own width box. Normal editing inside a section still
only moves spacers (no reparent); blocks are re-wrapped only when section structure
changes, and the caret is preserved the same way the column reflow does (remember a
(block, char-offset) pair, re-place it after).

## Save

`cleanBody` already strips spacers / column wrappers; extend it to unwrap
`.docxedit-section` too, so the saved model is the flat block sequence and the section
breaks (data-docx-sectpr / data-odt-*) round-trip unchanged.

## Phasing

1. **docx engine**: readers emit `data-rdoc-secbreak` (docx: parse the in-paragraph
   sectPr pgSz/pgMar/orient); engine `repaginateSections()` renders variable
   size/orientation/margins per section, single column, with centred cards + caret
   preservation. Verify A4 portrait -> A3 landscape. Keep the single-section spacer and
   columns paths for documents with no mid-doc break.
2. **odt parity**: odt reader resolves each section's master-page geometry and emits the
   same `data-rdoc-secbreak`.
3. **per-section columns** (combine section wrappers with the column page wrappers) and
   **per-section header/footer** (use that section's header/footer refs).
4. **authoring**: a control to insert a section break and set the section's page setup;
   the ruler edits the section at the caret.

## Out of scope (for now)

- Vertical (tategaki) sections mixed with horizontal.
- Different first-page / even-odd page setup within a section.
- Unequal-width columns.
