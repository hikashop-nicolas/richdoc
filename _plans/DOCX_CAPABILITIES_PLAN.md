# richdoc - docx/odt capability gaps: plan

## Context

The engine edits text, common run/paragraph formatting, comments, track changes,
images, header/footer, page margins (ruler), zoom and pagination, and round-trips
everything it does not model as passthrough. Several real Office capabilities are
preserved-but-not-editable, missing controls, or dropped on edit. This plan
catalogues them and sequences the work. Each workstream is independently
shippable; nothing here changes the passthrough guarantee (untouched content
still round-trips byte-for-byte).

Two families of gap:
- **Editing fidelity**: tables, fields/footnotes/symbols, a few run/paragraph
  properties, list nesting.
- **Page format**: more page sizes, orientation, and vertical (Japanese) writing.

## Current state (verified in code)

### Page geometry
- `PageGeometry` (core/types.ts) is `{ widthPx, heightPx, margin }` in CSS px.
  It is fully arbitrary, so **any** page size and orientation already *renders*
  correctly (a landscape file has `widthPx > heightPx` and draws right). The
  model is not the limit.
- `PageSizeName` is only `"a4" | "letter"`; `PAGE_SIZES` (core/page.ts) holds
  just those two. This registry is only the *default* used when a file declares
  no geometry.
- docx read `parsePageGeometry` reads `w:pgSz` `@w`/`@h` and `w:pgMar`. It does
  not read `w:orient` (unnecessary, w/h are explicit) nor any text direction.
- docx write `applyPageMargins` writes **only `w:pgMar`**. It never writes
  `w:pgSz`, so a size or orientation change cannot be persisted today.
- odt read pulls margins from the `style:page-layout`. odt write
  `applyPageMargins` writes **only `fo:margin-*`**; never `fo:page-width/height`
  or `style:print-orientation`.
- No writing-mode anywhere: not read, not written, not rendered. Pagination,
  rulers and zoom all assume horizontal top-to-bottom flow.
- Omnitext exposes `pageSize: "a4" | "letter"` (settings + New dialog); blank
  templates support the same two.

### Editing fidelity (docx read.ts)
- **Tables** (`w:tbl`): rendered as a read-only HTML table
  (`.docx-table`, `contenteditable=false`), preserved as passthrough.
- **Inline passthrough**: a run is "modeled" only if its children are in
  `MODELED_RUN_CHILDREN`; otherwise (`w:sym` symbols, `w:fldChar` fields like
  page numbers / TOC / cross-refs, `w:footnoteReference` / `w:endnoteReference`)
  it renders read-only and round-trips the original XML.
- **Block passthrough**: any block that is not a modeled `w:p`/`w:tbl` renders as
  a read-only `.docx-pass-block`.
- **Run formatting**: `Fmt` models `b,i,u,strike,color,highlight,shading,
  sizeHalfPt,font`. So strikethrough *is* modeled on read+write, but there is no
  toolbar button. Subscript/superscript (`w:vertAlign`) is **not** in `Fmt`, so
  it is dropped when an edited run is regenerated.
- **Paragraph formatting**: alignment is modeled; indentation (`w:ind`), line
  spacing and space-before/after (`w:spacing`) have no control and are normalized
  away when a paragraph is regenerated.
- **Lists**: ordered/unordered are best-effort; nesting levels and
  restart/continue numbering are not modeled.

## Workstreams

Effort S/M/L, risk low/med/high. Ordered later under "Sequencing".

### W1. More page sizes (and persist the size)
Add the common Office sizes and make a size change actually save.
- Expose the **full** set. Extend `PageSizeName` and `PAGE_SIZES` with: A3, A4,
  A5, B4 (JIS), B5 (JIS), Letter, Legal, Tabloid, Executive. Dimensions
  (portrait, px @96dpi): A3 1123x1587, A5 559x794, B4 971x1374, B5 687x971,
  Legal 816x1344, Tabloid 1056x1632, Executive 696x1008 (A4 794x1123 and
  Letter 816x1056 exist).
- **UI (settled)**: New-popup only, like orientation. **No in-editor size
  picker.** A page-size choice in Omnitext's New docx/odt popup makes the blank
  template generate the chosen size; an opened file keeps whatever size it
  declares (already rendered correctly from its geometry).
- **No richdoc write path needed**: because size is set only at creation, the
  blank template carries `w:pgSz` / `fo:page-width/height`, and richdoc preserves
  it on save (the write path rewrites margins only; the section properties /
  page-layout pass through). So no `w:pgSz` write in richdoc.
- **Omnitext**: extend the `Paper` map in blank-templates.ts (docx twips + odt
  cm/in) and the New-dialog size dropdown to the full list. The richdoc
  `PageSizeName`/`PAGE_SIZES` only need extending if the `defaultPageSize`
  fallback (for files with no geometry) should also offer the new sizes; optional.
- Effort S, risk low (Omnitext-only).

### W2. Orientation (portrait/landscape)
- **UI (settled)**: New-popup only. No in-editor orientation control. A
  portrait/landscape choice in Omnitext's New docx/odt popup makes the blank
  template generate the chosen orientation.
- **Render**: already correct (w/h are explicit, so a landscape doc draws right).
- **Persist**: an opened landscape document keeps its orientation on save for
  free, because the write path preserves the existing `w:sectPr` / page-layout as
  passthrough. So richdoc needs **no** orientation write of its own; the only
  interaction is that W1's size write must preserve the current orientation.
- **Omnitext blank templates**: generate `w:pgSz @w:orient="landscape"` + swapped
  w/h (docx) and `style:print-orientation="landscape"` + swapped page-width/height
  (odt) when landscape is chosen.
- Effort S, risk low.
- Builds directly on W1's size-write code; do them together.
- Effort S, risk low.

### W3. Vertical writing (Japanese tategaki) - the hard one
Render and round-trip top-to-bottom, right-to-left text.
- **Read direction**: docx `w:sectPr/w:textDirection @w:val` (East-Asian vertical
  is `tbRl`/`tbRlV`); odt `style:writing-mode` on the page-layout (`tb-rl`).
  Confirm exact tokens against a real Word/LibreOffice vertical fixture before
  coding (note: `w:textDirection` also appears at table-cell and paragraph level;
  v1 targets the section/page level).
- **Render**: set CSS `writing-mode: vertical-rl` on the page body; the inline
  axis becomes vertical and columns advance right-to-left.
- **Pagination**: the fill axis flips. `paginate()` is axis-agnostic (it takes
  block "heights" and a page "content height"), so feed it the *inline-size* of
  each block and the page's content *width*; pages then advance along x (right to
  left). The page-view, spacers, page cards and the rulers all need an
  axis-aware variant. This is the bulk of the risk.
- **Write**: emit the section text direction (docx) / writing-mode (odt).
- **Scope (settled)**: ship full edit support in one go (read, render, edit,
  write), not a display-only first pass. Gate behind a capability flag
  (`verticalText`) so a format/UI can hide it, and add a vertical-writing toggle
  to Omnitext's New docx/odt popup (alongside the size/orientation option) so a
  new document can start vertical.
- Effort L, risk high. Own milestone, sequenced last.

### W4. Editable tables - the biggest editing gap
Promote `w:tbl` from read-only passthrough to an editable model.
- **Read**: render `w:tbl` to a real editable `<table>` (cells contenteditable),
  carrying enough structure (grid, cell spans, cell/row props) to rebuild OOXML.
- **Edit (settled scope)**: cell text editing for free (contenteditable); add UI
  for insert/delete row and column **and merge/split cells** (cell spans map to
  `w:gridSpan`/`w:vMerge` in docx and `table:number-columns-spanned`/
  `table:covered-table-cell` in odt).
- **Write**: serialize the edited table back to `w:tbl`/`w:tr`/`w:tc` (docx) and
  `table:table` (odt), preserving untouched cell content as passthrough where we
  do not model the cell internals.
- Decide the fidelity line: model grid + cell paragraphs; keep complex cell
  content (nested tables, shapes) as passthrough inside the cell.
- Effort L, risk med-high (round-trip fidelity).

### W5. Run formatting pass - quick polish
- Add a **strikethrough** toolbar button (the `Fmt.strike` path already exists,
  read and write).
- Model **subscript/superscript**: add `vertAlign` to `Fmt`, read `w:vertAlign`
  (`superscript`/`subscript`) and odt `style:text-position`, render as
  `<sup>`/`<sub>`, write back; add two toolbar buttons.
- Effort S, risk low.

### W6. Paragraph formatting
- **Indentation**: increase/decrease-indent buttons; read/write `w:ind`
  (docx) and `fo:margin-left`/`fo:text-indent` (odt).
- **Line spacing**: a small picker (1.0/1.15/1.5/2.0); read/write `w:spacing`
  (`@w:line`/`@w:lineRule`) and odt `fo:line-height`.
- Effort M, risk low.

### W7. List fidelity (optional/later)
- Nesting levels and restart/continue numbering.
- Effort M, risk med.

## Suggested sequencing

1. **W5** (run formatting) - small, visible polish, no architecture risk.
2. **W1 + W2** (page sizes + orientation) - one effort; adds the missing `w:pgSz`
   / `fo:page-*` write path that both need.
3. **W6** (paragraph indent + line spacing).
4. **W4** (editable tables) - the big editing win; own milestone.
5. **W3** (vertical writing) - highest risk, own milestone, last.

W7 slots in whenever lists become a priority.

## Per-format mechanics (read / write reference)

| Capability | docx (OOXML) | odt (ODF) |
|---|---|---|
| Page size | `w:sectPr/w:pgSz @w @h` (twips) | `style:page-layout-properties @fo:page-width @fo:page-height` |
| Orientation | `w:pgSz @w:orient` + swapped w/h | `@style:print-orientation` + swapped w/h |
| Vertical text | `w:sectPr/w:textDirection @w:val="tbRl"` | `@style:writing-mode="tb-rl"` |
| Table | `w:tbl / w:tr / w:tc` | `table:table / table:table-row / table:table-cell` |
| Strikethrough | `w:rPr/w:strike` (modeled) | `style:text-properties @style:text-line-through-style` |
| Sub/superscript | `w:rPr/w:vertAlign @w:val` | `@style:text-position` |
| Indent | `w:pPr/w:ind` | `@fo:margin-left` / `@fo:text-indent` |
| Line spacing | `w:pPr/w:spacing @w:line @w:lineRule` | `@fo:line-height` |

## Tests

- A golden round-trip fixture per capability: parse then serialize an unedited
  file with that feature and assert the relevant part is byte/structure-stable.
- Unit: size/orientation twips<->px and cm<->px conversions from a `sectPr` /
  `page-layout` fixture.
- Unit: `paginate()` already DOM-free; for W3 add an axis-flipped case
  (inline-size fill, right-to-left page advance) with synthetic block sizes.
- Mount (jsdom, injectable measurement): table edit inserts/removes a row and the
  serialized `w:tbl` reflects it; a size change writes `w:pgSz`.
- Browser: each feature verified on a real file (Word/LibreOffice), since jsdom
  cannot measure layout (rulers, pagination, vertical flow).

## Risks and open questions

- **Vertical writing is the thorny one**: it reshapes pagination, rulers and zoom
  into an axis-aware form, and exact Word parity for mixed horizontal/vertical
  runs is out of scope for v1 (section-level vertical only). Confirm the OOXML
  `w:textDirection` tokens/placement against a real vertical Word doc first.
- **Table round-trip fidelity**: keep complex/unmodeled cell content as
  passthrough so we never corrupt a table we cannot fully model.
- **Paragraph regen normalizes direct formatting**: W5/W6 reduce but do not
  eliminate this; the honest-scope note stays.
- The write path currently persists only margins; W1 is the prerequisite that
  makes "geometry" mean size+orientation, not just margins.

## Decisions (settled)

1. **Page sizes**: expose the full list (A3, A4, A5, B4, B5, Letter, Legal,
   Tabloid, Executive).
2. **Size + orientation UI**: New-popup only (Omnitext New docx/odt dialog). No
   in-editor controls; an opened file keeps its declared size/orientation.
3. **Table v1 scope**: text cells + row/col insert-delete **and** cell
   merge/split.
4. **Vertical writing**: ship edit support directly (no display-only first pass);
   it also adds a vertical-writing option to the New popup.
