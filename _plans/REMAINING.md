# richdoc - remaining gaps

The shipped feature set is in the [README](../README.md#features). This file tracks
only what is **not** done yet.

**The passthrough guarantee.** Any element the reader does not model is preserved
byte-for-byte: it is stashed in `data-docx-xml` / `data-odt-xml` (block / run / table /
image level) and re-emitted on save. So an unmodelled feature is *preserved untouched*,
not lost. The only things actually lost are in "Lost when you edit" below, where the
surrounding context (a paragraph, or the document body) is regenerated from the edited
HTML.

---

## Lost when you edit

These do not round-trip once the surrounding context is edited, because that context is
regenerated from the HTML.

- **Wrapped-image offset rendering.** A wrapped image's exact offset (docx `posOffset`, odt
  `svg:x` / `svg:y`) now round-trips per axis instead of snapping to alignment on save, but the
  editor still *renders* square / tight wrap by float at the nearest alignment rather than at the
  precise offset, because arbitrary-offset text wrapping is not expressible in CSS. Choosing an
  alignment in the image toolbar switches such an image to the alignment model. Behind / front
  images are positioned exactly (CSS-positioned and draggable).

---

## Preserved losslessly but not yet editable

These round-trip untouched today. Adding an insert/edit UI would make them authorable;
they are realistic to do, just not yet built.

- Legacy symbol-font glyphs (`w:sym`): fully rendered and round-tripped verbatim. Symbol, Webdings
  and Wingdings 2/3 map to Unicode and Wingdings renders via the bundled open MaterialDings font, so
  none needs the proprietary font installed; any other (rare, custom) symbol font falls back to the
  named font and displays only where it is installed. The special-character picker deliberately
  *inserts* plain Unicode (the modern equivalent) rather than emitting legacy `w:sym`, so there is
  nothing left to do here; this entry just records that authoring-as-`w:sym` is intentionally absent.
- Complex fields - PAGE / NUMPAGES / TOC, the cross-reference / caption fields
  (REF / PAGEREF / SEQ), and the date / time / author / file-name info fields are authored;
  other, less-common fields (document title / subject, ASK / input, ...) are preserved but
  not insertable.
- Page borders, page-number restart / format, line numbering and page vertical alignment are
  all authorable through Page setup. A few format asymmetries: the page-number restart "start
  at N", the line-number "start at" / "restart each section", and page vertical alignment are
  docx-only (odt has no home for them); line numbers render in-editor only in the single-column
  horizontal layout, and page vertical alignment is previewed only for a single-page document
  (centre / bottom) - other layouts round-trip the setting but do not render it.

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

## Notes

- The odt adapter mirrors docx throughout; where a remaining gap differs by format it is
  noted above.
- Single-section documents keep their full section properties (columns, borders,
  page-number restart, etc.); per-section authoring is implemented, and untouched sections
  still round-trip byte-for-byte. Editing a section's geometry preserves a custom column
  layout (unequal widths, separator line) when the column count is unchanged; only changing
  the count rebuilds equal-width columns.
