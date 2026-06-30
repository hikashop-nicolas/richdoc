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

- **Vertical + multi-column mid-section.** A whole-document vertical multi-column layout
  works; a *mid-document* section that is BOTH vertical AND multi-column renders as a
  single vertical flow.
- **Style authoring depth.** Paragraph shading, borders and tab stops are authorable both as direct
  formatting and as part of a named style, round-tripping on both formats (docx pPr `w:shd` /
  `w:pBdr` / `w:tabs`, odt `fo:background-color` / `fo:border-*` / `style:tab-stops`); a style's tab
  stops ride a `--rdoc-tabstops` custom property so styled paragraphs render them. The border picker
  offers colour, line style (solid / dashed / dotted / double) and width, with box / top / bottom /
  top-and-bottom side presets; an arbitrary single side (e.g. left only) still round-trips on import
  but is not offered as a preset.
- **Image layout fine detail.** Square / tight wrap uses alignment only (a file's exact
  `posOffset` for a wrapped image is not honored); behind / front offsets map to a
  CSS-positioned element rather than a true layout engine.
- **Vertical (tategaki) tab alignment.** Custom tab-stop alignment renders in horizontal
  text; vertical text keeps the default grid.

---

## Preserved losslessly but not yet editable

These round-trip untouched today. Adding an insert/edit UI would make them authorable;
they are realistic to do, just not yet built.

- Legacy symbol-font glyphs (`w:sym`): rendered and round-tripped verbatim, but not authored as
  `w:sym` from the UI (the special-character picker inserts plain Unicode, the modern equivalent).
  Symbol, Webdings and Wingdings 2/3 map to Unicode and Wingdings renders via the bundled open
  MaterialDings font, so none needs the proprietary font installed; any other (rare, custom) symbol
  font still falls back to the named font and displays only where it is installed.
- Complex fields - PAGE / NUMPAGES / TOC, the cross-reference / caption fields
  (REF / PAGEREF / SEQ), and the date / time / author / file-name info fields are authored;
  other, less-common fields (document title / subject, ASK / input, ...) are preserved but
  not insertable.
- Page borders and page-number restart - preserved on the trailing section, not
  authorable.

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
  still round-trip byte-for-byte.
