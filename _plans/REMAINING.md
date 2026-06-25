# richdoc - remaining gaps

The DOCX_CAPABILITIES_PLAN workstreams W1-W7 are all shipped (page sizes,
orientation, vertical/RTL writing with paginated tategaki + header/footer + rulers,
editable tables with borders/resize/merge/indent, run formatting, paragraph indent +
line spacing, paragraph space before/after, and nested ordered/unordered lists), plus extras the plan did not list
(cell border colour/style/width, table column/row/indent resize with graduated rulers
+ magnet, and fields: page number / count / table of contents). What is left:

Named paragraph and character styles are fully editable: the block dropdown and a
character-style dropdown list the document's styles, apply them (w:pStyle / w:rStyle,
odt text:style-name), render each via injected CSS, and round-trip. A style settings
dialog (name, bold/italic/underline/strike, alignment, size, text colour, background
colour, font) authors new styles and edits an existing style's definition in place
(updating every paragraph/run that uses it); both are persisted to the stylesheet on
save (add-or-update by id).

## Modelled-but-incomplete
- **List restart/continue numbering**: nesting levels and ordered/bullet kind now
  round-trip (w:numPr/w:ilvl/w:numId, odt per-level list styles), and the indent /
  outdent buttons create real nesting. Restart-at-N and continue-previous-list are
  still not modelled; every ordered list restarts at 1.
- **Tab stops**: custom tab stops (`w:tabs`) are dropped on paragraph regen.
- **Style editing depth**: the style dialog covers the common run/paragraph properties
  (align, indent + spacing via passthrough, weight/italic/underline/strike, colour,
  background, size, font). basedOn inheritance is flattened when a style is edited, and
  tab stops, borders and the long tail of style properties are not exposed.

## Preserved-but-not-editable (round-trip via passthrough; no insert/edit UI)
These survive a save untouched but cannot be created or modified in the editor:
- Footnotes / endnotes (`w:footnoteReference` / `w:endnoteReference`).
- Symbols (`w:sym`).
- Equations / math (OMML `m:oMath`, odt formula objects).
- Content controls / structured document tags (`w:sdt`).
- Multi-column section layouts (`w:cols`) and multiple sections.
- Text boxes / shapes / charts / SmartArt (drawing objects beyond inline images).
- Bookmarks and cross-references as an editable concept (insert UI); existing ones
  pass through. Inserted fields cover PAGE / NUMPAGES / TOC only.

## Notes
- The passthrough guarantee holds for all of the above: untouched content round-trips
  byte-for-byte; only regenerated paragraphs lose the unmodelled direct properties.
- None of these are blocking for the common word-processing use cases now covered.
