# Bookmarks + cross-references

## Status

- **Phase 1 DONE** (bookmarks + heading/bookmark cross-refs, both formats, round-trip + engine
  resolution browser-verified).
- **Phase 2 DONE** (captions + figure/table cross-refs, both formats, browser-verified). What shipped:
  - Caption model: a paragraph carrying `data-rdoc-caption="figure|table"` with a `seq` field
    (`<span class="docx-field docx-field-seq" data-field="seq" data-seq="Figure|Table">`). The engine's
    decorateFields renumbers each sequence per type in document order, alongside PAGE/NUMPAGES.
  - Image: the alt prompt became an "Image options" dialog (alt + caption). Table: an
    "Add / edit caption" item in the cell menu (prompt). Both go through caption.ts
    (buildCaption/applyCaption/captionText/captionAfter/topBlock).
  - Cross-ref dialog gained Figure + Table target-type radios; the bookmark-wrap helper was
    generalised from headings to any block (headings AND captions) as `blockBmName`.
  - Round-trip: docx `SEQ <id> \* ARABIC` fldSimple + a "Caption" pStyle; odt `text:sequence`. A
    paragraph is re-tagged as a caption on read from the presence of the seq field.
  - Caption helpers live in `src/core/feature/caption.ts`.
- Design refinement vs the original plan: a heading/caption cross-ref does NOT stamp `data-rdoc-bm`
  on the element (that would not serialize, since neither odf nor ooxml treats a heading/caption as a
  bookmark). Instead it WRAPS the element's content in a real `docx-bookmark` / `docx-bookmark-end`
  anchor pair (`_Ref<n>`), so it round-trips through the existing bookmark write path and the engine
  reads the text from the wrapped range.
- **Deferred item 1 DONE**: complex (fldChar begin/instr/separate/result/end) fields are now modelled
  on docx read, not just fldSimple. A state machine in the inline reader collects the field, then maps
  REF/PAGEREF -> xref and SEQ -> caption seq (cached result becomes the shown text); PAGE/TOC/DATE/etc.
  stay verbatim passthrough. They write back as fldSimple (complex -> simple, Word-equivalent). odt has
  no complex-field concept, so this is docx-only.
- Deferred (still TODO): reference formats beyond text/page (label+number vs caption-text-only,
  above/below, paragraph number); range bookmark fidelity across block boundaries; cross-refs to equations.

## Where we are today

- Bookmarks are **passthrough only**: docx `w:bookmarkStart`/`w:bookmarkEnd` and odt
  `text:bookmark`/`text:bookmark-start`/`-end` fall through to the opaque `docx-pass` span, so
  they survive a save but are invisible and unreferenceable.
- Cross-references (docx `REF`/`PAGEREF` complex fields or `w:hyperlink w:anchor`; odt
  `text:bookmark-ref`) are not modelled at all (passthrough).
- Fields ARE modelled: `<span class="docx-field" data-field="PAGE">`, inserted from the toolbar
  fields menu, written as `w:fldSimple` (docx) / `text:page-number` (odt), and recomputed live by
  `decorateFields()` on every reflow (the TOC field rebuilds its rows from the live headings - the
  closest analog to an auto-updating cross-reference).
- Hyperlinks round-trip as `<a href>`; internal anchor links are not handled yet.

## Scope

Cross-references target four kinds, chosen by a **target-type radio** in the insert dialog that then
lists that kind's targets: **Heading**, **Bookmark**, **Figure** (a captioned image), **Table** (a
captioned table). Built in two phases:

- **Phase 1 - bookmarks + heading/bookmark cross-refs.** Insert a named bookmark at the caret (point)
  or wrapping the selection (range); parse/render existing ones as zero-width markers so they
  round-trip and become jump targets. Cross-ref dialog with the target-type radio (Heading + Bookmark
  populated), format = the target's **text** or its **page number**; clickable computed span,
  recomputed live, round-trips. Heading refs auto-mint a hidden bookmark on the heading.
- **Phase 2 - captions + figure/table cross-refs.** Give images an alt + **caption** popup (extending
  today's alt control) and tables a caption; captions are numbered paragraphs ("Figure 1", "Table 2")
  auto-numbered like a SEQ field. Then the Figure / Table radio options list the captions, and a
  cross-ref to one shows its label+number / caption text / page number.

Deferred but NOT forgotten (TODO in this file): reference **formats** beyond text / page number -
"above/below", paragraph number, "label and number" vs "caption text only" granularity. Also: exact
fidelity of a range bookmark whose span crosses block boundaries; cross-refs to equations.

## Model (HTML)

- Bookmark marker: `<a class="docx-bookmark" data-rdoc-bm="<name>" contenteditable="false"></a>`
  (zero-width). A range bookmark keeps a start marker (carrying the name) and an end marker
  (`data-rdoc-bm-end="<name>"`); a point bookmark is just the start. The spanned text stays in the
  flow between them, so ranges survive.
- Cross-reference: `<a class="docx-xref" data-rdoc-xref="<name>" data-rdoc-xref-fmt="text|page"
  contenteditable="false">computed</a>`. The text content is recomputed by decorateFields; clicking
  it scrolls to the bookmark / heading.

Headings referenced by a cross-ref get an auto-bookmark (a `_Ref<n>` name minted on demand and
stamped on the heading via `data-rdoc-bm`), so every cross-ref ultimately targets a bookmark name -
matching how Word stores heading references.

## Engine (editor.ts)

- `decorateFields` gains a cross-ref pass: for each `.docx-xref`, find its target (a `[data-rdoc-bm]`
  marker or a heading carrying the name), and set its text to the target's text (format "text") or
  its page number (format "page", computed from `offsetTop / pageStep` like the TOC). Guard with a
  per-element signature like the TOC so it does not loop reflow.
- Click handler (delegated on the body, next to the existing TOC-row handler): a `.docx-xref` scrolls
  its target into view.
- The two inserts: a bookmark marker at the caret (or wrapping the selection), and a cross-ref span
  whose initial text is computed immediately.

## UI (toolbar insert menu)

- **Insert bookmark**: prompt for a name (reusing the prompt pattern), insert the marker.
- **Insert cross-reference**: a dialog (like the note popup) with a **target-type radio**
  (Heading / Bookmark / Figure / Table) that repopulates a target list, plus a format choice
  (text / page number); inserts the xref span. Phase 1 wires Heading + Bookmark; Phase 2 adds
  Figure + Table once captions exist.
- **Image caption** (Phase 2): the image alt control becomes a small popup with alt text + a caption
  field; entering a caption inserts/updates a numbered caption paragraph after the image.
- **Table caption** (Phase 2): a caption control in the table toolbar, inserting/updating a numbered
  caption paragraph by the table.
- Both insert controls live in the existing insert group / fields menu area.

## Captions (Phase 2)

A caption is a paragraph `data-rdoc-caption="figure|table"` holding a label + an auto-number + text,
e.g. `Figure <span class="docx-field" data-field="seq" data-seq="Figure">1</span>: ...`. decorateFields
numbers them per type in document order (like PAGE/TOC). docx: a normal paragraph with a `SEQ Figure`
field (often styled "Caption"); odt: a paragraph with a `text:sequence` (the "Illustration"/"Table"
sequence). A captioned image/table referenced by a cross-ref gets an auto-bookmark on its caption,
same mechanism as headings.

## Round-trip

- docx read: `w:bookmarkStart`(name) -> bookmark marker, `w:bookmarkEnd` -> end marker; a `REF` /
  `PAGEREF` field (fldSimple or fldChar+instrText) -> xref span (fmt page for PAGEREF). A
  `w:hyperlink w:anchor` to a bookmark also maps to a jump.
- docx write: markers -> `w:bookmarkStart`/`w:bookmarkEnd` (minting ids); xref span -> a
  `w:fldSimple w:instr=" REF name \h "` / `" PAGEREF name \h "` holding the computed text.
- odt read: `text:bookmark`/`-start`/`-end` -> markers; `text:bookmark-ref`
  (text:reference-format text|page, text:ref-name) -> xref span.
- odt write: markers -> `text:bookmark-start`/`-end` (or a point `text:bookmark`); xref span ->
  `text:bookmark-ref` with the matching reference-format + ref-name.

## Tests

- Bookmark round-trips (point + range) in both formats.
- Cross-ref round-trips (text + page) in both formats; references the right bookmark.
- A heading cross-ref mints a bookmark on the heading and round-trips.
- decorateFields recomputes the xref text/page (engine mount test; page is layout-dependent so the
  text-format recompute is the unit-testable part).
