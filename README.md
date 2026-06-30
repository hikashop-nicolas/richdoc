# richdoc

One client-side, framework-agnostic rich-document editor for **.docx** and
**.odt**, in the browser. A single shared engine (paginated page view, toolbar,
contenteditable, comments panel, track changes, images, header/footer, margin
rulers, zoom, passthrough) is driven by a small per-format adapter that does the
irreducible work: parse bytes into an editable model and serialize the model
back, preserving everything it does not model.

No server, no upload: the file is read, edited and rebuilt entirely on the
user's machine.

## Features

A full word-processor surface in the browser, the same for both formats unless noted.

- **Text & paragraphs:** headings (H1-H3); nested ordered / unordered lists whose
  numbering restarts, continues or starts at N; bold, italic, underline, strikethrough,
  super/subscript; text colour, highlight, shading; font family and size; paragraph
  alignment, indent, line spacing, space before/after, paragraph shading (background fill), and
  paragraph borders (box / edge presets).
- **Named styles:** apply, author and edit paragraph and character styles (including shading, borders
  and tab stops), preserving each style's other properties and its inheritance.
- **Tables:** editable cells, cell borders (colour / style / width), column / row /
  indent resize, merge, and in-cell formatting.
- **Page layout:** page size, margins, orientation, a page border (style, width, colour),
  page numbering (format, plus restart "start at N" on docx) and line numbering (interval +
  restart) in a paginated page view, with horizontal + vertical margin rulers that follow the
  caret's page, zoom, and a page setup dialog.
- **Sections:** mid-document section breaks with mixed per-section page size,
  orientation, margins and columns; insert-section-break; per-section page setup; and
  distinct per-section headers/footers with a link-to-previous toggle.
- **Columns:** multi-column sections with true per-page balanced columns (reset per
  page, like Word).
- **Headers & footers:** default plus first-page and even/odd variants, per section,
  with live page-number, page-count and table-of-contents fields.
- **Fields:** insert page number / count / "X of Y", a table of contents, and date / time /
  author / file-name fields (the page fields recompute on reflow; date/author/etc. are snapshots).
  All round-trip (docx `w:fldSimple`, odt `text:date` / `text:author-name` / ...).
- **Vertical writing:** vertical (tategaki) and RTL layout, paginated, including
  multi-column vertical text (stacked bands) and a vertical ruler.
- **Images:** inline and floating / anchored with text wrap (square, tight,
  top-and-bottom, behind, in front), alignment, alt text and wrap padding, driven by an
  on-select toolbar. A wrapped image's exact offset round-trips per axis.
- **Furigana / ruby**, rendered natively (reading above in horizontal text, to the
  right in vertical).
- **Tabs & tab stops:** real tab characters plus custom tab stops with left / center /
  right / decimal alignment and dot leaders, authored on the ruler and aligned in both
  horizontal and vertical (tategaki) text.
- **Comments** (with replies, reactions and resolve) and **track changes**
  (insert / delete / formatting) with a suggesting mode.
- **Bookmarks, cross-references and captions:** reference a bookmark, heading or
  caption by text, page or direction (recomputed on reflow); figure / table / equation
  captions with auto-numbered sequences; internal links; and a collapsible outline pane.
- **Equations / math:** author in LaTeX with a live preview, displayed as native 2D
  MathML and round-tripped (docx OMML, odt formula objects); imported equations stay
  editable via best-effort LaTeX recovery.
- **Footnotes & endnotes:** insert, edit and delete, placed per page (covering
  multi-column, vertical and per-section layouts).
- **Special characters:** a picker inserts common symbols, arrows, Greek letters, currency and
  accented Latin as plain text (so they round-trip on both formats). Legacy Word symbol-font glyphs
  (`w:sym`) are rendered too (Symbol, Webdings and Wingdings 2/3 mapped to Unicode, Wingdings via a
  bundled open font) and preserved.
- **Find & replace** with case, whole-word and regex options.
- **Passthrough preservation:** anything not modelled is kept byte-for-byte and
  re-emitted on save, so editing never drops unsupported content.

What is *not* yet done is tracked in [`_plans/REMAINING.md`](_plans/REMAINING.md).

## Usage

```js
import { createEditor } from "richdoc";

// Sniffs the format (.docx vs .odt) and mounts the right adapter on the engine.
const editor = createEditor(container, bytes, { author: "Jane" });

const out = await editor.getBytes(); // edited file, or the original bytes if untouched
editor.isDirty();
editor.destroy();
```

The package root also exports `createDocxEditor`, `createOdtEditor`,
`sniffFormat(bytes)` and `setLocale(...)`. Each adapter module additionally
exports its low-level parse/serialize functions (`docxToParts` / `htmlToDocx`,
`odtToHtml` / `htmlToOdt`).

## Architecture

The engine knows nothing about OOXML or ODF; an adapter knows nothing about the
UI. An adapter implements three seams plus a capability set: `read()`, `write()`,
`newCommentMarkers()` and `capabilities`. The engine is split by feature so the
orchestrator stays small.

```
src/core/
  editor.ts          createRichEditor(container, adapter, options): the orchestrator
                       (DOM, the pagination render loop, lifecycle, feature wiring)
  feature/
    toolbar.ts         formatting controls + the overflow "more" row
    comments.ts        comments side panel + comment-edit bookkeeping
    track-changes.ts   suggestion mode (ins/del marks, accept/reject)
    images.ts          image select / resize / delete / insert
    page-view.ts       margin rulers + zoom + the centred canvas
  page.ts            page geometry + the pure paginate()
  types.ts           the Adapter / RichDoc / CommentEdits / Capabilities contract
  util.ts, i18n.ts   shared helpers + the multilingual dictionary
src/adapters/
  docx/              OOXML, split into shared / read / write / index
  odt/               OpenDocument, split into shared / read / write / index
src/index.ts         createEditor(): format sniff + adapter selection
```

## Capabilities per format

The toolbar only offers what a format's serializer can persist; the rest is
hidden (existing such content is still preserved on save as passthrough).

| Feature | .docx | .odt |
|---|---|---|
| Bold / italic / underline, headings, lists, links | yes | yes |
| Text/background colour, font family/size, alignment | yes | yes |
| Images (insert, resize, delete) | yes | yes |
| Header / footer (edit, and create by double-clicking the margin) | yes | yes |
| Track changes (suggesting, accept/reject) | yes | yes |
| Comments | panel, replies, reactions, resolve | panel + resolve |
| Manual page breaks | yes | no |
| Passthrough preservation of unmodelled content | yes | yes |

A capability flag gates its toolbar control, so a feature appears for a format
only once that format's `write()` path can persist it.

## Develop

```bash
npm install
npm test          # vitest: round-trip + mount tests
npm run typecheck
npm run build     # tsc -> dist + copy the shared CSS
node_modules/.bin/vite   # then open /demo/ to try it on a real file
```

## License

MIT, see [LICENSE](LICENSE).

The bundled MaterialDings webfont (used to render Wingdings `w:sym` glyphs without the proprietary
font) is © 2018 Accusoft Corporation under the SIL Open Font License 1.1; see
[src/adapters/docx/MaterialDings.LICENSE.md](src/adapters/docx/MaterialDings.LICENSE.md).
