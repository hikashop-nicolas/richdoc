# Footnotes / endnotes

Status: in progress. Today they are passthrough (refs + footnotes.xml preserved, not rendered
or editable). Goal: render, edit, insert, delete, round-trip , both formats.

## Model

The engine is format-agnostic. The readers normalise to:

- An inline **reference** in the body: `<sup class="docx-fnref" data-fn-id="ID" data-fn-kind="footnote|endnote" contenteditable="false">N</sup>`. The number `N` is assigned by the engine in document order (refs are renumbered on every reflow), so it is display-only; `ID` is the stable key.
- A **note store** handed to the engine: `RichDoc.notes?: { id, kind, html }[]`. The engine keeps the editable note bodies and writes them back through the adapter.

docx: refs are `w:footnoteReference` / `w:endnoteReference` (a run, currently inlinePassthrough); bodies live in `word/footnotes.xml` / `endnotes.xml` keyed by id (skip the id<=0 separator / continuationSeparator notes). odt: a `text:note` is inline (citation + note-body); the reader splits it into a ref + a note-store entry.

## Rendering

Footnotes render at the bottom of the page that holds their reference; endnotes in one block at the end of the document. The paginated single-section path reserves per-page footnote space (paginate() gains an optional per-page content-height; the engine iterates twice: paginate, assign refs to pages, measure each page's notes, re-paginate with the reserve). Other layout modes (columns / vertical / sections) fall back to a notes region at the document end for now.

A note body is an editable band (like header/footer) cloned into the page's footnote area; editing syncs to the note store; the area is in an overlay (never saved into the body).

## Authoring

- Insert footnote/endnote at the caret: mint an id, insert the ref, add an empty note, focus its body. Toolbar button + shortcut.
- Delete: removing the ref drops the note.
- Renumbering is automatic (display order).

## Save

- docx: rebuild footnotes.xml / endnotes.xml from the note store (preserving the separator notes); refs -> w:footnoteReference / w:endnoteReference; register the parts + content-types + rels + settings if newly created.
- odt: each ref + its note-store body -> an inline `text:note` (note-class footnote/endnote) at the ref position.

## Phasing

1. Read + render references (numbered) + the note bodies at the document end, editable, round-trip. Both formats. (foundation)
2. Per-page-bottom placement for the single-section paginated view (the paginate() reserve).
3. Insert / delete authoring + shortcut; endnote variant.
