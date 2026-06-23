# Pagination: pages as decoration over one continuous editor

Goal: make the editor show real, reflowing pages (page cards, gaps, repeated
header/footer, page numbers) while keeping a single continuous contenteditable
column underneath. Editing stays simple and robust; pages are a computed visual
layer that reflows as content changes.

This lives entirely in the shared engine (src/core/editor.ts plus the CSS), so both
docx and odt get pages the moment they mount. It is the reason to do this before odt
parity: pagination is engine-level, odt parity is adapter-level.

## Principle: pagination is a view, never persisted

The file stores content plus page geometry (w:pgSz / w:pgMar in OOXML; page-layout in
ODF). It never stores "block N is on page 3"; Word, LibreOffice and Google Docs all
re-paginate on open. So:

- We compute pages for display only. We never write pagination back.
- Getting a break position wrong can never corrupt a file. The feature is purely
  additive and reversible.
- Our breaks approximate; they will not match Word's exact positions (different line
  breaking and hyphenation). That is acceptable for a view and is not a regression.

## Current layout (what we build on)

```
.docxedit-scroll                 scroll container, centers content
  .docxedit-canvas               [ leftSpacer | .docxedit-page | rightArea ]
    .docxedit-page               one white card, width min(816px,100%), AUTO height
      .docxedit-header           optional editable band (once, top)
      .docxedit-doc              the body contenteditable (continuous)
      .docxedit-footer           optional editable band (once, bottom)
    rightArea > .docxedit-comments   comment cards positioned by anchor Y
```

Today there is one card of infinite height. Pagination turns that single card into N
page-card decorations at computed Y offsets, with the body still flowing as one
contenteditable on top.

## Target structure

```
.docxedit-pageview               position:relative wrapper (replaces the single card)
  .docxedit-pagelayer            z-index 0, absolutely positioned page cards:
    .docxedit-pagecard (xN)        each at top = i*(pageH+gap), the white bg + shadow,
                                    header clone (top), footer clone + page number (bottom)
  .docxedit-doc                  z-index 1, the ONE body contenteditable, positioned at
                                    the page content box (left = marginL, width = content),
                                    containing inert spacer divs at each page boundary
```

The body is never split into per-page DOM nodes. The caret, selection and undo keep
working exactly as now because it stays one contenteditable; the inert spacers are
contenteditable=false islands the browser skips for caret.

## The paginator (pure, testable)

Keep the math separate from the DOM so it is unit-testable without layout (jsdom returns
0 for getBoundingClientRect, so we must not depend on a live DOM in tests).

```
// pure function: given block heights and page metrics, return where pages break
paginate(blocks: { height: number }[], page: {
  contentHeight: number;   // page height minus top+bottom margins (and header/footer zones)
  spacerHeight: number;    // footer zone + inter-page gap + next header zone
}): { breakBeforeIndex: number[]; spacerBefore: Map<number, number> }
```

Single greedy pass at BLOCK granularity (v1 never splits a paragraph):
- Walk top-level blocks, accumulating used height within the current page content box.
- When the next block does not fit, record a break before it, reset the running height,
  and remember the spacer to insert before it.
- A block taller than a full page starts its own page and overflows downward (v1
  limitation, logged, not silently hidden).

Because the spacer height is deterministic, one pass suffices; no fixed-point iteration.

The DOM layer does three things: measure (read each top-level block's height via
getBoundingClientRect), call paginate, then apply (insert/update inert spacer divs in
the flow and lay out the page-card layer). Measurement and application are the only
parts that touch the DOM.

## Reflow triggers

Re-paginate (debounced, inside requestAnimationFrame):
- after first render,
- on document.fonts.ready and on each embedded-font load (heights shift when fonts swap),
- on image load inside the body,
- on container resize (ResizeObserver, already present for comment cards),
- on input, incrementally: re-run from the edited block's page onward, not the whole doc.

After every re-pagination, call positionCards() so comment cards follow their moved
anchors (the comment panel already reflows; it just needs to run after).

## Page geometry into the contract

Add an optional page block to RichDoc (src/core/types.ts):

```
page?: {
  widthPx: number; heightPx: number;            // physical page size
  margin: { top: number; right: number; bottom: number; left: number };
}
```

Adapters fill it; the engine falls back to a default when absent. The default is A4
(this tool is for anyone, and A4 is the world default). Omnitext exposes two settings
that flow in through EditorOptions:
- defaultPageSize: "a4" | "letter" (default "a4"), used only when the file has no
  geometry. Omnitext's settings page gets a page-size selector for this.
- paginated: boolean (default true), the paginated-vs-pageless view mode. Omnitext's
  settings page gets a toggle for this. Paginated is the default when geometry exists.
- docx adapter: read w:sectPr/w:pgSz (w:w, w:h, twips) and w:pgMar (top/right/bottom/
  left, twips). 1 twip = 1/1440 inch; px at 96 dpi = twips / 1440 * 96 = twips / 15.
- odt adapter: read the page-layout (fo:page-width, fo:page-height, fo:margin-*) from
  styles; later. odt also has no header/footer parsed yet (parts.header/footer are ""),
  so odt gets body pagination first, header/footer repeat follows with odt parity.

The fixed 816px width and 72px side padding in docxedit.css become driven by this
geometry (CSS variables) so each document uses its real page size and margins.

## Header / footer repetition (decided: click-to-edit clones)

In a paginated view the header/footer repeat on every page, but there is ONE canonical
header and ONE canonical footer in the model. Chosen model (Word-like):

- Click-to-edit clones: the header/footer render as read-only clones inside every page
  card. Clicking a clone makes that one the live editable element (the canonical band);
  the others stay clones; on blur the edit re-syncs to all clones. Only one is editable
  at a time. The clones are non-editable and excluded from save (they carry no data-*
  the serializer reads); only the canonical band feeds adapter.write via parts.

This is the thorniest part (caret edge cases when promoting a clone to editable), so it
is Phase 2, after body pagination is solid.

## Phasing

- Phase 0 (DONE, commit ec721d1): geometry. RichDoc.page; docx adapter fills it from
  pgSz/pgMar; engine drives page width/margins via CSS vars; A4 default fallback. The
  page renders at the document's true size and margins.
- Phase 1 (DONE, commit e8fa4ff): body pagination. core/page.ts paginate() (pure,
  unit-tested); core/editor.ts repaginate()/reflow()/cleanBody(); page-card layer,
  spacers, reflow triggers (fonts/images/resize/debounced edit); EditorOptions.paginated
  (default true) toggles to pageless. Shipped a bit more than planned: read-only repeated
  header/footer clones + page numbers are already in (so paginated view shows the
  letterhead), leaving only their EDITING for Phase 2.
- Phase 2 (DONE, richdoc f5708a6 + omnitext 6d9b6da): header/footer clones are editable
  (option B, click-to-edit via editBand + an editingBand guard); Omnitext settings
  (page-size selector, paginated toggle) wired through EditorOptions. Comment-panel
  reflow across pages verified.
- Phase 3a (DONE, richdoc 1330c91): respect explicit page breaks (forceBreakBefore in
  paginate(); manual w:pageBreakBefore / w:br forces a new page, markers hidden in the
  paginated view) and tighten page-top alignment (the page-starting block's top margin is
  zeroed so it sits flush under the header instead of drifting).
- Phase 3b (later, not v1): mid-paragraph and table splitting, widow/orphan control,
  incremental (not full) re-pagination on edit, and a true paginated PDF/print export.

## Known limitations (remaining, for Phase 3b)

- Block-level breaks only: a paragraph/table taller than a page overflows its card.
- Re-pagination is full (debounced), not incremental; fine for typical docs.
- Assumes the page fits the viewport width (no responsive shrink handling yet).

Fixed since v1: header/footer editing in paginated mode (Phase 2, click-to-edit); the
page-top drift (Phase 3a); explicit page breaks now honored (Phase 3a).

## Tests

- Unit: paginate() is pure; feed synthetic block heights and page metrics, assert break
  indices and spacer heights for the fit/overflow/oversized-block cases. No DOM needed.
- Unit: twips-to-px geometry parsing in the docx adapter from a sectPr fixture.
- Mount (jsdom, with a stubbed measurement returning fixed heights): pageview builds the
  expected number of cards; spacers inserted at the expected positions; pageless toggle
  removes them. jsdom cannot really measure, so the DOM layer takes an injectable
  height-measure function that the test overrides.

## Risks and limitations (set expectations)

- Measurement timing: heights change as fonts and images load; we re-paginate on those
  events. A flash of pre-font layout is possible on first paint.
- Performance on large documents: full re-pagination must stay off the input path; only
  the incremental from-edited-page pass runs on typing, debounced.
- Content taller than a page (big image or table) overflows its card in v1.
- Exact Word parity is out of scope; our pages approximate.
- Zoom (if added later) wraps the pageview; pagination math stays in unscaled px.
- Header/footer per-page editing is the thorniest part; v1 uses display-only clones.

## Decisions (settled)

1. Default page size when the file has no geometry: A4. Omnitext's settings page adds a
   page-size selector (A4 default, Letter option), passed in as EditorOptions.defaultPageSize.
2. Paginated by default (when geometry exists). Omnitext's settings page adds a
   paginated/pageless toggle, passed in as EditorOptions.paginated.
3. Header/footer editing: option B, click-a-clone-to-edit (Word-like). Phase 2.
4. PDF/print export: out of scope for now. This tool views, edits and saves files; it is
   not a converter. Print/PDF export is a plausible future phase, not part of this effort.
