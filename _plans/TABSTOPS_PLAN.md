# Tab stops: positioning + alignment + ruler authoring (C2)

## Where we are today

- A paragraph's custom tab stops are already **fully parsed** and round-trip:
  read into `data-rdoc-tabstops` on the block as JSON `[{pos:px, val:left|center|right|decimal, leader?:"dot"}]`
  (docx `w:tabs`/`w:tab`, odt `style:tab-stops`), written back from the same JSON.
- A tab character is an atomic span `.docx-tab` (`data-docx-tab="1"`, contenteditable=false, text `\t`).
- Rendering uses CSS `tab-size: 48px` only: tabs land on the **default 0.5in grid**, the stored
  positions are ignored, and right/center/decimal alignment renders as left.
- The ruler (`page-view.ts`) has a px-per-page coord system (`CM`, zoom `z`), reads the active
  page/section geometry, tracks the caret (`activePage`/`syncRulers`), and drags margin handles via
  `bindDrag` (snap-to-0.5cm magnet, Alt to bypass). This is the pattern tab markers will mirror.

So both gaps are pure view/authoring; the data model + round-trip need no change.

## Phase 1 - render tabs at their stops, with alignment

A per-paragraph layout pass (runs after reflow, on edit, and on zoom/width change) sets each
`.docx-tab` span's width so the following content aligns to the governing stop. Only paragraphs
with `data-rdoc-tabstops` are touched; everything else keeps the CSS `tab-size` default.

Per paragraph, left-to-right, for each tab span:
1. Read the tab's live left x relative to the paragraph content box (`getBoundingClientRect`), and
   its line (so wrapped paragraphs work: the governing x is the tab's x **within its own line**).
2. Pick the governing stop = first stop with `pos > x`. If none, fall back to the next default-grid
   multiple (Word's behaviour past the last stop).
3. Measure the segment after the tab (a `Range` from just after this tab to just before the next tab
   / line end) to get its width, and for `decimal` the x of the first `.`/locale separator.
4. Compute the tab width:
   - left: `stop - x`
   - right: `stop - segWidth - x`
   - center: `stop - segWidth/2 - x`
   - decimal: `stop - decimalOffset - x`
   Clamp to >= a 1px minimum; if the content would overflow the stop, advance to the next stop
   (Word's rule), so a tab never moves backwards.
5. Set the span: `display:inline-block; width:Npx`. For a `leader:"dot"` stop, fill the span with a
   baseline dot leader (a clipped repeating-dots background sized to the width).

Re-runs are cheap (only tabbed paragraphs) and idempotent. Hook points: the end of `reflow`, and
`applyZoom`. Measurement uses the same scaled coordinate space the ruler already deals with.

Edge cases: a tab before the first stop with no governing stop uses the default grid; multiple tabs
on one line each resolve against their own live x; an empty segment (tab at line end) gets width to
the stop. Vertical (tategaki) text: defer (tabs in vertical are rare); keep the default-grid render
there for now and note it.

## Phase 2 - ruler authoring

The horizontal ruler gains tab-stop markers for the **caret paragraph** (read its
`data-rdoc-tabstops`; markers update as the caret moves between paragraphs).

- **Type selector**: a small box at the ruler's left corner that cycles left -> center -> right ->
  decimal (Word-style); it sets the type used for newly added stops and shows the current type glyph.
- **Add**: clicking an empty spot on the ruler strip adds a stop of the selected type at that x.
- **Move**: drag a marker (snap-to-0.5cm magnet, Alt to bypass), like the margin handles.
- **Retype**: click an existing marker cycles its alignment type.
- **Remove**: drag a marker down off the ruler (or double-click it).
- Each marker draws a distinct glyph per type (L / inverted-T / reversed-L / decimal-dot).

Edits apply to all blocks in the selection (the caret's paragraph when collapsed), rebuild each
block's `data-rdoc-tabstops`, then `mark()` + reflow so Phase 1 re-renders. Markers live in the
ruler overlay layer, positioned `pos * z` within the page's content area (left margin origin).

## Out of scope / deferred

- Tab leaders other than dots (underscore/hyphen leaders) - render as dots for now; the stored
  `leader` value still round-trips.
- Per-line stop authoring differences - stops are a paragraph property (correct); only rendering is
  per-line.
- Vertical (tategaki) tab alignment - default-grid render retained, documented.

## Tests

- Render: a paragraph with left/right/center/decimal stops positions the following segment correctly
  (assert the tab span widths / segment left edges via the layout pass on a mounted editor).
- Round-trip unchanged (existing tests stay green).
- Ruler: adding/moving/removing/retyping a stop rewrites `data-rdoc-tabstops` and survives save.
