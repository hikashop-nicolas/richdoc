# True per-page column balancing

Status: DONE (both phases shipped). Columns read, render with per-page balancing,
preserve the caret on edit, strip wrappers on save, and round-trip. Remaining: a
control to author the column count (tracked in REMAINING.md C2).

Goal: render a multi-column section as real paginated pages, each with N balanced
columns that reset per page (like Word), instead of one continuous columned sheet.

## Why the continuous fallback exists

The paginator (`core/page.ts` + `repaginate` in `core/editor.ts`) lays the body out
as ONE single-column element and inserts inert spacer divs to push blocks to each
page's content top; page cards are a separate layer behind. CSS multi-column on that
one element makes columns span the whole document (col 1 = entire doc, then col 2),
not per page. So columns were rendered continuously with pagination turned off.

## The idea: one multi-column box per page

Fragment the flow into one multi-column wrapper per page and let the browser balance
each wrapper's columns:

```html
<div class="docxedit-doc">                  <!-- the single contentEditable flow -->
  <div class="docx-colpage" style="height:CONTENT_H">  …page 1 blocks… </div>
  <div class="docx-colpage" style="height:CONTENT_H; margin-top:INTER_PAGE"> …page 2… </div>
</div>
```

`column-fill: balance` does the actual balancing for free. The engine only decides
*which blocks belong on which page* — turning "reimplement Word's column layout" into
"bucket blocks into pages".

## Bucketing by the browser's own column flow

No separate width measurement: let the browser tell us where a page is full.

1. Collect the body's top-level blocks (unwrap any existing `.docx-colpage`).
2. Create a wrapper: `column-count: N; column-fill: auto; height: CONTENT_H;
   width: CONTENT_W; overflow: hidden`. With `auto` + a fixed height, content fills
   column 1 to the page height, then column 2, …; surplus content overflows into an
   (N+1)th column to the right, which grows `scrollWidth`.
3. Move blocks in one at a time. After each, if `scrollWidth > clientWidth + EPS`
   (content spilled past N columns) and the wrapper holds more than one block, move
   that block back out, finalize the wrapper, and start a new one with it.
4. A single block taller than N columns stays put (oversized; clipped) and the next
   block starts a new page — mirrors the current oversized-block handling.
5. Finalize each wrapper by switching it to `column-fill: balance` (it already fit in
   ≤ N columns ≤ page height, so balance only equalises them).

Manual page breaks and `column-span: all` figures/tables close the current wrapper.

## Page cards + geometry

One card per wrapper. With `doc` padding = content insets and each wrapper height =
CONTENT_H and non-first `margin-top = bottomInset + PAGE_GAP + topInset`, wrapper p
sits at `contentTop + p·pageStep`, aligned to card p's content area (same math as the
single-column path). `page.minHeight = cardCount·pageStep − PAGE_GAP`.

## Editing (caret)

Reflow reparents blocks (into wrappers), unlike the spacer approach which never moves
nodes. Because the SAME block nodes are moved (not cloned), node identity survives:

- Reflow stays debounced and suspended while a band is being edited (`editingBand`).
- Save `getSelection().getRangeAt(0)` (clone) before reflow; reparent via `appendChild`
  (moves nodes, keeps them attached); re-add the saved range after. The range's nodes
  are still valid because they were moved, not destroyed.
- `cleanBody()` strips `.docx-colpage` wrappers (as it strips spacers today), so the
  saved/round-tripped model is unchanged and `w:cols` still round-trips.

## Phasing

1. `repaginateColumns(n)`: unwrap → bucket via overflow detection → wrappers + page
   cards + header/footer clones. CSS for `.docx-colpage`. Re-enable pagination for
   columned docs (drop the continuous fallback). Verify in the browser.
2. Caret save/restore across the reflow; settle/oversized edge cases; tests
   (bucketing splits across pages; cleanBody strips wrappers; round-trip intact).

## Out of scope (for now)

- Vertical (tategaki) + columns (columns assume horizontal pages).
- Authoring the column count (separate control; columns now read/render/round-trip).
- Unequal-width columns (only equal-width `w:cols num` / `fo:column-count`).
