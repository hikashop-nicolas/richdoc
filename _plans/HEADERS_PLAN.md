# First-page and even/odd headers & footers

## Where we are today

- Only the **default** header/footer is read, rendered, edited and round-tripped. docx `refRid()` /
  `refTarget()` explicitly pick `w:type="default"` and drop `first` / `even`; `w:titlePg` (sectPr)
  and `w:evenAndOddHeaders` (settings.xml) are never read. odt `collectMasterBands()` reads
  `style:header` / `style:footer` only, never `style:header-left` / `style:footer-left`, and the
  first-page dual-master pattern is not modelled.
- The variant parts/elements are *preserved as opaque files* (passthrough), so they survive a save;
  they are just not selectable, editable, or page-aware.
- Per page, each paginate path clones the single `header` / `footer` source (mkClone) with no
  page-number awareness. A per-section "link to previous" chip already exists (mkLinkChip).

## Scope (v1)

Document-level variants of the **main** header/footer (the trailing-section default):

- **Different first page** (`w:titlePg` / odt first-page master): page 1 shows its own header/footer.
- **Different odd & even pages** (`w:evenAndOddHeaders` / odt `*-left`): even pages show their own.

Out of scope for v1 (documented, still round-trips): per-section first/even variants (rare; the
per-section default + link chip stays as-is), and first/even in the pageless view (no pages, so it
shows the default).

## Model (src/core/types.ts)

Extend RichDoc with the two flags and the variant bands, keeping `header`/`footer` as the default:

```
titlePage?: boolean;   // different first page
evenOdd?: boolean;     // different odd & even pages
headerFirst?: { html: string; path?: string };
headerEven?:  { html: string; path?: string };
footerFirst?: { html: string; path?: string };
footerEven?:  { html: string; path?: string };
```

(`path` is the adapter write-back key, like headerPath. Default stays in header/headerPath.)

## Engine (src/core/editor.ts)

- Build up to three editable source bands per role (default / first / even), each with the existing
  hf-clone focus/input/save wiring; only the bands the flags enable exist.
- A `pickHeader(pageIndex)` / `pickFooter(pageIndex)` helper: `titlePage && page 0` -> first;
  else `evenOdd && (page index is even-numbered page, i.e. 1-based even)` -> even; else default.
- In every paginate path's per-page loop (single / columns / vertical / vertical-columns), clone the
  picked band instead of the fixed one. setCloneFields stays. Editing a clone writes back to its
  own source band (the clone already knows its source via the existing wiring; extend it to carry
  which variant it is).
- Page-number fields keep working per page (setCloneFields already takes the page number).

## UI (page setup dialog, src/core/feature/page-view.ts)

Two checkboxes in the Page setup dialog: "Different first page" and "Different odd & even pages".
Toggling on creates the variant band(s) pre-filled from the default (Word's behaviour); toggling off
drops them (back to default everywhere). Writes the flags into the geometry/model + reflow.

## docx round-trip (src/adapters/docx)

- read: collect all three ref types from the trailing sectPr (extend refRid/refTarget to take a
  type); read `w:titlePg`; read `w:evenAndOddHeaders` from word/settings.xml (parse it as a part).
- write: mint a part per present variant with the right `w:type`; write the matching
  `w:headerReference`/`w:footerReference`; set `w:titlePg` in the sectPr when titlePage; set/clear
  `w:evenAndOddHeaders` in settings.xml (mint settings.xml + its content-type/rel if absent).

## odt round-trip (src/adapters/odt)

- even/odd: read/write `style:header-left` / `style:footer-left` in the master page, and set the
  page-layout's `style:page-usage` so the left variant is used on even pages. The default
  `style:header` serves odd/right pages.
- first page: ODF 1.3 has `style:header-first` / `style:footer-first` as master-page children,
  exactly parallel to `style:header-left` (no separate next-style-name master needed). Read/write
  them the same way as even/odd. DONE. (Edge: an enabled-but-empty first-page header is not
  representable, ODF keys the variant off the element's presence.)

## Tests

- docx: a doc with titlePg + first/even parts reads three editable bands; toggling the flags on a
  plain doc mints the parts + sets titlePg / evenAndOddHeaders; round-trip keeps all three.
- odt: header-left round-trips; page-usage set for even/odd.
- Render: page 1 shows first, an even page shows even, others show default (browser-verified, since
  the per-page pick is layout-time).
