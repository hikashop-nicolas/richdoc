import type { PageGeometry, PageSizeName } from "./types";

// Page geometry helpers. All measurements are CSS pixels at 96 dpi (1in = 96px).
// Pagination is a view concern: these sizes drive how a page is drawn, never what is saved.

const MM_PX = 96 / 25.4; // millimetres to px at 96 dpi
const IN_PX = 96; // inches to px

/** Named default page sizes, used when a document declares no geometry. */
export const PAGE_SIZES: Record<PageSizeName, { widthPx: number; heightPx: number }> = {
  a4: { widthPx: Math.round(210 * MM_PX), heightPx: Math.round(297 * MM_PX) }, // 794 x 1123
  letter: { widthPx: Math.round(8.5 * IN_PX), heightPx: Math.round(11 * IN_PX) }, // 816 x 1056
};

/** A default geometry for a named size with 1-inch margins all round. */
export function defaultPageGeometry(size: PageSizeName = "a4"): PageGeometry {
  const s = PAGE_SIZES[size];
  return { widthPx: s.widthPx, heightPx: s.heightPx, margin: { top: 96, right: 96, bottom: 96, left: 96 } };
}

// --- The paginator -----------------------------------------------------------
// Pure, DOM-free, so it unit-tests without layout. Given the measured height of each
// top-level block and the page metrics, it returns where to insert spacer gaps so the
// flow lands at each page's content top, plus how many page cards to draw.

export interface PageMetrics {
  /** Vertical distance between consecutive card tops: page height + the inter-page gap. */
  pageStep: number;
  /** Usable content height per page: page height minus top and bottom margins. */
  contentHeight: number;
  /** Extra space reserved at the bottom of a given page (e.g. for its footnotes), reducing how
      much body content fits there. */
  reserveOf?: (page: number) => number;
}

export interface PaginationResult {
  /** blockIndex -> px height of an inert spacer inserted before that block. */
  spacerBefore: Map<number, number>;
  /** number of page cards to draw. */
  cardCount: number;
  /** the page each block landed on (same length as `heights`). */
  pageOfBlock: number[];
  /** each block's top offset from its page's content top, in px (same length as `heights`). */
  offsetInPage: number[];
}

const EPS = 1; // sub-pixel overflow tolerance

/**
 * Greedy block-level fill: never splits a block. A block taller than a page overflows
 * its card and the next block resumes at the following card boundary (no negative gaps).
 * Indices in `forceBreakBefore` start a new page even if they would otherwise fit (used
 * for explicit/manual page breaks in the document).
 */
export function paginate(heights: number[], m: PageMetrics, forceBreakBefore?: ReadonlySet<number>): PaginationResult {
  const spacerBefore = new Map<number, number>();
  const pageOfBlock: number[] = [];
  const offsetInPage: number[] = [];
  if (m.contentHeight <= 0 || m.pageStep <= 0)
    return { spacerBefore, cardCount: 1, pageOfBlock: heights.map(() => 0), offsetInPage: heights.map(() => 0) };

  let localY = 0; // doc-local y where the next block goes (y=0 is page 0 content top)
  let page = 0;
  for (let i = 0; i < heights.length; i++) {
    const h = heights[i] ?? 0;
    const pageBottom = page * m.pageStep + m.contentHeight - (m.reserveOf?.(page) ?? 0);
    const pageTop = page * m.pageStep;
    const forced = forceBreakBefore?.has(i) ?? false;
    if (localY > pageTop && (forced || localY + h > pageBottom + EPS)) {
      // Start a new page: a forced break always advances; otherwise only when the block
      // does not fit. Advance to the next card boundary at or below the current y (covers
      // a prior oversized block that overran).
      const nextPage = Math.max(page + 1, Math.ceil(localY / m.pageStep));
      const spacer = nextPage * m.pageStep - localY;
      if (spacer > 0) spacerBefore.set(i, spacer);
      page = nextPage;
      localY = page * m.pageStep;
    }
    pageOfBlock[i] = page;
    offsetInPage[i] = localY - page * m.pageStep;
    localY += h;
  }
  const cardCount = Math.max(page + 1, Math.ceil(localY / m.pageStep) || 1);
  return { spacerBefore, cardCount, pageOfBlock, offsetInPage };
}
