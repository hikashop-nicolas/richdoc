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
