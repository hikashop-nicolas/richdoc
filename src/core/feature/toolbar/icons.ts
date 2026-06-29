// Toolbar icons: the inline SVG strings (and the two parametric icon builders) used across the
// toolbar modules. Pure presentation, no behaviour, so they live apart from the wiring.

/** Alignment glyph: three stacked bars whose x-offset/width encode left/center/right/justify. */
export const alignIcon = (rows: [number, number][]): string =>
  `<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">${rows
    .map(([x, w], k) => `<rect x="${x}" y="${3 + k * 4}" width="${w}" height="1.6" rx=".6"/>`)
    .join("")}</svg>`;

/** Indent / outdent glyph; dir > 0 points right (indent), else left (outdent). */
export const indentIcon = (dir: number): string =>
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
  '<rect x="2" y="2.5" width="12" height="1.6" rx=".6"/><rect x="7" y="6" width="7" height="1.6" rx=".6"/>' +
  '<rect x="7" y="9.4" width="7" height="1.6" rx=".6"/><rect x="2" y="12.9" width="12" height="1.6" rx=".6"/>' +
  (dir > 0 ? '<path d="M2 6.2l2.6 1.9L2 10z"/>' : '<path d="M4.6 6.2L2 8.1l2.6 1.9z"/>') + "</svg>";

export const bulletIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
  '<circle cx="2.3" cy="4" r="1.3"/><circle cx="2.3" cy="11" r="1.3"/>' +
  '<rect x="6" y="3.2" width="9" height="1.6" rx=".6"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6"/></svg>';

export const numberIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">' +
  '<text x="0.3" y="6.2" font-size="6" font-family="sans-serif" fill="currentColor">1</text>' +
  '<text x="0.3" y="13.4" font-size="6" font-family="sans-serif" fill="currentColor">2</text>' +
  '<rect x="6" y="3.2" width="9" height="1.6" rx=".6" fill="currentColor"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6" fill="currentColor"/></svg>';

export const linkIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true">' +
  '<path d="M6.6 9.4l2.8-2.8"/><path d="M7.2 4.6l1-1a2.4 2.4 0 0 1 3.4 3.4l-1 1"/><path d="M8.8 11.4l-1 1a2.4 2.4 0 0 1-3.4-3.4l1-1"/></svg>';

export const pbIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<rect x="3" y="1.5" width="10" height="4" rx=".5"/><rect x="3" y="10.5" width="10" height="4" rx=".5"/>' +
  '<line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="2 1.6"/></svg>';

// Footnote: lines of text with a small superscript mark.
// A bookmark: a ribbon/pennant shape.
export const bookmarkIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"><path d="M4 2.5h8v11l-4-3-4 3z"/></svg>';
// A cross-reference: a link arrow pointing to a target line.
export const xrefIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4h6M2 8h4"/><path d="M9 11h4M11 9l2 2-2 2"/></svg>';
// A caption: a framed item with a short label line beneath it.
export const captionIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<rect x="2.5" y="2" width="11" height="7" rx="1"/><path d="M5 12h6M4 14.5h8" stroke-linecap="round"/></svg>';
// An equation: a radical sign over a baseline (√‾).
export const equationIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M1.5 9l1.8 0 1.7 4 2.8-10H14"/></svg>';
// A note (footnote or endnote): text lines with a small superscript mark.
export const footnoteIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" aria-hidden="true">' +
  '<path d="M2 4h7M2 7.5h7M2 11h5"/><text x="11" y="7" font-size="7" fill="currentColor" stroke="none">1</text></svg>';

// Furigana / ruby: a base glyph (字-like cross) with small reading marks above.
export const furiganaIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" aria-hidden="true">' +
  '<circle cx="5" cy="3" r=".5" fill="currentColor" stroke="none"/><circle cx="8" cy="3" r=".5" fill="currentColor" stroke="none"/><circle cx="11" cy="3" r=".5" fill="currentColor" stroke="none"/>' +
  '<path d="M8 6.5v7M4.5 9.5h7"/></svg>';

export const imgIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
  '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

export const cmtIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<path d="M2 3.5h12v8H6l-3 2.5V11.5H2z"/><line x1="5" y1="6.2" x2="11" y2="6.2"/><line x1="5" y1="8.6" x2="9" y2="8.6"/></svg>';

export const supIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
  '<text x="0" y="14" font-size="11" font-family="serif">x</text><text x="8" y="7" font-size="7" font-family="serif">2</text></svg>';

export const subIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
  '<text x="0" y="11" font-size="11" font-family="serif">x</text><text x="8" y="16" font-size="7" font-family="serif">2</text></svg>';

export const lineSpacingIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M6 4h8M6 8h8M6 12h8"/><path d="M2.6 3.6v8.8M1.4 4.8 2.6 3.6 3.8 4.8M1.4 11.2 2.6 12.4 3.8 11.2"/></svg>';

export const tableIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><line x1="1.5" y1="6.2" x2="14.5" y2="6.2"/>' +
  '<line x1="1.5" y1="9.9" x2="14.5" y2="9.9"/><line x1="6" y1="2.5" x2="6" y2="13.5"/><line x1="10.3" y1="2.5" x2="10.3" y2="13.5"/></svg>';

export const fieldIcon =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
  '<path d="M5 2.5C3.5 2.5 3.5 5 3.5 8s0 5.5-1.5 5.5"/><path d="M11 2.5c1.5 0 1.5 2.5 1.5 5.5s0 5.5 1.5 5.5"/></svg>';

/** Down-caret appended to a group button to mark it as a dropdown. */
export const caret = '<svg width="8" height="8" viewBox="0 0 8 8" fill="currentColor" aria-hidden="true"><path d="M1 2.5 4 5.5 7 2.5z"/></svg>';

/** Cluster-button glyphs: a serif "A" for the formatting group, a "+" for the insert group. */
export const styleGroupSvg = '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 13 7 3h2l3 10h-2l-.66-2.4H6.66L6 13H4zm3.1-4.2h1.8L8 5.4 7.1 8.8z"/></svg>';
export const insertGroupSvg = '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M7 2h2v5h5v2H9v5H7V9H2V7h5z"/></svg>';
