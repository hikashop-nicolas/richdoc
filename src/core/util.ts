// Small format-agnostic helpers shared by the engine and the format adapters.
import type { ImageLayout, ImageWrap } from "./types";

/** Base64-encode bytes in chunks (avoids call-stack limits on large images/fonts). */
export function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(bin);
}

/** Normalize any CSS colour (#rgb, #rrggbb, rgb()) to a 6-hex uppercase string. */
export function toHex6(c: string | undefined): string | undefined {
  if (!c) return undefined;
  const s = c.trim();
  let m = /^#?([0-9a-fA-F]{6})$/.exec(s);
  if (m) return m[1]!.toUpperCase();
  m = /^#?([0-9a-fA-F]{3})$/.exec(s);
  if (m) return m[1]!.split("").map((x) => x + x).join("").toUpperCase();
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (m) return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase();
  return undefined;
}

/** A block's per-side CSS border, read for serialization. */
export interface BlockBorderSide { side: "top" | "right" | "bottom" | "left"; px: number; style: string; hex: string }
/** Per-side borders set on a block (the browser expands the border-<side> shorthand to these
 *  longhands). Sides with no border are omitted; colour is normalized to 6-hex (default black). */
export function blockBorders(el: HTMLElement): BlockBorderSide[] {
  const out: BlockBorderSide[] = [];
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const style = el.style.getPropertyValue(`border-${side}-style`);
    if (!style || style === "none" || style === "hidden") continue;
    const px = Math.max(1, Math.round(parseFloat(el.style.getPropertyValue(`border-${side}-width`)) || 1));
    const hex = toHex6(el.style.getPropertyValue(`border-${side}-color`)) ?? "000000";
    out.push({ side, px, style, hex });
  }
  return out;
}

/** Parse a CSS border value ("1px solid #000000") into its parts, or null if absent / borderless. */
export function parseCssBorder(v: string | undefined): { px: number; style: string; hex: string } | null {
  if (!v) return null;
  const m = /^([\d.]+)px\s+(\w+)\s+#?([0-9a-fA-F]{6})/.exec(v.trim());
  if (!m || m[2] === "none" || m[2] === "hidden") return null;
  return { px: Math.max(1, Math.round(parseFloat(m[1]!))), style: m[2]!, hex: m[3]!.toUpperCase() };
}

/** CSS font-size (pt or px) to OOXML/ODF half-points. */
export function fontSizeToHalfPt(v: string | undefined): number | undefined {
  if (!v) return undefined;
  let m = /^([\d.]+)pt$/.exec(v.trim());
  if (m) return Math.round(parseFloat(m[1]!) * 2);
  m = /^([\d.]+)px$/.exec(v.trim());
  if (m) return Math.round(parseFloat(m[1]!) * 0.75 * 2);
  return undefined;
}

/** First family name from a CSS font-family list, unquoted. */
export const firstFontFamily = (v: string | undefined): string | undefined =>
  v ? (v.split(",")[0] ?? "").trim().replace(/^['"]|['"]$/g, "") || undefined : undefined;

/** Read an <img>'s floating layout from its data attributes, or null when it is inline. This
 *  is the format-agnostic contract: the editor sets the attrs, each adapter maps them to its
 *  own XML (docx wp:anchor, odt draw:frame). */
export function imageLayoutFromEl(el: Element): ImageLayout | null {
  const wrap = el.getAttribute("data-rdoc-wrap") as ImageWrap | null;
  if (!wrap) return null;
  const a = el.getAttribute("data-rdoc-align");
  // data-rdoc-wrapdist is "t,r,b,l" in px (the gap kept clear of text on each side).
  const d = (el.getAttribute("data-rdoc-wrapdist") ?? "").split(",").map((n) => Number(n));
  const dist = d.length === 4 && d.every((n) => Number.isFinite(n)) ? { t: d[0]!, r: d[1]!, b: d[2]!, l: d[3]! } : undefined;
  return {
    wrap,
    align: a === "right" || a === "center" ? a : "left",
    x: Number(el.getAttribute("data-rdoc-x")) || 0,
    y: Number(el.getAttribute("data-rdoc-y")) || 0,
    dist,
  };
}

/** Build the data-rdoc-* attributes (and inline style) an adapter's reader puts on a floating
 *  <img> from its layout. Shared so docx and odt render floating images identically. */
export function imageLayoutAttrs(layout: ImageLayout | null): string {
  if (!layout) return "";
  let out = ` data-rdoc-wrap="${layout.wrap}" data-rdoc-align="${layout.align}"`;
  const styles: string[] = [];
  if (layout.wrap === "behind" || layout.wrap === "front") {
    out += ` data-rdoc-x="${layout.x}" data-rdoc-y="${layout.y}"`;
    styles.push(`left:${layout.x}px`, `top:${layout.y}px`);
  }
  if (layout.dist) {
    const d = layout.dist;
    out += ` data-rdoc-wrapdist="${d.t},${d.r},${d.b},${d.l}"`;
    if (layout.wrap === "square" || layout.wrap === "tight" || layout.wrap === "topbottom") {
      styles.push(`margin:${d.t}px ${d.r}px ${d.b}px ${d.l}px`); // the file's wrap padding
    }
  }
  if (styles.length) out += ` style="${styles.join(";")}"`;
  return out;
}
