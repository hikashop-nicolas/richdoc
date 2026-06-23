// Small format-agnostic helpers shared by the engine and the format adapters.

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
