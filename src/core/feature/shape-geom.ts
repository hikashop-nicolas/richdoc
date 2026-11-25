// The outline of a drawn shape (an arrow, a star, a callout ...): an SVG laid in the shape's box,
// painted from the box's own fill and border, which stay the model the shape bar edits and the
// writer reads. The box's own background and border are hidden by CSS ([data-geom]).

/** Shading for 3-D looking faces: darker below 0, lighter above. */
export const SHADE: Record<string, number> = { darken: -0.4, darkenLess: -0.2, lighten: 0.4, lightenLess: 0.2 };

/** A colour (#rrggbb or rgb()) darkened toward black (k < 0) or lightened toward white (k > 0). */
export function shadeColor(color: string, k: number): string {
  let rgb: number[] | undefined;
  const hex = /^#?([0-9a-f]{6})$/i.exec(color.trim());
  if (hex) { const n = parseInt(hex[1]!, 16); rgb = [n >> 16, (n >> 8) & 255, n & 255]; }
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color);
  if (m) rgb = [Number(m[1]), Number(m[2]), Number(m[3])];
  if (!rgb || !k) return color;
  const ch = (v: number) => Math.round(k < 0 ? v * (1 + k) : v + (255 - v) * k);
  return "#" + rgb.map((v) => ch(v).toString(16).padStart(2, "0")).join("");
}

/** Repaint a box's outline SVG from the box's fill and border, after the shape bar changed them. */
export function paintShapeGeom(box: HTMLElement): void {
  const svg = box.querySelector<SVGElement>(":scope > svg.rdoc-shape-geom");
  if (!svg) return;
  const st = box.style;
  const fill = st.backgroundColor && st.backgroundColor !== "transparent" ? st.backgroundColor : "";
  const bw = st.borderTopStyle && st.borderTopStyle !== "none" ? parseFloat(st.borderTopWidth) || 0 : 0;
  const line = bw ? st.borderTopColor || "#000000" : "";
  svg.style.left = svg.style.top = `${-bw}px`;
  svg.style.width = svg.style.height = "";
  svg.style.width = `calc(100% + ${2 * bw}px)`;
  svg.style.height = `calc(100% + ${2 * bw}px)`;
  for (const p of Array.from(svg.querySelectorAll<SVGElement>("path"))) {
    const head = p.hasAttribute("data-head");
    const mode = p.getAttribute("data-fill") ?? "norm";
    if (head) {
      p.setAttribute(p.getAttribute("data-head") === "open" ? "stroke" : "fill", line || "none");
      if (p.getAttribute("data-head") === "open") p.setAttribute("stroke-width", String(bw || 1));
      continue;
    }
    p.setAttribute("fill", mode === "none" || !fill ? "none" : shadeColor(fill, SHADE[mode] ?? 0));
    if (p.getAttribute("data-stroke") === "0" || !line) p.setAttribute("stroke", "none");
    else { p.setAttribute("stroke", line); p.setAttribute("stroke-width", String(bw)); }
  }
}
