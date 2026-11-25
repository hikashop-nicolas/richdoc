// Floating pictures and shapes (behind or in front of the text): where they are placed from, and
// the one order they stack in. Each keeps its place in that order as data-rdoc-z (pictures) or
// data-shape-z (shapes), the file's own value (docx relativeHeight, odt draw:z-index), and is
// painted by its rank in its layer. Its offsets are from its references across (data-*-relh) and
// down (data-*-relv): the page, its margins or its paragraph.

const FLOAT = ":is(.docx-shape, .odt-shape):is([data-shape-wrap='front'], [data-shape-wrap='behind']), img:is([data-rdoc-wrap='front'], [data-rdoc-wrap='behind'])";

/** Flag a shape as edited ("text", "look", "pos", "wrap", "z", "adj"), keeping every flag once set. */
export function markShapeDirty(shape: Element, kind: "text" | "look" | "pos" | "wrap" | "z" | "adj"): void {
  const cur = new Set((shape.getAttribute("data-shape-dirty") ?? "").split(/\s+/).filter(Boolean));
  cur.add(kind);
  shape.setAttribute("data-shape-dirty", [...cur].join(" "));
}

const isImg = (el: Element) => el.tagName === "IMG";
export const layerOf = (el: Element): string => el.getAttribute(isImg(el) ? "data-rdoc-wrap" : "data-shape-wrap") ?? "";
export const zOf = (el: Element): number => Number(el.getAttribute(isImg(el) ? "data-rdoc-z" : "data-shape-z")) || 0;

/** Set a float's order (an odt drawing also keeps it in its edit, which the writer reads). */
export function setZ(el: HTMLElement, z: number): void {
  if (isImg(el)) { el.setAttribute("data-rdoc-z", String(z)); return; }
  el.setAttribute("data-shape-z", String(z));
  if (el.matches(".odt-shape")) {
    let e: Record<string, unknown> = {};
    try { e = JSON.parse(el.getAttribute("data-shape-edit") || "{}"); } catch { /* a fresh edit */ }
    e.z = z;
    el.setAttribute("data-shape-edit", JSON.stringify(e));
  }
  markShapeDirty(el, "z");
}

/** Every float in the regions (not the members of a shape group). */
export function floats(regions: HTMLElement[]): HTMLElement[] {
  return regions.flatMap((r) => Array.from(r.querySelectorAll<HTMLElement>(FLOAT))).filter((el) => !el.parentElement?.closest(".docx-shape, .odt-shape"));
}

/** Paint the floats in their order: in front of the text above it, behind it below it. */
export function stack(regions: HTMLElement[]): void {
  const layers: Record<string, HTMLElement[]> = { front: [], behind: [] };
  for (const el of floats(regions)) layers[layerOf(el)]?.push(el);
  for (const [w, list] of Object.entries(layers)) {
    const sorted = list.map((el, i) => ({ el, i, z: zOf(el) })).sort((a, b) => a.z - b.z || a.i - b.i);
    sorted.forEach(({ el }, k) => {
      const painted = el.matches(".odt-shape") ? el.querySelector<HTMLElement>(":scope > .odt-shape-box") ?? el : el;
      painted.style.zIndex = String(w === "front" ? 4 + k : k - sorted.length);
    });
  }
}

/** The order just above everything that has one (shapes in the text keep theirs for later). */
export function topZ(regions: HTMLElement[]): number {
  const all = regions.flatMap((r) => Array.from(r.querySelectorAll<HTMLElement>("[data-shape-z], img[data-rdoc-z]")));
  return Math.max(0, ...all.map(zOf), ...floats(regions).map(zOf)) + 1;
}

/** Bring a float above every other one (1) or below them all (-1), and repaint. */
export function restack(el: HTMLElement, dir: 1 | -1, regions: HTMLElement[]): void {
  if (dir > 0) setZ(el, topZ(regions));
  else {
    const others = floats(regions).filter((o) => o !== el);
    if (!others.length) return;
    const low = Math.min(...others.map(zOf));
    // No room below: lift the others one step (orders are whole numbers from 0).
    if (low < 1) for (const o of others) setZ(o, zOf(o) + 1);
    setZ(el, Math.max(0, low - 1));
  }
  stack(regions);
}

/** Whether a point is on a letter of the text (not the gaps between lines and words). */
export function onText(x: number, y: number): boolean {
  const r = document.caretRangeFromPoint?.(x, y);
  const n = r?.startContainer;
  if (!r || !n || n.nodeType !== 3) return false;
  const len = (n as Text).length;
  for (const [a, b] of [[r.startOffset - 1, r.startOffset], [r.startOffset, r.startOffset + 1]] as const) {
    if (a < 0 || b > len) continue;
    const q = document.createRange();
    q.setStart(n, a);
    q.setEnd(n, b);
    for (const c of Array.from(q.getClientRects())) if (x >= c.left && x <= c.right && y >= c.top && y <= c.bottom) return true;
  }
  return false;
}

/** The float behind the text under a point that is not on a letter: what a click there picks. */
export function behindAt(x: number, y: number, regions: HTMLElement[]): HTMLElement | null {
  if (onText(x, y)) return null;
  for (const hit of document.elementsFromPoint(x, y)) {
    const el = (hit.tagName === "IMG" ? hit : hit.closest(".docx-shape, .odt-shape")) as HTMLElement | null;
    if (el && regions.some((r) => r.contains(el)) && layerOf(el) === "behind") return el;
  }
  return null;
}

/** Where a float's references put its offset 0,0, in its containing block's px (as left / top). */
export function originOf(el: HTMLElement, zoom: number): { x: number; y: number } {
  const pre = isImg(el) ? "data-rdoc" : "data-shape";
  const docEl = el.closest<HTMLElement>(".docxedit-doc");
  const cb = (isImg(el) ? el : el.querySelector<HTMLElement>(":scope > .odt-shape-box") ?? el).offsetParent as HTMLElement | null;
  if (!docEl || !cb) return { x: 0, y: 0 };
  const z = zoom || 1;
  const ds = getComputedStyle(docEl), dr = docEl.getBoundingClientRect();
  const cr = cb.getBoundingClientRect(), cs = getComputedStyle(cb);
  // The containing block's padding box, in the document's padding box.
  const cbX = (cr.left - dr.left) / z + (parseFloat(cs.borderLeftWidth) || 0) - (parseFloat(ds.borderLeftWidth) || 0);
  const cbY = (cr.top - dr.top) / z + (parseFloat(cs.borderTopWidth) || 0) - (parseFloat(ds.borderTopWidth) || 0);
  const ref = (axis: "h" | "v", page: number, margin: number, para: number) => {
    const r = el.getAttribute(`${pre}-rel${axis}`) ?? (axis === "h" ? "margin" : "para");
    return r === "page" ? page : r === "para" ? para : margin;
  };
  // In page view, the page its anchor is on (or the one it names): its top, from the doc's.
  const step = Number(docEl.dataset.rdocPageStep) || 0;
  let pageTop = 0;
  if (step) {
    const named = Number(el.getAttribute("data-shape-page"));
    const anchorTop = ((el.parentElement ?? el).getBoundingClientRect().top - dr.top) / z;
    pageTop = (named ? named - 1 : Math.max(0, Math.floor(anchorTop / step))) * step;
  }
  const marginTop = docEl.dataset.rdocMarginTop ? Number(docEl.dataset.rdocMarginTop) : parseFloat(ds.paddingTop) || 0;
  return {
    x: ref("h", 0, parseFloat(ds.paddingLeft) || 0, cbX) - cbX,
    y: ref("v", pageTop, pageTop + marginTop, cbY) - cbY,
  };
}

/** Place a floating picture from its offsets (data-rdoc-x / -y). */
export function placeImg(img: HTMLElement, zoom: number): void {
  if (!/^(behind|front)$/.test(img.getAttribute("data-rdoc-wrap") ?? "")) return;
  const o = originOf(img, zoom);
  img.style.left = `${Math.round((Number(img.getAttribute("data-rdoc-x")) || 0) + o.x)}px`;
  img.style.top = `${Math.round((Number(img.getAttribute("data-rdoc-y")) || 0) + o.y)}px`;
}

/** A dashed outline of a float, following the pointer while a wrapped one is dragged. */
export function ghostOf(el: HTMLElement): { move(dx: number, dy: number): void; remove(): void } {
  const r = (el.querySelector<HTMLElement>(":scope > .docx-shape-box, :scope > .odt-shape-box") ?? el).getBoundingClientRect();
  const g = document.createElement("div");
  g.className = "rdoc-float-ghost";
  Object.assign(g.style, { left: `${r.left}px`, top: `${r.top}px`, width: `${r.width}px`, height: `${r.height}px` });
  document.body.appendChild(g);
  return {
    move: (dx, dy) => { g.style.transform = `translate(${dx}px, ${dy}px)`; },
    remove: () => g.remove(),
  };
}

const FLOW_BLOCK = "p, h1, h2, h3, h4, h5, h6, li";

/** Drop a wrapped float where its outline's top-left lands: it leads the paragraph there (its
 *  float starts with it) and is aligned by where across that paragraph it was dropped. Returns
 *  the alignment, or null when there is no paragraph there. */
export function dropWrapped(el: HTMLElement, x: number, y: number, w: number, regions: HTMLElement[]): "left" | "center" | "right" | null {
  const inRegion = (n: Element) => regions.some((r) => r.contains(n));
  let block: HTMLElement | null = null;
  for (const hit of document.elementsFromPoint(x + 2, y + 2)) {
    if (el.contains(hit)) continue;
    const b = hit.closest<HTMLElement>(FLOW_BLOCK);
    if (b && inRegion(b) && !b.closest(".docx-shape, .odt-shape")) { block = b; break; }
  }
  if (!block) {
    const c = document.caretRangeFromPoint?.(x + 2, y + 2);
    const n = c?.startContainer;
    block = (n?.nodeType === 1 ? (n as Element) : n?.parentElement)?.closest<HTMLElement>(FLOW_BLOCK) ?? null;
    if (block && (!inRegion(block) || el.contains(block))) block = null;
  }
  if (!block) return null;
  const br = block.getBoundingClientRect();
  const mid = x + w / 2 - br.left;
  const align = mid < br.width / 3 ? "left" : mid > (2 * br.width) / 3 ? "right" : "center";
  if (block.firstChild !== el) block.insertBefore(el, block.firstChild);
  return align;
}
