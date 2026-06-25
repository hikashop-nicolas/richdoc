// Page view: the on-screen page framing. Owns the margin rulers (Google-Docs-style
// draggable handles that rewrite the geometry margins) and zoom (visual scale plus the
// slider/percentage controls), assembles the centred canvas around the page box, and tracks
// whether the margins were edited (so the engine writes them back). Extracted from the
// engine; the geometry, layout callbacks and DOM refs come in as deps.
import { t } from "../i18n";
import type { EditorOptions, PageGeometry } from "../types";

export interface PageViewDeps {
  page: HTMLElement;
  pagebox: HTMLElement;
  canvas: HTMLElement;
  leftSpacer: HTMLElement;
  rightArea: HTMLElement;
  scroll: HTMLElement;
  geometry: PageGeometry;
  options: EditorOptions;
  vertical: boolean;
  applyGeometry: () => void;
  mark: () => void;
  positionCards: () => void;
  reflow: () => void;
  scheduleReflow: () => void;
}

export function setupPageView(deps: PageViewDeps) {
  const { page, pagebox, canvas, leftSpacer, rightArea, scroll, geometry, options, vertical, applyGeometry, mark, positionCards, reflow, scheduleReflow } = deps;
  let geometryDirty = false;

  // --- Per-page rulers ------------------------------------------------------
  // One ruler set (a horizontal ruler above + a vertical ruler to the left) is drawn per
  // rendered page, each sized to that page and graduated in cm, via an overlay synced to the
  // page elements on every reflow/zoom. On single-section / columned documents the handles
  // drag the document margins; documents with mixed per-section pages get a correctly sized
  // ruler per section (per-section margin authoring is a later step).
  const pageWrap = document.createElement("div");
  pageWrap.className = "docxedit-pagewrap";
  pageWrap.append(pagebox);
  const rulerLayer = document.createElement("div"); // per-page ruler overlay, positioned over the pages
  rulerLayer.className = "docxedit-rulerlayer";
  rulerLayer.setAttribute("aria-hidden", "true");
  pageWrap.appendChild(rulerLayer);
  canvas.append(leftSpacer, pageWrap, rightArea);
  scroll.appendChild(canvas);

  const CM = 96 / 2.54;
  const RT = 16; // ruler thickness in screen px
  const MIN_CONTENT = 96; // ~1in of content must remain between opposing margins
  const tickGradient = (dir: "to right" | "to bottom", z: number): string =>
    `repeating-linear-gradient(${dir}, #6b7682 0, #6b7682 1px, transparent 1px, transparent ${CM * z}px),` +
    `repeating-linear-gradient(${dir}, #aab2bc 0, #aab2bc 1px, transparent 1px, transparent ${(CM / 2) * z}px)`;

  let dragging = false;
  const rulerTip = document.createElement("div");
  rulerTip.className = "docxedit-ruler-tip";
  rulerTip.hidden = true;
  scroll.appendChild(rulerTip);
  const showTipAt = (clientX: number, clientY: number, px: number) => {
    rulerTip.textContent = `${(px / CM).toFixed(2)} cm`;
    rulerTip.hidden = false;
    rulerTip.style.left = `${Math.round(clientX - rulerTip.offsetWidth / 2)}px`;
    rulerTip.style.top = `${Math.round(clientY - rulerTip.offsetHeight - 8)}px`;
  };

  interface RulerSet {
    root: HTMLElement;
    h: { r: HTMLElement; fill: HTMLElement; ticks: HTMLElement; left: HTMLElement; right: HTMLElement };
    v: { r: HTMLElement; fill: HTMLElement; ticks: HTMLElement; top: HTMLElement; bottom: HTMLElement };
  }
  const mkBar = (cls: string) => {
    const r = document.createElement("div");
    r.className = `docxedit-ruler ${cls}`;
    const fill = document.createElement("div");
    fill.className = "docxedit-ruler-fill";
    const ticks = document.createElement("div");
    ticks.className = "docxedit-ruler-ticks";
    r.append(fill, ticks);
    return { r, fill, ticks };
  };
  const mkHandle = (cls: string, label: string, parent: HTMLElement) => {
    const h = document.createElement("div");
    h.className = `docxedit-ruler-handle ${cls}`;
    h.title = label;
    h.setAttribute("role", "slider");
    h.setAttribute("aria-label", label);
    parent.appendChild(h);
    return h;
  };
  // Drag a margin handle: map the pointer to unscaled page px, clamp, and live-update the
  // document geometry (handles only show on editable, document-geometry pages).
  const bindDrag = (handle: HTMLElement, axis: "h" | "v", side: "left" | "right" | "top" | "bottom") => {
    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      dragging = true;
      const ruler = handle.parentElement!;
      try { handle.setPointerCapture(e.pointerId); } catch { /* no active pointer */ }
      const onMove = (ev: PointerEvent) => {
        const z = effectiveZoom();
        const rect = ruler.getBoundingClientRect();
        let pos = (axis === "h" ? ev.clientX - rect.left : ev.clientY - rect.top) / z;
        if (!ev.altKey) {
          const step = CM / 2; // magnet to 0.5cm graduations (Alt bypasses)
          const snapped = Math.round(pos / step) * step;
          if (Math.abs(snapped - pos) * z < 5) pos = snapped;
        }
        const m = geometry.margin, W = geometry.widthPx, H = geometry.heightPx;
        if (side === "left") m.left = Math.max(0, Math.min(pos, W - m.right - MIN_CONTENT));
        else if (side === "right") m.right = Math.max(0, Math.min(W - pos, W - m.left - MIN_CONTENT));
        else if (side === "top") m.top = Math.max(0, Math.min(pos, H - m.bottom - MIN_CONTENT));
        else m.bottom = Math.max(0, Math.min(H - pos, H - m.top - MIN_CONTENT));
        showTipAt(ev.clientX, ev.clientY, pos);
        geometryDirty = true;
        applyGeometry();
        syncRulers();
        scheduleReflow();
        mark();
      };
      const onUp = (ev: PointerEvent) => {
        try { handle.releasePointerCapture(ev.pointerId); } catch { /* not captured */ }
        handle.removeEventListener("pointermove", onMove);
        handle.removeEventListener("pointerup", onUp);
        dragging = false;
        rulerTip.hidden = true;
        reflow();
      };
      handle.addEventListener("pointermove", onMove);
      handle.addEventListener("pointerup", onUp);
    });
  };
  const makeRulerSet = (): RulerSet => {
    const root = document.createElement("div");
    root.className = "docxedit-rulerset";
    const h = mkBar("docxedit-ruler-h");
    const v = mkBar("docxedit-ruler-v");
    const set: RulerSet = {
      root,
      h: { ...h, left: mkHandle("docxedit-rh-h", t("marginLeft"), h.r), right: mkHandle("docxedit-rh-h", t("marginRight"), h.r) },
      v: { ...v, top: mkHandle("docxedit-rh-v", t("marginTop"), v.r), bottom: mkHandle("docxedit-rh-v", t("marginBottom"), v.r) },
    };
    root.append(h.r, v.r);
    bindDrag(set.h.left, "h", "left");
    bindDrag(set.h.right, "h", "right");
    bindDrag(set.v.top, "v", "top");
    bindDrag(set.v.bottom, "v", "bottom");
    // Hover anywhere on a ruler: show the cursor's distance from the page's top-left corner.
    h.r.addEventListener("mousemove", (e) => { if (!dragging) showTipAt(e.clientX, e.clientY, (e.clientX - h.r.getBoundingClientRect().left) / effectiveZoom()); });
    v.r.addEventListener("mousemove", (e) => { if (!dragging) showTipAt(e.clientX, e.clientY, (e.clientY - v.r.getBoundingClientRect().top) / effectiveZoom()); });
    for (const bar of [h.r, v.r]) bar.addEventListener("mouseleave", () => { if (!dragging) rulerTip.hidden = true; });
    return set;
  };

  // A page's unscaled geometry: a per-section box carries its own size/margins (display only);
  // a plain page card uses the document geometry (editable handles).
  const pageGeomOf = (pg: HTMLElement): { w: number; h: number; m: { top: number; right: number; bottom: number; left: number }; editable: boolean } => {
    const sg = pg.getAttribute("data-rdoc-secgeom");
    if (sg) {
      try {
        const g = JSON.parse(sg);
        return { w: g.w, h: g.h, m: { top: g.mt, right: g.mr, bottom: g.mb, left: g.ml }, editable: false };
      } catch { /* fall through */ }
    }
    return { w: geometry.widthPx, h: geometry.heightPx, m: geometry.margin, editable: true };
  };
  // The page the caret (or last selection) is on: the section box the caret sits in, else the
  // card whose vertical band the caret falls in, else the first page.
  const activePage = (): HTMLElement | null => {
    const all = Array.from(page.querySelectorAll<HTMLElement>(".docxedit-pagecard, .docxedit-secpage"));
    if (!all.length) return null;
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.anchorNode : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
    if (el && page.contains(el)) {
      const box = el.closest<HTMLElement>(".docxedit-secpage");
      if (box) return box;
      const cr = sel!.getRangeAt(0).getBoundingClientRect();
      if (cr.top || cr.bottom) {
        const hit = all.find((c) => { const r = c.getBoundingClientRect(); return cr.top >= r.top - 2 && cr.top <= r.bottom + 2; });
        if (hit) return hit;
      }
    }
    return all[0]!;
  };

  let set: RulerSet | null = null;
  const syncRulers = () => {
    const pg = vertical ? null : activePage();
    if (!pg) { rulerLayer.style.display = "none"; return; } // vertical pages / nothing: rulers off
    rulerLayer.style.display = "";
    if (!set) { set = makeRulerSet(); rulerLayer.appendChild(set.root); }
    const z = effectiveZoom();
    const wrapRect = pageWrap.getBoundingClientRect();
    const r = pg.getBoundingClientRect();
    const x = r.left - wrapRect.left;
    const y = r.top - wrapRect.top;
    const g = pageGeomOf(pg);
    // horizontal ruler above the page
    set.h.r.style.cssText = `left:${x}px;top:${y - RT}px;width:${r.width}px;height:${RT}px`;
    set.h.ticks.style.backgroundImage = tickGradient("to right", z);
    set.h.fill.style.left = `${g.m.left * z}px`;
    set.h.fill.style.right = `${g.m.right * z}px`;
    set.h.left.style.left = `${g.m.left * z}px`;
    set.h.right.style.left = `${(g.w - g.m.right) * z}px`;
    // vertical ruler left of the page
    set.v.r.style.cssText = `left:${x - RT}px;top:${y}px;width:${RT}px;height:${r.height}px`;
    set.v.ticks.style.backgroundImage = tickGradient("to bottom", z);
    set.v.fill.style.top = `${g.m.top * z}px`;
    set.v.fill.style.bottom = `${g.m.bottom * z}px`;
    set.v.top.style.top = `${g.m.top * z}px`;
    set.v.bottom.style.top = `${(g.h - g.m.bottom) * z}px`;
    // handles only on editable (document-geometry) pages
    for (const hd of [set.h.left, set.h.right, set.v.top, set.v.bottom]) hd.style.display = g.editable ? "" : "none";
  };
  const updateRulers = syncRulers; // applyZoom and reflow call this
  // Move the ruler to whatever page the caret is on (coalesced; setTimeout survives backgrounding).
  let rulerSyncTimer = 0;
  const onSelForRuler = () => {
    window.clearTimeout(rulerSyncTimer);
    rulerSyncTimer = window.setTimeout(syncRulers, 30);
  };
  document.addEventListener("selectionchange", onSelForRuler);

  // --- Zoom -----------------------------------------------------------------
  // Scale the page visually; userZoom null means fit-to-width. The body's layout (and so
  // pagination) stays unscaled because the transform does not affect offsetTop/Height.
  let userZoom: number | null = options.zoom ?? null;
  const fitZoom = (): number => {
    // Vertical pages grow along x, so fit to the page's fixed height; horizontal fit to width.
    const avail = (vertical ? scroll.clientHeight : scroll.clientWidth) - 56;
    const dim = vertical ? geometry.heightPx : geometry.widthPx;
    return Math.max(0.2, Math.min(1, avail / Math.max(1, dim)));
  };
  const effectiveZoom = (): number => userZoom ?? fitZoom();
  // A slider (like pdfedit) plus an editable percentage field the user can type into.
  const zoomSlider = document.createElement("input");
  zoomSlider.type = "range";
  zoomSlider.min = "30";
  zoomSlider.max = "250";
  zoomSlider.step = "5";
  zoomSlider.value = "100";
  zoomSlider.className = "docxedit-zoom-slider";
  zoomSlider.title = t("zoom");
  zoomSlider.setAttribute("aria-label", t("zoom"));
  zoomSlider.addEventListener("input", () => setZoom(Number(zoomSlider.value) / 100));
  const zoomLabel = document.createElement("input");
  zoomLabel.type = "text";
  zoomLabel.inputMode = "numeric";
  zoomLabel.className = "docxedit-zoom-label";
  zoomLabel.title = t("zoom");
  zoomLabel.setAttribute("aria-label", t("zoom"));
  zoomLabel.value = "100%";
  const commitZoom = () => {
    const n = parseInt(zoomLabel.value.replace(/[^0-9]/g, ""), 10);
    if (Number.isFinite(n) && n > 0) setZoom(n / 100);
    else applyZoom(); // invalid entry: restore the shown value
  };
  zoomLabel.addEventListener("focus", () => zoomLabel.select());
  zoomLabel.addEventListener("change", commitZoom);
  zoomLabel.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commitZoom();
      zoomLabel.blur();
    }
  });
  const applyZoom = () => {
    const z = effectiveZoom();
    page.style.transformOrigin = "top left";
    page.style.transform = `scale(${z})`;
    // Size the box to the page's real footprint (handles mixed per-section page widths too).
    // A vertical page grows in width and is fixed in height; a horizontal one is the reverse.
    pagebox.style.width = `${Math.round(page.offsetWidth * z)}px`;
    pagebox.style.height = `${Math.round((vertical ? geometry.heightPx : page.offsetHeight) * z)}px`;
    if (document.activeElement !== zoomLabel) zoomLabel.value = `${Math.round(z * 100)}%`;
    zoomSlider.value = String(Math.round(z * 100));
    updateRulers();
  };
  const setZoom = (z: number | null) => {
    userZoom = z == null ? null : Math.max(0.3, Math.min(2.5, Math.round(z * 100) / 100));
    applyZoom();
    positionCards();
  };

  const teardown = () => {
    document.removeEventListener("selectionchange", onSelForRuler);
    window.clearTimeout(rulerSyncTimer);
  };
  return { applyZoom, effectiveZoom, setZoom, zoomSlider, zoomLabel, isGeometryDirty: () => geometryDirty, teardown };
}
