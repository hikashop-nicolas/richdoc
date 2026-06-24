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

  // --- Margin rulers --------------------------------------------------------
  // Google-Docs-style rulers frame the page: a horizontal one (top) carries draggable
  // left/right margin handles, a vertical one (left) carries top/bottom handles. They sit
  // in a wrapper around the page so they track the scaled footprint, and dragging a handle
  // rewrites geometry.margin, re-renders, re-paginates, and writes back via getBytes.
  const pageWrap = document.createElement("div");
  pageWrap.className = "docxedit-pagewrap";
  const mkRuler = (cls: string) => {
    const r = document.createElement("div");
    r.className = `docxedit-ruler ${cls}`;
    const fill = document.createElement("div");
    fill.className = "docxedit-ruler-fill";
    const ticks = document.createElement("div"); // graduation marks (drawn above the fill)
    ticks.className = "docxedit-ruler-ticks";
    r.append(fill, ticks);
    return { r, fill, ticks };
  };
  const mkHandle = (cls: string, label: string) => {
    const h = document.createElement("div");
    h.className = `docxedit-ruler-handle ${cls}`;
    h.title = label;
    h.setAttribute("role", "slider");
    h.setAttribute("aria-label", label);
    return h;
  };
  const hRuler = mkRuler("docxedit-ruler-h");
  const vRuler = mkRuler("docxedit-ruler-v");
  const hLeft = mkHandle("docxedit-rh-h", t("marginLeft"));
  const hRight = mkHandle("docxedit-rh-h", t("marginRight"));
  const vTop = mkHandle("docxedit-rh-v", t("marginTop"));
  const vBottom = mkHandle("docxedit-rh-v", t("marginBottom"));
  hRuler.r.append(hLeft, hRight);
  vRuler.r.append(vTop, vBottom);
  pageWrap.append(hRuler.r, vRuler.r, pagebox);
  canvas.append(leftSpacer, pageWrap, rightArea);
  scroll.appendChild(canvas);

  // Graduations: a minor tick every 0.5cm and a darker major tick every 1cm, scaled with zoom.
  const CM = 96 / 2.54;
  const tickGradient = (dir: "to right" | "to bottom", z: number): string =>
    `repeating-linear-gradient(${dir}, #6b7682 0, #6b7682 1px, transparent 1px, transparent ${CM * z}px),` +
    `repeating-linear-gradient(${dir}, #aab2bc 0, #aab2bc 1px, transparent 1px, transparent ${(CM / 2) * z}px)`;
  const updateRulers = () => {
    const z = effectiveZoom();
    const g = geometry, m = g.margin;
    hRuler.r.style.width = `${g.widthPx * z}px`;
    vRuler.r.style.height = `${g.heightPx * z}px`;
    hRuler.ticks.style.backgroundImage = tickGradient("to right", z);
    vRuler.ticks.style.backgroundImage = tickGradient("to bottom", z);
    // Vertical: the page grows along x, so align the (one-page-wide) horizontal ruler with the
    // rightmost page (page 1, the reading start). The vertical ruler maps to the fixed height.
    hRuler.r.style.left = vertical ? `${22 + Math.max(0, page.offsetWidth - g.widthPx) * z}px` : "";
    hRuler.fill.style.left = `${m.left * z}px`;
    hRuler.fill.style.right = `${m.right * z}px`;
    vRuler.fill.style.top = `${m.top * z}px`;
    vRuler.fill.style.bottom = `${m.bottom * z}px`;
    hLeft.style.left = `${m.left * z}px`;
    hRight.style.left = `${(g.widthPx - m.right) * z}px`;
    vTop.style.top = `${m.top * z}px`;
    vBottom.style.top = `${(g.heightPx - m.bottom) * z}px`;
  };

  // Drag a handle: map the pointer position within its ruler to unscaled page px, clamp so
  // opposing margins keep a minimum content band, then live-update (debounced reflow).
  const MIN_CONTENT = 96; // ~1in of content must remain between opposing margins
  const dragHandle = (handle: HTMLElement, axis: "h" | "v", side: "left" | "right" | "top" | "bottom") => {
    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      const ruler = axis === "h" ? hRuler.r : vRuler.r;
      try { handle.setPointerCapture(e.pointerId); } catch { /* no active pointer */ }
      const onMove = (ev: PointerEvent) => {
        const z = effectiveZoom();
        const rect = ruler.getBoundingClientRect();
        let pos = (axis === "h" ? ev.clientX - rect.left : ev.clientY - rect.top) / z;
        // Magnet: snap to the nearest 0.5cm graduation when within ~5 screen px (hold Alt to bypass).
        if (!ev.altKey) {
          const step = CM / 2;
          const snapped = Math.round(pos / step) * step;
          if (Math.abs(snapped - pos) * z < 5) pos = snapped;
        }
        const m = geometry.margin, W = geometry.widthPx, H = geometry.heightPx;
        if (side === "left") m.left = Math.max(0, Math.min(pos, W - m.right - MIN_CONTENT));
        else if (side === "right") m.right = Math.max(0, Math.min(W - pos, W - m.left - MIN_CONTENT));
        else if (side === "top") m.top = Math.max(0, Math.min(pos, H - m.bottom - MIN_CONTENT));
        else m.bottom = Math.max(0, Math.min(H - pos, H - m.top - MIN_CONTENT));
        geometryDirty = true;
        applyGeometry();
        updateRulers();
        scheduleReflow();
        mark();
      };
      const onUp = (ev: PointerEvent) => {
        try { handle.releasePointerCapture(ev.pointerId); } catch { /* not captured */ }
        handle.removeEventListener("pointermove", onMove);
        handle.removeEventListener("pointerup", onUp);
        reflow();
      };
      handle.addEventListener("pointermove", onMove);
      handle.addEventListener("pointerup", onUp);
    });
  };
  dragHandle(hLeft, "h", "left");
  dragHandle(hRight, "h", "right");
  dragHandle(vTop, "v", "top");
  dragHandle(vBottom, "v", "bottom");

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
  // A slider (like pdfedit) plus a clickable percentage that resets to fit-to-width.
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
  const zoomLabel = document.createElement("button");
  zoomLabel.type = "button";
  zoomLabel.className = "docxedit-zoom-label";
  zoomLabel.title = t("zoomReset");
  zoomLabel.textContent = "100%";
  zoomLabel.addEventListener("mousedown", (e) => e.preventDefault());
  zoomLabel.addEventListener("click", () => setZoom(null));
  const applyZoom = () => {
    const z = effectiveZoom();
    page.style.transformOrigin = "top left";
    page.style.transform = `scale(${z})`;
    // A vertical page grows in width and is fixed in height; a horizontal one is the reverse.
    pagebox.style.width = `${Math.round((vertical ? page.offsetWidth : geometry.widthPx) * z)}px`;
    pagebox.style.height = `${Math.round((vertical ? geometry.heightPx : page.offsetHeight) * z)}px`;
    zoomLabel.textContent = `${Math.round(z * 100)}%`;
    zoomSlider.value = String(Math.round(z * 100));
    updateRulers();
  };
  const setZoom = (z: number | null) => {
    userZoom = z == null ? null : Math.max(0.3, Math.min(2.5, Math.round(z * 100) / 100));
    applyZoom();
    positionCards();
  };

  return { applyZoom, effectiveZoom, setZoom, zoomSlider, zoomLabel, isGeometryDirty: () => geometryDirty };
}
