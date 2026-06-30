// Page view: the on-screen page framing. Owns the margin rulers (Google-Docs-style
// draggable handles that rewrite the geometry margins) and zoom (visual scale plus the
// slider/percentage controls), assembles the centred canvas around the page box, and tracks
// whether the margins were edited (so the engine writes them back). Extracted from the
// engine; the geometry, layout callbacks and DOM refs come in as deps.
import { t } from "../i18n";
import type { Capabilities, EditorOptions, PageGeometry, SecGeom } from "../types";

export interface PageViewDeps {
  page: HTMLElement;
  pagebox: HTMLElement;
  canvas: HTMLElement;
  leftSpacer: HTMLElement;
  rightArea: HTMLElement;
  scroll: HTMLElement;
  geometry: PageGeometry;
  options: EditorOptions;
  caps: Capabilities;
  /** Live writing-direction read (vertical/tategaki), so the dialog can switch direction. */
  getVertical: () => boolean;
  applyGeometry: () => void;
  mark: () => void;
  positionCards: () => void;
  reflow: () => void;
  scheduleReflow: () => void;
  /** Flag the document geometry as edited so it is written back on save. */
  markGeometryDirty: () => void;
  /** Geometry of the section the caret is in (the document geometry if not in a sub-section). */
  readSectionGeom: () => SecGeom;
  /** Apply a geometry to the caret's section (mutates the model; the caller reflows). */
  writeSectionGeom: (g: SecGeom) => void;
  /** Insert a next-page section break after the caret's paragraph (mutates the model). */
  insertSectionBreak: () => void;
  /** Turn the document's first-page / even-odd header & footer variants on or off. */
  toggleHFVariant: (variant: "first" | "even", on: boolean) => void;
}

export function setupPageView(deps: PageViewDeps) {
  const { page, pagebox, canvas, leftSpacer, rightArea, scroll, geometry, options, caps, getVertical, applyGeometry, mark, positionCards, reflow, scheduleReflow, markGeometryDirty, readSectionGeom, writeSectionGeom, insertSectionBreak, toggleHFVariant } = deps;
  const vertical = () => getVertical();

  // --- Tab-stop authoring state ---------------------------------------------
  // Tab stops live as data-rdoc-tabstops JSON on each paragraph block; the ruler reads the caret
  // block's stops to draw markers, and edits apply to every selected block (the caret block when
  // collapsed), like the other paragraph formatting.
  interface TabStop { pos: number; val: string; leader?: string }
  const TAB_TYPES = ["left", "center", "right", "decimal"] as const;
  let newTabType: (typeof TAB_TYPES)[number] = "left"; // the type given to stops added by clicking the ruler
  const TAB_BLOCK_SEL = "p,h1,h2,h3,h4,h5,h6,li,blockquote";
  const tabNear = (a: number, b: number) => Math.abs(a - b) < 6; // px tolerance to match a marker to a stop
  const tabBlocks = (): HTMLElement[] => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return [];
    const blockOf = (n: Node | null): HTMLElement | null => {
      const el = n ? (n.nodeType === 3 ? n.parentElement : (n as Element)) : null;
      const b = el?.closest?.(TAB_BLOCK_SEL) as HTMLElement | null;
      return b && page.contains(b) && !b.closest(".docxedit-note, .docxedit-hf-clone") ? b : null;
    };
    const range = sel.getRangeAt(0);
    const start = blockOf(range.startContainer);
    const end = blockOf(range.endContainer);
    const out: HTMLElement[] = start ? [start] : [];
    let n: HTMLElement | null = start;
    while (n && n !== end) { n = n.nextElementSibling as HTMLElement | null; if (n && n.matches(TAB_BLOCK_SEL)) out.push(n); }
    if (end && !out.includes(end)) out.push(end);
    return out;
  };
  const parseStops = (b: HTMLElement): TabStop[] => {
    try { const s = JSON.parse(b.getAttribute("data-rdoc-tabstops") || "[]"); return Array.isArray(s) ? s : []; } catch { return []; }
  };
  // Apply a mutation to the selected blocks' tab stops, rewrite the attr, and re-render. `live`
  // re-lays the tabs only (cheap, during a drag); otherwise a full reflow settles pagination.
  const applyTabEdit = (mutate: (stops: TabStop[]) => TabStop[], live = false): void => {
    const blocks = tabBlocks();
    if (!blocks.length) return;
    for (const b of blocks) {
      const stops = mutate(parseStops(b)).filter((s) => s.pos > 0).sort((a, c) => a.pos - c.pos);
      if (stops.length) b.setAttribute("data-rdoc-tabstops", JSON.stringify(stops));
      else {
        b.removeAttribute("data-rdoc-tabstops"); // no stops left: reset the tab spans to the default grid
        for (const tb of Array.from(b.querySelectorAll<HTMLElement>(".docx-tab"))) { tb.classList.remove("docx-tab-laid", "docx-tab-leader"); tb.style.width = ""; }
      }
    }
    mark();
    if (live) layoutTabs(); else reflow();
  };

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
    h: { r: HTMLElement; fill: HTMLElement; ticks: HTMLElement; left: HTMLElement; right: HTMLElement; tabmarks: HTMLElement };
    v: { r: HTMLElement; fill: HTMLElement; ticks: HTMLElement; top: HTMLElement; bottom: HTMLElement };
    typeSel: HTMLElement;
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
  // Drag a margin handle: map the pointer to unscaled page px, clamp, and live-update the margins
  // of whichever page the ruler is on, the document geometry or the caret's section.
  const bindDrag = (handle: HTMLElement, axis: "h" | "v", side: "left" | "right" | "top" | "bottom") => {
    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      dragging = true;
      const ruler = handle.parentElement!;
      const pg = activePage();
      const onSection = !!pg && pg.classList.contains("docxedit-secpage");
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
        // Work against the active page's own geometry (section box or the document).
        const base = pg ? pageGeomOf(pg) : { w: geometry.widthPx, h: geometry.heightPx, m: { ...geometry.margin }, section: false };
        const m = { ...base.m }, W = base.w, H = base.h;
        if (side === "left") m.left = Math.max(0, Math.min(pos, W - m.right - MIN_CONTENT));
        else if (side === "right") m.right = Math.max(0, Math.min(W - pos, W - m.left - MIN_CONTENT));
        else if (side === "top") m.top = Math.max(0, Math.min(pos, H - m.bottom - MIN_CONTENT));
        else m.bottom = Math.max(0, Math.min(H - pos, H - m.top - MIN_CONTENT));
        showTipAt(ev.clientX, ev.clientY, pos);
        if (onSection) {
          // Commit to the section's geometry; live-update the box's padding + ruler attr so the
          // preview tracks the drag (a full re-bucket happens on pointerup).
          const cur = readSectionGeom();
          const g = { ...cur, mt: m.top, mr: m.right, mb: m.bottom, ml: m.left };
          writeSectionGeom(g);
          pg!.setAttribute("data-rdoc-secgeom", JSON.stringify(g));
          pg!.style.padding = `${g.mt}px ${g.mr}px ${g.mb}px ${g.ml}px`;
          syncRulers();
          mark();
        } else {
          geometry.margin.top = m.top; geometry.margin.right = m.right; geometry.margin.bottom = m.bottom; geometry.margin.left = m.left;
          markGeometryDirty();
          applyGeometry();
          syncRulers();
          scheduleReflow();
          mark();
        }
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
    const tabmarks = document.createElement("div");
    tabmarks.className = "docxedit-tabmarks";
    h.r.appendChild(tabmarks);
    // The tab-type selector in the corner where the two rulers meet: cycles the type used for stops
    // added by clicking the ruler (left -> center -> right -> decimal), Word-style.
    const typeSel = document.createElement("button");
    typeSel.type = "button";
    typeSel.title = t("tabStopType");
    typeSel.setAttribute("aria-label", t("tabStopType"));
    const paintTypeSel = () => { typeSel.className = `docxedit-tabtype is-${newTabType}`; };
    paintTypeSel();
    typeSel.addEventListener("click", () => { newTabType = TAB_TYPES[(TAB_TYPES.indexOf(newTabType) + 1) % TAB_TYPES.length]!; paintTypeSel(); });
    const set: RulerSet = {
      root,
      h: { ...h, tabmarks, left: mkHandle("docxedit-rh-h", t("marginLeft"), h.r), right: mkHandle("docxedit-rh-h", t("marginRight"), h.r) },
      v: { ...v, top: mkHandle("docxedit-rh-v", t("marginTop"), v.r), bottom: mkHandle("docxedit-rh-v", t("marginBottom"), v.r) },
      typeSel,
    };
    root.append(h.r, v.r, typeSel);
    bindDrag(set.h.left, "h", "left");
    bindDrag(set.h.right, "h", "right");
    bindDrag(set.v.top, "v", "top");
    bindDrag(set.v.bottom, "v", "bottom");
    // Press an empty spot on the horizontal ruler's content area to add a tab stop of the current
    // type. Handled on mousedown (with preventDefault) rather than click: pressing the non-editable
    // ruler would otherwise collapse the caret out of the paragraph before a click could read it, so
    // we keep the selection and read the caret's block while it is still there.
    h.r.addEventListener("mousedown", (e) => {
      const tgt = e.target as HTMLElement;
      if (tgt.closest(".docxedit-ruler-handle, .docxedit-tabmark")) return; // handles/markers run their own pointer logic
      e.preventDefault(); // keep the document selection in the paragraph
      if (dragging) return;
      const pg = activePage();
      if (!pg) return;
      const g = pageGeomOf(pg), z = effectiveZoom();
      let pos = (e.clientX - h.r.getBoundingClientRect().left) / z - g.m.left;
      if (pos <= 0 || pos >= g.w - g.m.left - g.m.right) return; // only within the content area
      const step = CM / 2;
      const snapped = Math.round((g.m.left + pos) / step) * step - g.m.left;
      if (Math.abs(snapped - pos) * z < 5) pos = snapped;
      pos = Math.round(pos);
      applyTabEdit((s) => [...s.filter((x) => !tabNear(x.pos, pos)), { pos, val: newTabType }]);
      syncRulers();
    });
    // Hover anywhere on a ruler: show the cursor's distance from the page's top-left corner.
    h.r.addEventListener("mousemove", (e) => { if (!dragging) showTipAt(e.clientX, e.clientY, (e.clientX - h.r.getBoundingClientRect().left) / effectiveZoom()); });
    v.r.addEventListener("mousemove", (e) => { if (!dragging) showTipAt(e.clientX, e.clientY, (e.clientY - v.r.getBoundingClientRect().top) / effectiveZoom()); });
    for (const bar of [h.r, v.r]) bar.addEventListener("mouseleave", () => { if (!dragging) rulerTip.hidden = true; });
    return set;
  };
  // A draggable tab-stop marker on the horizontal ruler. Drag to move (snap magnet, Alt bypasses),
  // click (no move) to cycle its type, drag it down off the ruler to remove it.
  const mkTabMark = (stop: TabStop, g: { m: { left: number } }, z: number, hRect: DOMRect): HTMLElement => {
    const m = document.createElement("div");
    m.className = `docxedit-tabmark is-${stop.val}${stop.leader ? " is-leader" : ""}`;
    m.style.left = `${(g.m.left + stop.pos) * z}px`;
    m.title = t("tabStop");
    m.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      dragging = true;
      let moved = false, lastY = e.clientY, cur = stop.pos;
      const startX = e.clientX;
      try { m.setPointerCapture(e.pointerId); } catch { /* no pointer */ }
      const onMove = (ev: PointerEvent) => {
        lastY = ev.clientY;
        if (Math.abs(ev.clientX - startX) > 3) moved = true;
        if (!moved) return;
        let pos = (ev.clientX - hRect.left) / z - g.m.left;
        if (!ev.altKey) { const step = CM / 2; const sn = Math.round((g.m.left + pos) / step) * step - g.m.left; if (Math.abs(sn - pos) * z < 5) pos = sn; }
        pos = Math.round(Math.max(0, pos));
        applyTabEdit((s) => s.map((x) => (tabNear(x.pos, cur) ? { ...x, pos } : x)), true);
        cur = pos;
        m.style.left = `${(g.m.left + cur) * z}px`;
        showTipAt(ev.clientX, ev.clientY, g.m.left + cur);
      };
      const onUp = (ev: PointerEvent) => {
        try { m.releasePointerCapture(ev.pointerId); } catch { /* not captured */ }
        m.removeEventListener("pointermove", onMove);
        m.removeEventListener("pointerup", onUp);
        dragging = false;
        rulerTip.hidden = true;
        if (!moved) { const i = (TAB_TYPES.indexOf(stop.val as (typeof TAB_TYPES)[number]) + 1) % TAB_TYPES.length; applyTabEdit((s) => s.map((x) => (tabNear(x.pos, cur) ? { ...x, val: TAB_TYPES[i]! } : x))); }
        else if (lastY > hRect.bottom + 14) applyTabEdit((s) => s.filter((x) => !tabNear(x.pos, cur))); // dragged off: remove
        else reflow();
        syncRulers();
      };
      m.addEventListener("pointermove", onMove);
      m.addEventListener("pointerup", onUp);
    });
    return m;
  };
  // Draw the caret paragraph's tab stops as markers on the horizontal ruler.
  const renderTabMarks = (g: { m: { left: number } }, z: number, hRect: DOMRect): void => {
    if (!set) return;
    set.h.tabmarks.replaceChildren();
    const block = tabBlocks()[0];
    if (!block) return;
    for (const stop of parseStops(block)) if (stop.pos > 0) set.h.tabmarks.appendChild(mkTabMark(stop, g, z, hRect));
  };

  // A page's unscaled geometry: a per-section box carries its own size/margins (edited through the
  // section's geometry JSON); a plain page card uses the document geometry. Both are draggable.
  const pageGeomOf = (pg: HTMLElement): { w: number; h: number; m: { top: number; right: number; bottom: number; left: number }; section: boolean } => {
    const sg = pg.getAttribute("data-rdoc-secgeom");
    if (sg) {
      try {
        const g = JSON.parse(sg);
        return { w: g.w, h: g.h, m: { top: g.mt, right: g.mr, bottom: g.mb, left: g.ml }, section: true };
      } catch { /* fall through */ }
    }
    return { w: geometry.widthPx, h: geometry.heightPx, m: geometry.margin, section: false };
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
      // The caret is in the one flow, so map its position to a page card. A collapsed caret (e.g.
      // in an empty paragraph) can report a zero rect, so fall back to its element's rect. Vertical
      // pages advance along x (right to left), so match on the caret's horizontal position there.
      const cr = sel!.getRangeAt(0).getBoundingClientRect();
      if (vertical()) {
        const cx = cr.left || cr.right || el.getBoundingClientRect().left;
        const hit = all.find((c) => { const r = c.getBoundingClientRect(); return cx >= r.left - 2 && cx <= r.right + 2; });
        if (hit) return hit;
      } else {
        const cy = cr.top || cr.bottom || el.getBoundingClientRect().top;
        const hit = all.find((c) => { const r = c.getBoundingClientRect(); return cy >= r.top - 2 && cy <= r.bottom + 2; });
        if (hit) return hit;
      }
    }
    return all[0]!;
  };

  let set: RulerSet | null = null;
  const syncRulers = () => {
    const pg = activePage();
    if (!pg) { rulerLayer.style.display = "none"; return; } // nothing rendered yet: rulers off
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
    for (const hd of [set.h.left, set.h.right, set.v.top, set.v.bottom]) hd.style.display = ""; // draggable on every page
    // tab-type selector in the corner + the caret paragraph's tab-stop markers
    set.typeSel.style.cssText = `left:${x - RT}px;top:${y - RT}px;width:${RT}px;height:${RT}px`;
    renderTabMarks(g, z, set.h.r.getBoundingClientRect());
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
    const avail = (vertical() ? scroll.clientHeight : scroll.clientWidth) - 56;
    const dim = vertical() ? geometry.heightPx : geometry.widthPx;
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
  // --- Tab stops: render tabs at their paragraph's custom stops, with alignment ---------------
  // Only paragraphs carrying data-rdoc-tabstops are touched; the rest keep the CSS default grid.
  // Each tab span's width is set so the segment after it aligns to the governing stop (left: starts
  // at it, right: ends at it, center: centred, decimal: the '.'/',' sits on it). Positions are read
  // live, left-to-right, so successive tabs and wrapped lines each resolve against the tab's own x.
  const DEFAULT_TAB = 48; // unscaled px, the default 0.5in grid used past the last stop
  // The unscaled x where a segment of width segW (decimal point at `dec`) must start to satisfy a stop.
  const alignStart = (val: string, pos: number, segW: number, dec: number): number =>
    val === "right" ? pos - segW : val === "center" ? pos - segW / 2 : val === "decimal" ? pos - dec : pos;
  const layoutTabs = () => {
    const z = effectiveZoom();
    if (!(z > 0)) return;
    // Paragraphs with their own tab stops, plus styled paragraphs whose named style defines them
    // (a style's stops ride a --rdoc-tabstops custom property that the injected style CSS sets).
    const targets = new Map<HTMLElement, string>();
    for (const b of Array.from(page.querySelectorAll<HTMLElement>("[data-rdoc-tabstops]"))) targets.set(b, b.getAttribute("data-rdoc-tabstops") || "[]");
    for (const b of Array.from(page.querySelectorAll<HTMLElement>("[data-rdoc-style]"))) {
      if (targets.has(b) || !b.querySelector(".docx-tab")) continue; // own stops win
      const cv = getComputedStyle(b).getPropertyValue("--rdoc-tabstops").trim();
      if (cv) targets.set(b, cv);
    }
    for (const [block, raw] of targets) {
      const tabs = Array.from(block.querySelectorAll<HTMLElement>(".docx-tab"));
      if (!tabs.length) continue;
      if (getComputedStyle(block).writingMode !== "horizontal-tb") continue; // vertical: keep default grid
      let stops: { pos: number; val: string; leader?: string }[];
      try { stops = JSON.parse(raw); } catch { continue; }
      stops = (Array.isArray(stops) ? stops : []).filter((s) => s && s.pos > 0).sort((a, b) => a.pos - b.pos);
      // .docx-tab-laid (in the stylesheet) makes the span an inline-block on the baseline; only the
      // computed width is per-tab, so it is the one inline value we set.
      for (const tab of tabs) { tab.classList.add("docx-tab-laid"); tab.style.width = "0px"; }
      const marginLeft = parseFloat(getComputedStyle(block).marginLeft) || 0; // the paragraph indent (unscaled)
      const originX = block.getBoundingClientRect().left - marginLeft * z; // screen x of the margin origin
      for (let i = 0; i < tabs.length; i++) {
        const tab = tabs[i]!;
        const xUnscaled = (tab.getBoundingClientRect().left - originX) / z; // tab left, from the margin
        const next = tabs[i + 1] ?? null;
        // measure the segment after this tab (up to the next tab / line end) for non-left alignment
        const rng = document.createRange();
        rng.setStartAfter(tab);
        if (next) rng.setEndBefore(next); else rng.setEnd(block, block.childNodes.length);
        const segRect = rng.getBoundingClientRect();
        const segW = segRect.width / z;
        let dec = segW; // decimal point offset within the segment (right edge if none found)
        if (stops.some((s) => s.val === "decimal")) {
          try {
            const w = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
            let n: Node | null;
            while ((n = w.nextNode())) {
              const afterTab = !!(tab.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING);
              const beforeNext = !next || !!(next.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_PRECEDING);
              if (!afterTab || !beforeNext) continue;
              const m = (n.textContent || "").search(/[.,]/);
              if (m >= 0) { const r2 = document.createRange(); r2.setStart(n, m); r2.setEnd(n, m + 1); dec = (r2.getBoundingClientRect().left - segRect.left) / z; break; }
            }
          } catch { /* leave dec = segW */ }
        }
        let targetStart: number | null = null;
        let leader: string | undefined;
        for (const s of stops) {
          if (s.pos <= xUnscaled + 0.5) continue;
          const ts = alignStart(s.val, s.pos, segW, dec);
          if (ts >= xUnscaled + 0.5) { targetStart = ts; leader = s.leader; break; } // first stop that fits
        }
        if (targetStart == null) targetStart = Math.ceil((xUnscaled + 0.5) / DEFAULT_TAB) * DEFAULT_TAB;
        tab.style.width = `${Math.max(0, targetStart - xUnscaled)}px`;
        tab.classList.toggle("docx-tab-leader", !!leader);
      }
    }
  };

  const applyZoom = () => {
    const z = effectiveZoom();
    page.style.transformOrigin = "top left";
    page.style.transform = `scale(${z})`;
    // Size the box to the page's real footprint (handles mixed per-section page widths too).
    // A vertical page grows in width and is fixed in height; a horizontal one is the reverse.
    pagebox.style.width = `${Math.round(page.offsetWidth * z)}px`;
    pagebox.style.height = `${Math.round((vertical() ? geometry.heightPx : page.offsetHeight) * z)}px`;
    if (document.activeElement !== zoomLabel) zoomLabel.value = `${Math.round(z * 100)}%`;
    zoomSlider.value = String(Math.round(z * 100));
    updateRulers();
    layoutTabs();
  };
  const setZoom = (z: number | null) => {
    userZoom = z == null ? null : Math.max(0.3, Math.min(2.5, Math.round(z * 100) / 100));
    applyZoom();
    positionCards();
  };

  // --- Page setup -----------------------------------------------------------
  // A dialog to set the document's page size, orientation, margins and column count. It edits the
  // the section the caret is in (or the document geometry when not inside a sub-section).
  const PAGE_SIZES: Record<string, [number, number]> = {
    a4: [794, 1123], a5: [559, 794], a3: [1123, 1587], letter: [816, 1056], legal: [816, 1344], tabloid: [1056, 1632],
  };
  const SIZE_LABELS: Record<string, string> = { a4: "A4", a5: "A5", a3: "A3", letter: "Letter", legal: "Legal", tabloid: "Tabloid" };
  const MARGIN_PRESETS: Record<string, { top: number; right: number; bottom: number; left: number }> = {
    normal: { top: 96, right: 96, bottom: 96, left: 96 },
    narrow: { top: 48, right: 48, bottom: 48, left: 48 },
    moderate: { top: 96, right: 72, bottom: 96, left: 72 },
    wide: { top: 96, right: 192, bottom: 96, left: 192 },
  };
  const near = (a: number, b: number, tol = 2) => Math.abs(a - b) <= tol;
  const matchSize = (g: SecGeom): { key: string; landscape: boolean } => {
    for (const [key, [pw, ph]] of Object.entries(PAGE_SIZES)) {
      if (near(g.w, pw) && near(g.h, ph)) return { key, landscape: false };
      if (near(g.w, ph) && near(g.h, pw)) return { key, landscape: true };
    }
    return { key: "custom", landscape: g.w > g.h };
  };
  const matchMargins = (g: SecGeom): string => {
    for (const [key, p] of Object.entries(MARGIN_PRESETS))
      if (near(g.mt, p.top, 1) && near(g.mr, p.right, 1) && near(g.mb, p.bottom, 1) && near(g.ml, p.left, 1)) return key;
    return "custom";
  };
  const psOverlay = document.createElement("div");
  psOverlay.className = "docxedit-dialog-overlay";
  psOverlay.hidden = true;
  const psPanel = document.createElement("div");
  psPanel.className = "docxedit-dialog docxedit-pagesetup";
  psOverlay.appendChild(psPanel);
  const psTitle = document.createElement("div");
  psTitle.className = "docxedit-dialog-title";
  psTitle.textContent = t("pageSetup");
  const mkSelectRow = (labelText: string, opts: [string, string][]): { row: HTMLElement; sel: HTMLSelectElement } => {
    const row = document.createElement("label");
    row.className = "docxedit-dialog-row docxedit-pagesetup-row";
    const span = document.createElement("span");
    span.className = "docxedit-pagesetup-label";
    span.textContent = labelText;
    const sel = document.createElement("select");
    for (const [v, lbl] of opts) sel.add(new Option(lbl, v));
    row.append(span, sel);
    return { row, sel };
  };
  const { row: sizeRow, sel: sizeSel } = mkSelectRow(t("pageSize"), [...Object.keys(PAGE_SIZES).map((k) => [k, SIZE_LABELS[k]!] as [string, string]), ["custom", t("custom")]]);
  const { row: orientRow, sel: orientSel } = mkSelectRow(t("orientation"), [["portrait", t("portrait")], ["landscape", t("landscape")]]);
  // A labelled cm field (caption above the input) so the meaning stays visible while typing.
  const cmField = (label: string, min: string): { wrap: HTMLElement; input: HTMLInputElement } => {
    const wrap = document.createElement("label");
    wrap.className = "docxedit-pagesetup-field";
    const cap = document.createElement("span");
    cap.className = "docxedit-pagesetup-fieldlabel";
    cap.textContent = label;
    const input = document.createElement("input");
    input.type = "number"; input.min = min; input.step = "0.1"; input.className = "docxedit-dialog-size";
    wrap.append(cap, input);
    return { wrap, input };
  };
  const customRow = document.createElement("div");
  customRow.className = "docxedit-dialog-row docxedit-pagesetup-custom";
  const fW = cmField(t("pageWidthCm"), "1");
  const fH = cmField(t("pageHeightCm"), "1");
  const wIn = fW.input, hIn = fH.input;
  customRow.append(fW.wrap, fH.wrap);
  const { row: marginRow, sel: marginSel } = mkSelectRow(t("margins"), [["normal", t("marginNormal")], ["narrow", t("marginNarrow")], ["moderate", t("marginModerate")], ["wide", t("marginWide")], ["custom", t("custom")]]);
  // Custom margins: four labelled cm fields revealed when the margin preset is "custom".
  const marginCustomRow = document.createElement("div");
  marginCustomRow.className = "docxedit-dialog-row docxedit-pagesetup-custom";
  const fMT = cmField(`${t("edgeTop")} (cm)`, "0");
  const fMR = cmField(`${t("edgeRight")} (cm)`, "0");
  const fMB = cmField(`${t("edgeBottom")} (cm)`, "0");
  const fML = cmField(`${t("edgeLeft")} (cm)`, "0");
  const mtIn = fMT.input, mrIn = fMR.input, mbIn = fMB.input, mlIn = fML.input;
  marginCustomRow.append(fMT.wrap, fMR.wrap, fMB.wrap, fML.wrap);
  const { row: colRow, sel: colSel } = mkSelectRow(t("columns"), [["1", "1"], ["2", "2"], ["3", "3"]]);
  const { row: dirRow, sel: dirSel } = mkSelectRow(t("textDirection"), [["horizontal", t("dirHorizontal")], ["vertical", t("dirVertical")], ["rtl", t("dirRtl")]]);
  // Document-level header/footer variant toggles (apply to the whole document, not just one section).
  const mkCheckRow = (label: string): { row: HTMLElement; input: HTMLInputElement } => {
    const row = document.createElement("label");
    row.className = "docxedit-dialog-row docxedit-pagesetup-check";
    const input = document.createElement("input");
    input.type = "checkbox";
    const span = document.createElement("span");
    span.textContent = label;
    row.append(input, span);
    return { row, input };
  };
  const { row: firstRow, input: firstCheck } = mkCheckRow(t("differentFirstPage"));
  const { row: evenRow, input: evenCheck } = mkCheckRow(t("differentEvenOdd"));
  const syncCustom = () => {
    const customSize = sizeSel.value === "custom";
    customRow.hidden = !customSize;
    orientRow.hidden = customSize; // custom width/height already imply the orientation
    marginCustomRow.hidden = marginSel.value !== "custom";
  };
  sizeSel.addEventListener("change", syncCustom);
  marginSel.addEventListener("change", syncCustom);
  const psCancel = document.createElement("button");
  psCancel.type = "button"; psCancel.className = "docxedit-menu-item"; psCancel.textContent = t("cancel");
  const psApply = document.createElement("button");
  psApply.type = "button"; psApply.className = "docxedit-menu-item docxedit-dialog-primary"; psApply.textContent = t("apply");
  const psActions = document.createElement("div");
  psActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  psActions.append(psCancel, psApply);
  psPanel.append(psTitle, sizeRow, customRow, orientRow, marginRow, marginCustomRow, colRow, ...(caps.verticalText ? [dirRow] : []), firstRow, evenRow, psActions);
  scroll.appendChild(psOverlay);
  const closePageSetup = () => { psOverlay.hidden = true; };
  const openPageSetup = () => {
    const g = readSectionGeom(); // the section the caret is in (document geometry if none)
    const m = matchSize(g);
    sizeSel.value = m.key;
    orientSel.value = m.landscape ? "landscape" : "portrait";
    wIn.value = (g.w / CM).toFixed(2);
    hIn.value = (g.h / CM).toFixed(2);
    marginSel.value = matchMargins(g);
    mtIn.value = (g.mt / CM).toFixed(2);
    mrIn.value = (g.mr / CM).toFixed(2);
    mbIn.value = (g.mb / CM).toFixed(2);
    mlIn.value = (g.ml / CM).toFixed(2);
    colSel.value = String(g.cols && g.cols > 1 ? g.cols : 1);
    dirSel.value = g.vertical ? "vertical" : g.rtl ? "rtl" : "horizontal";
    firstCheck.checked = !!geometry.titlePage; // document-level, not per-section
    evenCheck.checked = !!geometry.evenOdd;
    syncCustom();
    psOverlay.hidden = false;
  };
  const applyPageSetup = () => {
    const cur = readSectionGeom();
    let w = cur.w, h = cur.h;
    if (sizeSel.value === "custom") {
      const cw = parseFloat(wIn.value), ch = parseFloat(hIn.value);
      if (cw > 0) w = Math.round(cw * CM);
      if (ch > 0) h = Math.round(ch * CM);
    } else {
      [w, h] = PAGE_SIZES[sizeSel.value]!;
    }
    // For a preset size the orientation drives the swap; custom typed dims already imply it.
    if (sizeSel.value !== "custom" && (orientSel.value === "landscape") !== w > h) [w, h] = [h, w];
    let mt = cur.mt, mr = cur.mr, mb = cur.mb, ml = cur.ml;
    if (marginSel.value === "custom") {
      const cm = (v: string, fallback: number) => { const n = parseFloat(v); return n >= 0 ? Math.round(n * CM) : fallback; };
      mt = cm(mtIn.value, cur.mt); mr = cm(mrIn.value, cur.mr); mb = cm(mbIn.value, cur.mb); ml = cm(mlIn.value, cur.ml);
    } else {
      const p = MARGIN_PRESETS[marginSel.value]!;
      mt = p.top; mr = p.right; mb = p.bottom; ml = p.left;
    }
    const c = parseInt(colSel.value, 10) || 1;
    const g: SecGeom = { w, h, mt, mr, mb, ml, cols: c > 1 ? c : undefined, colGap: c > 1 ? (cur.colGap ?? 36) : undefined };
    if (caps.verticalText) { g.vertical = dirSel.value === "vertical" || undefined; g.rtl = dirSel.value === "rtl" || undefined; }
    writeSectionGeom(g);
    // Document-level header/footer variants: only act on a change (so existing bands are kept).
    if (!!geometry.titlePage !== firstCheck.checked) toggleHFVariant("first", firstCheck.checked);
    if (!!geometry.evenOdd !== evenCheck.checked) toggleHFVariant("even", evenCheck.checked);
    reflow();
    applyZoom();
    mark();
    closePageSetup();
  };
  psApply.addEventListener("click", applyPageSetup);
  psCancel.addEventListener("click", closePageSetup);
  psOverlay.addEventListener("click", (e) => { if (e.target === psOverlay) closePageSetup(); });
  const pageSetupBtn = document.createElement("button");
  pageSetupBtn.type = "button";
  pageSetupBtn.className = "docxedit-pagesetup-btn";
  pageSetupBtn.title = t("pageSetup");
  pageSetupBtn.setAttribute("aria-label", t("pageSetup"));
  pageSetupBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="3" y="1.5" width="10" height="13" rx="1"/><path d="M5 5h6M5 8h6M5 11h4"/></svg>`;
  pageSetupBtn.addEventListener("click", openPageSetup);

  // Insert section break (only when the format supports section authoring).
  let sectionBreakBtn: HTMLButtonElement | null = null;
  if (caps.sections) {
    sectionBreakBtn = document.createElement("button");
    sectionBreakBtn.type = "button";
    sectionBreakBtn.className = "docxedit-pagesetup-btn docxedit-sectionbreak-btn";
    sectionBreakBtn.title = t("insertSectionBreak");
    sectionBreakBtn.setAttribute("aria-label", t("insertSectionBreak"));
    sectionBreakBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="3" y="1.5" width="10" height="5" rx="1"/><rect x="3" y="9.5" width="10" height="5" rx="1"/><path d="M1 8h2M5 8h2M9 8h2M13 8h2" stroke-dasharray="0"/></svg>`;
    sectionBreakBtn.addEventListener("click", () => {
      insertSectionBreak();
      reflow();
      applyZoom();
      mark();
    });
  }

  const teardown = () => {
    document.removeEventListener("selectionchange", onSelForRuler);
    window.clearTimeout(rulerSyncTimer);
  };
  return { applyZoom, effectiveZoom, setZoom, zoomSlider, zoomLabel, pageSetupBtn, sectionBreakBtn, teardown };
}
