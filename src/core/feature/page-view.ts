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
      // single-section / columns: the caret is in the one flow, so map its vertical position to
      // a page card. A collapsed caret (e.g. in an empty paragraph) can report a zero rect, so
      // fall back to its element's rect.
      const cr = sel!.getRangeAt(0).getBoundingClientRect();
      const cy = cr.top || cr.bottom || el.getBoundingClientRect().top;
      const hit = all.find((c) => { const r = c.getBoundingClientRect(); return cy >= r.top - 2 && cy <= r.bottom + 2; });
      if (hit) return hit;
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

  // --- Page setup -----------------------------------------------------------
  // A dialog to set the document's page size, orientation, margins and column count. It edits the
  // document geometry (= the trailing section); per-section authoring comes later.
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
  const matchSize = (): { key: string; landscape: boolean } => {
    const w = geometry.widthPx, h = geometry.heightPx;
    for (const [key, [pw, ph]] of Object.entries(PAGE_SIZES)) {
      if (near(w, pw) && near(h, ph)) return { key, landscape: false };
      if (near(w, ph) && near(h, pw)) return { key, landscape: true };
    }
    return { key: "custom", landscape: w > h };
  };
  const matchMargins = (): string => {
    const m = geometry.margin;
    for (const [key, p] of Object.entries(MARGIN_PRESETS))
      if (near(m.top, p.top, 1) && near(m.right, p.right, 1) && near(m.bottom, p.bottom, 1) && near(m.left, p.left, 1)) return key;
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
  const syncCustom = () => {
    customRow.hidden = sizeSel.value !== "custom";
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
  psPanel.append(psTitle, sizeRow, customRow, orientRow, marginRow, marginCustomRow, colRow, psActions);
  scroll.appendChild(psOverlay);
  const closePageSetup = () => { psOverlay.hidden = true; };
  const openPageSetup = () => {
    const m = matchSize();
    sizeSel.value = m.key;
    orientSel.value = m.landscape ? "landscape" : "portrait";
    wIn.value = (geometry.widthPx / CM).toFixed(2);
    hIn.value = (geometry.heightPx / CM).toFixed(2);
    marginSel.value = matchMargins();
    mtIn.value = (geometry.margin.top / CM).toFixed(2);
    mrIn.value = (geometry.margin.right / CM).toFixed(2);
    mbIn.value = (geometry.margin.bottom / CM).toFixed(2);
    mlIn.value = (geometry.margin.left / CM).toFixed(2);
    colSel.value = String(geometry.columns && geometry.columns > 1 ? geometry.columns : 1);
    syncCustom();
    psOverlay.hidden = false;
  };
  const applyPageSetup = () => {
    let w = geometry.widthPx, h = geometry.heightPx;
    if (sizeSel.value === "custom") {
      const cw = parseFloat(wIn.value), ch = parseFloat(hIn.value);
      if (cw > 0) w = Math.round(cw * CM);
      if (ch > 0) h = Math.round(ch * CM);
    } else {
      [w, h] = PAGE_SIZES[sizeSel.value]!;
    }
    if ((orientSel.value === "landscape") !== w > h) [w, h] = [h, w];
    geometry.widthPx = w;
    geometry.heightPx = h;
    if (marginSel.value === "custom") {
      const cm = (v: string, fallback: number) => { const n = parseFloat(v); return n >= 0 ? Math.round(n * CM) : fallback; };
      geometry.margin = {
        top: cm(mtIn.value, geometry.margin.top), right: cm(mrIn.value, geometry.margin.right),
        bottom: cm(mbIn.value, geometry.margin.bottom), left: cm(mlIn.value, geometry.margin.left),
      };
    } else {
      geometry.margin = { ...MARGIN_PRESETS[marginSel.value]! };
    }
    const c = parseInt(colSel.value, 10) || 1;
    geometry.columns = c > 1 ? c : undefined;
    if (geometry.columns && !geometry.columnGapPx) geometry.columnGapPx = 36;
    geometryDirty = true;
    applyGeometry();
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

  const teardown = () => {
    document.removeEventListener("selectionchange", onSelForRuler);
    window.clearTimeout(rulerSyncTimer);
  };
  return { applyZoom, effectiveZoom, setZoom, zoomSlider, zoomLabel, pageSetupBtn, isGeometryDirty: () => geometryDirty, teardown };
}
