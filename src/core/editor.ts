// Shared rich-document editor engine. Renders the editable surface, toolbar, comments
// panel, track changes, image/page chrome and passthrough; the format-specific parse,
// serialize and comment markers come from an Adapter (see core/types.ts). docx is the
// reference adapter; odt reuses this same engine.

import { t } from "./i18n";
import { defaultPageGeometry, paginate } from "./page";
import type { Adapter, EditorOptions, RichEditor, RichDoc } from "./types";
import { setupComments } from "./feature/comments";
import { setupImages } from "./feature/images";
import { setupPageView } from "./feature/page-view";
import { setupToolbar } from "./feature/toolbar";
import { setupTableEdit } from "./feature/table-edit";
import "../adapters/docx/docxedit.css";

export function createRichEditor(container: HTMLElement, adapter: Adapter, options: EditorOptions = {}): RichEditor {
  const original = adapter.original;
  const caps = adapter.capabilities;
  let dirty = false;

  const wrap = document.createElement("div");
  wrap.className = "docxedit-wrap";
  const toolbar = document.createElement("div");
  toolbar.className = "docxedit-toolbar";
  toolbar.setAttribute("role", "toolbar");
  toolbar.setAttribute("aria-label", t("toolbar"));
  const scroll = document.createElement("div");
  scroll.className = "docxedit-scroll";
  const page = document.createElement("div");
  page.className = "docxedit-page";
  const doc = document.createElement("div");
  doc.className = "docxedit-doc";
  doc.contentEditable = "true";
  doc.spellcheck = false;
  doc.setAttribute("role", "textbox");
  doc.setAttribute("aria-multiline", "true");
  doc.setAttribute("aria-label", t("documentText"));

  let parts: RichDoc = { body: "<p><br></p>", header: "", footer: "", comments: [] };
  try {
    parts = adapter.read();
  } catch (e) {
    console.warn("richdoc: failed to parse document", e);
  }
  doc.innerHTML = parts.body || "<p><br></p>";

  // Render the document in its own embedded typefaces, if the adapter supplied any.
  const fontUrls: string[] = parts.fontUrls ?? [];
  if (parts.fontCss) {
    const fs = document.createElement("style");
    fs.textContent = parts.fontCss;
    wrap.appendChild(fs);
  }
  if (parts.defaultFont) page.style.setProperty("--docxedit-doc-font", `"${parts.defaultFont.replace(/"/g, "")}"`);

  // Page geometry: render at the document's real size and margins, or the default size
  // (A4 unless options override) when the file declares none. Height is unused until
  // pagination (Phase 1); width and margins apply now.
  const geometry = parts.page ?? defaultPageGeometry(options.defaultPageSize ?? "a4");
  const applyGeometry = () => {
    page.style.setProperty("--rdoc-page-width", `${geometry.widthPx}px`);
    page.style.setProperty("--rdoc-page-height", `${geometry.heightPx}px`);
    page.style.setProperty("--rdoc-margin-top", `${geometry.margin.top}px`);
    page.style.setProperty("--rdoc-margin-right", `${geometry.margin.right}px`);
    page.style.setProperty("--rdoc-margin-bottom", `${geometry.margin.bottom}px`);
    page.style.setProperty("--rdoc-margin-left", `${geometry.margin.left}px`);
  };
  applyGeometry();
  // Vertical (Japanese tategaki): fixed-height page, columns top-to-bottom advancing right to
  // left, so the page grows along x and page cards advance right-to-left (see repaginateVertical).
  // Horizontal RTL (Arabic/Hebrew) is just direction:rtl and keeps the normal layout.
  const isVertical = caps.verticalText && !!geometry.vertical;
  if (isVertical) page.classList.add("is-vertical");
  else if (caps.verticalText && geometry.rtl) page.classList.add("is-rtl");

  const band = (cls: string, label: string, html: string): HTMLElement | null => {
    if (!html) return null;
    const el = document.createElement("div");
    el.className = cls;
    el.contentEditable = "true";
    el.spellcheck = false;
    el.setAttribute("role", "textbox");
    el.setAttribute("aria-multiline", "true");
    el.setAttribute("aria-label", label);
    el.innerHTML = html;
    return el;
  };
  let header = band("docxedit-header", t("header"), parts.header);
  let footer = band("docxedit-footer", t("footer"), parts.footer);
  // A band counts as empty when it has no text and no image, so an abandoned new one
  // (created by a double-click but never typed in) is dropped instead of being saved.
  const isBandEmpty = (el: HTMLElement): boolean => !el.textContent?.trim() && !el.querySelector("img");

  // Paginated view: one continuous editable body (doc) on top of a layer of page-card
  // decorations, with inert spacer gaps inserted at page boundaries. Pageless view keeps
  // the body and header/footer stacked in one card (the previous behaviour).
  const paginated = options.paginated ?? true;
  const pagelayer = document.createElement("div"); // page cards, behind the body
  pagelayer.className = "docxedit-pagelayer";
  pagelayer.setAttribute("aria-hidden", "true");
  const hflayer = document.createElement("div"); // header/footer clones, above the body (clickable)
  hflayer.className = "docxedit-hflayer";
  // Off-screen holder so header/footer can be measured (and kept as the save source)
  // without showing as in-flow bands in paginated mode.
  const measure = document.createElement("div");
  measure.className = "docxedit-measure";
  if (paginated) {
    page.classList.add("is-paginated");
    page.append(pagelayer, doc, hflayer);
    if (header) measure.appendChild(header);
    if (footer) measure.appendChild(footer);
    page.appendChild(measure);
  } else {
    if (header) page.appendChild(header);
    page.appendChild(doc);
    if (footer) page.appendChild(footer);
  }

  // Keep the page centred, with the comments column in the right margin (Google-Docs
  // style): an empty left spacer balances the right comments area so the page stays centred.
  const canvas = document.createElement("div");
  canvas.className = "docxedit-canvas";
  const leftSpacer = document.createElement("div");
  leftSpacer.className = "docxedit-margin";
  const rightArea = document.createElement("div");
  rightArea.className = "docxedit-margin";
  const cmtPanel = document.createElement("div");
  cmtPanel.className = "docxedit-comments";
  rightArea.appendChild(cmtPanel);
  // The page is scaled inside a box sized to the scaled footprint, so zoom is a visual
  // transform that leaves the body's layout (which pagination measures) unscaled.
  const pagebox = document.createElement("div");
  pagebox.className = "docxedit-pagebox";
  pagebox.appendChild(page);

  // The editable regions (body + header/footer). Toolbar actions target whichever last
  // had focus, so formatting works inside the header and footer too.
  const regions = [doc, header, footer].filter(Boolean) as HTMLElement[];
  let activeEl: HTMLElement = doc;

  // Comment id allocation, the side panel, edit bookkeeping and click-to-open live in the
  // comments module; mark() is defined here so it (and the page view) can be handed it.
  const mark = () => {
    dirty = true;
    options.onChange?.();
    scheduleReflow(); // content changed: re-paginate (debounced)
  };
  const { addThreadCard, positionCards, setActiveComment, allocId, freshParaId, getEdits } =
    setupComments({ wrap, cmtPanel, pagebox, options, caps, mark });

  // Page view: rulers (margin handles) + zoom + the centred canvas around the page box.
  const { applyZoom, effectiveZoom, zoomSlider, zoomLabel, isGeometryDirty } = setupPageView({
    page, pagebox, canvas, leftSpacer, rightArea, scroll, geometry, options,
    vertical: !!(geometry.vertical && caps.verticalText),
    applyGeometry, mark, positionCards, reflow: () => reflow(), scheduleReflow: () => scheduleReflow(),
  });
  wrap.append(toolbar, scroll);
  container.appendChild(wrap);

  // --- Pagination -----------------------------------------------------------
  // Measure top-level block heights, compute page breaks, then draw page cards behind the
  // body and insert inert spacer gaps so the flow lands at each page's content box. The
  // body stays one contenteditable; spacers are stripped before saving (see cleanBody).
  const PAGE_GAP = 24;
  // The header/footer clones are editable in place (so the browser places the caret at the
  // click point natively). Edits sync into the canonical band (the save source) and, on
  // blur, the pages re-clone its content. editingBand suspends reflow during an edit.
  let editingBand: HTMLElement | null = null;

  // Vertical (tategaki) pagination: the fill axis is x (right to left). Feed paginate() each
  // block's inline-size (its width in vertical-rl) and the page's content width; page cards
  // advance right-to-left and spacer gaps are widths. The doc is an inline-block, right-aligned,
  // fixed at the page height and growing leftward (see the .is-vertical.is-paginated CSS).
  const manualBreaks = (kids: HTMLElement[]): Set<number> => {
    const set = new Set<number>();
    const isMarker = (el: Element) => el.classList.contains("docx-pagebreak") && el.getAttribute("data-docx-pagebreak") === "manual";
    kids.forEach((k, i) => {
      if (isMarker(k)) { if (i + 1 < kids.length) set.add(i + 1); }
      else if (i > 0 && k.querySelector('.docx-pagebreak[data-docx-pagebreak="manual"]')) set.add(i);
    });
    return set;
  };
  const repaginateVertical = () => {
    for (const s of Array.from(doc.querySelectorAll(":scope > .docxedit-pagespacer"))) s.remove();
    pagelayer.replaceChildren();
    hflayer.replaceChildren(); // header/footer not paginated in vertical v1
    const { left, right } = geometry.margin;
    const contentExtent = geometry.widthPx - left - right; // usable page width along the fill axis
    const pageStep = geometry.widthPx + PAGE_GAP;
    doc.style.height = `${geometry.heightPx}px`;
    doc.style.padding = `${geometry.margin.top}px ${right}px ${geometry.margin.bottom}px ${left}px`;
    for (const el of Array.from(doc.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");

    const kids = Array.from(doc.children).filter((c) => !c.classList.contains("docxedit-pagespacer")) as HTMLElement[];
    // progression (right-to-left) size: right edge of block i minus right edge of block i+1.
    const rights = kids.map((k) => k.offsetLeft + k.offsetWidth);
    const sizes = kids.map((k, i) => (i < kids.length - 1 ? rights[i]! - rights[i + 1]! : k.offsetWidth));

    const { spacerBefore, cardCount } = paginate(sizes, { pageStep, contentHeight: contentExtent }, manualBreaks(kids));
    for (const [idx, w] of spacerBefore) {
      const sp = document.createElement("div");
      sp.className = "docxedit-pagespacer";
      sp.contentEditable = "false";
      sp.setAttribute("aria-hidden", "true");
      sp.style.width = `${w}px`;
      doc.insertBefore(sp, kids[idx]!);
      kids[idx]!.classList.add("docxedit-pagetop");
    }
    for (let p = 0; p < cardCount; p++) {
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.left = "auto";
      card.style.right = `${p * pageStep}px`;
      card.style.top = "0";
      card.style.width = `${geometry.widthPx}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
    }
    page.style.width = `${cardCount * pageStep - PAGE_GAP}px`;
    page.style.minHeight = `${geometry.heightPx}px`;
  };

  const repaginate = () => {
    if (!paginated || editingBand) return;
    if (isVertical) return repaginateVertical();
    for (const s of Array.from(doc.querySelectorAll(":scope > .docxedit-pagespacer"))) s.remove();
    pagelayer.replaceChildren();
    hflayer.replaceChildren();

    // measure header/footer at full page width (their own padding provides the margins)
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = header ? header.offsetHeight : 0;
    const footerH = footer ? footer.offsetHeight : 0;
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    const contentHeight = geometry.heightPx - contentTop - contentBottomInset;
    const pageStep = geometry.heightPx + PAGE_GAP;

    // place the body inside each page's content box (overrides the CSS-var padding)
    doc.style.padding = `${contentTop}px ${geometry.margin.right}px ${contentBottomInset}px ${geometry.margin.left}px`;

    // clear the previous page-top markers so they don't skew this measurement
    for (const el of Array.from(doc.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");

    // delta of offsetTop captures each block's height plus its collapsed inter-block margin
    const kids = Array.from(doc.children).filter((c) => !c.classList.contains("docxedit-pagespacer")) as HTMLElement[];
    const tops = kids.map((k) => k.offsetTop);
    const heights = kids.map((k, i) => (i < kids.length - 1 ? tops[i + 1]! - tops[i]! : k.offsetHeight));

    // honor explicit page breaks. A manual break renders either as its own marker element
    // (break before the next block) or inside a block (break before that block).
    const forceBreakBefore = new Set<number>();
    const isManualMarker = (el: Element) =>
      el.classList.contains("docx-pagebreak") && el.getAttribute("data-docx-pagebreak") === "manual";
    kids.forEach((k, i) => {
      if (isManualMarker(k)) {
        if (i + 1 < kids.length) forceBreakBefore.add(i + 1);
      } else if (i > 0 && k.querySelector('.docx-pagebreak[data-docx-pagebreak="manual"]')) {
        forceBreakBefore.add(i);
      }
    });

    const { spacerBefore, cardCount } = paginate(heights, { pageStep, contentHeight }, forceBreakBefore);

    for (const [idx, h] of spacerBefore) {
      const sp = document.createElement("div");
      sp.className = "docxedit-pagespacer";
      sp.contentEditable = "false";
      sp.setAttribute("aria-hidden", "true");
      sp.style.height = `${h}px`;
      doc.insertBefore(sp, kids[idx]!);
      // drop the page-starting block's top margin so it aligns to the page content top
      kids[idx]!.classList.add("docxedit-pagetop");
    }

    // Editable clone of the canonical band: clicking it places the caret natively at the
    // click point; edits sync into the canonical (save source); on blur the pages re-clone.
    const mkClone = (src: HTMLElement, topPx: number): HTMLElement => {
      const c = src.cloneNode(true) as HTMLElement;
      c.removeAttribute("role");
      c.removeAttribute("aria-label");
      c.classList.add("docxedit-hf-clone");
      c.contentEditable = "true";
      c.spellcheck = false;
      c.style.cssText = `top:${topPx}px;left:0;width:100%`;
      c.title = t("editHeaderFooter");
      c.addEventListener("focus", () => {
        editingBand = src;
        activeEl = c;
        c.classList.add("is-editing");
      });
      c.addEventListener("input", () => {
        src.innerHTML = c.innerHTML; // keep the canonical (save source) current
        mark();
      });
      c.addEventListener("blur", () => {
        c.classList.remove("is-editing");
        editingBand = null;
        // A band created by double-click but left empty is dropped, not saved. Defer the
        // check so moving between page-clones of the same band does not count as leaving.
        if (src.dataset.pending === "1") {
          setTimeout(() => {
            if ((document.activeElement as HTMLElement | null)?.classList.contains("docxedit-hf-clone")) return;
            if (isBandEmpty(src)) {
              if (src === header) header = null;
              else if (src === footer) footer = null;
              src.remove();
              activeEl = doc;
            } else {
              delete src.dataset.pending; // it has content now: keep it
            }
            reflow();
          }, 0);
          return;
        }
        reflow();
      });
      return c;
    };

    for (let p = 0; p < cardCount; p++) {
      const base = p * pageStep;
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.top = `${base}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
      if (header) hflayer.appendChild(mkClone(header, base + geometry.margin.top));
      if (footer) hflayer.appendChild(mkClone(footer, base + geometry.heightPx - contentBottomInset));
    }

    page.style.minHeight = `${cardCount * pageStep - PAGE_GAP}px`;
  };

  // Body HTML for saving: the live doc minus pagination artifacts (inert spacers and the
  // transient page-top class the engine adds for alignment).
  const cleanBody = (): string => {
    if (!doc.querySelector(".docxedit-pagespacer, .docxedit-pagetop")) return doc.innerHTML;
    const tmp = doc.cloneNode(true) as HTMLElement;
    for (const s of Array.from(tmp.querySelectorAll(".docxedit-pagespacer"))) s.remove();
    for (const el of Array.from(tmp.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");
    return tmp.innerHTML;
  };

  let afterReflow = () => {}; // assigned once the toolbar exists (toggles change buttons)
  const reflow = () => {
    if (editingBand) return; // don't yank the band currently being edited
    repaginate();
    applyZoom();
    positionCards();
    afterReflow();
  };
  let reflowTimer = 0;
  const scheduleReflow = () => {
    window.clearTimeout(reflowTimer);
    reflowTimer = window.setTimeout(reflow, 150);
  };

  // --- Create header/footer on demand --------------------------------------
  // Double-clicking a page's top (or bottom) margin, where none exists yet, starts an
  // empty header (or footer) and drops the caret in it. It is only persisted if typed in
  // (see the clone's blur handler), so an abandoned one never gets added to the file.
  const createBand = (kind: "header" | "footer") => {
    const el = document.createElement("div");
    el.className = kind === "header" ? "docxedit-header" : "docxedit-footer";
    el.contentEditable = "true";
    el.spellcheck = false;
    el.setAttribute("role", "textbox");
    el.setAttribute("aria-multiline", "true");
    el.setAttribute("aria-label", kind === "header" ? t("header") : t("footer"));
    el.innerHTML = "<p><br></p>";
    el.dataset.pending = "1"; // unsaved until typed in
    measure.appendChild(el);
    if (kind === "header") header = el;
    else footer = el;
    reflow(); // clones the new band into the page layer
    const sel = `.docxedit-hf-clone.docxedit-${kind}`;
    (hflayer.querySelector(sel) as HTMLElement | null)?.focus();
  };
  if (paginated && caps.headerFooter && !isVertical) {
    page.addEventListener("dblclick", (e) => {
      if ((e.target as HTMLElement).closest(".docxedit-hf-clone")) return; // editing an existing band
      if (header && footer) return;
      const z = effectiveZoom();
      const within = ((e.clientY - page.getBoundingClientRect().top) / z) % (geometry.heightPx + PAGE_GAP);
      if (within > geometry.heightPx) return; // in the gap between two pages
      if (!header && within < geometry.margin.top) createBand("header");
      else if (!footer && within > geometry.heightPx - geometry.margin.bottom) createBand("footer");
    });
  }

  for (const thread of parts.comments) addThreadCard(thread);
  // Lay out pages + comment cards now and once the layout settles (rAF is throttled in
  // background tabs, so also use a timeout), and again whenever heights can shift: after
  // fonts and images load, on container width change, and (debounced) on every edit.
  reflow();
  requestAnimationFrame(reflow);
  setTimeout(reflow, 150);
  // Vertical reads right-to-left, so start scrolled to the rightmost page (page 1).
  if (isVertical) {
    const scrollRight = () => { scroll.scrollLeft = scroll.scrollWidth; };
    requestAnimationFrame(scrollRight);
    setTimeout(scrollRight, 200);
  }
  if (document.fonts?.ready) document.fonts.ready.then(reflow).catch(() => {});
  for (const img of Array.from(doc.querySelectorAll("img"))) img.addEventListener("load", reflow);
  let lastWidth = scroll.clientWidth;
  const repositionObserver = new ResizeObserver(() => {
    const w = scroll.clientWidth;
    if (w !== lastWidth) {
      lastWidth = w;
      reflow();
    } else {
      positionCards();
    }
  });
  repositionObserver.observe(scroll);
  for (const r of regions) r.addEventListener("input", scheduleReflow);

  // Image select/resize/delete and insert-from-file live in the images feature module.
  const { insertImage } = setupImages({ wrap, scroll, regions, mark, getActiveEl: () => activeEl });
  // Emit inline CSS (text-align, font-weight, ...) the serializer reads back, not legacy tags.
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch {
    /* not supported; legacy tags still round-trip */
  }

  for (const r of regions) {
    r.addEventListener("input", mark);
    r.addEventListener("focusin", (e) => {
      // A table cell is its own editing host nested in the region; target it directly so
      // toolbar commands act inside the cell instead of stealing focus back to the region.
      const cell = (e.target as HTMLElement | null)?.closest?.(".docx-cell") as HTMLElement | null;
      activeEl = cell && r.contains(cell) ? cell : r;
    });
  }
  // The toolbar (formatting controls + track-changes wiring + the overflow row) lives in
  // its own module; it returns the change-button toggle to hook into the reflow cycle.
  const { updateChangeButtons, teardown: teardownToolbar } = setupToolbar({
    toolbar, wrap, doc, regions, caps, options, parts, adapter, getActiveEl: () => activeEl, mark,
    positionCards, addThreadCard, setActiveComment, allocId, freshParaId, insertImage, zoomSlider, zoomLabel,
  });
  afterReflow = updateChangeButtons;
  updateChangeButtons();

  // Table editing toolbar (shown while the caret is in a table cell).
  const tableEdit = caps.tables ? setupTableEdit({ wrap, scroll, mark, scheduleReflow }) : null;

  return {
    isDirty() {
      return dirty;
    },
    async getBytes() {
      if (!dirty) return original.slice();
      const editedParts: { path: string; html: string }[] = [];
      // Use the band's own part path when it came from the file; a band created in-editor
      // has none, so fall back to the "header"/"footer" sentinel the adapters create from.
      if (header && !isBandEmpty(header)) editedParts.push({ path: parts.headerPath ?? "header", html: header.innerHTML });
      if (footer && !isBandEmpty(footer)) editedParts.push({ path: parts.footerPath ?? "footer", html: footer.innerHTML });
      return adapter.write(cleanBody(), editedParts, getEdits(), isGeometryDirty() ? geometry : undefined);
    },
    destroy() {
      for (const u of fontUrls) URL.revokeObjectURL(u);
      window.clearTimeout(reflowTimer);
      repositionObserver.disconnect();
      teardownToolbar();
      tableEdit?.teardown();
      wrap.remove();
    },
  };
}
