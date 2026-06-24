// Shared rich-document editor engine. Renders the editable surface, toolbar, comments
// panel, track changes, image/page chrome and passthrough; the format-specific parse,
// serialize and comment markers come from an Adapter (see core/types.ts). docx is the
// reference adapter; odt reuses this same engine.

import { t } from "./i18n";
import { defaultPageGeometry, paginate } from "./page";
import { bytesToBase64 } from "./util";
import type { Adapter, EditorOptions, RichEditor, RichDoc } from "./types";
import { setupComments } from "./feature/comments";
import { setupPageView } from "./feature/page-view";
import { setupTrackChanges } from "./feature/track-changes";
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
    page.style.setProperty("--rdoc-margin-top", `${geometry.margin.top}px`);
    page.style.setProperty("--rdoc-margin-right", `${geometry.margin.right}px`);
    page.style.setProperty("--rdoc-margin-bottom", `${geometry.margin.bottom}px`);
    page.style.setProperty("--rdoc-margin-left", `${geometry.margin.left}px`);
  };
  applyGeometry();

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

  const repaginate = () => {
    if (!paginated || editingBand) return;
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
  if (paginated && caps.headerFooter) {
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

  // Image select: click an image to select it, drag the corner handle to resize, and use
  // the delete button (top-right) or the Delete/Backspace key to remove it.
  let selImg: HTMLImageElement | null = null;
  const imgHandle = document.createElement("div");
  imgHandle.className = "docxedit-img-handle is-hidden";
  const imgDel = document.createElement("button");
  imgDel.type = "button";
  imgDel.className = "docxedit-img-del is-hidden";
  imgDel.textContent = "✕";
  imgDel.title = t("deleteImage");
  wrap.append(imgHandle, imgDel);
  const placeHandle = () => {
    if (!selImg || !wrap.contains(selImg)) {
      imgHandle.classList.add("is-hidden");
      imgDel.classList.add("is-hidden");
      return;
    }
    const ir = selImg.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    // positions are runtime-computed; they must follow the selected image
    imgHandle.style.left = `${ir.right - wr.left - 6}px`;
    imgHandle.style.top = `${ir.bottom - wr.top - 6}px`;
    imgHandle.classList.remove("is-hidden");
    imgDel.style.left = `${ir.right - wr.left - 11}px`;
    imgDel.style.top = `${ir.top - wr.top - 11}px`;
    imgDel.classList.remove("is-hidden");
  };
  const selectImg = (img: HTMLImageElement | null) => {
    if (selImg) selImg.classList.remove("sel");
    selImg = img;
    if (selImg) selImg.classList.add("sel");
    placeHandle();
  };
  const deleteSelImg = () => {
    if (!selImg) return;
    selImg.remove();
    selImg = null;
    placeHandle();
    mark();
  };
  imgDel.addEventListener("mousedown", (e) => e.preventDefault());
  imgDel.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteSelImg();
  });
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement) === imgDel) return;
    const img = (e.target as HTMLElement).closest?.("img") as HTMLImageElement | null;
    selectImg(img && wrap.contains(img) ? img : null);
  });
  for (const r of regions)
    r.addEventListener("keydown", (e) => {
      if (selImg && (e.key === "Delete" || e.key === "Backspace")) {
        e.preventDefault();
        deleteSelImg();
      }
    });
  scroll.addEventListener("scroll", placeHandle);
  // Persist a resize: new images use the width/height attributes; existing ones carry the
  // original drawing in data-docx-xml, so update its extent (EMU) to the new size.
  const persistImgSize = (img: HTMLImageElement) => {
    const xml = img.getAttribute("data-docx-xml");
    const w = Number(img.getAttribute("width")) || 0;
    const h = Number(img.getAttribute("height")) || 0;
    if (!xml || !w) return;
    try {
      const frag = new DOMParser().parseFromString(xml, "application/xml");
      for (const tag of ["wp:extent", "a:ext"]) {
        const el = frag.getElementsByTagName(tag)[0];
        if (el) {
          el.setAttribute("cx", String(Math.round(w * 9525)));
          el.setAttribute("cy", String(Math.round(h * 9525)));
        }
      }
      img.setAttribute("data-docx-xml", new XMLSerializer().serializeToString(frag.documentElement!));
    } catch {
      /* leave as-is */
    }
  };
  imgHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    const img = selImg;
    if (!img) return;
    const startX = e.clientX;
    const rect = img.getBoundingClientRect();
    const startW = rect.width;
    const ratio = rect.height / startW || 1;
    const onMove = (ev: MouseEvent) => {
      const w = Math.max(16, Math.round(startW + (ev.clientX - startX)));
      img.setAttribute("width", String(w));
      img.setAttribute("height", String(Math.round(w * ratio)));
      placeHandle();
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      persistImgSize(img);
      mark();
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
  // Emit inline CSS (text-align, font-weight, ...) the serializer reads back, not legacy tags.
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch {
    /* not supported; legacy tags still round-trip */
  }

  for (const r of regions) {
    r.addEventListener("input", mark);
    r.addEventListener("focusin", () => {
      activeEl = r;
    });
  }
  const exec = (cmd: string, val?: string) => {
    activeEl.focus();
    document.execCommand(cmd, false, val);
    mark();
  };
  // Wrap the current selection in a span carrying one CSS property (for font size, which
  // has no execCommand equivalent in CSS mode). No-op on a collapsed selection.
  const styleSel = (prop: string, val: string) => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (range.collapsed) return;
    const span = document.createElement("span");
    (span.style as unknown as Record<string, string>)[prop] = val;
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
    } catch {
      return;
    }
    sel.removeAllRanges();
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.addRange(r2);
    mark();
  };
  const btn = (label: string, title: string, fn: () => void, cls = "") => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    if (cls) b.className = cls;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", fn);
    return b;
  };
  const sep = () => {
    const s = document.createElement("span");
    s.className = "sep";
    return s;
  };
  const alignIcon = (rows: [number, number][]): string =>
    `<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">${rows
      .map(([x, w], k) => `<rect x="${x}" y="${3 + k * 4}" width="${w}" height="1.6" rx=".6"/>`)
      .join("")}</svg>`;
  const iconBtn = (svg: string, title: string, fn: () => void) => {
    const b = btn("", title, fn);
    b.innerHTML = svg;
    return b;
  };
  const bulletIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<circle cx="2.3" cy="4" r="1.3"/><circle cx="2.3" cy="11" r="1.3"/>' +
    '<rect x="6" y="3.2" width="9" height="1.6" rx=".6"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6"/></svg>';
  const numberIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">' +
    '<text x="0.3" y="6.2" font-size="6" font-family="sans-serif" fill="currentColor">1</text>' +
    '<text x="0.3" y="13.4" font-size="6" font-family="sans-serif" fill="currentColor">2</text>' +
    '<rect x="6" y="3.2" width="9" height="1.6" rx=".6" fill="currentColor"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6" fill="currentColor"/></svg>';
  const linkIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true">' +
    '<path d="M6.6 9.4l2.8-2.8"/><path d="M7.2 4.6l1-1a2.4 2.4 0 0 1 3.4 3.4l-1 1"/><path d="M8.8 11.4l-1 1a2.4 2.4 0 0 1-3.4-3.4l1-1"/></svg>';

  const block = document.createElement("select");
  block.title = t("paragraphStyle");
  block.setAttribute("aria-label", t("paragraphStyle"));
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(t(key), v));
  }
  block.addEventListener("mousedown", () => activeEl.focus());
  block.addEventListener("change", () => exec("formatBlock", block.value));

  // A select whose first option is a non-selectable title; firing fn(value) on change.
  const pickerSelect = (title: string, opts: [string, string][], fn: (v: string) => void): HTMLSelectElement => {
    const s = document.createElement("select");
    s.title = title;
    s.setAttribute("aria-label", title);
    const head = new Option(title, "");
    head.disabled = true;
    head.selected = true;
    s.add(head);
    for (const [v, label] of opts) s.add(new Option(label, v));
    s.addEventListener("mousedown", () => activeEl.focus());
    s.addEventListener("change", () => {
      if (s.value) fn(s.value); // keep the chosen value shown (do not reset to the placeholder)
    });
    return s;
  };

  // Text colour: a native colour input that applies w:color via foreColor (CSS mode).
  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.value = "#000000";
  colorInput.title = t("textColor");
  colorInput.setAttribute("aria-label", t("textColor"));
  colorInput.className = "docxedit-color";
  colorInput.addEventListener("mousedown", () => activeEl.focus());
  colorInput.addEventListener("input", () => {
    beginFormatChange();
    exec("foreColor", colorInput.value);
  });

  // Background colour: a free colour picker (maps to w:highlight when it matches a named
  // highlight exactly, otherwise to arbitrary w:shd shading). A button clears it.
  const bgWrap = document.createElement("span");
  bgWrap.className = "docxedit-bg";
  const bgInput = document.createElement("input");
  bgInput.type = "color";
  bgInput.value = "#ffff00";
  bgInput.title = t("highlight");
  bgInput.setAttribute("aria-label", t("highlight"));
  bgInput.className = "docxedit-color";
  bgInput.addEventListener("mousedown", () => activeEl.focus());
  bgInput.addEventListener("input", () => {
    beginFormatChange();
    exec("hiliteColor", bgInput.value);
  });
  const bgClear = btn("⌫", t("none"), () => exec("hiliteColor", "transparent"), "docxedit-bg-clear");
  bgWrap.append(bgInput, bgClear);

  const FONTS = ["Arial", "Calibri", "Century", "Courier New", "Georgia", "Times New Roman", "Verdana"];
  const fontSel = pickerSelect(t("font"), FONTS.map((f) => [f, f] as [string, string]), (v) => {
    beginFormatChange();
    exec("fontName", v);
  });

  const SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48"];
  const sizeSel = pickerSelect(t("size"), SIZES.map((s) => [s, s] as [string, string]), (v) => {
    beginFormatChange();
    styleSel("fontSize", `${v}pt`);
  });

  // A new (empty) document gets a default font + size, shown in the pickers and applied so
  // typing starts in them. Existing documents keep their own fonts; the pickers stay blank.
  if (caps.fontControls && !doc.textContent?.trim()) {
    const defFont = parts.defaultFont && FONTS.includes(parts.defaultFont) ? parts.defaultFont : "Arial";
    const defSize = "11";
    fontSel.value = defFont;
    sizeSel.value = defSize;
    doc.style.fontFamily = `'${defFont}', sans-serif`;
    doc.style.fontSize = `${defSize}pt`;
  }

  const insertPageBreak = () => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const el = document.createElement("span");
    el.className = "docx-pagebreak";
    el.contentEditable = "false";
    el.setAttribute("data-docx-pagebreak", "manual");
    el.setAttribute("data-label", t("pageBreak"));
    range.insertNode(el);
    range.setStartAfter(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
  };
  const pbIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="3" y="1.5" width="10" height="4" rx=".5"/><rect x="3" y="10.5" width="10" height="4" rx=".5"/>' +
    '<line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="2 1.6"/></svg>';

  // Insert an image: read a file, show it via a data URL, and let the serializer embed it.
  const insertImage = () => {
    const sel = window.getSelection();
    const savedRange = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp";
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const buf = new Uint8Array(await file.arrayBuffer());
      const dataUrl = `data:${file.type};base64,${bytesToBase64(buf)}`;
      const probe = new Image();
      probe.onload = () => {
        const maxW = 600;
        let w = probe.naturalWidth || 200;
        let h = probe.naturalHeight || 200;
        if (w > maxW) {
          h = Math.round((h * maxW) / w);
          w = maxW;
        }
        const img = document.createElement("img");
        img.src = dataUrl;
        img.setAttribute("width", String(w));
        img.setAttribute("height", String(h));
        const range = savedRange ?? (() => {
          const r = document.createRange();
          r.selectNodeContents(activeEl);
          r.collapse(false);
          return r;
        })();
        range.collapse(false);
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
        const s2 = window.getSelection();
        if (s2) {
          s2.removeAllRanges();
          s2.addRange(range);
        }
        mark();
      };
      probe.src = dataUrl;
    });
    fileInput.click();
  };
  const imgIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
    '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

  // Add a comment over the current selection: wrap it in comment-range markers and a
  // reference marker that carries the text, so the serializer can build comments.xml.
  const addComment = () => {
    activeEl.focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed) {
      const tip = document.createElement("div");
      tip.className = "docxedit-cmt-pop docxedit-cmt-tip";
      tip.textContent = t("commentSelect");
      wrap.appendChild(tip);
      setTimeout(() => tip.remove(), 1800);
      return;
    }
    const text = prompt(t("commentPrompt"));
    if (!text) return;
    const id = allocId();
    const author = options.author || "Author";
    const date = options.now || new Date().toISOString();
    const paraId = freshParaId();
    const { start, end, ref } = adapter.newCommentMarkers({ id, author, date, text, paraId });
    const range = sel.getRangeAt(0);
    const visual = document.createElement("span");
    visual.className = "docx-comment";
    visual.setAttribute("data-comment-id", id);
    visual.appendChild(range.extractContents());
    range.insertNode(visual);
    const parent = visual.parentNode;
    if (parent) {
      parent.insertBefore(start, visual);
      parent.insertBefore(end, visual.nextSibling);
      parent.insertBefore(ref, end.nextSibling);
    }
    addThreadCard({ id, author, date, text, reactions: [], replies: [], paraId, resolved: false });
    setActiveComment(id);
    positionCards();
    mark();
  };
  const cmtIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<path d="M2 3.5h12v8H6l-3 2.5V11.5H2z"/><line x1="5" y1="6.2" x2="11" y2="6.2"/><line x1="5" y1="8.6" x2="9" y2="8.6"/></svg>';

  // Track changes (suggestion mode) lives in its own feature module; it builds its own
  // toolbar buttons and exposes beginFormatChange for the bold/italic/underline buttons.
  const { beginFormatChange, suggestBtn, acceptAllBtn, rejectAllBtn, updateChangeButtons } = setupTrackChanges({
    doc, wrap, regions, options, mark, positionCards, getActiveEl: () => activeEl, iconBtn, btn,
  });

  // Toolbar: shared controls always shown; image/comment/page-break/track-changes are
  // gated by the adapter's capabilities so a format can hide what it cannot serialize.
  const linkBtn = iconBtn(linkIcon, t("linkAria"), () => {
    const url = prompt(t("linkPrompt"), "https://");
    if (url === null) return;
    if (url === "") exec("unlink");
    else exec("createLink", url);
  });
  // Accept-all / reject-all apply to tracked changes; shown only when the doc has some.
  afterReflow = updateChangeButtons;
  updateChangeButtons();
  const items: (Node | null)[] = [
    btn("B", t("bold"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold"),
    btn("I", t("italic"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic"),
    btn("U", t("underline"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline"),
    caps.textColor ? colorInput : null,
    caps.textColor ? bgWrap : null,
    sep(),
    block,
    caps.fontControls ? fontSel : null,
    caps.fontControls ? sizeSel : null,
    sep(),
    iconBtn(bulletIcon, t("bulleted"), () => exec("insertUnorderedList")),
    iconBtn(numberIcon, t("numbered"), () => exec("insertOrderedList")),
    caps.alignment ? sep() : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), t("alignLeft"), () => exec("justifyLeft")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), t("alignCenter"), () => exec("justifyCenter")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), t("alignRight"), () => exec("justifyRight")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), t("alignJustify"), () => exec("justifyFull")) : null,
    sep(),
    caps.images ? iconBtn(imgIcon, t("insertImage"), insertImage) : null,
    caps.comments ? iconBtn(cmtIcon, t("addComment"), addComment) : null,
    caps.pageBreak ? iconBtn(pbIcon, t("insertPageBreak"), insertPageBreak) : null,
    linkBtn,
    caps.trackChanges ? sep() : null,
    caps.trackChanges ? suggestBtn : null,
    caps.trackChanges ? acceptAllBtn : null,
    caps.trackChanges ? rejectAllBtn : null,
    sep(),
    zoomSlider,
    zoomLabel,
  ];
  // Overflow menu: the toolbar is a single row; items that do not fit move into a "…"
  // popover so nothing is lost on narrow widths. The popover lives inside the toolbar so
  // the toolbar's button/sep styling (descendant selectors) still applies to pocketed items.
  const toolbarItems = items.filter((n): n is HTMLElement => n != null);
  const moreBtn = document.createElement("button");
  moreBtn.type = "button";
  moreBtn.className = "docxedit-tb-more";
  moreBtn.textContent = "⋯";
  moreBtn.title = t("moreTools");
  moreBtn.setAttribute("aria-label", t("moreTools"));
  // The popover carries the toolbar class so pocketed items keep their styling, and lives
  // in wrap (not the toolbar) so the toolbar's overflow:hidden does not clip it.
  const overflow = document.createElement("div");
  overflow.className = "docxedit-toolbar docxedit-tb-overflow";
  overflow.hidden = true;
  moreBtn.addEventListener("mousedown", (e) => e.preventDefault());
  moreBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    overflow.style.top = `${toolbar.offsetHeight}px`;
    overflow.hidden = !overflow.hidden;
  });
  const closeOverflow = (e: MouseEvent) => {
    if (!overflow.hidden && !overflow.contains(e.target as Node) && e.target !== moreBtn) overflow.hidden = true;
  };
  document.addEventListener("click", closeOverflow);
  toolbar.append(...toolbarItems, moreBtn);
  wrap.appendChild(overflow);

  const layoutToolbar = () => {
    overflow.hidden = true;
    for (const it of toolbarItems) toolbar.insertBefore(it, moreBtn); // pull everything back in
    moreBtn.style.display = "none";
    if (toolbar.scrollWidth <= toolbar.clientWidth + 1) return; // it all fits
    moreBtn.style.display = "";
    for (let i = toolbarItems.length - 1; i >= 0; i--) {
      if (toolbar.scrollWidth <= toolbar.clientWidth + 1) break;
      overflow.insertBefore(toolbarItems[i], overflow.firstChild); // pocket trailing items, in order
    }
  };
  layoutToolbar();
  requestAnimationFrame(layoutToolbar);
  setTimeout(layoutToolbar, 150);
  const toolbarObserver = new ResizeObserver(() => layoutToolbar());
  toolbarObserver.observe(toolbar);

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
      toolbarObserver.disconnect();
      document.removeEventListener("click", closeOverflow);
      wrap.remove();
    },
  };
}
