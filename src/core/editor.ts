// Shared rich-document editor engine. Renders the editable surface, toolbar, comments
// panel, track changes, image/page chrome and passthrough; the format-specific parse,
// serialize and comment markers come from an Adapter (see core/types.ts). docx is the
// reference adapter; odt reuses this same engine.

import { t } from "./i18n";
import { defaultPageGeometry, paginate } from "./page";
import type { Adapter, EditorOptions, RichEditor, RichDoc, SecGeom, Note } from "./types";
import { setupComments } from "./feature/comments";
import { setupImages } from "./feature/images";
import { setupImageLayout } from "./feature/image-layout";
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
  // Named paragraph styles: their appearance is driven by injected CSS keyed on data-rdoc-style,
  // so applying a style is just setting the attribute (no inline formatting to compute).
  if (parts.styleCss) {
    const ss = document.createElement("style");
    ss.textContent = parts.styleCss;
    wrap.appendChild(ss);
  }
  // Styles the user authors in-session: collected for save, and their CSS appended live so a
  // new style takes effect immediately (same data-rdoc-style / data-rdoc-cstyle keying).
  const newStyles: import("./types").NewStyle[] = [];
  const newStyleCss = document.createElement("style");
  wrap.appendChild(newStyleCss);
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
    const cols = geometry.columns && geometry.columns > 1 ? geometry.columns : 1;
    page.classList.toggle("is-columns", cols > 1);
    page.style.setProperty("--rdoc-columns", String(cols));
    page.style.setProperty("--rdoc-column-gap", `${geometry.columnGapPx ?? 36}px`);
    // Writing direction (re-applied here so it can be toggled live): vertical Japanese tategaki
    // (fixed-height page, columns advancing right-to-left) or horizontal RTL.
    page.classList.toggle("is-vertical", isVertical());
    page.classList.toggle("is-rtl", !isVertical() && caps.verticalText && !!geometry.rtl);
  };
  // Vertical (Japanese tategaki): fixed-height page, columns top-to-bottom advancing right to
  // left, so the page grows along x and page cards advance right-to-left (see repaginateVertical).
  // Horizontal RTL (Arabic/Hebrew) is just direction:rtl and keeps the normal layout. Live (a
  // function) so the Page setup dialog can switch direction at runtime.
  // Whole-page vertical layout (the optimised tategaki path). A document with section breaks is
  // laid out per-section instead (each box carries its own direction), so this is false there.
  const isVertical = (): boolean => caps.verticalText && !!geometry.vertical && !doc.querySelector("[data-rdoc-secbreak], [data-rdoc-secstart]");
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
  // First-page / even-page header & footer variants (the default lives in header/footer). Each is an
  // editable source band shown on the matching pages; null when the document doesn't declare it.
  let headerFirst = parts.headerFirst ? band("docxedit-header", t("header"), parts.headerFirst.html) : null;
  let footerFirst = parts.footerFirst ? band("docxedit-footer", t("footer"), parts.footerFirst.html) : null;
  let headerEven = parts.headerEven ? band("docxedit-header", t("header"), parts.headerEven.html) : null;
  let footerEven = parts.footerEven ? band("docxedit-footer", t("footer"), parts.footerEven.html) : null;
  // The source band to show for a role on a given 0-based page: the first-page variant on page 0
  // (titlePage), the even variant on even-numbered pages (evenOdd; page index 1 = page 2), else the
  // default; a missing variant falls back to the default.
  const pickHF = (role: "header" | "footer", page: number): HTMLElement | null => {
    const def = role === "header" ? header : footer;
    if (geometry.titlePage && page === 0) return (role === "header" ? headerFirst : footerFirst) ?? def;
    if (geometry.evenOdd && page % 2 === 1) return (role === "header" ? headerEven : footerEven) ?? def;
    return def;
  };
  // The space a role reserves: the tallest of its present variants, so no page's variant overlaps
  // the body even though the paginator uses one uniform content height for all pages.
  const hfHeight = (role: "header" | "footer"): number => {
    const first = geometry.titlePage ? (role === "header" ? headerFirst : footerFirst) : null;
    const even = geometry.evenOdd ? (role === "header" ? headerEven : footerEven) : null;
    return Math.max(0, ...[role === "header" ? header : footer, first, even].map((b) => b?.offsetHeight ?? 0));
  };
  // Distinct per-section header/footer source bands, keyed by the key a section's boundary
  // paragraph carries (data-rdoc-sec*key). Each is editable + saved back to its own part.
  const secBands = new Map<string, { el: HTMLElement; path: string }>();
  for (const [key, { html, path }] of Object.entries(parts.sectionBands ?? {})) {
    const el = band("docxedit-header", t("header"), html);
    if (el) secBands.set(key, { el, path });
  }
  // Footnote / endnote bodies, keyed by id. Each is an editable block placed in the notes area
  // at the bottom of the page holding its reference; editing it is the save source.
  const noteBands = new Map<string, { el: HTMLElement; kind: "footnote" | "endnote" }>();
  for (const n of parts.notes ?? []) {
    const el = document.createElement("div");
    el.className = "docxedit-note";
    el.contentEditable = "true";
    el.spellcheck = false;
    el.innerHTML = n.html || "<p><br></p>";
    el.addEventListener("focus", () => { activeEl = el; });
    el.addEventListener("input", () => mark());
    noteBands.set(n.id, { el, kind: n.kind });
  }
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
    for (const b of [header, footer, headerFirst, footerFirst, headerEven, footerEven]) if (b) measure.appendChild(b);
    for (const { el } of secBands.values()) measure.appendChild(el); // off-screen, for measuring
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

  // Page geometry write-back is gated on this flag (set when the document geometry is edited via
  // the ruler or Page setup). Owned here so section authoring can set it too.
  let geometryDirty = false;

  // --- Section authoring (insert breaks + per-section page setup) -----------------------------
  // The flat block sequence (descending the column/section wrappers a layout pass may have added).
  const flatBlocks = (): HTMLElement[] => {
    const out: HTMLElement[] = [];
    for (const c of Array.from(doc.children) as HTMLElement[]) {
      if (c.classList.contains("docxedit-secpage") || c.classList.contains("docxedit-colpage")) out.push(...(Array.from(c.children) as HTMLElement[]));
      else if (!c.classList.contains("docxedit-pagespacer") && !c.classList.contains("docxedit-pagetop")) out.push(c);
    }
    return out;
  };
  const caretFlatBlock = (): HTMLElement | null => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const node = sel.getRangeAt(0).startContainer;
    return flatBlocks().find((b) => b === node || b.contains(node)) ?? null;
  };
  // Which paragraph carries the caret's section geometry, per the format's convention. null = the
  // caret is in the document (trailing-final / leading-first) section, whose geometry is `geometry`.
  const sectionGeomEl = (): HTMLElement | null => {
    if (!caps.sections) return null;
    const blocks = flatBlocks();
    const b = caretFlatBlock();
    const idx = b ? blocks.indexOf(b) : -1;
    if (idx < 0) return null;
    if (caps.sections === "trailing") {
      for (let i = idx; i < blocks.length; i++) if (blocks[i]!.hasAttribute("data-rdoc-secbreak")) return blocks[i]!;
    } else {
      for (let i = idx; i >= 0; i--) if (blocks[i]!.hasAttribute("data-rdoc-secstart")) return blocks[i]!;
    }
    return null;
  };
  const parseSecGeom = (el: HTMLElement): SecGeom => {
    const attr = caps.sections === "leading" ? "data-rdoc-secstart" : "data-rdoc-secbreak";
    try { return mergeSecGeom(JSON.parse(el.getAttribute(attr) ?? "")); } catch { return docGeom(); }
  };
  const readSectionGeom = (): SecGeom => {
    const el = sectionGeomEl();
    return el ? parseSecGeom(el) : docGeom();
  };
  // Apply a geometry to the caret's section (document geometry, or a mid-document section's
  // boundary paragraph). Mutates the model only; the caller reflows.
  const writeSectionGeom = (g: SecGeom): void => {
    const el = sectionGeomEl();
    if (!el) {
      geometry.widthPx = g.w;
      geometry.heightPx = g.h;
      geometry.margin = { top: g.mt, right: g.mr, bottom: g.mb, left: g.ml };
      geometry.columns = g.cols && g.cols > 1 ? g.cols : undefined;
      geometry.columnGapPx = geometry.columns ? (g.colGap ?? 36) : undefined;
      geometry.vertical = g.vertical;
      geometry.rtl = g.rtl;
      geometryDirty = true;
      applyGeometry();
      return;
    }
    const attr = caps.sections === "leading" ? "data-rdoc-secstart" : "data-rdoc-secbreak";
    el.setAttribute(attr, JSON.stringify(g));
    el.setAttribute("data-rdoc-secedited", "1"); // regenerate this section's props on save (vs passthrough)
  };
  let secMasterSeq = 0;
  // Insert a next-page section break after the caret's paragraph; the new section inherits the
  // current section's geometry until the user changes it. Mutates the model only.
  const insertSectionBreak = (): void => {
    if (!caps.sections) return;
    const blocks = flatBlocks();
    const b = caretFlatBlock() ?? blocks[blocks.length - 1];
    if (!b) return;
    const geom = JSON.stringify(readSectionGeom());
    if (caps.sections === "trailing") {
      b.setAttribute("data-rdoc-secbreak", geom); // b ends the section before the break
      b.setAttribute("data-rdoc-secedited", "1");
      if (b === blocks[blocks.length - 1]) {
        const p = document.createElement("p");
        p.innerHTML = "<br>";
        b.after(p); // give the new section a paragraph to type in
      }
    } else {
      let c = blocks[blocks.indexOf(b) + 1] ?? null; // the new section starts at the block after b
      if (!c) { c = document.createElement("p"); c.innerHTML = "<br>"; b.after(c); }
      c.setAttribute("data-rdoc-secstart", geom);
      c.setAttribute("data-rdoc-secedited", "1");
      c.setAttribute("data-odt-masterpage", `rdoc-sec-${++secMasterSeq}`);
      c.setAttribute("data-odt-break-before", "page");
    }
  };
  // Toggle a section's header/footer "link to previous": linked = inherits the document default;
  // independent = its own editable band (pre-filled with the inherited content), saved to its own
  // part (docx) / master (odt). Only non-main sections (with a boundary paragraph) can toggle.
  let secBandSeq = 0;
  const toggleSectionBand = (boundaryEl: HTMLElement, role: "header" | "footer"): void => {
    const attr = `data-rdoc-sec${role}key`;
    const existing = boundaryEl.getAttribute(attr);
    if (existing) {
      boundaryEl.removeAttribute(attr); // relink: drop the override, fall back to the default
      secBands.delete(existing);
    } else {
      const inherited = (role === "header" ? header : footer)?.innerHTML || "<p><br></p>";
      let key: string, path: string;
      if (caps.sections === "leading") {
        const master = boundaryEl.getAttribute("data-odt-masterpage");
        if (!master) return; // a leading section always has a master; nothing to attach to otherwise
        key = `${role === "header" ? "oh" : "of"}:${master}`;
        path = `${role}@${master}`;
      } else {
        key = `new${role}:${++secBandSeq}`; // a fresh part minted on save
        path = key;
      }
      const el = band("docxedit-header", t(role), inherited);
      if (!el) return;
      measure.appendChild(el);
      secBands.set(key, { el, path });
      boundaryEl.setAttribute(attr, key);
    }
    if (caps.sections === "trailing") boundaryEl.setAttribute("data-rdoc-secedited", "1"); // regenerate the sectPr ref
    reflow();
    mark();
  };

  // Page view: rulers (margin handles) + zoom + the centred canvas around the page box.
  // Turn the document's first-page / even-odd header & footer variants on or off. On: ensure the
  // variant source bands exist (empty, like Word, so an enabled-but-empty first header just leaves
  // page 1 blank); off: keep any band in memory but stop showing/saving it (pickHF + hfHeight +
  // getBytes all gate on the flag, so re-enabling restores the content).
  const toggleHFVariant = (variant: "first" | "even", on: boolean): void => {
    const ensure = (cur: HTMLElement | null, role: "header" | "footer"): HTMLElement | null => {
      if (!on || cur) return cur;
      const el = band(role === "header" ? "docxedit-header" : "docxedit-footer", t(role), "<p><br></p>");
      if (el) measure.appendChild(el);
      return el;
    };
    if (variant === "first") { geometry.titlePage = on || undefined; headerFirst = ensure(headerFirst, "header"); footerFirst = ensure(footerFirst, "footer"); }
    else { geometry.evenOdd = on || undefined; headerEven = ensure(headerEven, "header"); footerEven = ensure(footerEven, "footer"); }
    geometryDirty = true;
    mark();
    reflow();
  };
  const { applyZoom, effectiveZoom, zoomSlider, zoomLabel, pageSetupBtn, sectionBreakBtn, teardown: teardownPageView } = setupPageView({
    page, pagebox, canvas, leftSpacer, rightArea, scroll, geometry, options, caps,
    getVertical: () => isVertical(),
    applyGeometry, mark, positionCards, reflow: () => reflow(), scheduleReflow: () => scheduleReflow(),
    markGeometryDirty: () => { geometryDirty = true; },
    readSectionGeom, writeSectionGeom, insertSectionBreak, toggleHFVariant,
  });
  // Bottom bar: paragraph/character style pickers on the left, zoom on the right. Keeps the
  // top toolbar focused on formatting/insert controls.
  const bottomBar = document.createElement("div");
  bottomBar.className = "docxedit-bottombar";
  const bottomLeft = document.createElement("div");
  bottomLeft.className = "docxedit-bottombar-left";
  const bottomRight = document.createElement("div");
  bottomRight.className = "docxedit-bottombar-right";
  if (sectionBreakBtn) bottomRight.append(sectionBreakBtn);
  bottomRight.append(pageSetupBtn, zoomSlider, zoomLabel);
  bottomBar.append(bottomLeft, bottomRight);
  wrap.append(toolbar, scroll, bottomBar);
  container.appendChild(wrap);

  // Footnote / endnote area below the pages. References are renumbered in document order and each
  // referenced note's editable body is shown here (per-page-bottom placement is a later refinement).
  // The document's footnote style (font/size/line-height/colour), applied inline to each note area
  // so footnotes render at the document's footnote size instead of the stylesheet's fallback.
  const noteAreaCss = parts.notes?.length ? (parts.noteCss ?? "") : "";
  const noteslayer = document.createElement("div");
  noteslayer.className = "docxedit-noteslayer";
  if (noteAreaCss) noteslayer.style.cssText = noteAreaCss;
  noteslayer.hidden = true;
  scroll.appendChild(noteslayer);
  let footnotesPerPage = false; // the single-section path renders footnotes at each page bottom
  // Number every reference in document order (footnotes and endnotes as separate sequences).
  const numberRefs = (): { ref: HTMLElement; kind: "footnote" | "endnote"; num: number; id: string }[] => {
    let fn = 0, en = 0;
    return Array.from(doc.querySelectorAll<HTMLElement>(".docx-fnref")).map((ref) => {
      const kind = ref.getAttribute("data-fn-kind") === "endnote" ? "endnote" : "footnote";
      const num = kind === "endnote" ? ++en : ++fn;
      ref.textContent = String(num);
      return { ref, kind, num, id: ref.getAttribute("data-fn-id") ?? "" };
    });
  };
  // A "N. <editable note body>" row for the notes area / a page's footnote area.
  const noteRow = (num: number, id: string): HTMLElement | null => {
    const nb = noteBands.get(id);
    if (!nb) return null;
    const row = document.createElement("div");
    row.className = "docxedit-note-row";
    const n = document.createElement("span");
    n.className = "docxedit-note-num";
    n.contentEditable = "false";
    n.textContent = `${num}.`;
    row.append(n, nb.el);
    return row;
  };
  // --- Per-page footnote areas, shared by every paginated layout -------------------------------
  // The page bottom (horizontal) or left edge (vertical) holds the page's referenced footnotes,
  // and the body reserves the space so it does not overlap. The reserve = a base (the separator
  // border + padding) plus each note's measured thickness (row height horizontally, row width in
  // a vertical band).
  const FN_BASE = 8;
  const footnoteRefs = () => numberRefs().filter((r) => r.kind === "footnote" && noteBands.has(r.id));
  // Build each footnote's row once and measure it laid out as it will render (its font, the area's
  // width/height, paragraph margins), so the reserve matches the rendered size. `extent` is the
  // area's fixed cross-size: its width horizontally, its height (column length) in a vertical band.
  const measureFootnotes = (footRefs: ReturnType<typeof footnoteRefs>, vertical: boolean, extent: number) => {
    const fnRow = new Map<string, HTMLElement>();
    const fnSize = new Map<string, number>();
    if (footRefs.length) {
      const probe = document.createElement("div");
      probe.className = vertical ? "docxedit-fnarea is-vertical" : "docxedit-fnarea";
      probe.style.cssText = `${vertical ? `top:0;right:0;height:${extent}px` : `top:0;left:0;width:${extent}px`};visibility:hidden;${noteAreaCss}`;
      hflayer.appendChild(probe);
      for (const fr of footRefs) { const row = noteRow(fr.num, fr.id); if (row) { fnRow.set(fr.id, row); probe.appendChild(row); } }
      for (const fr of footRefs) { const row = fnRow.get(fr.id); if (row) fnSize.set(fr.id, vertical ? row.offsetWidth : row.offsetHeight); }
      probe.remove();
    }
    return { fnRow, fnSize };
  };
  // The reserve per page: FN_BASE plus the thickness of every footnote whose reference is on it.
  const footnoteReserve = (footRefs: ReturnType<typeof footnoteRefs>, fnSize: Map<string, number>, pageOf: (ref: HTMLElement) => number): number[] => {
    const r: number[] = [];
    for (const fr of footRefs) { const p = pageOf(fr.ref); if (p < 0) continue; r[p] = (r[p] || FN_BASE) + (fnSize.get(fr.id) ?? 0); }
    return r;
  };
  // Group the footnotes by their page and draw one area per page; `areaCss(page)` positions it.
  const drawFootnoteAreas = (footRefs: ReturnType<typeof footnoteRefs>, fnRow: Map<string, HTMLElement>, vertical: boolean, pageOf: (ref: HTMLElement) => number, areaCss: (page: number) => string) => {
    const byPage = new Map<number, ReturnType<typeof footnoteRefs>>();
    for (const fr of footRefs) { const p = pageOf(fr.ref); if (p < 0) continue; (byPage.get(p) ?? byPage.set(p, []).get(p)!).push(fr); }
    for (const [p, frs] of byPage) {
      const area = document.createElement("div");
      area.className = vertical ? "docxedit-fnarea is-vertical" : "docxedit-fnarea";
      area.style.cssText = `${areaCss(p)};${noteAreaCss}`;
      for (const fr of frs) { const row = fnRow.get(fr.id); if (row) area.appendChild(row); }
      hflayer.appendChild(area);
    }
  };
  // Doc-end notes area: endnotes always, plus footnotes when the layout does not place them per page.
  const renderNotes = () => {
    const rows: HTMLElement[] = [];
    for (const r of numberRefs()) {
      if (r.kind === "footnote" && footnotesPerPage) continue;
      const row = noteRow(r.num, r.id);
      if (row) rows.push(row);
    }
    noteslayer.replaceChildren(...rows);
    noteslayer.hidden = rows.length === 0;
  };
  // Insert a footnote / endnote at the caret: a reference in the body + a new empty editable note.
  let noteSeq = 0;
  const insertNote = (kind: "footnote" | "endnote", text = "") => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const node = range.startContainer;
    const el = node.nodeType === 3 ? node.parentElement : (node as HTMLElement);
    if (!el || !doc.contains(el)) return; // references live in the body
    const id = `rdoc-note-new-${++noteSeq}`;
    const sup = document.createElement("sup");
    sup.className = "docx-fnref";
    sup.setAttribute("data-fn-id", id);
    sup.setAttribute("data-fn-kind", kind);
    sup.contentEditable = "false";
    sup.textContent = "*";
    range.collapse(false);
    range.insertNode(sup);
    const nb = document.createElement("div");
    nb.className = "docxedit-note";
    nb.contentEditable = "true";
    nb.spellcheck = false;
    const p = document.createElement("p");
    if (text.trim()) p.textContent = text; else p.innerHTML = "<br>";
    nb.appendChild(p);
    nb.addEventListener("focus", () => { activeEl = nb; });
    nb.addEventListener("input", () => mark());
    noteBands.set(id, { el: nb, kind });
    mark();
    reflow();
    setTimeout(() => { nb.focus(); placeCaret(nb, 0); }, 0);
  };

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

  // Editable clone of a header/footer band, positioned by posCss (top/left/right/width). Clicking
  // it places the caret natively; edits sync into the canonical band (the save source); on blur the
  // pages re-clone. Shared by horizontal and vertical pagination.
  const mkClone = (src: HTMLElement, posCss: string): HTMLElement => {
    const c = src.cloneNode(true) as HTMLElement;
    c.removeAttribute("role");
    c.removeAttribute("aria-label");
    c.classList.add("docxedit-hf-clone");
    c.contentEditable = "true";
    c.spellcheck = false;
    c.style.cssText = posCss;
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
      // A band created by double-click but left empty is dropped, not saved. Defer the check so
      // moving between page-clones of the same band does not count as leaving.
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
  // A corner chip on a non-main section's header/footer toggling "link to previous" (shared
  // default vs the section's own band). Lives in the overlay, outside the editable clone.
  const mkLinkChip = (role: "header" | "footer", boundaryEl: HTMLElement, posCss: string): HTMLElement => {
    const linked = !boundaryEl.getAttribute(`data-rdoc-sec${role}key`);
    const c = document.createElement("button");
    c.type = "button";
    c.className = `docxedit-hf-link${linked ? "" : " is-on"}`;
    c.title = linked ? t("sectionUnlink") : t("sectionLink");
    c.style.cssText = posCss;
    c.innerHTML = linked
      ? `<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6.5 9.5l3-3M6 5.5L7.5 4a2.5 2.5 0 0 1 3.5 3.5L9.5 9M10 10.5L8.5 12A2.5 2.5 0 0 1 5 8.5L6.5 7"/></svg>`
      : `<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6 5.5L7.5 4a2.5 2.5 0 0 1 3.5 3.5L9.5 9M10 10.5L8.5 12A2.5 2.5 0 0 1 5 8.5L6.5 7M2 2l12 12"/></svg>`;
    c.addEventListener("mousedown", (e) => e.preventDefault()); // keep the caret/selection
    c.addEventListener("click", (e) => { e.preventDefault(); toggleSectionBand(boundaryEl, role); });
    return c;
  };
  // Fields (page number / count / table of contents). Clone bands get per-page values; body
  // fields and the TOC are computed from the laid-out positions after pagination.
  const setCloneFields = (el: HTMLElement, page: number, total: number): void => {
    for (const f of Array.from(el.querySelectorAll<HTMLElement>(".docx-field"))) {
      const k = f.getAttribute("data-field");
      if (k === "PAGE") f.textContent = String(page);
      else if (k === "NUMPAGES") f.textContent = String(total);
    }
  };
  // Clone the header and footer variant that applies to page `p` (0-based) into the overlay, with
  // the given positioning. Shared by every paginate path so first/even variants render everywhere.
  const cloneHF = (p: number, total: number, headerCss: string, footerCss: string): void => {
    const hSrc = pickHF("header", p);
    if (hSrc) { const hc = mkClone(hSrc, headerCss); setCloneFields(hc, p + 1, total); hflayer.appendChild(hc); }
    const fSrc = pickHF("footer", p);
    if (fSrc) { const fc = mkClone(fSrc, footerCss); setCloneFields(fc, p + 1, total); hflayer.appendChild(fc); }
  };
  // The display text for a cross-reference to `target`: a heading's text, a range bookmark's spanned
  // text, or the bookmark name as a fallback for a point bookmark.
  const xrefTargetText = (target: HTMLElement, name: string): string => {
    if (/^H[1-6]$/.test(target.tagName)) return (target.textContent ?? "").trim() || name;
    const id = target.getAttribute("data-rdoc-bm-id");
    const end = id ? Array.from(doc.querySelectorAll<HTMLElement>(".docx-bookmark-end")).find((e) => e.getAttribute("data-rdoc-bm-id") === id) : null;
    if (end && target.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING) {
      const r = document.createRange();
      r.setStartAfter(target);
      r.setEndBefore(end);
      const txt = r.toString().trim();
      if (txt) return txt;
    }
    return name;
  };
  const tocSig = new WeakMap<Element, string>();
  const decorateFields = (cardCount: number, pageStep: number, vertical: boolean): void => {
    for (const f of Array.from(doc.querySelectorAll<HTMLElement>('.docx-field[data-field="NUMPAGES"]'))) f.textContent = String(cardCount);
    for (const f of Array.from(doc.querySelectorAll<HTMLElement>('.docx-field[data-field="PAGE"]')))
      f.textContent = String(vertical ? 1 : Math.max(1, Math.floor(f.offsetTop / pageStep) + 1));
    // Cross-references: recompute each xref's text from its target (a heading or bookmark carrying
    // the matching data-rdoc-bm) - its text, or its page number for the "page" format.
    const bmTarget = (name: string) => Array.from(doc.querySelectorAll<HTMLElement>("[data-rdoc-bm]")).find((e) => e.getAttribute("data-rdoc-bm") === name) ?? null;
    for (const x of Array.from(doc.querySelectorAll<HTMLElement>(".docx-xref"))) {
      const name = x.getAttribute("data-rdoc-xref");
      const target = name ? bmTarget(name) : null;
      if (!target) continue;
      let text: string;
      if (x.getAttribute("data-rdoc-xref-fmt") === "page") text = vertical ? "1" : String(Math.max(1, Math.floor(target.offsetTop / pageStep) + 1));
      else text = xrefTargetText(target, name!);
      if (text && x.textContent !== text) x.textContent = text;
    }
    let needReflow = false;
    for (const toc of Array.from(doc.querySelectorAll<HTMLElement>(".docx-field-toc"))) {
      const headings = Array.from(doc.querySelectorAll<HTMLElement>("h1,h2,h3")).filter((h) => !h.closest(".docx-field-toc"));
      const pageOf = (el: HTMLElement) => (vertical ? "" : String(Math.max(1, Math.floor(el.offsetTop / pageStep) + 1)));
      const sig = `${cardCount}|` + headings.map((h) => `${h.tagName}:${h.textContent}:${pageOf(h)}`).join("|");
      if (tocSig.get(toc) === sig) continue; // unchanged: don't rebuild (and don't loop reflow)
      tocSig.set(toc, sig);
      needReflow = true;
      toc.replaceChildren();
      if (!headings.length) {
        const e = document.createElement("div");
        e.className = "docx-field-toc-empty";
        e.textContent = t("tocEmpty");
        toc.appendChild(e);
        continue;
      }
      const title = document.createElement("div");
      title.className = "docx-field-toc-title";
      title.textContent = t("tocTitle");
      toc.appendChild(title);
      for (const h of headings) {
        const row = document.createElement("div");
        row.className = `docx-field-toc-row toc-${h.tagName.toLowerCase()}`;
        const txt = document.createElement("span");
        txt.className = "docx-field-toc-text";
        txt.textContent = h.textContent || "";
        const pg = document.createElement("span");
        pg.className = "docx-field-toc-page";
        pg.textContent = pageOf(h);
        row.append(txt, pg);
        toc.appendChild(row);
      }
    }
    if (needReflow) scheduleReflow(); // TOC height changed; one more pass settles pagination
  };

  const repaginateVertical = () => {
    for (const s of Array.from(doc.querySelectorAll(":scope > .docxedit-pagespacer"))) s.remove();
    pagelayer.replaceChildren();
    hflayer.replaceChildren();
    const { left, right } = geometry.margin;
    const contentExtent = geometry.widthPx - left - right; // usable page width along the fill axis
    const pageStep = geometry.widthPx + PAGE_GAP;
    // Header/footer are horizontal bands at the page top/bottom (measured at page width); the
    // body columns fill the height between them.
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = hfHeight("header");
    const footerH = hfHeight("footer");
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    doc.style.height = `${geometry.heightPx}px`;
    doc.style.width = ""; // clear any inline width from a prior vertical-columns layout
    doc.style.padding = `${contentTop}px ${right}px ${contentBottomInset}px ${left}px`;
    // Unwrap vertical band wrappers from a prior multi-column vertical layout.
    for (const w of Array.from(doc.querySelectorAll<HTMLElement>(`.${VBAND}`))) {
      while (w.firstChild) doc.insertBefore(w.firstChild, w);
      w.remove();
    }
    for (const el of Array.from(doc.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");

    const kids = Array.from(doc.children).filter((c) => !c.classList.contains("docxedit-pagespacer")) as HTMLElement[];
    // progression (right-to-left) size: right edge of block i minus right edge of block i+1.
    const rights = kids.map((k) => k.offsetLeft + k.offsetWidth);
    const sizes = kids.map((k, i) => (i < kids.length - 1 ? rights[i]! - rights[i + 1]! : k.offsetWidth));

    // Footnotes: a vertical band along each page's left edge (the end of the right-to-left flow),
    // mirroring the horizontal per-page area. Its thickness (block-axis width) reserves room via the
    // paginator, so body columns stop before it. Notes are measured laid out vertical-rl, at the
    // column height, so the reserve matches how they render.
    const colHeight = geometry.heightPx - contentTop - contentBottomInset;
    const footRefs = footnoteRefs();
    footnotesPerPage = true;
    const kidIndexOf = (ref: HTMLElement) => kids.findIndex((k) => k.contains(ref));
    const { fnRow, fnSize } = measureFootnotes(footRefs, true, colHeight);
    const reserveFor = (pob: number[]): number[] => footnoteReserve(footRefs, fnSize, (ref) => { const ki = kidIndexOf(ref); return ki < 0 ? -1 : pob[ki]!; });
    let reserve = footRefs.length ? reserveFor(paginate(sizes, { pageStep, contentHeight: contentExtent }, manualBreaks(kids)).pageOfBlock) : [];
    const { spacerBefore, cardCount, pageOfBlock } = paginate(sizes, { pageStep, contentHeight: contentExtent, reserveOf: (p) => reserve[p] || 0 }, manualBreaks(kids));
    reserve = footRefs.length ? reserveFor(pageOfBlock) : [];
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
      const rightPx = p * pageStep;
      cloneHF(p, cardCount, `top:${geometry.margin.top}px;right:${rightPx}px;width:${geometry.widthPx}px`, `top:${geometry.heightPx - contentBottomInset}px;right:${rightPx}px;width:${geometry.widthPx}px`);
    }

    // Footnotes: per page, a vertical band at the left edge of the content (just inside the left
    // margin), holding that page's referenced notes with a separator on the body side.
    drawFootnoteAreas(footRefs, fnRow, true, (ref) => { const ki = kidIndexOf(ref); return ki < 0 ? -1 : pageOfBlock[ki]!; },
      (p) => `top:${contentTop}px;right:${p * pageStep + (geometry.widthPx - geometry.margin.left) - (reserve[p] || 0)}px;width:${reserve[p] || 0}px;height:${colHeight}px`);

    page.style.width = `${cardCount * pageStep - PAGE_GAP}px`;
    page.style.minHeight = `${geometry.heightPx}px`;
    decorateFields(cardCount, pageStep, true);
  };

  // Vertical (tategaki) multi-column pagination. CSS multicol does not fragment vertical-rl text
  // into stacked bands, so each "column" is laid out manually: a page is divided into N horizontal
  // bands stacked top-to-bottom, each a vertical-rl region; blocks are bucketed into bands by
  // block-axis (width) overflow, bands fill a page top-down, and pages advance right-to-left. The
  // caret is preserved across the reparent like the column reflow.
  const VBAND = "docxedit-vband";
  const repaginateVerticalColumns = () => {
    const blocks: HTMLElement[] = [];
    for (const child of Array.from(doc.children) as HTMLElement[]) {
      if (child.classList.contains(VBAND)) blocks.push(...(Array.from(child.children) as HTMLElement[]));
      else if (!child.classList.contains("docxedit-pagespacer")) blocks.push(child);
    }
    const sel = window.getSelection();
    let caretBlock: HTMLElement | null = null;
    let caretOffset = 0;
    if (sel && sel.rangeCount) {
      const r = sel.getRangeAt(0);
      const blk = blocks.find((b) => b === r.startContainer || b.contains(r.startContainer));
      if (blk) { caretBlock = blk; caretOffset = charOffsetIn(blk, r.startContainer, r.startOffset); }
    }
    const N = geometry.columns && geometry.columns > 1 ? geometry.columns : 2;
    const gap = geometry.columnGapPx ?? 36;
    const { left, right } = geometry.margin;
    const contentExtent = geometry.widthPx - left - right;
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = hfHeight("header");
    const footerH = hfHeight("footer");
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    const contentHeight = geometry.heightPx - contentTop - contentBottomInset;
    const bandHeight = (contentHeight - (N - 1) * gap) / N;
    const pageStep = geometry.widthPx + PAGE_GAP;
    const EPS = 2;

    doc.style.padding = "0";
    pagelayer.replaceChildren();
    hflayer.replaceChildren();
    // A page reserves a strip at its left edge (where the right-to-left flow runs out) for footnotes
    // by narrowing every band on that page; the footnote band fills the freed strip.
    const newBand = (bandWidth: number): HTMLElement => {
      const b = document.createElement("div");
      b.className = VBAND;
      b.style.cssText = `position:absolute;writing-mode:vertical-rl;width:${bandWidth}px;height:${bandHeight}px;overflow:hidden`;
      return b;
    };
    // vertical-rl overflows leftward, so scrollWidth stays equal to clientWidth; detect overflow by
    // the just-added block crossing the band's left edge instead.
    const vOver = (b: HTMLElement, block: HTMLElement): boolean => block.getBoundingClientRect().left < b.getBoundingClientRect().left - EPS;
    const bucketBands = (reserveOf: (bandIndex: number) => number): HTMLElement[] => {
      const bs: HTMLElement[] = [];
      let band = newBand(contentExtent - reserveOf(0));
      doc.appendChild(band);
      bs.push(band);
      for (const block of blocks) {
        band.appendChild(block);
        if (vOver(band, block) && band.children.length > 1) {
          band.removeChild(block);
          band.style.overflow = "clip";
          band = newBand(contentExtent - reserveOf(bs.length));
          doc.appendChild(band);
          bs.push(band);
          band.appendChild(block);
        }
      }
      band.style.overflow = "clip";
      return bs;
    };
    // Two passes when there are footnotes: find each note's page (a page is N consecutive bands),
    // then re-bucket with that page's left-strip reserve applied to every band on it.
    const footRefs = footnoteRefs();
    footnotesPerPage = true;
    const { fnRow, fnSize } = measureFootnotes(footRefs, true, contentHeight);
    const pageOfBand = (bs: HTMLElement[]) => (ref: HTMLElement) => { const i = bs.findIndex((b) => b.contains(ref)); return i < 0 ? -1 : Math.floor(i / N); };
    let bands = bucketBands(() => 0);
    let reserve: number[] = [];
    if (footRefs.length) {
      reserve = footnoteReserve(footRefs, fnSize, pageOfBand(bands));
      bands = bucketBands((i) => reserve[Math.floor(i / N)] || 0);
      reserve = footnoteReserve(footRefs, fnSize, pageOfBand(bands));
    }
    for (const old of Array.from(doc.children) as HTMLElement[]) {
      if ((old.classList.contains(VBAND) && !bands.includes(old)) || old.classList.contains("docxedit-pagespacer")) old.remove();
    }

    const cardCount = Math.ceil(bands.length / N);
    bands.forEach((b, i) => {
      const p = Math.floor(i / N), c = i % N;
      b.style.right = `${p * pageStep + right}px`;
      b.style.top = `${contentTop + c * (bandHeight + gap)}px`;
    });
    doc.style.width = `${cardCount * pageStep - PAGE_GAP}px`;
    doc.style.height = `${geometry.heightPx}px`;
    for (let p = 0; p < cardCount; p++) {
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.left = "auto";
      card.style.right = `${p * pageStep}px`;
      card.style.top = "0";
      card.style.width = `${geometry.widthPx}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
      cloneHF(p, cardCount, `top:${geometry.margin.top}px;right:${p * pageStep}px;width:${geometry.widthPx}px`, `top:${geometry.heightPx - contentBottomInset}px;right:${p * pageStep}px;width:${geometry.widthPx}px`);
    }
    // Footnotes: a vertical band down each page's left edge, beside the narrowed text bands.
    drawFootnoteAreas(footRefs, fnRow, true, pageOfBand(bands),
      (p) => `top:${contentTop}px;right:${p * pageStep + right + (contentExtent - (reserve[p] || 0))}px;width:${reserve[p] || 0}px;height:${contentHeight}px`);
    page.style.width = `${cardCount * pageStep - PAGE_GAP}px`;
    page.style.minHeight = `${geometry.heightPx}px`;
    decorateFields(cardCount, pageStep, true);
    if (caretBlock && caretBlock.isConnected) placeCaret(caretBlock, caretOffset);
  };

  // Multi-column pagination: wrap the body's blocks into one balanced multi-column box per
  // page. The browser's own column flow decides where a page fills (a wrapper sized to the page
  // content with column-fill:auto overflows past N columns when full), so the engine only buckets
  // blocks into pages; column-fill:balance then equalises each page's columns. The same block
  // nodes are moved (not cloned), so the caret survives a reparent (saved + restored here).
  const COLPAGE = "docxedit-colpage";
  // The caret's character offset within a block (walking text nodes), so it can be re-placed
  // after the block is reparented into a column wrapper (moving a node drops the live selection).
  const charOffsetIn = (block: Node, container: Node, offset: number): number => {
    if (container.nodeType !== 3) return 0; // element container (empty block): start
    let total = 0;
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let n: Node | null;
    while ((n = walker.nextNode())) {
      if (n === container) return total + offset;
      total += (n.textContent ?? "").length;
    }
    return total;
  };
  const placeCaret = (block: HTMLElement, target: number): void => {
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let n: Node | null;
    let acc = 0;
    let last: Node | null = null;
    while ((n = walker.nextNode())) {
      const len = (n.textContent ?? "").length;
      if (acc + len >= target) {
        last = n;
        break;
      }
      acc += len;
      last = n;
    }
    const s = window.getSelection();
    if (!s) return;
    const r = document.createRange();
    if (last) r.setStart(last, Math.min((last.textContent ?? "").length, Math.max(0, target - acc)));
    else r.setStart(block, 0);
    r.collapse(true);
    s.removeAllRanges();
    s.addRange(r);
  };
  const repaginateColumns = () => {
    // Flatten: collect the body's blocks, unwrapping any wrappers from a previous pass.
    const blocks: HTMLElement[] = [];
    for (const child of Array.from(doc.children) as HTMLElement[]) {
      if (child.classList.contains(COLPAGE)) blocks.push(...(Array.from(child.children) as HTMLElement[]));
      else if (!child.classList.contains("docxedit-pagespacer")) blocks.push(child);
    }
    // Remember the caret as a (block, char-offset) pair so it survives the reparent below.
    const sel = window.getSelection();
    let caretBlock: HTMLElement | null = null;
    let caretOffset = 0;
    if (sel && sel.rangeCount) {
      const r = sel.getRangeAt(0);
      const blk = blocks.find((b) => b === r.startContainer || b.contains(r.startContainer));
      if (blk) {
        caretBlock = blk;
        caretOffset = charOffsetIn(blk, r.startContainer, r.startOffset);
      }
    }
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = hfHeight("header");
    const footerH = hfHeight("footer");
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    const contentHeight = geometry.heightPx - contentTop - contentBottomInset;
    const contentWidth = geometry.widthPx - geometry.margin.left - geometry.margin.right;
    const pageStep = geometry.heightPx + PAGE_GAP;
    const interPage = contentBottomInset + PAGE_GAP + contentTop; // gap between consecutive wrappers
    doc.style.height = ""; // clear any inline height left by a prior vertical layout
    doc.style.padding = `${contentTop}px ${geometry.margin.right}px ${contentBottomInset}px ${geometry.margin.left}px`;

    // A page reserves a bottom strip for its footnotes (border-box padding keeps the wrapper the
    // same outer height, so the fixed page grid stays aligned); the columns then fill above it.
    const newWrapper = (first: boolean, reserve: number): HTMLElement => {
      const w = document.createElement("div");
      w.className = COLPAGE;
      w.style.boxSizing = "border-box";
      w.style.height = `${contentHeight}px`;
      w.style.width = `${contentWidth}px`;
      if (reserve) w.style.paddingBottom = `${reserve}px`;
      w.style.columnFill = "auto"; // column-by-column during bucketing; balanced when finalized
      w.style.overflow = "hidden"; // a scroll container, so scrollWidth reveals the (N+1)th column
      if (!first) w.style.marginTop = `${interPage}px`;
      return w;
    };
    // Finalize a page: balance its columns and switch to overflow:clip. clip clips without being
    // a scroll container, so the browser stops scrolling the wrapper to chase the caret between
    // columns (the bug overflow:hidden caused); caret scrolling then uses the real viewport.
    const finalize = (w: HTMLElement): void => {
      w.style.columnFill = "balance";
      w.style.overflow = "clip";
    };
    const isManual = (el: Element) => el.classList.contains("docx-pagebreak") && el.getAttribute("data-docx-pagebreak") === "manual";

    // Bucket the blocks into page wrappers, each reserving reserveOf(pageIndex) at its bottom for
    // footnotes. Re-bucket without ever detaching a block from the document: each appendChild moves
    // a block straight from its old parent into the new wrapper, so the node holding the caret stays
    // connected and the selection is not collapsed.
    const EPS = 2;
    const bucket = (reserveOf: (i: number) => number): HTMLElement[] => {
      const ws: HTMLElement[] = [];
      let wrap = newWrapper(true, reserveOf(0));
      doc.appendChild(wrap);
      ws.push(wrap);
      for (const block of blocks) {
        if (isManual(block) && wrap.children.length) {
          finalize(wrap);
          wrap = newWrapper(false, reserveOf(ws.length));
          doc.appendChild(wrap);
          ws.push(wrap);
          wrap.appendChild(block);
          continue;
        }
        wrap.appendChild(block);
        // scrollWidth grows when content spills into an (N+1)th column: this page is full.
        if (wrap.scrollWidth > wrap.clientWidth + EPS && wrap.children.length > 1) {
          wrap.removeChild(block);
          finalize(wrap);
          wrap = newWrapper(false, reserveOf(ws.length));
          doc.appendChild(wrap);
          ws.push(wrap);
          wrap.appendChild(block); // a lone oversized block stays even if it still overflows
        }
      }
      finalize(wrap);
      return ws;
    };
    // Two passes when there are footnotes: the first finds which page each note lands on, the second
    // re-buckets with that page's reserve baked in so the body stops above the footnote strip.
    const footRefs = footnoteRefs();
    footnotesPerPage = true;
    const { fnRow, fnSize } = measureFootnotes(footRefs, false, contentWidth);
    const pageOf = (ws: HTMLElement[]) => (ref: HTMLElement) => ws.findIndex((w) => w.contains(ref));
    let wrappers = bucket(() => 0);
    let reserve: number[] = [];
    if (footRefs.length) {
      reserve = footnoteReserve(footRefs, fnSize, pageOf(wrappers));
      wrappers = bucket((i) => reserve[i] || 0);
      reserve = footnoteReserve(footRefs, fnSize, pageOf(wrappers));
    }
    // Drop every wrapper that is not part of the final layout, plus leftover spacers.
    for (const old of Array.from(doc.children) as HTMLElement[]) {
      if ((old.classList.contains(COLPAGE) && !wrappers.includes(old)) || old.classList.contains("docxedit-pagespacer")) old.remove();
    }

    pagelayer.replaceChildren();
    hflayer.replaceChildren();
    const cardCount = wrappers.length;
    for (let p = 0; p < cardCount; p++) {
      const base = p * pageStep;
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.top = `${base}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
      cloneHF(p, cardCount, `top:${base + geometry.margin.top}px;left:0;width:100%`, `top:${base + geometry.heightPx - contentBottomInset}px;left:0;width:100%`);
    }
    // Footnotes: a full-width area below all columns, just above the footer, on each page.
    drawFootnoteAreas(footRefs, fnRow, false, pageOf(wrappers),
      (p) => `top:${p * pageStep + (geometry.heightPx - contentBottomInset) - (reserve[p] || 0)}px;left:${geometry.margin.left}px;width:${contentWidth}px`);
    page.style.minHeight = `${cardCount * pageStep - PAGE_GAP}px`;
    decorateFields(cardCount, pageStep, false);
    if (caretBlock && caretBlock.isConnected) placeCaret(caretBlock, caretOffset);
  };

  // Per-section pagination: a document with mid-document section breaks can mix page sizes /
  // orientation / margins / columns. Each section's blocks are bucketed into editable page boxes
  // sized to that section (the box overflows when full, by height for one column or by the
  // (N+1)th column for several), and the boxes are centred + stacked. The caret is preserved as
  // a (block, char-offset) pair across the reparent, like the column reflow.
  const SECPAGE = "docxedit-secpage";
  const docGeom = (): SecGeom => ({ w: geometry.widthPx, h: geometry.heightPx, mt: geometry.margin.top, mr: geometry.margin.right, mb: geometry.margin.bottom, ml: geometry.margin.left, cols: geometry.columns, colGap: geometry.columnGapPx, vertical: geometry.vertical, rtl: geometry.rtl });
  // Resolve a section's geometry from its (possibly partial) JSON: size + margins fall back to the
  // document, but section-specific fields (columns, direction) are taken only from the section, so
  // a section that omits them does NOT inherit the document's columns / writing direction.
  const mergeSecGeom = (j: Partial<SecGeom>): SecGeom => {
    const d = docGeom();
    return { w: j.w ?? d.w, h: j.h ?? d.h, mt: j.mt ?? d.mt, mr: j.mr ?? d.mr, mb: j.mb ?? d.mb, ml: j.ml ?? d.ml, cols: j.cols, colGap: j.colGap, vertical: j.vertical, rtl: j.rtl };
  };
  const repaginateSections = () => {
    const blocks: HTMLElement[] = [];
    const collect = (parent: HTMLElement) => {
      for (const c of Array.from(parent.children) as HTMLElement[]) {
        if (c.classList.contains(VBAND)) collect(c); // descend vertical band wrappers to the real blocks
        else blocks.push(c);
      }
    };
    for (const child of Array.from(doc.children) as HTMLElement[]) {
      if (child.classList.contains(SECPAGE)) collect(child);
      else if (!child.classList.contains("docxedit-pagespacer")) blocks.push(child);
    }
    const sel = window.getSelection();
    let caretBlock: HTMLElement | null = null;
    let caretOffset = 0;
    if (sel && sel.rangeCount) {
      const r = sel.getRangeAt(0);
      const blk = blocks.find((b) => b === r.startContainer || b.contains(r.startContainer));
      if (blk) {
        caretBlock = blk;
        caretOffset = charOffsetIn(blk, r.startContainer, r.startOffset);
      }
    }
    // Segment into sections. Two conventions: docx marks the LAST paragraph of a section
    // (data-rdoc-secbreak = that section's geometry); odt marks the FIRST paragraph of a new
    // section (data-rdoc-secstart = the new section's geometry). The leading run uses the
    // document geometry.
    const parseGeom = (s: string | null): SecGeom => {
      try { return mergeSecGeom(JSON.parse(s ?? "")); } catch { return docGeom(); }
    };
    // Each section resolves its own header/footer source band: a keyed per-section band when its
    // boundary paragraph names one (data-rdoc-sec*key), else the document default. Distinct
    // sections therefore edit + save distinct header/footer parts.
    const resolveBand = (key: string | null, fallback: HTMLElement | null): HTMLElement | null =>
      key && secBands.has(key) ? secBands.get(key)!.el : fallback;
    // boundaryEl is the paragraph carrying the section's keys (null for the main section, which
    // has no break and inherits the document header/footer + cannot toggle its link).
    type Section = { blocks: HTMLElement[]; geom: SecGeom; headerEl: HTMLElement | null; footerEl: HTMLElement | null; boundaryEl: HTMLElement | null };
    const sections: Section[] = [];
    let run: HTMLElement[] = [];
    let runGeom = docGeom();
    let runHKey: string | null = null;
    let runFKey: string | null = null;
    let runBoundary: HTMLElement | null = null;
    const pushRun = () => {
      if (run.length) sections.push({ blocks: run, geom: runGeom, headerEl: resolveBand(runHKey, header), footerEl: resolveBand(runFKey, footer), boundaryEl: runBoundary });
      run = [];
    };
    for (const b of blocks) {
      const start = b.getAttribute("data-rdoc-secstart");
      if (start) {
        pushRun(); // close the previous run with its own keys
        runGeom = parseGeom(start); // this block begins a section with this geometry
        runHKey = b.getAttribute("data-rdoc-secheaderkey");
        runFKey = b.getAttribute("data-rdoc-secfooterkey");
        runBoundary = b; // leading convention: the section starts here
      }
      run.push(b);
      const brk = b.getAttribute("data-rdoc-secbreak");
      if (brk) {
        // Trailing convention: this block ends the section and carries its keys.
        sections.push({ blocks: run, geom: parseGeom(brk), headerEl: resolveBand(b.getAttribute("data-rdoc-secheaderkey"), header), footerEl: resolveBand(b.getAttribute("data-rdoc-secfooterkey"), footer), boundaryEl: b });
        run = [];
        runGeom = docGeom();
        runHKey = null;
        runFKey = null;
        runBoundary = null;
      }
    }
    pushRun();

    const maxW = Math.max(geometry.widthPx, ...sections.map((s) => s.geom.w));
    pagelayer.replaceChildren();
    hflayer.replaceChildren();
    page.style.width = `${maxW}px`;
    doc.style.cssText = `display:flex;flex-direction:column;align-items:center;gap:${PAGE_GAP}px;padding:0`;

    const EPS = 2;
    // A section reserves space at its top/bottom for its header/footer band (0 when absent/empty)
    // so the body does not overlap them.
    const bandH = (el: HTMLElement | null, contentWidth: number): number => {
      if (!el || isBandEmpty(el)) return 0;
      measure.style.width = `${contentWidth}px`;
      return el.offsetHeight;
    };
    const newBox = (g: SecGeom, hH: number, fH: number, plain = false, reserve = 0): HTMLElement => {
      const b = document.createElement("div");
      b.className = SECPAGE;
      b.style.width = `${g.w}px`;
      b.style.height = `${g.h}px`;
      b.style.boxSizing = "border-box";
      // Reserve header/footer space, plus the footnote strip at the block-axis end: the bottom for a
      // horizontal section, the left edge for a vertical one (where its right-to-left flow runs out).
      const padB = g.mb + fH + (g.vertical ? 0 : reserve);
      const padL = g.ml + (g.vertical ? reserve : 0);
      b.style.padding = `${g.mt + hH}px ${g.mr}px ${padB}px ${padL}px`;
      b.setAttribute("data-rdoc-secgeom", JSON.stringify(g)); // for the per-page ruler
      b.style.overflow = "hidden"; // a scroll container during bucketing; clipped on finalize
      if (plain) return b; // a frame holding manually-laid bands (vertical multi-column)
      if (g.vertical) b.style.writingMode = "vertical-rl"; // tategaki section: text fills top-down, rl
      else if (g.rtl) b.style.direction = "rtl";
      if (g.cols && g.cols > 1) {
        b.style.columnCount = String(g.cols);
        b.style.columnGap = `${g.colGap ?? 36}px`;
        b.style.columnFill = "auto";
      }
      return b;
    };
    type BoxMeta = { box: HTMLElement; g: SecGeom; hH: number; fH: number; headerEl: HTMLElement | null; footerEl: HTMLElement | null; boundaryEl: HTMLElement | null };
    // Vertical + multi-column section: lay N vertical-rl bands stacked in each page box (CSS multicol
    // does not fragment vertical text), bucketing blocks by block-axis overflow like the columns path.
    // `reserveOf(boxIndex)` narrows the bands to leave the page's left footnote strip free.
    const layoutVCols = (sec: Section, cw: number, hH: number, fH: number, meta: Omit<BoxMeta, "box">, boxMeta: BoxMeta[], reserveOf: (i: number) => number): void => {
      const g = sec.geom, N = g.cols!, gap = g.colGap ?? 36;
      const bandHeight = (g.h - (g.mt + hH) - (g.mb + fH) - (N - 1) * gap) / N;
      const mkBand = (slot: number, bandWidth: number): HTMLElement => {
        const w = document.createElement("div");
        w.className = VBAND;
        w.style.cssText = `writing-mode:vertical-rl;width:${bandWidth}px;height:${bandHeight}px;overflow:hidden${slot < N - 1 ? `;margin-bottom:${gap}px` : ""}`;
        return w;
      };
      let reserve = reserveOf(boxMeta.length);
      let box = newBox(g, hH, fH, true, reserve);
      box.style.overflow = "clip";
      doc.appendChild(box);
      boxMeta.push({ box, ...meta });
      // vertical-rl overflows leftward, so scrollWidth stays equal to clientWidth; detect overflow
      // by the just-added block crossing the band's left edge instead.
      const vOver = (b: HTMLElement, block: HTMLElement): boolean => block.getBoundingClientRect().left < b.getBoundingClientRect().left - EPS;
      let slot = 0, band = mkBand(0, cw - reserve);
      box.appendChild(band);
      for (const block of sec.blocks) {
        band.appendChild(block);
        if (vOver(band, block) && band.children.length > 1) {
          band.removeChild(block);
          band.style.overflow = "clip";
          if (++slot >= N) { reserve = reserveOf(boxMeta.length); box = newBox(g, hH, fH, true, reserve); box.style.overflow = "clip"; doc.appendChild(box); boxMeta.push({ box, ...meta }); slot = 0; }
          band = mkBand(slot, cw - reserve);
          box.appendChild(band);
          band.appendChild(block);
        }
      }
      band.style.overflow = "clip";
    };
    const finalize = (b: HTMLElement, g: SecGeom): void => {
      b.style.overflow = "clip";
      if (g.cols && g.cols > 1) b.style.columnFill = "balance";
    };
    // A vertical or multi-column box fills along the block axis (width); a plain box fills by height.
    const overflowed = (b: HTMLElement, g: SecGeom): boolean =>
      g.vertical || (g.cols && g.cols > 1) ? b.scrollWidth > b.clientWidth + EPS : b.scrollHeight > b.clientHeight + EPS;

    // Bucket every section into its page boxes, each reserving reserveOf(boxIndex) for footnotes.
    const buildBoxes = (reserveOf: (i: number) => number): BoxMeta[] => {
      const boxMeta: BoxMeta[] = [];
      for (const sec of sections) {
        const cw = sec.geom.w - sec.geom.ml - sec.geom.mr;
        const hH = bandH(sec.headerEl, cw);
        const fH = bandH(sec.footerEl, cw);
        const meta = { g: sec.geom, hH, fH, headerEl: sec.headerEl, footerEl: sec.footerEl, boundaryEl: sec.boundaryEl };
        if (sec.geom.vertical && (sec.geom.cols ?? 0) > 1) { layoutVCols(sec, cw, hH, fH, meta, boxMeta, reserveOf); continue; }
        let box = newBox(sec.geom, hH, fH, false, reserveOf(boxMeta.length));
        doc.appendChild(box);
        boxMeta.push({ box, ...meta });
        for (const block of sec.blocks) {
          box.appendChild(block);
          if (overflowed(box, sec.geom) && box.children.length > 1) {
            box.removeChild(block);
            finalize(box, sec.geom);
            box = newBox(sec.geom, hH, fH, false, reserveOf(boxMeta.length));
            doc.appendChild(box);
            boxMeta.push({ box, ...meta });
            box.appendChild(block); // a lone oversized block stays even if it still overflows
          }
        }
        finalize(box, sec.geom);
      }
      return boxMeta;
    };
    // Two passes when there are footnotes: bucket once to find each note's box (measuring it at that
    // box's content extent and direction), then re-bucket with each box's reserve baked into its
    // padding so the body stops before the footnote strip.
    const footRefs = footnoteRefs();
    footnotesPerPage = true;
    const colHeightOf = (m: BoxMeta) => m.g.h - (m.g.mt + m.hH) - (m.g.mb + m.fH);
    const boxIndexOf = (bm: BoxMeta[]) => (ref: HTMLElement) => bm.findIndex((m) => m.box.contains(ref));
    const fnRow = new Map<string, HTMLElement>();
    const fnSize = new Map<string, number>();
    let boxMeta = buildBoxes(() => 0);
    let reserve: number[] = [];
    if (footRefs.length) {
      boxMeta.forEach((m) => {
        const refs = footRefs.filter((fr) => m.box.contains(fr.ref));
        if (!refs.length) return;
        const meas = measureFootnotes(refs, !!m.g.vertical, m.g.vertical ? colHeightOf(m) : m.g.w - m.g.ml - m.g.mr);
        for (const [id, row] of meas.fnRow) fnRow.set(id, row);
        for (const [id, sz] of meas.fnSize) fnSize.set(id, sz);
      });
      reserve = footnoteReserve(footRefs, fnSize, boxIndexOf(boxMeta));
      boxMeta = buildBoxes((i) => reserve[i] || 0);
      reserve = footnoteReserve(footRefs, fnSize, boxIndexOf(boxMeta));
    }
    for (const old of Array.from(doc.children) as HTMLElement[]) {
      if ((old.classList.contains(SECPAGE) && !boxMeta.some((m) => m.box === old)) || old.classList.contains(VBAND) || old.classList.contains("docxedit-pagespacer")) old.remove();
    }
    // Header/footer clones, positioned over each section box (page-numbered cumulatively). Each box
    // clones its own section's bands; a non-main section also gets a corner chip to link/unlink.
    const dx = doc.offsetLeft, dy = doc.offsetTop;
    const total = boxMeta.length;
    boxMeta.forEach(({ box, g, fH, headerEl, footerEl, boundaryEl }, i) => {
      const cw = g.w - g.ml - g.mr;
      const left = dx + box.offsetLeft + g.ml;
      // A section using the document default gets the first/even variant for its page; a section with
      // its own band keeps it. (i is the cumulative 0-based page index.)
      const hSrc = headerEl === header ? pickHF("header", i) : headerEl;
      const fSrc = footerEl === footer ? pickHF("footer", i) : footerEl;
      if (hSrc && !isBandEmpty(hSrc)) {
        const top = dy + box.offsetTop + g.mt;
        const hc = mkClone(hSrc, `top:${top}px;left:${left}px;width:${cw}px`);
        setCloneFields(hc, i + 1, total);
        hflayer.appendChild(hc);
        if (boundaryEl) hflayer.appendChild(mkLinkChip("header", boundaryEl, `top:${top}px;left:${left + cw - 18}px`));
      }
      if (fSrc && !isBandEmpty(fSrc)) {
        const top = dy + box.offsetTop + g.h - g.mb - fH;
        const fc = mkClone(fSrc, `top:${top}px;left:${left}px;width:${cw}px`);
        setCloneFields(fc, i + 1, total);
        hflayer.appendChild(fc);
        if (boundaryEl) hflayer.appendChild(mkLinkChip("footer", boundaryEl, `top:${top}px;left:${left + cw - 18}px`));
      }
    });
    // Footnotes: per section box, an area at its block-axis end (bottom for a horizontal section,
    // a left band for a vertical one), holding the notes whose references land in that box.
    boxMeta.forEach((m, i) => {
      const refs = footRefs.filter((fr) => m.box.contains(fr.ref));
      if (!refs.length) return;
      const g = m.g, cw = g.w - g.ml - g.mr, left = dx + m.box.offsetLeft + g.ml;
      const area = document.createElement("div");
      area.className = g.vertical ? "docxedit-fnarea is-vertical" : "docxedit-fnarea";
      area.style.cssText = g.vertical
        ? `top:${dy + m.box.offsetTop + g.mt + m.hH}px;left:${left}px;width:${reserve[i] || 0}px;height:${colHeightOf(m)}px;${noteAreaCss}`
        : `top:${dy + m.box.offsetTop + g.h - g.mb - m.fH - (reserve[i] || 0)}px;left:${left}px;width:${cw}px;${noteAreaCss}`;
      for (const fr of refs) { const row = fnRow.get(fr.id); if (row) area.appendChild(row); }
      hflayer.appendChild(area);
    });
    page.style.minHeight = "";
    decorateFields(0, 0, false); // refresh field text (page-number fields show no count here)
    if (caretBlock && caretBlock.isConnected) placeCaret(caretBlock, caretOffset);
  };

  const repaginate = () => {
    if (!paginated || editingBand) return;
    footnotesPerPage = false; // only the single-section path places footnotes per page
    if (doc.querySelector("[data-rdoc-secbreak], [data-rdoc-secstart]")) return repaginateSections();
    if (isVertical()) return (geometry.columns ?? 0) > 1 ? repaginateVerticalColumns() : repaginateVertical();
    // Unwrap any vertical band wrappers left by a prior vertical-columns layout + reset doc sizing.
    for (const w of Array.from(doc.querySelectorAll<HTMLElement>(`.${VBAND}`))) {
      while (w.firstChild) doc.insertBefore(w.firstChild, w);
      w.remove();
    }
    doc.style.width = "";
    if ((geometry.columns ?? 0) > 1) return repaginateColumns();
    // Single column: unwrap any column/section wrappers left by a previous layout (e.g. after a
    // page-setup change drops the column count back to 1) so the blocks are flat again.
    for (const w of Array.from(doc.querySelectorAll<HTMLElement>(`.${COLPAGE}, .${SECPAGE}`))) {
      while (w.firstChild) doc.insertBefore(w.firstChild, w);
      w.remove();
    }
    for (const s of Array.from(doc.querySelectorAll(":scope > .docxedit-pagespacer"))) s.remove();
    pagelayer.replaceChildren();
    hflayer.replaceChildren();

    // measure header/footer at full page width (their own padding provides the margins)
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = hfHeight("header");
    const footerH = hfHeight("footer");
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    const contentHeight = geometry.heightPx - contentTop - contentBottomInset;
    const pageStep = geometry.heightPx + PAGE_GAP;

    // place the body inside each page's content box (overrides the CSS-var padding)
    doc.style.height = ""; // clear any inline height left by a prior vertical layout
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

    // Footnotes: measure each referenced footnote body at content width, then reserve that space at
    // the bottom of the page its reference lands on (paginate accounts for it). Endnotes go to the
    // doc-end notes area (renderNotes), not per page.
    const cw = geometry.widthPx - geometry.margin.left - geometry.margin.right;
    const footRefs = footnoteRefs();
    footnotesPerPage = true;
    const kidIndexOf = (ref: HTMLElement) => kids.findIndex((k) => k.contains(ref));
    const { fnRow, fnSize } = measureFootnotes(footRefs, false, cw);
    const reserveFor = (pob: number[]): number[] => footnoteReserve(footRefs, fnSize, (ref) => { const ki = kidIndexOf(ref); return ki < 0 ? -1 : pob[ki]!; });
    let reserve = footRefs.length ? reserveFor(paginate(heights, { pageStep, contentHeight }, forceBreakBefore).pageOfBlock) : [];
    const { spacerBefore, cardCount, pageOfBlock } = paginate(heights, { pageStep, contentHeight, reserveOf: (p) => reserve[p] || 0 }, forceBreakBefore);
    reserve = footRefs.length ? reserveFor(pageOfBlock) : [];

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

    for (let p = 0; p < cardCount; p++) {
      const base = p * pageStep;
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.top = `${base}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
      cloneHF(p, cardCount, `top:${base + geometry.margin.top}px;left:0;width:100%`, `top:${base + geometry.heightPx - contentBottomInset}px;left:0;width:100%`);
    }

    // Footnotes: a per-page area just above the footer, holding the page's referenced notes.
    drawFootnoteAreas(footRefs, fnRow, false, (ref) => { const ki = kidIndexOf(ref); return ki < 0 ? -1 : pageOfBlock[ki]!; },
      (p) => `top:${p * pageStep + (geometry.heightPx - contentBottomInset) - (reserve[p] || 0)}px;left:${geometry.margin.left}px;width:${cw}px`);

    page.style.minHeight = `${cardCount * pageStep - PAGE_GAP}px`;
    decorateFields(cardCount, pageStep, false);
  };

  // Body HTML for saving: the live doc minus pagination artifacts (inert spacers and the
  // transient page-top class the engine adds for alignment).
  const cleanBody = (): string => {
    if (!doc.querySelector(".docxedit-pagespacer, .docxedit-pagetop, .docxedit-colpage, .docxedit-secpage, .docxedit-vband")) return doc.innerHTML;
    const tmp = doc.cloneNode(true) as HTMLElement;
    for (const s of Array.from(tmp.querySelectorAll(".docxedit-pagespacer"))) s.remove();
    for (const el of Array.from(tmp.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");
    // Unwrap the per-page column / per-section page / vertical band boxes, lifting blocks to the body.
    for (const w of Array.from(tmp.querySelectorAll(".docxedit-colpage, .docxedit-secpage, .docxedit-vband"))) {
      while (w.firstChild) w.parentNode!.insertBefore(w.firstChild, w);
      w.remove();
    }
    return tmp.innerHTML;
  };

  let afterReflow = () => {}; // assigned once the toolbar exists (toggles change buttons)
  const reflow = () => {
    if (editingBand) return; // don't yank the band currently being edited
    repaginate();
    renderNotes();
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
      // Vertical pages share one y band (they stack along x); horizontal pages stack along y.
      const within = isVertical()
        ? (e.clientY - page.getBoundingClientRect().top) / z
        : ((e.clientY - page.getBoundingClientRect().top) / z) % (geometry.heightPx + PAGE_GAP);
      if (within < 0 || within > geometry.heightPx) return; // outside a page (in the gap)
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
  if (isVertical()) {
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

  // Image select/resize/delete + insert live in the images module; the layout toolbar (wrap
  // mode, alignment, alt, drag-to-position) is a sibling module driven by the selection.
  let images: ReturnType<typeof setupImages>;
  const imageLayout = caps.images
    ? setupImageLayout({ wrap, doc, scroll, mark, getZoom: effectiveZoom, reposition: () => images.repositionHandles() })
    : null;
  images = setupImages({ wrap, scroll, regions, mark, getActiveEl: () => activeEl, onSelect: imageLayout?.onSelect });
  const { insertImage } = images;
  // Emit inline CSS (text-align, font-weight, ...) the serializer reads back, not legacy tags.
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch {
    /* not supported; legacy tags still round-trip */
  }

  // Clicking a table-of-contents row scrolls to its heading (rows match headings in order); clicking
  // a cross-reference jumps to its target bookmark / heading.
  doc.addEventListener("click", (e) => {
    const xref = (e.target as HTMLElement).closest?.(".docx-xref") as HTMLElement | null;
    if (xref) {
      const name = xref.getAttribute("data-rdoc-xref");
      const target = name ? Array.from(doc.querySelectorAll<HTMLElement>("[data-rdoc-bm]")).find((el) => el.getAttribute("data-rdoc-bm") === name) : null;
      target?.scrollIntoView({ block: "center", behavior: "smooth" });
      return;
    }
    const row = (e.target as HTMLElement).closest?.(".docx-field-toc-row") as HTMLElement | null;
    if (!row || !row.parentElement) return;
    const i = Array.from(row.parentElement.querySelectorAll(".docx-field-toc-row")).indexOf(row);
    const headings = Array.from(doc.querySelectorAll<HTMLElement>("h1,h2,h3")).filter((h) => !h.closest(".docx-field-toc"));
    headings[i]?.scrollIntoView({ block: "center", behavior: "smooth" });
  });

  // Remove a footnote/endnote by deleting its reference mark: the mark is an atomic,
  // non-selectable superscript, so Backspace just after it (or Delete just before it) would
  // otherwise be swallowed. Removing the reference drops the note from the view and from the save
  // (getBytes keeps only notes whose reference is still present).
  const removeAdjacentNoteRef = (e: KeyboardEvent): void => {
    if (e.key !== "Backspace" && e.key !== "Delete") return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (!range.collapsed) return;
    const c = range.startContainer, o = range.startOffset;
    const adj = e.key === "Backspace"
      ? (c.nodeType === 1 ? (o > 0 ? c.childNodes[o - 1] : null) : (o === 0 ? c.previousSibling : null))
      : (c.nodeType === 1 ? c.childNodes[o] : (o === (c.textContent?.length ?? 0) ? c.nextSibling : null));
    if (adj && adj.nodeType === 1 && (adj as HTMLElement).classList.contains("docx-fnref")) {
      e.preventDefault();
      (adj as HTMLElement).remove();
      mark(); // reflow drops the orphaned note; the save omits it
    }
  };
  for (const r of regions) {
    r.addEventListener("input", mark);
    r.addEventListener("keydown", removeAdjacentNoteRef);
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
    positionCards, addThreadCard, setActiveComment, allocId, freshParaId, insertImage, styleBar: bottomLeft,
    newStyles, newStyleCss, vertical: isVertical(),
    insertSectionBreak: sectionBreakBtn ? () => sectionBreakBtn.click() : null,
    insertNote: (kind, text) => insertNote(kind, text),
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
      // First/even variants: saved whenever their flag is on (toggling off drops them), even when
      // empty, so an enabled-but-blank variant means a blank first/even page (which odt can only
      // represent by writing the empty element). The sentinel paths tell the adapter the slot.
      if (geometry.titlePage && headerFirst) editedParts.push({ path: parts.headerFirst?.path ?? "header:first", html: headerFirst.innerHTML });
      if (geometry.titlePage && footerFirst) editedParts.push({ path: parts.footerFirst?.path ?? "footer:first", html: footerFirst.innerHTML });
      if (geometry.evenOdd && headerEven) editedParts.push({ path: parts.headerEven?.path ?? "header:even", html: headerEven.innerHTML });
      if (geometry.evenOdd && footerEven) editedParts.push({ path: parts.footerEven?.path ?? "footer:even", html: footerEven.innerHTML });
      for (const { el, path } of secBands.values()) if (!isBandEmpty(el)) editedParts.push({ path, html: el.innerHTML }); // distinct per-section bands
      // Footnote / endnote bodies still referenced in the body, in document order.
      const refIds = new Set(Array.from(doc.querySelectorAll(".docx-fnref")).map((r) => r.getAttribute("data-fn-id")));
      const notes: Note[] = [];
      for (const [id, nb] of noteBands) if (refIds.has(id)) notes.push({ id, kind: nb.kind, html: nb.el.innerHTML });

      return adapter.write(cleanBody(), editedParts, getEdits(), geometryDirty ? geometry : undefined, newStyles, notes);
    },
    destroy() {
      for (const u of fontUrls) URL.revokeObjectURL(u);
      window.clearTimeout(reflowTimer);
      repositionObserver.disconnect();
      teardownToolbar();
      tableEdit?.teardown();
      imageLayout?.teardown();
      teardownPageView();
      wrap.remove();
    },
  };
}
