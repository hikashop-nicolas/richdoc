// Toolbar: the formatting controls and their wiring. Builds the button/select factories,
// the bold/italic/underline + list + alignment commands (execCommand), the paragraph-style,
// font, size and colour pickers, link / page-break / image / comment actions, hosts the
// track-changes setup (which needs the button factories), and lays the single row out with
// an overflow "more" popover. The feature behaviours it drives come in as deps.
import { t } from "../i18n";
import { firstFontFamily, fontSizeToHalfPt } from "../util";
import { setupTrackChanges } from "./track-changes";
import type { Adapter, Capabilities, CommentThread, EditorOptions, RichDoc } from "../types";

export interface ToolbarDeps {
  toolbar: HTMLElement;
  wrap: HTMLElement;
  doc: HTMLElement;
  regions: HTMLElement[];
  caps: Capabilities;
  options: EditorOptions;
  parts: RichDoc;
  adapter: Adapter;
  getActiveEl: () => HTMLElement;
  mark: () => void;
  positionCards: () => void;
  addThreadCard: (thread: CommentThread) => HTMLElement;
  setActiveComment: (id: string | null) => void;
  allocId: () => string;
  freshParaId: () => string;
  insertImage: () => void;
  zoomSlider: HTMLElement;
  zoomLabel: HTMLElement;
}

export function setupToolbar(deps: ToolbarDeps) {
  const {
    toolbar, wrap, doc, regions, caps, options, parts, adapter, getActiveEl, mark, positionCards,
    addThreadCard, setActiveComment, allocId, freshParaId, insertImage, zoomSlider, zoomLabel,
  } = deps;

  const exec = (cmd: string, val?: string) => {
    getActiveEl().focus();
    document.execCommand(cmd, false, val);
    mark();
    syncToolbarState();
  };
  // Wrap the current selection in a span carrying one CSS property (for font size, which
  // has no execCommand equivalent in CSS mode). No-op on a collapsed selection.
  const styleSel = (prop: string, val: string) => {
    getActiveEl().focus();
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
    syncToolbarState();
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
  block.addEventListener("mousedown", () => getActiveEl().focus());
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
    s.addEventListener("mousedown", () => getActiveEl().focus());
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
  colorInput.addEventListener("mousedown", () => getActiveEl().focus());
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
  bgInput.addEventListener("mousedown", () => getActiveEl().focus());
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
    getActiveEl().focus();
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
  const imgIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
    '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

  // Add a comment over the current selection: wrap it in comment-range markers and a
  // reference marker that carries the text, so the serializer can build comments.xml.
  const addComment = () => {
    getActiveEl().focus();
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
    doc, wrap, regions, options, mark, positionCards, getActiveEl, iconBtn, btn,
  });

  // Toolbar: shared controls always shown; image/comment/page-break/track-changes are
  // gated by the adapter's capabilities so a format can hide what it cannot serialize.
  const linkBtn = iconBtn(linkIcon, t("linkAria"), () => {
    const url = prompt(t("linkPrompt"), "https://");
    if (url === null) return;
    if (url === "") exec("unlink");
    else exec("createLink", url);
  });
  const supIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<text x="0" y="14" font-size="11" font-family="serif">x</text><text x="8" y="7" font-size="7" font-family="serif">2</text></svg>';
  const subIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<text x="0" y="11" font-size="11" font-family="serif">x</text><text x="8" y="16" font-size="7" font-family="serif">2</text></svg>';
  // Named so their pressed state can be reflected from the caret (see syncToolbarState).
  const boldBtn = btn("B", t("bold"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold");
  const italicBtn = btn("I", t("italic"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic");
  const underlineBtn = btn("U", t("underline"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline");
  const strikeBtn = btn("S", t("strikethrough"), () => { beginFormatChange(); exec("strikeThrough"); }, "docxedit-tb-strike");
  const supBtn = iconBtn(supIcon, t("superscript"), () => { beginFormatChange(); exec("superscript"); });
  const subBtn = iconBtn(subIcon, t("subscript"), () => { beginFormatChange(); exec("subscript"); });
  const alignLeftBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), t("alignLeft"), () => exec("justifyLeft")) : null;
  const alignCenterBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), t("alignCenter"), () => exec("justifyCenter")) : null;
  const alignRightBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), t("alignRight"), () => exec("justifyRight")) : null;
  const alignJustifyBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), t("alignJustify"), () => exec("justifyFull")) : null;

  // Reflect the caret's formatting in the controls: the font / size / paragraph-style
  // pickers track the text under the caret, and bold/italic/underline/alignment show their
  // pressed state. Runs (rAF-coalesced) on every selection change and after a command.
  const queryState = (cmd: string): boolean => {
    try { return document.queryCommandState(cmd); } catch { return false; }
  };
  const syncToolbarState = () => {
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.anchorNode : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
    if (!el || !regions.some((r) => r.contains(el))) return; // only while editing a region
    const cs = getComputedStyle(el);
    if (caps.fontControls) {
      const fam = firstFontFamily(cs.fontFamily);
      fontSel.value = fam && FONTS.includes(fam) ? fam : "";
      const hp = fontSizeToHalfPt(cs.fontSize);
      const pt = hp ? String(Math.round(hp / 2)) : "";
      sizeSel.value = SIZES.includes(pt) ? pt : "";
    }
    const tag = el.closest("h1,h2,h3,p")?.tagName ?? "";
    block.value = tag === "H1" || tag === "H2" || tag === "H3" ? tag : "P";
    const setOn = (b: HTMLElement | null, on: boolean) => b?.classList.toggle("is-on", on);
    setOn(boldBtn, queryState("bold"));
    setOn(italicBtn, queryState("italic"));
    setOn(underlineBtn, queryState("underline"));
    setOn(strikeBtn, queryState("strikeThrough"));
    setOn(supBtn, queryState("superscript"));
    setOn(subBtn, queryState("subscript"));
    setOn(alignLeftBtn, queryState("justifyLeft"));
    setOn(alignCenterBtn, queryState("justifyCenter"));
    setOn(alignRightBtn, queryState("justifyRight"));
    setOn(alignJustifyBtn, queryState("justifyFull"));
  };
  // Coalesce bursts of selectionchange (e.g. during a drag-select) with a short timer.
  // setTimeout (not rAF) so it still fires when the tab is in the background.
  let syncTimer = 0;
  const scheduleSync = () => {
    window.clearTimeout(syncTimer);
    syncTimer = window.setTimeout(syncToolbarState, 16);
  };
  document.addEventListener("selectionchange", scheduleSync);

  const items: (Node | null)[] = [
    boldBtn,
    italicBtn,
    underlineBtn,
    strikeBtn,
    supBtn,
    subBtn,
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
    alignLeftBtn,
    alignCenterBtn,
    alignRightBtn,
    alignJustifyBtn,
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

  scheduleSync(); // reflect the initial caret position once mounted

  const teardown = () => {
    toolbarObserver.disconnect();
    document.removeEventListener("click", closeOverflow);
    document.removeEventListener("selectionchange", scheduleSync);
    window.clearTimeout(syncTimer);
  };
  return { updateChangeButtons, teardown };
}
