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
  const BUILTIN_BLOCKS = new Set(["P", "H1", "H2", "H3"]);
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(t(key), v));
  }
  // The document's own named paragraph styles (Title, Quote, ...), value = style id with a
  // "s:" prefix so it never collides with the built-in block tags.
  const namedStyles = parts.paragraphStyles ?? [];
  if (namedStyles.length) {
    const grp = document.createElement("optgroup");
    grp.label = t("documentStyles");
    for (const s of namedStyles) grp.appendChild(new Option(s.name, `s:${s.id}`));
    block.appendChild(grp);
  }
  block.addEventListener("mousedown", () => getActiveEl().focus());
  block.addEventListener("change", () => {
    const v = block.value;
    if (BUILTIN_BLOCKS.has(v)) {
      // a built-in block clears any named style on the affected paragraphs
      exec("formatBlock", v);
      for (const b of selectedBlocks()) b.removeAttribute("data-rdoc-style");
      mark();
    } else if (v.startsWith("s:")) {
      // a named paragraph style: drop heading state, then tag the blocks with the style id
      getActiveEl().focus();
      document.execCommand("formatBlock", false, "P");
      for (const b of selectedBlocks()) b.setAttribute("data-rdoc-style", v.slice(2));
      mark();
      syncToolbarState();
    }
  });

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

  // Fonts/sizes the document actually uses but that are not in the defaults are added to
  // the pickers, so the caret-sync can show them instead of falling back to the placeholder.
  const BASE_FONTS = ["Arial", "Calibri", "Century", "Courier New", "Georgia", "Times New Roman", "Verdana"];
  const BASE_SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48"];
  const docFonts = new Set<string>();
  const docSizes = new Set<string>();
  for (const el of Array.from(doc.querySelectorAll<HTMLElement>("[style]"))) {
    const fam = firstFontFamily(el.style.fontFamily);
    if (fam) docFonts.add(fam);
    const hp = fontSizeToHalfPt(el.style.fontSize);
    if (hp) docSizes.add(String(Math.round(hp / 2)));
  }
  if (parts.defaultFont) docFonts.add(parts.defaultFont);
  const FONTS = [...BASE_FONTS, ...[...docFonts].filter((f) => !BASE_FONTS.includes(f)).sort()];
  const SIZES = [...new Set([...BASE_SIZES, ...docSizes])].sort((a, b) => Number(a) - Number(b));
  const fontSel = pickerSelect(t("font"), FONTS.map((f) => [f, f] as [string, string]), (v) => {
    beginFormatChange();
    exec("fontName", v);
  });

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

  // Paragraph indent + line spacing operate on the block(s) intersecting the selection.
  const BLOCK_SEL = "p,h1,h2,h3,h4,h5,h6,li,blockquote,div";
  const selectedBlocks = (): HTMLElement[] => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return [];
    const region = getActiveEl();
    const blockOf = (n: Node | null): HTMLElement | null => {
      const el = n ? (n.nodeType === 3 ? n.parentElement : (n as Element)) : null;
      const b = el?.closest?.(BLOCK_SEL) as HTMLElement | null;
      return b && region.contains(b) && !b.classList.contains("docxedit-doc") ? b : null;
    };
    const range = sel.getRangeAt(0);
    const start = blockOf(range.startContainer);
    const end = blockOf(range.endContainer);
    const out: HTMLElement[] = start ? [start] : [];
    let n = start;
    while (n && n !== end) {
      n = n.nextElementSibling as HTMLElement | null;
      if (n && n.matches(BLOCK_SEL)) out.push(n);
    }
    if (end && !out.includes(end)) out.push(end);
    return out;
  };
  const indentIcon = (dir: number): string =>
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<rect x="2" y="2.5" width="12" height="1.6" rx=".6"/><rect x="7" y="6" width="7" height="1.6" rx=".6"/>' +
    '<rect x="7" y="9.4" width="7" height="1.6" rx=".6"/><rect x="2" y="12.9" width="12" height="1.6" rx=".6"/>' +
    (dir > 0 ? '<path d="M2 6.2l2.6 1.9L2 10z"/>' : '<path d="M4.6 6.2L2 8.1l2.6 1.9z"/>') + "</svg>";
  const STEP = 48; // 0.5in
  const marginAdjust = (b: HTMLElement, dir: 1 | -1) => {
    const next = Math.max(0, (parseFloat(b.style.marginLeft) || 0) + dir * STEP);
    b.style.marginLeft = next ? `${next}px` : "";
  };
  // Indent a list item by sinking it into a sublist on its previous sibling (real nesting,
  // so it round-trips as w:ilvl / a deeper text:list rather than a margin).
  const indentLi = (li: HTMLElement) => {
    const prev = li.previousElementSibling as HTMLElement | null;
    if (!prev || prev.tagName !== "LI") return marginAdjust(li, 1); // first item can't nest
    const parentTag = (li.parentElement?.tagName ?? "UL").toLowerCase();
    let sub = prev.lastElementChild as HTMLElement | null;
    if (!sub || (sub.tagName !== "UL" && sub.tagName !== "OL")) {
      sub = li.ownerDocument.createElement(parentTag);
      prev.appendChild(sub);
    }
    sub.appendChild(li);
  };
  const outdentLi = (li: HTMLElement) => {
    const parentList = li.parentElement as HTMLElement | null;
    const grandLi = parentList?.parentElement as HTMLElement | null;
    if (!parentList || !grandLi || grandLi.tagName !== "LI") return marginAdjust(li, -1); // already top level
    const outerList = grandLi.parentElement as HTMLElement;
    // following siblings stay nested under the outdented item, preserving their depth
    let next: Element | null = li.nextElementSibling;
    if (next) {
      const newSub = li.ownerDocument.createElement(parentList.tagName.toLowerCase());
      while (next) {
        const after: Element | null = next.nextElementSibling;
        newSub.appendChild(next);
        next = after;
      }
      li.appendChild(newSub);
    }
    outerList.insertBefore(li, grandLi.nextElementSibling);
    if (!parentList.children.length) parentList.remove();
  };
  const adjustIndent = (dir: 1 | -1) => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) {
      if (b.tagName === "LI") dir === 1 ? indentLi(b) : outdentLi(b);
      else marginAdjust(b, dir);
    }
    mark();
  };
  const indentBtn = iconBtn(indentIcon(1), t("indent"), () => adjustIndent(1));
  const outdentBtn = iconBtn(indentIcon(-1), t("outdent"), () => adjustIndent(-1));

  // Line spacing: an icon button (like the others) opening a small menu of presets.
  const lineSpacingIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M6 4h8M6 8h8M6 12h8"/><path d="M2.6 3.6v8.8M1.4 4.8 2.6 3.6 3.8 4.8M1.4 11.2 2.6 12.4 3.8 11.2"/></svg>';
  const lineSpacingMenu = document.createElement("div");
  lineSpacingMenu.className = "docxedit-menu";
  lineSpacingMenu.hidden = true;
  const setLineHeight = (v: string): void => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) b.style.lineHeight = v;
    mark();
    lineSpacingMenu.hidden = true;
  };
  const lsButtons = ([["1", "1.0"], ["1.15", "1.15"], ["1.5", "1.5"], ["2", "2.0"]] as const).map(([v, label]) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.textContent = label;
    b.dataset.v = v;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      setLineHeight(v);
    });
    return b;
  });
  lineSpacingMenu.append(...lsButtons);
  // Paragraph space before/after: a default add, or an explicit 0 to remove inherited spacing.
  const SPACE_ADD = 14; // px, ~10.5pt, a sensible "add space" default
  const divider = document.createElement("div");
  divider.className = "docxedit-menu-sep";
  const setSpace = (side: "marginTop" | "marginBottom", on: boolean): void => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) b.style[side] = on ? `${SPACE_ADD}px` : "0px";
    mark();
    lineSpacingMenu.hidden = true;
  };
  const spaceItem = (side: "marginTop" | "marginBottom"): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.dataset.side = side;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      setSpace(side, b.dataset.on !== "1"); // toggle off when currently on
    });
    return b;
  };
  const spaceBeforeItem = spaceItem("marginTop");
  const spaceAfterItem = spaceItem("marginBottom");
  lineSpacingMenu.append(divider, spaceBeforeItem, spaceAfterItem);
  // Reflect the current block's state on the two toggle items (add vs remove).
  const refreshSpaceItems = (): void => {
    const b = selectedBlocks()[0];
    const on = (prop: "marginTop" | "marginBottom"): boolean => (parseFloat(b?.style[prop] ?? "") || 0) > 0;
    const beforeOn = on("marginTop");
    const afterOn = on("marginBottom");
    spaceBeforeItem.dataset.on = beforeOn ? "1" : "0";
    spaceAfterItem.dataset.on = afterOn ? "1" : "0";
    spaceBeforeItem.textContent = t(beforeOn ? "removeSpaceBefore" : "addSpaceBefore");
    spaceAfterItem.textContent = t(afterOn ? "removeSpaceAfter" : "addSpaceAfter");
  };
  const lineSpacingBtn = iconBtn(lineSpacingIcon, t("lineSpacing"), () => {});
  lineSpacingBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = lineSpacingBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    lineSpacingMenu.style.left = `${r.left - wr.left}px`;
    lineSpacingMenu.style.top = `${r.bottom - wr.top + 2}px`;
    if (lineSpacingMenu.hidden) refreshSpaceItems();
    lineSpacingMenu.hidden = !lineSpacingMenu.hidden;
  });

  // --- Insert table: a button opening a grid picker (drag/hover to size, click to insert) ---
  const tableIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><line x1="1.5" y1="6.2" x2="14.5" y2="6.2"/>' +
    '<line x1="1.5" y1="9.9" x2="14.5" y2="9.9"/><line x1="6" y1="2.5" x2="6" y2="13.5"/><line x1="10.3" y1="2.5" x2="10.3" y2="13.5"/></svg>';
  const insertTable = (rows: number, cols: number) => {
    getActiveEl().focus();
    const table = document.createElement("table");
    table.className = "docx-table";
    table.contentEditable = "false";
    // Fill the available content width with equal columns (a <colgroup> the adapters read).
    const cs = getComputedStyle(doc);
    const avail = Math.max(120, doc.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0));
    const colW = Math.floor(avail / cols);
    const colgroup = document.createElement("colgroup");
    for (let c = 0; c < cols; c++) {
      const col = document.createElement("col");
      col.style.width = `${colW}px`;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
    for (let r = 0; r < rows; r++) {
      const tr = table.insertRow();
      for (let c = 0; c < cols; c++) {
        const td = tr.insertCell();
        const div = document.createElement("div");
        div.className = "docx-cell";
        div.contentEditable = "true";
        div.innerHTML = "<br>";
        td.appendChild(div);
      }
    }
    // Insert as a top-level block after the block holding the caret, then a paragraph after.
    const sel = window.getSelection();
    const trailing = document.createElement("p");
    trailing.innerHTML = "<br>";
    let anchor: HTMLElement | null = null;
    if (sel && sel.rangeCount) {
      let el: Element | null = ((n) => (n.nodeType === 3 ? n.parentElement : (n as Element)))(sel.getRangeAt(0).startContainer);
      while (el && el.parentElement !== doc) el = el.parentElement;
      anchor = el as HTMLElement | null;
    }
    if (anchor && anchor.parentElement === doc) {
      doc.insertBefore(table, anchor.nextSibling);
      doc.insertBefore(trailing, table.nextSibling);
    } else {
      doc.append(table, trailing);
    }
    mark();
    const firstCell = table.querySelector(".docx-cell") as HTMLElement | null;
    if (firstCell) {
      const r = document.createRange();
      r.selectNodeContents(firstCell);
      r.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(r);
      firstCell.focus();
    }
  };
  const GRID_ROWS = 8;
  const GRID_COLS = 10;
  const tablePicker = document.createElement("div");
  tablePicker.className = "docxedit-table-picker";
  tablePicker.hidden = true;
  const tableGrid = document.createElement("div");
  tableGrid.className = "docxedit-table-grid";
  const tableLabel = document.createElement("div");
  tableLabel.className = "docxedit-table-label";
  tableLabel.textContent = t("insertTable");
  const squares: HTMLElement[] = [];
  const highlight = (rr: number, cc: number) => {
    for (let r = 0; r < GRID_ROWS; r++) for (let c = 0; c < GRID_COLS; c++) squares[r * GRID_COLS + c].classList.toggle("on", r <= rr && c <= cc);
    tableLabel.textContent = `${cc + 1} × ${rr + 1}`;
  };
  for (let r = 0; r < GRID_ROWS; r++) {
    for (let c = 0; c < GRID_COLS; c++) {
      const sq = document.createElement("div");
      sq.className = "docxedit-table-sq";
      sq.addEventListener("mouseenter", () => highlight(r, c));
      sq.addEventListener("mousedown", (e) => {
        e.preventDefault();
        insertTable(r + 1, c + 1);
        tablePicker.hidden = true;
      });
      tableGrid.appendChild(sq);
      squares.push(sq);
    }
  }
  tablePicker.append(tableGrid, tableLabel);
  const tableBtn = iconBtn(tableIcon, t("insertTable"), () => {});
  tableBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = tableBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    tablePicker.style.left = `${r.left - wr.left}px`;
    tablePicker.style.top = `${r.bottom - wr.top + 2}px`;
    tablePicker.hidden = !tablePicker.hidden;
    if (!tablePicker.hidden) highlight(-1, -1);
  });

  // --- Insert field: page number / count / table of contents (docx fields) ---------------
  const fieldIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<path d="M5 2.5C3.5 2.5 3.5 5 3.5 8s0 5.5-1.5 5.5"/><path d="M11 2.5c1.5 0 1.5 2.5 1.5 5.5s0 5.5 1.5 5.5"/></svg>';
  const insertAtCaret = (node: Node): void => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    getActiveEl().dispatchEvent(new Event("input", { bubbles: true })); // sync header/footer clone + reflow
  };
  const fieldSpan = (instr: string, text: string): HTMLElement => {
    const s = document.createElement("span");
    s.className = "docx-field";
    s.contentEditable = "false";
    s.setAttribute("data-field", instr);
    s.textContent = text;
    return s;
  };
  const insertField = (instr: string) => insertAtCaret(fieldSpan(instr, "1"));
  const insertPageXofY = () => {
    const frag = document.createDocumentFragment();
    frag.append(fieldSpan("PAGE", "1"), document.createTextNode(" / "), fieldSpan("NUMPAGES", "1"));
    insertAtCaret(frag);
  };
  const insertTOC = () => {
    const toc = document.createElement("div");
    toc.className = "docx-field-toc";
    toc.contentEditable = "false";
    toc.setAttribute("data-field", "TOC");
    toc.innerHTML = `<div class="docx-field-toc-empty">${t("tocEmpty")}</div>`;
    const trailing = document.createElement("p");
    trailing.innerHTML = "<br>";
    // The TOC belongs in the body: insert after the caret's top-level block, else at the top.
    const sel = window.getSelection();
    let anchor: HTMLElement | null = null;
    if (sel && sel.rangeCount && doc.contains(sel.getRangeAt(0).startContainer)) {
      let el: Element | null = ((n) => (n.nodeType === 3 ? n.parentElement : (n as Element)))(sel.getRangeAt(0).startContainer);
      while (el && el.parentElement !== doc) el = el.parentElement;
      anchor = el as HTMLElement | null;
    }
    if (anchor && anchor.parentElement === doc) {
      doc.insertBefore(toc, anchor.nextSibling);
      doc.insertBefore(trailing, toc.nextSibling);
    } else {
      doc.insertBefore(trailing, doc.firstChild);
      doc.insertBefore(toc, doc.firstChild);
    }
    doc.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const fieldsMenu = document.createElement("div");
  fieldsMenu.className = "docxedit-menu";
  fieldsMenu.hidden = true;
  for (const it of [
    { label: t("fieldPageNumber"), fn: () => insertField("PAGE") },
    { label: t("fieldPageCount"), fn: () => insertField("NUMPAGES") },
    { label: t("fieldPageXofY"), fn: insertPageXofY },
    { label: t("fieldToc"), fn: insertTOC },
  ]) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.textContent = it.label;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      it.fn();
      fieldsMenu.hidden = true;
    });
    fieldsMenu.appendChild(b);
  }
  const fieldsBtn = iconBtn(fieldIcon, t("insertField"), () => {});
  fieldsBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = fieldsBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    fieldsMenu.style.left = `${r.left - wr.left}px`;
    fieldsMenu.style.top = `${r.bottom - wr.top + 2}px`;
    fieldsMenu.hidden = !fieldsMenu.hidden;
  });

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
    const styled = el.closest("[data-rdoc-style]") as HTMLElement | null;
    const styleId = styled && regions.some((r) => r.contains(styled)) ? styled.getAttribute("data-rdoc-style") : null;
    const tag = el.closest("h1,h2,h3,p")?.tagName ?? "";
    if (styleId && namedStyles.some((s) => s.id === styleId)) block.value = `s:${styleId}`;
    else block.value = tag === "H1" || tag === "H2" || tag === "H3" ? tag : "P";
    const lh = (el.closest(BLOCK_SEL) as HTMLElement | null)?.style.lineHeight ?? "";
    for (const b of lsButtons) b.classList.toggle("is-on", b.dataset.v === lh);
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
    outdentBtn,
    indentBtn,
    lineSpacingBtn,
    sep(),
    caps.images ? iconBtn(imgIcon, t("insertImage"), insertImage) : null,
    caps.tables ? tableBtn : null,
    caps.fields ? fieldsBtn : null,
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
    if (!tablePicker.hidden && !tablePicker.contains(e.target as Node) && !tableBtn.contains(e.target as Node)) tablePicker.hidden = true;
    if (!lineSpacingMenu.hidden && !lineSpacingMenu.contains(e.target as Node) && !lineSpacingBtn.contains(e.target as Node)) lineSpacingMenu.hidden = true;
    if (!fieldsMenu.hidden && !fieldsMenu.contains(e.target as Node) && !fieldsBtn.contains(e.target as Node)) fieldsMenu.hidden = true;
  };
  document.addEventListener("click", closeOverflow);
  toolbar.append(...toolbarItems, moreBtn);
  wrap.appendChild(overflow);
  wrap.appendChild(tablePicker);
  wrap.appendChild(lineSpacingMenu);
  wrap.appendChild(fieldsMenu);

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
