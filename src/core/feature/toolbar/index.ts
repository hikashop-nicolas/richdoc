// Toolbar: the formatting controls and their wiring. Builds the button/select factories, the
// bold/italic/underline + list + alignment commands (execCommand), the font / size / colour
// pickers, link / page-break / image / table / field / comment actions, and the caret-state
// sync, then lays the row out with progressive grouping + an overflow popover. Named styles,
// the floating bar, the keyboard shortcuts and the icons each live in a sibling module.
import { t } from "../../i18n";
import { firstFontFamily, fontSizeToHalfPt } from "../../util";
import { setupTrackChanges } from "../track-changes";
import { setupStyles } from "./styles";
import { setupFloatBar } from "./float-bar";
import { setupShortcuts } from "./shortcuts";
import {
  alignIcon, indentIcon, bulletIcon, numberIcon, linkIcon, pbIcon, imgIcon, cmtIcon,
  supIcon, subIcon, lineSpacingIcon, tableIcon, fieldIcon, furiganaIcon, footnoteIcon, bookmarkIcon, xrefIcon, caret, styleGroupSvg, insertGroupSvg,
} from "./icons";
import type { Adapter, Capabilities, CommentThread, EditorOptions, NewStyle, RichDoc } from "../../types";

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
  styleBar: HTMLElement; // the bottom-bar left slot for the style pickers
  newStyles: NewStyle[]; // styles authored in-session, collected for save
  newStyleCss: HTMLStyleElement; // live <style> for the appearance of in-session styles
  vertical: boolean; // vertical (tategaki) writing, for the floating bar layout
  insertSectionBreak: (() => void) | null; // section-break action (Ctrl+Shift+Enter), null if unsupported
  insertNote: (kind: "footnote" | "endnote", text: string) => void; // insert a footnote/endnote at the caret
}

export function setupToolbar(deps: ToolbarDeps) {
  const {
    toolbar, wrap, doc, regions, caps, options, parts, adapter, getActiveEl, mark, positionCards,
    addThreadCard, setActiveComment, allocId, freshParaId, insertImage, styleBar,
    newStyles, newStyleCss, vertical, insertSectionBreak, insertNote,
  } = deps;

  // An element counts as an editing host when it is inside the body/header/footer regions, or
  // inside a header/footer clone or a footnote/endnote body. The latter two live outside the
  // `regions` array (in the page overlay), so formatting state, the floating bar and selection
  // capture must recognise them explicitly or they would only respond inside the body.
  const inEditingHost = (el: HTMLElement | null): boolean =>
    !!el && (regions.some((r) => r.contains(el)) || !!el.closest(".docxedit-hf-clone, .docxedit-note"));

  // Clicking a toolbar <select> can drop the editor's selection (especially a non-collapsed
  // one). Capture it on mousedown, restore it before a style action reads it.
  let savedSel: Range | null = null;
  const captureSel = (): void => {
    const s = window.getSelection();
    if (!s || !s.rangeCount) return;
    const r = s.getRangeAt(0);
    const n = r.startContainer;
    const el = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
    if (inEditingHost(el)) savedSel = r.cloneRange();
  };
  const restoreSel = (): void => {
    if (!savedSel) return;
    const s = window.getSelection();
    if (s) {
      s.removeAllRanges();
      s.addRange(savedSel);
    }
  };

  // Named styles are created later (they need queryState/selectedBlocks/syncToolbarState), but
  // syncToolbarState refers to them, so the holder is declared up front.
  let styles: ReturnType<typeof setupStyles> | null = null;

  const exec = (cmd: string, val?: string) => {
    getActiveEl().focus();
    document.execCommand(cmd, false, val);
    mark();
    syncToolbarState();
  };
  // Format a keyboard shortcut for tooltips, platform-aware (⌘ on Mac, Ctrl elsewhere).
  const isMac = /Mac|iPhone|iPad/.test((typeof navigator !== "undefined" && (navigator.platform || navigator.userAgent)) || "");
  const sc = (key: string, opts: { shift?: boolean; alt?: boolean } = {}): string =>
    isMac
      ? `${opts.alt ? "⌥" : ""}${opts.shift ? "⇧" : ""}⌘${key.toUpperCase()}`
      : `Ctrl+${opts.alt ? "Alt+" : ""}${opts.shift ? "Shift+" : ""}${key.toUpperCase()}`;
  const withSc = (title: string, key: string, opts?: { shift?: boolean; alt?: boolean }): string => `${title} (${sc(key, opts)})`;
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
  const iconBtn = (svg: string, title: string, fn: () => void) => {
    const b = btn("", title, fn);
    b.innerHTML = svg;
    return b;
  };

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

  // Track changes (suggestion mode) lives in its own feature module; it builds its own
  // toolbar buttons and exposes beginFormatChange for the bold/italic/underline buttons.
  const { beginFormatChange, suggestBtn, acceptAllBtn, rejectAllBtn, updateChangeButtons } = setupTrackChanges({
    doc, wrap, regions, options, mark, positionCards, getActiveEl, iconBtn, btn,
  });

  // Toolbar: shared controls always shown; image/comment/page-break/track-changes are
  // gated by the adapter's capabilities so a format can hide what it cannot serialize.
  const insertLink = () => {
    const url = prompt(t("linkPrompt"), "https://");
    if (url === null) return;
    if (url === "") exec("unlink");
    else exec("createLink", url);
  };
  const linkBtn = iconBtn(linkIcon, withSc(t("linkAria"), "K"), insertLink);
  // Furigana (ruby): wrap the selection in <ruby>base<rt>reading</rt></ruby>, or edit / remove the
  // reading when the caret is already inside a ruby.
  const insertFurigana = () => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const node = range.startContainer;
    const el = (node.nodeType === 3 ? node.parentElement : node) as HTMLElement | null;
    if (!el || !regions.some((r) => r.contains(el))) return;
    const existing = el.closest("ruby");
    if (existing) {
      const rt = existing.querySelector("rt");
      const reading = prompt(t("furiganaPrompt"), rt?.textContent ?? "");
      if (reading === null) return;
      mark();
      if (reading === "") {
        for (const r of Array.from(existing.querySelectorAll("rt"))) r.remove();
        while (existing.firstChild) existing.parentNode!.insertBefore(existing.firstChild, existing);
        existing.remove();
      } else if (rt) rt.textContent = reading;
      else { const r = document.createElement("rt"); r.textContent = reading; existing.appendChild(r); }
      return;
    }
    if (range.collapsed) return; // need a base selection to annotate
    const reading = prompt(t("furiganaPrompt"), "");
    if (!reading) return;
    const ruby = document.createElement("ruby");
    ruby.appendChild(range.extractContents());
    for (const r of Array.from(ruby.querySelectorAll("rt"))) r.remove(); // base only
    const rt = document.createElement("rt");
    rt.textContent = reading;
    ruby.appendChild(rt);
    range.insertNode(ruby);
    const after = document.createRange();
    after.setStartAfter(ruby);
    after.collapse(true);
    sel.removeAllRanges();
    sel.addRange(after);
    mark();
  };
  const furiganaBtn = iconBtn(furiganaIcon, withSc(t("furigana"), "F", { shift: true }), insertFurigana);
  // Named so their pressed state can be reflected from the caret (see syncToolbarState).
  const boldBtn = btn("B", withSc(t("bold"), "B"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold");
  const italicBtn = btn("I", withSc(t("italic"), "I"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic");
  const underlineBtn = btn("U", withSc(t("underline"), "U"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline");
  const strikeBtn = btn("S", t("strikethrough"), () => { beginFormatChange(); exec("strikeThrough"); }, "docxedit-tb-strike");
  const supBtn = iconBtn(supIcon, t("superscript"), () => { beginFormatChange(); exec("superscript"); });
  const subBtn = iconBtn(subIcon, t("subscript"), () => { beginFormatChange(); exec("subscript"); });
  const alignLeftBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), withSc(t("alignLeft"), "L"), () => exec("justifyLeft")) : null;
  const alignCenterBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), withSc(t("alignCenter"), "E"), () => exec("justifyCenter")) : null;
  const alignRightBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), withSc(t("alignRight"), "R"), () => exec("justifyRight")) : null;
  const alignJustifyBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), withSc(t("alignJustify"), "J"), () => exec("justifyFull")) : null;

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

  // --- Ordered-list numbering: restart / continue / start-at on the list at the caret -------
  const currentOl = (): HTMLOListElement | null => {
    const n = window.getSelection()?.anchorNode as Node | null;
    const el = n ? (n.nodeType === 3 ? n.parentElement : (n as Element)) : null;
    const ol = el?.closest?.("ol") as HTMLOListElement | null;
    return ol && regions.some((r) => r.contains(ol)) ? ol : null;
  };
  const setListStart = (ol: HTMLOListElement, start: number): void => {
    if (start > 1) ol.setAttribute("start", String(start));
    else ol.removeAttribute("start"); // restart at 1 = no start attribute
    mark();
    syncToolbarState();
  };
  const restartList = () => { const ol = currentOl(); if (ol) setListStart(ol, 1); };
  const continueList = () => {
    const ol = currentOl();
    if (!ol) return;
    let prev = ol.previousElementSibling; // the previous ordered list in the body
    while (prev && prev.tagName !== "OL") prev = prev.previousElementSibling;
    if (!prev) return;
    const prevStart = parseInt(prev.getAttribute("start") || "1", 10) || 1;
    const prevCount = Array.from(prev.children).filter((c) => c.tagName === "LI").length;
    setListStart(ol, prevStart + prevCount);
  };
  const startAtList = () => {
    const ol = currentOl();
    if (!ol) return;
    const v = prompt(t("listStartPrompt"), ol.getAttribute("start") || "1");
    if (v === null) return;
    const n = parseInt(v, 10);
    setListStart(ol, Number.isFinite(n) && n > 0 ? n : 1);
  };
  const listMenu = document.createElement("div");
  listMenu.className = "docxedit-menu";
  listMenu.hidden = true;
  for (const it of [
    { label: t("listRestart"), fn: restartList },
    { label: t("listContinue"), fn: continueList },
    { label: t("listStartAt"), fn: startAtList },
  ]) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.textContent = it.label;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      it.fn();
      listMenu.hidden = true;
    });
    listMenu.appendChild(b);
  }
  const listNumBtn = iconBtn(numberIcon + caret, t("listNumbering"), () => {});
  listNumBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = listNumBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    listMenu.style.left = `${r.left - wr.left}px`;
    listMenu.style.top = `${r.bottom - wr.top + 2}px`;
    listMenu.hidden = !listMenu.hidden;
  });

  // Reflect the caret's formatting in the controls: the font / size pickers track the text
  // under the caret, bold/italic/underline/alignment show their pressed state, and the named
  // style dropdowns are refreshed by the styles module. rAF-coalesced on selection change.
  const queryState = (cmd: string): boolean => {
    try { return document.queryCommandState(cmd); } catch { return false; }
  };
  const syncToolbarState = () => {
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.anchorNode : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as HTMLElement)) : null;
    if (!el || !inEditingHost(el)) return; // only while editing a host
    const cs = getComputedStyle(el);
    if (caps.fontControls) {
      const fam = firstFontFamily(cs.fontFamily);
      fontSel.value = fam && FONTS.includes(fam) ? fam : "";
      const hp = fontSizeToHalfPt(cs.fontSize);
      const pt = hp ? String(Math.round(hp / 2)) : "";
      sizeSel.value = SIZES.includes(pt) ? pt : "";
    }
    styles?.syncStyleState(el); // paragraph + character style dropdowns
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

  // Named paragraph / character styles: the two bottom-bar dropdowns + the authoring dialog.
  styles = setupStyles({
    wrap, regions, parts, styleBar, newStyles, newStyleCss, fonts: FONTS,
    getActiveEl, mark, exec, selectedBlocks, queryState, captureSel, restoreSel,
    resync: syncToolbarState, sc,
  });

  // Insert-note popup: one "Insert note" button opens it; the user types the note text and picks
  // footnote (default) or endnote. The body caret is captured when the button is clicked and
  // restored before inserting, so the note lands where the caret was.
  const noteOverlay = document.createElement("div");
  noteOverlay.className = "docxedit-dialog-overlay";
  noteOverlay.hidden = true;
  const notePanel = document.createElement("div");
  notePanel.className = "docxedit-dialog docxedit-noteinsert";
  const noteTitle = document.createElement("div");
  noteTitle.className = "docxedit-dialog-title";
  noteTitle.textContent = t("insertNote");
  const noteText = document.createElement("textarea");
  noteText.className = "docxedit-dialog-textarea";
  noteText.rows = 3;
  const kindRow = document.createElement("div");
  kindRow.className = "docxedit-dialog-row docxedit-noteinsert-kind";
  const mkRadio = (value: "footnote" | "endnote", label: string, checked: boolean) => {
    const lab = document.createElement("label");
    lab.className = "docxedit-noteinsert-opt";
    const input = document.createElement("input");
    input.type = "radio"; input.name = "rdoc-note-kind"; input.value = value; input.checked = checked;
    lab.append(input, document.createTextNode(` ${label}`));
    return { lab, input };
  };
  const fnOpt = mkRadio("footnote", t("footnote"), true);
  const enOpt = mkRadio("endnote", t("endnote"), false);
  kindRow.append(fnOpt.lab, enOpt.lab);
  const noteActions = document.createElement("div");
  noteActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const noteCancel = document.createElement("button");
  noteCancel.type = "button"; noteCancel.className = "docxedit-menu-item"; noteCancel.textContent = t("cancel");
  const noteInsertBtn = document.createElement("button");
  noteInsertBtn.type = "button"; noteInsertBtn.className = "docxedit-menu-item docxedit-dialog-primary"; noteInsertBtn.textContent = t("insert");
  noteActions.append(noteCancel, noteInsertBtn);
  notePanel.append(noteTitle, noteText, kindRow, noteActions);
  noteOverlay.appendChild(notePanel);
  wrap.appendChild(noteOverlay);
  const closeNote = () => { noteOverlay.hidden = true; };
  const openNoteDialog = () => {
    captureSel();
    noteText.value = "";
    fnOpt.input.checked = true;
    noteOverlay.hidden = false;
    setTimeout(() => noteText.focus(), 0);
  };
  noteOverlay.addEventListener("mousedown", (e) => { if (e.target === noteOverlay) closeNote(); });
  noteCancel.addEventListener("click", closeNote);
  noteInsertBtn.addEventListener("click", () => {
    const kind = enOpt.input.checked ? "endnote" : "footnote";
    const text = noteText.value;
    closeNote();
    restoreSel();
    insertNote(kind, text);
  });

  // --- Bookmarks + cross-references ---------------------------------------------------------------
  const nextBmId = (): number => {
    let max = 0;
    for (const e of Array.from(doc.querySelectorAll("[data-rdoc-bm-id]"))) max = Math.max(max, parseInt(e.getAttribute("data-rdoc-bm-id") || "0", 10) || 0);
    return max + 1;
  };
  const mkBmEl = (cls: string, name: string | null, id: string): HTMLElement => {
    const a = document.createElement("a");
    a.className = cls;
    if (name) a.setAttribute("data-rdoc-bm", name);
    a.setAttribute("data-rdoc-bm-id", id);
    a.contentEditable = "false";
    return a;
  };
  // Insert a bookmark: a point at the caret, or a range wrapping the selection (a start marker at the
  // selection start + an end marker at its end, so the spanned text stays referenceable).
  const insertBookmark = (): void => {
    captureSel();
    const name = (window.prompt(t("bookmarkPrompt")) || "").trim();
    if (!name) return;
    restoreSel();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const cont = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element);
    if (!cont || !doc.contains(cont)) return;
    const id = String(nextBmId());
    const start = mkBmEl("docx-bookmark", name, id);
    const end = mkBmEl("docx-bookmark-end", null, id);
    end.setAttribute("data-rdoc-bm-end", name); // odt pairs by name; docx by id
    if (range.collapsed) { range.insertNode(end); range.insertNode(start); }
    else { const r2 = range.cloneRange(); r2.collapse(false); r2.insertNode(end); range.collapse(true); range.insertNode(start); }
    mark();
  };
  // A heading's referenceable bookmark name: reuse one already wrapping it, else wrap its content in a
  // fresh _Ref bookmark so the cross-reference target survives a save (headings aren't bookmarks in odf/ooxml).
  const headingBmName = (h: HTMLElement): string => {
    const existing = h.querySelector<HTMLElement>(":scope > .docx-bookmark");
    if (existing) return existing.getAttribute("data-rdoc-bm") || "";
    const id = String(nextBmId());
    const name = `_Ref${id}`;
    const start = mkBmEl("docx-bookmark", name, id);
    const end = mkBmEl("docx-bookmark-end", null, id);
    end.setAttribute("data-rdoc-bm-end", name);
    h.insertBefore(start, h.firstChild);
    h.appendChild(end);
    return name;
  };
  // The cross-reference targets of a kind, in document order: headings (referenced via an auto-minted
  // bookmark) or named bookmarks.
  const xrefTargets = (kind: "heading" | "bookmark"): { label: string; el: HTMLElement }[] =>
    kind === "heading"
      ? Array.from(doc.querySelectorAll<HTMLElement>("h1,h2,h3")).filter((h) => !h.closest(".docx-field-toc")).map((h) => ({ label: (h.textContent || "").trim() || t("untitled"), el: h }))
      : Array.from(doc.querySelectorAll<HTMLElement>(".docx-bookmark")).map((b) => ({ label: b.getAttribute("data-rdoc-bm") || "", el: b }));
  const xrefOverlay = document.createElement("div");
  xrefOverlay.className = "docxedit-dialog-overlay";
  xrefOverlay.hidden = true;
  const xrefPanel = document.createElement("div");
  xrefPanel.className = "docxedit-dialog docxedit-xref";
  const xrefTitle = document.createElement("div");
  xrefTitle.className = "docxedit-dialog-title";
  xrefTitle.textContent = t("insertCrossRef");
  const typeRow = document.createElement("div");
  typeRow.className = "docxedit-dialog-row docxedit-xref-kind";
  const mkXrefRadio = (group: string, value: string, label: string, checked: boolean) => {
    const lab = document.createElement("label");
    lab.className = "docxedit-noteinsert-opt";
    const input = document.createElement("input");
    input.type = "radio"; input.name = group; input.value = value; input.checked = checked;
    lab.append(input, document.createTextNode(` ${label}`));
    return { lab, input };
  };
  const tHeading = mkXrefRadio("rdoc-xref-kind", "heading", t("refHeadings"), true);
  const tBookmark = mkXrefRadio("rdoc-xref-kind", "bookmark", t("refBookmarks"), false);
  typeRow.append(tHeading.lab, tBookmark.lab);
  const targetSel = document.createElement("select");
  targetSel.className = "docxedit-dialog-font";
  let xrefEls: HTMLElement[] = [];
  const fillTargets = () => {
    const kind = tBookmark.input.checked ? "bookmark" : "heading";
    const list = xrefTargets(kind);
    xrefEls = list.map((x) => x.el);
    targetSel.replaceChildren(...list.map((x, i) => { const o = document.createElement("option"); o.value = String(i); o.textContent = x.label; return o; }));
  };
  tHeading.input.addEventListener("change", fillTargets);
  tBookmark.input.addEventListener("change", fillTargets);
  const fmtRow = document.createElement("div");
  fmtRow.className = "docxedit-dialog-row docxedit-xref-fmt";
  const fmtText = mkXrefRadio("rdoc-xref-fmt", "text", t("refFormatText"), true);
  const fmtPage = mkXrefRadio("rdoc-xref-fmt", "page", t("refFormatPage"), false);
  fmtRow.append(fmtText.lab, fmtPage.lab);
  const xrefActions = document.createElement("div");
  xrefActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const xrefCancel = document.createElement("button");
  xrefCancel.type = "button"; xrefCancel.className = "docxedit-menu-item"; xrefCancel.textContent = t("cancel");
  const xrefInsertBtn = document.createElement("button");
  xrefInsertBtn.type = "button"; xrefInsertBtn.className = "docxedit-menu-item docxedit-dialog-primary"; xrefInsertBtn.textContent = t("insert");
  xrefActions.append(xrefCancel, xrefInsertBtn);
  xrefPanel.append(xrefTitle, typeRow, targetSel, fmtRow, xrefActions);
  xrefOverlay.appendChild(xrefPanel);
  wrap.appendChild(xrefOverlay);
  const closeXref = () => { xrefOverlay.hidden = true; };
  const openXrefDialog = () => {
    captureSel();
    tHeading.input.checked = true;
    fmtText.input.checked = true;
    fillTargets();
    xrefOverlay.hidden = false;
  };
  xrefOverlay.addEventListener("mousedown", (e) => { if (e.target === xrefOverlay) closeXref(); });
  xrefCancel.addEventListener("click", closeXref);
  xrefInsertBtn.addEventListener("click", () => {
    const target = xrefEls[Number(targetSel.value)];
    if (!target) { closeXref(); return; }
    // A heading is referenced through a bookmark wrapping its content; a bookmark target by its own name.
    const name = /^H[1-6]$/.test(target.tagName) ? headingBmName(target) : target.getAttribute("data-rdoc-bm") || "";
    if (!name) { closeXref(); return; }
    const fmt = fmtPage.input.checked ? "page" : "text";
    closeXref();
    restoreSel();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const cont = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element);
    if (!cont || !doc.contains(cont)) return;
    const xr = document.createElement("a");
    xr.className = "docx-xref";
    xr.setAttribute("data-rdoc-xref", name);
    xr.setAttribute("data-rdoc-xref-fmt", fmt);
    xr.contentEditable = "false";
    xr.textContent = fmt === "page" ? "1" : (target.textContent || name).trim().slice(0, 80) || name;
    range.collapse(false);
    range.insertNode(xr);
    range.setStartAfter(xr);
    mark();
  });

  // Two clusters collapse into a single dropdown button when the toolbar runs out of room:
  // the character-formatting controls ("style"), and the insert controls.
  const styleSrc: (HTMLElement | null)[] = [boldBtn, italicBtn, underlineBtn, strikeBtn, supBtn, subBtn, caps.textColor ? colorInput : null, caps.textColor ? bgWrap : null, caps.fontControls ? fontSel : null, caps.fontControls ? sizeSel : null];
  const insertSrc: (HTMLElement | null)[] = [caps.images ? iconBtn(imgIcon, t("insertImage"), insertImage) : null, caps.tables ? tableBtn : null, caps.fields ? fieldsBtn : null, caps.comments ? iconBtn(cmtIcon, t("addComment"), addComment) : null, caps.pageBreak ? iconBtn(pbIcon, t("insertPageBreak"), insertPageBreak) : null, linkBtn, iconBtn(footnoteIcon, t("insertNote"), openNoteDialog), iconBtn(bookmarkIcon, t("insertBookmark"), insertBookmark), iconBtn(xrefIcon, t("insertCrossRef"), openXrefDialog), caps.verticalText ? furiganaBtn : null];
  const styleControls = styleSrc.filter((n): n is HTMLElement => n != null);
  const insertControls = insertSrc.filter((n): n is HTMLElement => n != null);
  // A collapsible cluster: a slot that holds the controls inline or a group button + a popover.
  const makeGroup = (controls: HTMLElement[], svg: string, title: string) => {
    const slot = document.createElement("span");
    slot.className = "docxedit-tb-slot";
    const btn = iconBtn(svg + caret, title, () => {});
    btn.classList.add("docxedit-tb-group");
    const menu = document.createElement("div");
    menu.className = "docxedit-toolbar docxedit-tb-overflow docxedit-tb-groupmenu";
    menu.hidden = true;
    let collapsed = false;
    btn.addEventListener("mousedown", (e) => e.preventDefault());
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const r = btn.getBoundingClientRect();
      const wr = wrap.getBoundingClientRect();
      menu.style.top = `${r.bottom - wr.top + 2}px`;
      menu.style.left = `${r.left - wr.left}px`;
      menu.style.right = "auto";
      menu.hidden = !menu.hidden;
    });
    slot.replaceChildren(...controls); // start expanded
    const expand = () => {
      if (!collapsed) return;
      slot.replaceChildren(...controls);
      menu.hidden = true;
      collapsed = false;
    };
    const collapse = () => {
      if (collapsed || !controls.length) return;
      menu.replaceChildren(...controls);
      slot.replaceChildren(btn);
      collapsed = true;
    };
    return { slot, btn, menu, expand, collapse, has: controls.length > 0 };
  };
  const styleGroup = makeGroup(styleControls, styleGroupSvg, t("formatting"));
  const insertGroup = makeGroup(insertControls, insertGroupSvg, t("insertMenu"));

  const items: (Node | null)[] = [
    styleGroup.has ? styleGroup.slot : null,
    styleGroup.has ? sep() : null,
    iconBtn(bulletIcon, withSc(t("bulleted"), "L", { shift: true }), () => exec("insertUnorderedList")),
    iconBtn(numberIcon, withSc(t("numbered"), "7", { shift: true }), () => exec("insertOrderedList")),
    listNumBtn,
    caps.alignment ? sep() : null,
    alignLeftBtn,
    alignCenterBtn,
    alignRightBtn,
    alignJustifyBtn,
    outdentBtn,
    indentBtn,
    lineSpacingBtn,
    insertGroup.has ? sep() : null,
    insertGroup.has ? insertGroup.slot : null,
    caps.trackChanges ? sep() : null,
    caps.trackChanges ? suggestBtn : null,
    caps.trackChanges ? acceptAllBtn : null,
    caps.trackChanges ? rejectAllBtn : null,
  ];
  // Overflow menu: anything that still does not fit after the groups collapse moves into a "…"
  // popover so nothing is lost on very narrow widths.
  const toolbarItems = items.filter((n): n is HTMLElement => n != null);
  const moreBtn = document.createElement("button");
  moreBtn.type = "button";
  moreBtn.className = "docxedit-tb-more";
  moreBtn.textContent = "⋯";
  moreBtn.title = t("moreTools");
  moreBtn.setAttribute("aria-label", t("moreTools"));
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
    if (!styleGroup.menu.hidden && !styleGroup.menu.contains(e.target as Node) && !styleGroup.btn.contains(e.target as Node)) styleGroup.menu.hidden = true;
    if (!insertGroup.menu.hidden && !insertGroup.menu.contains(e.target as Node) && !insertGroup.btn.contains(e.target as Node)) insertGroup.menu.hidden = true;
    if (!tablePicker.hidden && !tablePicker.contains(e.target as Node) && !tableBtn.contains(e.target as Node)) tablePicker.hidden = true;
    if (!lineSpacingMenu.hidden && !lineSpacingMenu.contains(e.target as Node) && !lineSpacingBtn.contains(e.target as Node)) lineSpacingMenu.hidden = true;
    if (!fieldsMenu.hidden && !fieldsMenu.contains(e.target as Node) && !fieldsBtn.contains(e.target as Node)) fieldsMenu.hidden = true;
    if (!listMenu.hidden && !listMenu.contains(e.target as Node) && !listNumBtn.contains(e.target as Node)) listMenu.hidden = true;
  };
  document.addEventListener("click", closeOverflow);
  toolbar.append(...toolbarItems, moreBtn);
  wrap.append(overflow, styleGroup.menu, insertGroup.menu, tablePicker, lineSpacingMenu, fieldsMenu, listMenu);

  const fits = () => toolbar.scrollWidth <= toolbar.clientWidth + 1;
  const layoutToolbar = () => {
    overflow.hidden = true;
    for (const it of toolbarItems) toolbar.insertBefore(it, moreBtn); // pull everything back in
    moreBtn.style.display = "none";
    insertGroup.expand();
    styleGroup.expand();
    if (fits()) return;
    insertGroup.collapse(); // first collapse the insert cluster
    if (fits()) return;
    styleGroup.collapse(); // then the formatting cluster
    if (fits()) return;
    moreBtn.style.display = ""; // final fallback: pocket trailing items
    for (let i = toolbarItems.length - 1; i >= 0; i--) {
      if (fits()) break;
      overflow.insertBefore(toolbarItems[i], overflow.firstChild);
    }
  };
  layoutToolbar();
  requestAnimationFrame(layoutToolbar);
  setTimeout(layoutToolbar, 150);
  const toolbarObserver = new ResizeObserver(() => layoutToolbar());
  toolbarObserver.observe(toolbar);

  scheduleSync(); // reflect the initial caret position once mounted

  // The floating formatting bar and the keyboard shortcuts each live in a sibling module.
  const floatBar = setupFloatBar({ wrap, regions, inEditingHost, getActiveEl, beginFormatChange, exec, queryState, withSc, vertical });
  const shortcuts = setupShortcuts({ wrap, inEditingHost, caps, getActiveEl, exec, beginFormatChange, styleSel, selectedBlocks, mark, syncToolbarState, insertLink, insertFurigana, insertSectionBreak });

  const teardown = () => {
    toolbarObserver.disconnect();
    document.removeEventListener("click", closeOverflow);
    document.removeEventListener("selectionchange", scheduleSync);
    floatBar.teardown();
    shortcuts.teardown();
    window.clearTimeout(syncTimer);
  };
  return { updateChangeButtons, teardown };
}
