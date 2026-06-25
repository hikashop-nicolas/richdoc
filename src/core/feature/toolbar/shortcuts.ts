// Word-like keyboard shortcuts, bound to the editable regions (so single-key bindings such as
// Ctrl+R for right-align only fire while editing and leave the browser's own shortcuts alone
// everywhere else). The actual commands come from the toolbar as deps.
import type { Capabilities } from "../../types";

export interface ShortcutDeps {
  regions: HTMLElement[];
  caps: Capabilities;
  getActiveEl: () => HTMLElement;
  exec: (cmd: string, val?: string) => void;
  beginFormatChange: () => void;
  styleSel: (prop: string, val: string) => void;
  selectedBlocks: () => HTMLElement[];
  mark: () => void;
  syncToolbarState: () => void;
  insertLink: () => void;
}

export function setupShortcuts(deps: ShortcutDeps) {
  const { regions, caps, getActiveEl, exec, beginFormatChange, styleSel, selectedBlocks, mark, syncToolbarState, insertLink } = deps;

  const applyBlockTag = (tag: string) => {
    getActiveEl().focus();
    document.execCommand("formatBlock", false, tag);
    for (const b of selectedBlocks()) b.removeAttribute("data-rdoc-style");
    mark();
    syncToolbarState();
  };
  const changeFontSize = (delta: number) => {
    const node = window.getSelection()?.anchorNode as Node | null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as HTMLElement)) : null;
    let pt = 11;
    if (el) {
      const px = parseFloat(getComputedStyle(el).fontSize);
      if (px) pt = Math.round((px * 72) / 96);
    }
    beginFormatChange();
    styleSel("fontSize", `${Math.max(1, pt + delta)}pt`);
  };
  const onShortcut = (e: KeyboardEvent) => {
    if (!(e.metaKey || e.ctrlKey)) return;
    const k = e.key.toLowerCase();
    const run = (fn: () => void) => {
      e.preventDefault();
      fn();
    };
    if (e.altKey && !e.shiftKey) {
      if (k === "1") return run(() => applyBlockTag("H1"));
      if (k === "2") return run(() => applyBlockTag("H2"));
      if (k === "3") return run(() => applyBlockTag("H3"));
      if (k === "0") return run(() => applyBlockTag("P"));
      return;
    }
    if (e.shiftKey) {
      if (k === "l") return run(() => exec("insertUnorderedList")); // bulleted list
      if (k === "7") return run(() => exec("insertOrderedList")); // numbered list
      return;
    }
    switch (k) {
      case "b": return run(() => { beginFormatChange(); exec("bold"); });
      case "i": return run(() => { beginFormatChange(); exec("italic"); });
      case "u": return run(() => { beginFormatChange(); exec("underline"); });
      case "k": return run(insertLink);
      case "l": if (caps.alignment) return run(() => exec("justifyLeft")); return;
      case "e": if (caps.alignment) return run(() => exec("justifyCenter")); return;
      case "r": if (caps.alignment) return run(() => exec("justifyRight")); return;
      case "j": if (caps.alignment) return run(() => exec("justifyFull")); return;
      case "]": if (caps.fontControls) return run(() => changeFontSize(1)); return;
      case "[": if (caps.fontControls) return run(() => changeFontSize(-1)); return;
    }
  };
  for (const r of regions) r.addEventListener("keydown", onShortcut);

  // Tab inserts a real tab (an atomic, non-editable span carrying a tab character so copy/paste
  // yields a tab); Shift+Tab removes a tab immediately before the caret, else moves focus normally.
  const insertTab = () => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const span = document.createElement("span");
    span.className = "docx-tab";
    span.setAttribute("data-docx-tab", "1");
    span.contentEditable = "false";
    span.textContent = "\t";
    range.insertNode(span);
    range.setStartAfter(span);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
  };
  const onTab = (e: KeyboardEvent) => {
    if (e.key !== "Tab" || e.metaKey || e.ctrlKey || e.altKey) return;
    if (e.shiftKey) {
      const sel = window.getSelection();
      const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
      if (!range || !range.collapsed) return;
      const c = range.startContainer;
      const prev = c.nodeType === 1 && range.startOffset > 0 ? c.childNodes[range.startOffset - 1] : c.previousSibling;
      const el = prev && prev.nodeType === 1 ? (prev as HTMLElement) : null;
      if (el && el.getAttribute("data-docx-tab")) {
        e.preventDefault();
        el.remove();
        mark();
      }
      return; // no preceding tab: let Shift+Tab move focus as usual
    }
    e.preventDefault();
    insertTab();
  };
  for (const r of regions) r.addEventListener("keydown", onTab);

  const teardown = () => {
    for (const r of regions) {
      r.removeEventListener("keydown", onShortcut);
      r.removeEventListener("keydown", onTab);
    }
  };
  return { teardown };
}
