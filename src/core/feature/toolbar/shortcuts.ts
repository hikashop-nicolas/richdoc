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

  const teardown = () => {
    for (const r of regions) r.removeEventListener("keydown", onShortcut);
  };
  return { teardown };
}
