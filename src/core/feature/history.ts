// Trustworthy undo/redo for the document body. The browser's native contenteditable
// undo is invalidated by pagination (blocks reparented into column/section/vertical
// wrappers) and by every direct-DOM mutation path (table editing, image resize,
// styleSel spans, accept/reject, inserts), so the editor keeps its own stack of
// logical-body snapshots (cleanBody output) and restores them wholesale.
//
// Cost control: snapshots are strings of the cleaned body with data-URL image
// payloads interned into a shared pool (so a step costs text, not megabytes), and
// typing runs are captured lazily: commit("t") only marks a run open, and the
// snapshot is taken once when the run closes (idle timeout, focus leaving the body,
// a discrete operation, or undo itself). Discrete operations (mark() without a
// coalesce key) snapshot immediately.

export interface HistoryDeps {
  doc: HTMLElement;
  cleanBody: () => string;
  /** The caret's character offset within a block (the editor's own walker). */
  charOffsetIn: (block: Node, container: Node, offset: number) => number;
  /** Re-place a caret at a character offset within a block (the editor's own). */
  placeCaret: (block: HTMLElement, target: number) => void;
  /** Run the editor's own change plumbing after a restore (dirty + reflow). */
  onRestore: () => void;
}

interface Caret {
  block: number;
  off: number;
}
interface Snap {
  html: string; // interned cleaned body
  caret: Caret | null;
}

const CAP = 50;
const RUN_IDLE_MS = 400;
const WRAPPERS = ".docxedit-colpage, .docxedit-secpage, .docxedit-vband";

export function setupHistory({ doc, cleanBody, charOffsetIn, placeCaret, onRestore }: HistoryDeps) {
  // Image payload pool: data URLs are immutable, so every snapshot shares one copy.
  const imgPool = new Map<string, string>();
  const imgIds = new Map<string, string>();
  let imgSeq = 0;
  const intern = (html: string): string =>
    html.replace(/src="(data:[^"]+)"/g, (_m, url: string) => {
      let id = imgIds.get(url);
      if (!id) {
        id = String(++imgSeq);
        imgIds.set(url, id);
        imgPool.set(id, url);
      }
      return `src="rdoc-hist:${id}"`;
    });
  const unintern = (html: string): string =>
    html.replace(/src="rdoc-hist:(\d+)"/g, (_m, id: string) => `src="${imgPool.get(id) ?? ""}"`);

  // The body's blocks in logical (cleaned) order, looking through pagination wrappers,
  // so caret positions recorded on the live DOM resolve in the restored clean DOM too.
  const logicalBlocks = (): HTMLElement[] => {
    const out: HTMLElement[] = [];
    for (const child of Array.from(doc.children) as HTMLElement[]) {
      if (child.classList.contains("docxedit-pagespacer")) continue;
      if (child.matches(WRAPPERS)) out.push(...(Array.from(child.children) as HTMLElement[]));
      else out.push(child);
    }
    return out;
  };

  const captureCaret = (): Caret | null => {
    const s = window.getSelection();
    if (!s || !s.rangeCount) return null;
    const r = s.getRangeAt(0);
    if (!doc.contains(r.startContainer)) return null;
    const blocks = logicalBlocks();
    const idx = blocks.findIndex((b) => b.contains(r.startContainer));
    if (idx < 0) return null;
    return { block: idx, off: charOffsetIn(blocks[idx]!, r.startContainer, r.startOffset) };
  };

  const snap = (): Snap => ({ html: intern(cleanBody()), caret: captureCaret() });

  const undoStack: Snap[] = [];
  const redoStack: Snap[] = [];
  let base = snap(); // the state at the last recorded boundary
  let pendingKey: string | null = null;
  let restoring = false;
  let composing = false;
  let runTimer = 0;

  const pushIfChanged = (): void => {
    const s = snap();
    if (s.html === base.html) return;
    undoStack.push(base);
    if (undoStack.length > CAP) undoStack.shift();
    base = s;
    redoStack.length = 0;
  };

  const flushRun = (): void => {
    window.clearTimeout(runTimer);
    if (!pendingKey) return;
    pendingKey = null;
    pushIfChanged();
  };

  const armTimer = (): void => {
    window.clearTimeout(runTimer);
    runTimer = window.setTimeout(() => {
      if (composing) armTimer(); // never snapshot mid-IME-composition
      else flushRun();
    }, RUN_IDLE_MS);
  };

  const commit = (key: string | null): void => {
    if (restoring) return;
    if (key) {
      if (pendingKey && pendingKey !== key) flushRun();
      pendingKey = key;
      armTimer();
    } else {
      flushRun();
      pushIfChanged();
    }
  };

  const restore = (s: Snap): void => {
    restoring = true;
    try {
      doc.innerHTML = unintern(s.html);
      onRestore(); // dirty + immediate reflow (repaginates, notes, cards, fields)
      if (s.caret) {
        const b = logicalBlocks()[s.caret.block];
        if (b) {
          doc.focus();
          placeCaret(b, s.caret.off);
        }
      }
    } finally {
      restoring = false;
    }
  };

  const undo = (): boolean => {
    flushRun(); // capture an open typing run so redo can come back to it
    if (!undoStack.length) return false;
    redoStack.push(base);
    base = undoStack.pop()!;
    restore(base);
    return true;
  };

  const redo = (): boolean => {
    flushRun();
    if (!redoStack.length) return false;
    undoStack.push(base);
    base = redoStack.pop()!;
    restore(base);
    return true;
  };

  const onCompStart = () => {
    composing = true;
  };
  const onCompEnd = () => {
    composing = false;
  };
  const onFocusOut = () => flushRun();
  doc.addEventListener("compositionstart", onCompStart);
  doc.addEventListener("compositionend", onCompEnd);
  doc.addEventListener("focusout", onFocusOut);

  return {
    commit,
    undo,
    redo,
    canUndo: () => undoStack.length > 0 || pendingKey !== null,
    canRedo: () => redoStack.length > 0,
    teardown() {
      window.clearTimeout(runTimer);
      doc.removeEventListener("compositionstart", onCompStart);
      doc.removeEventListener("compositionend", onCompEnd);
      doc.removeEventListener("focusout", onFocusOut);
    },
  };
}
