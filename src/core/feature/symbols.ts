// Special-character / symbol picker: a popup grid of common characters, each inserted as plain text
// at the caret. Because they are ordinary Unicode text they round-trip in both docx and odt with no
// adapter work. The popup stays open so several can be inserted in a row; Esc or a click outside
// closes it.
import { t } from "../i18n";

export interface SymbolDeps {
  doc: HTMLElement; // the editable body (kept for parity with the other features)
  wrap: HTMLElement; // chrome the popup overlays
  mark: () => void; // flag dirty after an insert
  getActiveEl: () => HTMLElement; // the current editing host (body / note / header / cell)
  captureSel: () => void; // save the caret (so it survives focus moving to the popup)
  restoreSel: () => void; // restore the saved caret before inserting
}

// A curated, visually grouped set of the most useful special characters (punctuation, math, arrows,
// currency, Greek, marks, accented Latin). Ordering groups them; the grid wraps.
const SYMBOLS = Array.from(
  "…—–•·§¶†‡«»“”‘’′″" + // punctuation & quotes
    "×÷±∓≈≠≤≥≡∞∑∏∫√∂∈∉∅∝°µ½⅓¼¾" + // math
    "←→↑↓↔⇒⇐⇔↦" + // arrows
    "€£¥¢₹₽" + // currency
    "αβγδεζθλμνξπρστφχψω" + // greek lower
    "ΓΔΘΛΞΠΣΦΨΩ" + // greek upper
    "©®™°★☆✓✗№⚠♥♦♣♠♪" + // marks
    "æœøåçñßäöüé", // accented Latin
);

export function setupSymbols(deps: SymbolDeps) {
  const { wrap, mark, getActiveEl, captureSel, restoreSel } = deps;

  const overlay = document.createElement("div");
  overlay.className = "docxedit-dialog-overlay";
  overlay.hidden = true;
  const panel = document.createElement("div");
  panel.className = "docxedit-dialog docxedit-symbols";
  const title = document.createElement("div");
  title.className = "docxedit-dialog-title";
  title.textContent = t("insertSymbol");
  const grid = document.createElement("div");
  grid.className = "docxedit-symbol-grid";

  const insert = (ch: string): void => {
    const host = getActiveEl();
    host.focus();
    restoreSel();
    // Prefer execCommand (joins the native undo stack); fall back to a Range insert where it is
    // absent / a no-op (e.g. jsdom), so the caret still advances and the next pick continues here.
    let inserted = false;
    try { inserted = typeof document.execCommand === "function" && document.execCommand("insertText", false, ch); } catch { inserted = false; }
    if (!inserted) {
      const sel = window.getSelection();
      const r = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
      const cont = r && (r.startContainer.nodeType === 3 ? r.startContainer.parentElement : (r.startContainer as Element));
      if (r && cont && host.contains(cont)) {
        r.deleteContents();
        const node = document.createTextNode(ch);
        r.insertNode(node);
        r.setStartAfter(node);
        r.collapse(true);
        sel!.removeAllRanges();
        sel!.addRange(r);
      } else {
        // No usable caret in the host: append to its last block so the click still does something.
        const blocks = host.querySelectorAll("p, h1, h2, h3, li");
        (blocks[blocks.length - 1] ?? host).appendChild(document.createTextNode(ch));
      }
    }
    captureSel();
    mark();
  };

  for (const ch of SYMBOLS) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-symbol";
    b.textContent = ch;
    b.title = ch;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the editor's selection/focus
    b.addEventListener("click", () => insert(ch));
    grid.appendChild(b);
  }

  const actions = document.createElement("div");
  actions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "docxedit-menu-item";
  closeBtn.textContent = t("findClose");
  actions.appendChild(closeBtn);

  panel.append(title, grid, actions);
  overlay.appendChild(panel);
  wrap.appendChild(overlay);

  const close = (): void => { overlay.hidden = true; };
  const open = (): void => { captureSel(); overlay.hidden = false; };

  closeBtn.addEventListener("click", close);
  overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) close(); });
  const onKey = (e: KeyboardEvent): void => { if (!overlay.hidden && e.key === "Escape") { e.preventDefault(); close(); } };
  window.addEventListener("keydown", onKey);

  return {
    openDialog: open,
    teardown(): void { window.removeEventListener("keydown", onKey); overlay.remove(); },
  };
}
