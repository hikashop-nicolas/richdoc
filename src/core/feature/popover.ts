// Floating toolbars hold icon buttons; one with several settings opens a small panel just above
// the toolbar (below it when there is no room). One panel is open at a time; Escape or a press
// elsewhere closes it, and its button reports it as an expanded dialog.

export interface Popovers {
  /** A button that opens `content` in a panel above `bar`; `onOpen` refreshes it first. */
  add(btn: HTMLButtonElement, label: string, content: HTMLElement, onOpen?: () => void): void;
  close(focusButton?: boolean): void;
  /** Keep the open panel on its button after the toolbar moved. */
  place(): void;
  contains(n: Node | null): boolean;
  teardown(): void;
}

export function setupPopovers(wrap: HTMLElement, bar: HTMLElement): Popovers {
  const panels: { btn: HTMLButtonElement; el: HTMLElement }[] = [];
  let open: { btn: HTMLButtonElement; el: HTMLElement } | null = null;
  const place = () => {
    if (!open) return;
    const wr = wrap.getBoundingClientRect(), br = bar.getBoundingClientRect(), b = open.btn.getBoundingClientRect();
    const el = open.el;
    el.style.left = `${Math.max(4, Math.min(b.left - wr.left, wr.width - el.offsetWidth - 4))}px`;
    const above = br.top - wr.top - el.offsetHeight - 6;
    el.style.top = `${above >= 4 ? above : br.bottom - wr.top + 6}px`;
  };
  const close = (focusButton = false) => {
    if (!open) return;
    const { btn, el } = open;
    el.hidden = true;
    btn.setAttribute("aria-expanded", "false");
    btn.classList.remove("is-open");
    open = null;
    if (focusButton) btn.focus();
  };
  const onDown = (e: PointerEvent) => {
    const t = e.target as Node;
    if (open && !open.el.contains(t) && !open.btn.contains(t)) close();
  };
  document.addEventListener("pointerdown", onDown, true);
  return {
    add(btn, label, content, onOpen) {
      const el = document.createElement("div");
      el.className = "rdoc-pop";
      el.setAttribute("role", "dialog");
      el.setAttribute("aria-label", label);
      el.hidden = true;
      el.appendChild(content);
      el.addEventListener("keydown", (e) => {
        if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(true); }
      });
      wrap.appendChild(el);
      panels.push({ btn, el });
      btn.setAttribute("aria-haspopup", "dialog");
      btn.setAttribute("aria-expanded", "false");
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (open?.btn === btn) return close();
        close();
        onOpen?.();
        el.hidden = false;
        btn.setAttribute("aria-expanded", "true");
        btn.classList.add("is-open");
        open = { btn, el };
        place();
        // Keyboard users land on its first control; a pointer keeps the caret where it was.
        if (e.detail === 0) el.querySelector<HTMLElement>("button:not([hidden]), input:not([disabled])")?.focus();
      });
    },
    close,
    place,
    contains: (n) => !!n && panels.some((p) => p.el.contains(n)),
    teardown() {
      document.removeEventListener("pointerdown", onDown, true);
      for (const p of panels) p.el.remove();
    },
  };
}

/** Colours offered as swatches: black, white, then Office's theme accents and greys. */
const SWATCHES = ["#000000", "#ffffff", "#7f7f7f", "#c00000", "#ed7d31", "#ffc000", "#70ad47", "#4472c4", "#5b9bd5", "#7030a0"];

/** A colour chooser: swatches, any other colour, and "none". Returns its element and a setter. */
export function colorChooser(opts: { noneLabel: string; otherLabel: string; onPick: (c: string | null) => void }) {
  const el = document.createElement("div");
  el.className = "rdoc-swatches";
  el.setAttribute("role", "group");
  const cells: HTMLButtonElement[] = [];
  for (const c of SWATCHES) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "rdoc-swatch";
    b.style.background = c;
    b.title = c;
    b.setAttribute("aria-label", c);
    b.dataset.color = c;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => opts.onPick(c));
    cells.push(b);
    el.appendChild(b);
  }
  const none = document.createElement("button");
  none.type = "button";
  none.className = "rdoc-swatch rdoc-swatch-none";
  none.title = opts.noneLabel;
  none.setAttribute("aria-label", opts.noneLabel);
  none.addEventListener("mousedown", (e) => e.preventDefault());
  none.addEventListener("click", () => opts.onPick(null));
  const other = document.createElement("input");
  other.type = "color";
  other.className = "rdoc-swatch rdoc-swatch-other";
  other.title = opts.otherLabel;
  other.setAttribute("aria-label", opts.otherLabel);
  other.addEventListener("input", () => opts.onPick(other.value));
  el.append(none, other);
  return {
    el,
    set(c: string | null) {
      const hex = c?.toLowerCase() ?? null;
      for (const b of cells) b.setAttribute("aria-pressed", String(b.dataset.color === hex));
      none.setAttribute("aria-pressed", String(hex === null));
      if (hex) other.value = hex;
    },
  };
}
