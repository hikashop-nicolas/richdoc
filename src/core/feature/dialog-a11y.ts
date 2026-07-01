// Shared modal accessibility for the standard docxedit dialog (an overlay containing a
// .docxedit-dialog panel with a .docxedit-dialog-title). Marks it as a modal dialog for
// assistive tech, traps Tab within the panel, and restores focus to whatever was focused
// before it opened once the overlay leaves the DOM. Call once, right after the overlay is
// appended. Escape / overlay-click dismissal stays with each dialog's own handlers.
let dlgSeq = 0;

export function makeDialogAccessible(overlay: HTMLElement, initialFocus?: HTMLElement): void {
  const panel = overlay.querySelector<HTMLElement>(".docxedit-dialog");
  if (!panel) return;
  const prev = document.activeElement as HTMLElement | null; // the opener, restored on close
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  const title = panel.querySelector<HTMLElement>(".docxedit-dialog-title");
  if (title) {
    if (!title.id) title.id = `docxedit-dlg-title-${++dlgSeq}`;
    panel.setAttribute("aria-labelledby", title.id);
  }
  const focusables = (): HTMLElement[] =>
    [...panel.querySelectorAll<HTMLElement>('button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),a[href],[tabindex]:not([tabindex="-1"])')].filter(
      (el) => el.offsetParent !== null,
    );
  panel.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    const f = focusables();
    if (!f.length) return;
    const first = f[0]!;
    const last = f[f.length - 1]!;
    const a = document.activeElement;
    if (e.shiftKey && (a === first || !panel.contains(a))) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && a === last) {
      e.preventDefault();
      first.focus();
    }
  });
  // Only take focus if the caller has not already focused something inside the dialog.
  if (!panel.contains(document.activeElement)) (initialFocus ?? focusables()[0])?.focus();
  // Restore focus to the opener when the dialog is dismissed (its overlay removed).
  const root = overlay.parentNode;
  if (root && typeof MutationObserver !== "undefined") {
    const mo = new MutationObserver(() => {
      if (!overlay.isConnected) {
        mo.disconnect();
        prev?.focus?.();
      }
    });
    mo.observe(root, { childList: true });
  }
}
