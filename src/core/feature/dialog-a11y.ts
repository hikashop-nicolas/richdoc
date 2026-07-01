// Shared modal accessibility for the standard docxedit dialog (an overlay containing a
// .docxedit-dialog panel with a .docxedit-dialog-title). Marks it as a modal dialog for
// assistive tech, traps Tab within the panel, focuses the first field when it opens, and
// restores focus to the opener when it closes. Escape / overlay-click dismissal stays with
// each dialog's own handlers.
//
// Works for both dialog lifecycles used here: created-per-open (appended then removed) and
// persistent (created once, shown/hidden via the `hidden` attribute), by observing the
// overlay's visibility rather than assuming one or the other. Call once, right after the
// overlay is created/appended.
let dlgSeq = 0;

export function makeDialogAccessible(overlay: HTMLElement, initialFocus?: HTMLElement): void {
  const panel = overlay.querySelector<HTMLElement>(".docxedit-dialog");
  if (!panel) return;
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

  const isVisible = (): boolean => overlay.isConnected && !overlay.hidden && getComputedStyle(overlay).display !== "none";
  let opener: HTMLElement | null = null;
  let wasVisible = false;
  const check = (): void => {
    const vis = isVisible();
    if (vis && !wasVisible) {
      opener = document.activeElement as HTMLElement | null; // restored on close
      // Defer so any post-open focus handling from the opener runs first.
      setTimeout(() => {
        if (isVisible() && !panel.contains(document.activeElement)) (initialFocus ?? focusables()[0])?.focus();
      }, 0);
    } else if (!vis && wasVisible) {
      opener?.focus?.();
      opener = null;
      if (!overlay.isConnected) mo.disconnect(); // a per-open dialog is gone for good
    }
    wasVisible = vis;
  };
  const mo = new MutationObserver(check);
  mo.observe(overlay, { attributes: true, attributeFilter: ["hidden", "style", "class"] });
  if (overlay.parentNode) mo.observe(overlay.parentNode, { childList: true });
  check(); // handle a dialog that is already visible at call time (per-open pattern)
}
