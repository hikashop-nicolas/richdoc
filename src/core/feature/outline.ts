// Document outline (navigation) pane: a collapsible left sidebar listing the document's headings
// (H1-3) as a clickable, indented list. It is a pure view over the live DOM, rebuilt after each
// reflow but only when the heading set actually changed (a signature guard, like the TOC field).
// Clicking an entry scrolls that heading into view.
import { t } from "../i18n";

export interface OutlineDeps {
  doc: HTMLElement; // the editable body the headings live in
}

// A small "list / outline" glyph for the toggle button.
const OUTLINE_ICON =
  '<svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
  '<circle cx="2.4" cy="3.2" r="1.1"/><rect x="5" y="2.4" width="9" height="1.6" rx=".6"/>' +
  '<circle cx="4" cy="8" r="1.1"/><rect x="6.6" y="7.2" width="7.4" height="1.6" rx=".6"/>' +
  '<circle cx="5.6" cy="12.8" r="1.1"/><rect x="8.2" y="12" width="5.8" height="1.6" rx=".6"/></svg>';

export function setupOutline(deps: OutlineDeps) {
  const { doc } = deps;

  const pane = document.createElement("aside");
  pane.className = "docxedit-outline";
  pane.hidden = true;
  const title = document.createElement("div");
  title.className = "docxedit-outline-title";
  title.textContent = t("outline");
  const list = document.createElement("div");
  list.className = "docxedit-outline-list";
  pane.append(title, list);

  const toggleBtn = document.createElement("button");
  toggleBtn.type = "button";
  toggleBtn.className = "docxedit-bottombar-btn";
  toggleBtn.innerHTML = OUTLINE_ICON;
  toggleBtn.title = t("outline");
  toggleBtn.setAttribute("aria-label", t("outline"));

  const headings = (): HTMLElement[] =>
    Array.from(doc.querySelectorAll<HTMLElement>("h1,h2,h3")).filter((h) => !h.closest(".docx-field-toc"));

  let sig = "";
  const refresh = (): void => {
    if (pane.hidden) return; // nothing to draw while collapsed; rebuilt on open
    const hs = headings();
    const next = hs.map((h) => `${h.tagName}:${(h.textContent ?? "").trim()}`).join("|");
    if (next === sig) return; // headings unchanged: skip the rebuild
    sig = next;
    if (!hs.length) {
      const empty = document.createElement("div");
      empty.className = "docxedit-outline-empty";
      empty.textContent = t("tocEmpty");
      list.replaceChildren(empty);
      return;
    }
    list.replaceChildren(
      ...hs.map((h) => {
        const row = document.createElement("button");
        row.type = "button";
        row.className = `docxedit-outline-row docxedit-outline-${h.tagName.toLowerCase()}`;
        row.textContent = (h.textContent ?? "").trim() || t("untitled");
        row.title = row.textContent;
        row.addEventListener("click", () => h.scrollIntoView({ block: "center", behavior: "smooth" }));
        return row;
      }),
    );
  };

  let open = false;
  const setOpen = (v: boolean): void => {
    open = v;
    pane.hidden = !v;
    toggleBtn.classList.toggle("is-on", v);
    toggleBtn.setAttribute("aria-pressed", String(v));
    if (v) { sig = ""; refresh(); } // force a rebuild on open
  };
  toggleBtn.addEventListener("click", () => setOpen(!open));

  return {
    pane,
    toggleBtn,
    refresh,
    teardown(): void { pane.remove(); toggleBtn.remove(); },
  };
}
