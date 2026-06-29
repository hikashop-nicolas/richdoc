// Find & Replace: a floating panel that searches the editable body (not the page chrome, the cloned
// headers/footers, or the off-screen measure copy, so no duplicate or phantom matches) and replaces
// through execCommand so edits join the native undo stack and the editor's dirty/reflow tracking.
// Matches are highlighted with the CSS Custom Highlight API, so nothing in the document DOM is mutated
// just to show them. v1 matches within a single text node (a match split across inline elements, e.g.
// "<b>fo</b>o", is not found) and replaces literally (no regex backreferences).
import { t } from "../i18n";

export interface FindReplaceDeps {
  doc: HTMLElement; // the editable body to search / replace within
  wrap: HTMLElement; // the editor chrome the panel overlays
}

interface Match { node: Text; start: number; end: number }

// Minimal typing for the CSS Custom Highlight API (not in every TS lib.dom yet); feature-detected.
type HighlightLike = { add(r: Range): void };
type HighlightCtor = new (...ranges: Range[]) => HighlightLike;
const HighlightAPI = (globalThis as unknown as { Highlight?: HighlightCtor }).Highlight;
const hlRegistry = (globalThis as unknown as { CSS?: { highlights?: { set(k: string, h: HighlightLike): void; delete(k: string): void } } }).CSS?.highlights;

const escapeRegExp = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const ICON = (paths: string): string =>
  `<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${paths}</svg>`;
const SEARCH_ICON = ICON('<circle cx="7" cy="7" r="4.5"/><path d="M11 11l3 3"/>');

export function setupFindReplace(deps: FindReplaceDeps) {
  const { doc, wrap } = deps;

  // --- Panel DOM ---------------------------------------------------------------------------------
  const panel = document.createElement("div");
  panel.className = "docxedit-find";
  panel.hidden = true;
  const mkInput = (placeholder: string): HTMLInputElement => {
    const i = document.createElement("input");
    i.type = "text"; i.className = "docxedit-find-input"; i.placeholder = placeholder;
    return i;
  };
  const findInput = mkInput(t("findLabel"));
  const replaceInput = mkInput(t("findReplaceWith"));
  const count = document.createElement("span");
  count.className = "docxedit-find-count";
  const mkBtn = (label: string, title: string, fn: () => void, html = false): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button"; b.className = "docxedit-find-btn"; b.title = title; b.setAttribute("aria-label", title);
    if (html) b.innerHTML = label; else b.textContent = label;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the document selection
    b.addEventListener("click", (e) => { e.preventDefault(); fn(); });
    return b;
  };
  const mkToggle = (label: string, title: string): HTMLButtonElement => {
    const b = mkBtn(label, title, () => { b.classList.toggle("is-on"); search(); });
    b.classList.add("docxedit-find-toggle");
    return b;
  };
  const caseBtn = mkToggle("Aa", t("findMatchCase"));
  const wordBtn = mkToggle("〈W〉", t("findWholeWord"));
  const regexBtn = mkToggle(".*", t("findRegex"));
  const prevBtn = mkBtn("‹", t("findPrev"), () => navigate(-1));
  const nextBtn = mkBtn("›", t("findNext"), () => navigate(1));
  const closeBtn = mkBtn("×", t("findClose"), () => close());
  const replaceBtn = mkBtn(t("findReplaceOne"), t("findReplaceOne"), () => replaceOne());
  const replaceAllBtn = mkBtn(t("findReplaceAll"), t("findReplaceAll"), () => replaceAll());

  const findRow = document.createElement("div");
  findRow.className = "docxedit-find-row";
  const opts = document.createElement("span");
  opts.className = "docxedit-find-opts";
  opts.append(caseBtn, wordBtn, regexBtn);
  findRow.append(SEARCH_ICON ? Object.assign(document.createElement("span"), { className: "docxedit-find-ico", innerHTML: SEARCH_ICON }) : document.createElement("span"), findInput, opts, count, prevBtn, nextBtn, closeBtn);
  const replaceRow = document.createElement("div");
  replaceRow.className = "docxedit-find-row";
  replaceRow.append(replaceInput, replaceBtn, replaceAllBtn);
  panel.append(findRow, replaceRow);
  wrap.appendChild(panel);

  const toggleBtn = document.createElement("button");
  toggleBtn.type = "button";
  toggleBtn.className = "docxedit-bottombar-btn";
  toggleBtn.innerHTML = SEARCH_ICON;
  toggleBtn.title = t("findReplace");
  toggleBtn.setAttribute("aria-label", t("findReplace"));
  toggleBtn.addEventListener("click", () => (panel.hidden ? open() : close()));

  // --- Search ------------------------------------------------------------------------------------
  let matches: Match[] = [];
  let current = 0;

  const buildRegex = (): RegExp | null => {
    const q = findInput.value;
    if (!q) return null;
    const flags = "g" + (caseBtn.classList.contains("is-on") ? "" : "i");
    try {
      if (regexBtn.classList.contains("is-on")) return new RegExp(q, flags);
      const pat = escapeRegExp(q);
      return new RegExp(wordBtn.classList.contains("is-on") ? `\\b${pat}\\b` : pat, flags);
    } catch {
      return null; // invalid regex: treat as no matches
    }
  };
  const collect = (re: RegExp): Match[] => {
    const out: Match[] = [];
    const walker = document.createTreeWalker(doc, NodeFilter.SHOW_TEXT, {
      acceptNode(n) {
        // Skip empty nodes and the contents of computed / non-editable widgets (fields, xrefs,
        // bookmark markers, the TOC, passthrough spans), which should not be searched or replaced.
        if (!n.nodeValue) return NodeFilter.FILTER_REJECT;
        const el = (n as Text).parentElement;
        if (!el || el.closest('[contenteditable="false"], .docx-field, .docx-field-toc, .docx-xref, .docx-pass')) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
      const text = node.nodeValue!;
      re.lastIndex = 0;
      for (let m = re.exec(text); m; m = re.exec(text)) {
        if (m[0].length === 0) { re.lastIndex++; continue; } // never advance on an empty match
        out.push({ node, start: m.index, end: m.index + m[0].length });
      }
    }
    return out;
  };
  const rangeOf = (m: Match): Range | null => {
    if (!m.node.isConnected || m.end > (m.node.nodeValue?.length ?? 0)) return null;
    const r = document.createRange();
    try { r.setStart(m.node, m.start); r.setEnd(m.node, m.end); return r; } catch { return null; }
  };
  const renderHighlights = (): void => {
    if (!HighlightAPI || !hlRegistry) return;
    if (panel.hidden || !matches.length) { hlRegistry.delete("rdoc-find"); hlRegistry.delete("rdoc-find-current"); return; }
    const all = new HighlightAPI();
    const cur = new HighlightAPI();
    matches.forEach((m, i) => { const r = rangeOf(m); if (r) (i === current ? cur : all).add(r); });
    hlRegistry.set("rdoc-find", all);
    hlRegistry.set("rdoc-find-current", cur);
  };
  const updateCount = (): void => {
    count.textContent = !findInput.value ? "" : matches.length ? `${current + 1} / ${matches.length}` : t("findNone");
    count.classList.toggle("is-none", !!findInput.value && !matches.length);
  };
  const search = (): void => {
    const re = buildRegex();
    matches = re ? collect(re) : [];
    if (current >= matches.length) current = Math.max(0, matches.length - 1);
    renderHighlights();
    updateCount();
  };
  const navigate = (delta: number): void => {
    if (!matches.length) return;
    current = (current + delta + matches.length) % matches.length;
    renderHighlights();
    updateCount();
    matches[current]?.node.parentElement?.scrollIntoView({ block: "center", behavior: "smooth" });
  };

  // --- Replace (via execCommand, so it joins undo + dirty/reflow) ---------------------------------
  const replaceAt = (idx: number): boolean => {
    const r = matches[idx] ? rangeOf(matches[idx]!) : null;
    if (!r) return false;
    const sel = window.getSelection();
    doc.focus();
    sel?.removeAllRanges(); sel?.addRange(r);
    document.execCommand(replaceInput.value ? "insertText" : "delete", false, replaceInput.value);
    return true;
  };
  const replaceOne = (): void => { if (matches.length && replaceAt(current)) search(); };
  const replaceAll = (): void => {
    let guard = 0;
    while (matches.length && guard++ < 10000) {
      const before = matches.length;
      if (!replaceAt(0)) break;
      search();
      if (matches.length >= before) break; // replacement still matches: stop rather than loop forever
    }
  };

  // --- Open / close ------------------------------------------------------------------------------
  const open = (): void => {
    panel.hidden = false;
    toggleBtn.classList.add("is-on");
    const sel = window.getSelection();
    const picked = sel && !sel.isCollapsed && doc.contains(sel.anchorNode) ? sel.toString() : "";
    if (picked && !picked.includes("\n")) findInput.value = picked;
    current = 0;
    search();
    findInput.focus();
    findInput.select();
  };
  const close = (): void => {
    panel.hidden = true;
    toggleBtn.classList.remove("is-on");
    renderHighlights(); // clears, since hidden
    doc.focus();
  };

  // --- Wiring ------------------------------------------------------------------------------------
  let searchTimer = 0;
  const scheduleSearch = () => { window.clearTimeout(searchTimer); searchTimer = window.setTimeout(search, 120); };
  findInput.addEventListener("input", () => { current = 0; scheduleSearch(); });
  findInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); navigate(e.shiftKey ? -1 : 1); }
    else if (e.key === "Escape") { e.preventDefault(); close(); }
  });
  replaceInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); replaceOne(); }
    else if (e.key === "Escape") { e.preventDefault(); close(); }
  });
  // Re-scan when the document changes while the panel is open (highlights stay in sync).
  const onDocInput = () => { if (!panel.hidden) scheduleSearch(); };
  doc.addEventListener("input", onDocInput);
  // Ctrl/Cmd-F opens the panel when focus is inside the editor (replacing the degraded native find).
  const onKey = (e: KeyboardEvent) => {
    if ((e.metaKey || e.ctrlKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === "f") { e.preventDefault(); open(); }
  };
  wrap.addEventListener("keydown", onKey);

  return {
    panel,
    toggleBtn,
    open,
    teardown(): void {
      doc.removeEventListener("input", onDocInput);
      wrap.removeEventListener("keydown", onKey);
      hlRegistry?.delete("rdoc-find");
      hlRegistry?.delete("rdoc-find-current");
      panel.remove();
      toggleBtn.remove();
    },
  };
}
