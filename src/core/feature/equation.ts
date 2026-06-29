// Equation editor: insert and edit math, authored as LaTeX and stored as MathML (which the browser
// renders natively). LaTeX -> MathML uses temml, lazy-loaded only when the dialog opens. The MathML is
// the in-document representation; the docx adapter converts it to/from OMML. An equation is a single
// non-editable inline span: <span class="docx-eq" data-rdoc-eq data-latex="...">{<math>…</math>}</span>.
// Insert / edit / delete go through execCommand so they join the native undo stack and dirty tracking.
import { t } from "../i18n";

export interface EquationDeps {
  doc: HTMLElement; // the editable body
  wrap: HTMLElement; // chrome the dialog overlays
  mark: () => void; // flag dirty + schedule reflow after a DOM edit
  captureSel: () => void;
  restoreSel: () => void;
}

type Temml = { renderToString(latex: string, opts?: { xml?: boolean; displayMode?: boolean; throwOnError?: boolean }): string };
let temmlMod: Temml | null = null;
const loadTemml = async (): Promise<Temml> => {
  if (!temmlMod) {
    // temml's export shape varies through bundlers (namespace / default / nested default); pick the
    // level that actually exposes renderToString.
    const mod = (await import("temml")) as unknown as Record<string, unknown>;
    const cands = [mod, mod.default, (mod.default as Record<string, unknown> | undefined)?.default];
    const found = cands.find((c) => c && typeof (c as Partial<Temml>).renderToString === "function") as Temml | undefined;
    if (!found) throw new Error("temml: renderToString not found");
    temmlMod = found;
  }
  return temmlMod;
};
const toMathml = (temml: Temml, latex: string): string => temml.renderToString(latex, { xml: true, displayMode: false, throwOnError: false });

export function setupEquation(deps: EquationDeps) {
  const { doc, wrap, mark, captureSel, restoreSel } = deps;

  const overlay = document.createElement("div");
  overlay.className = "docxedit-dialog-overlay";
  overlay.hidden = true;
  const panel = document.createElement("div");
  panel.className = "docxedit-dialog docxedit-eq";
  const title = document.createElement("div");
  title.className = "docxedit-dialog-title";
  title.textContent = t("insertEquation");
  const input = document.createElement("textarea");
  input.className = "docxedit-eq-input";
  input.rows = 2;
  input.placeholder = "x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}";
  input.spellcheck = false;
  const preview = document.createElement("div");
  preview.className = "docxedit-eq-preview";
  const actions = document.createElement("div");
  actions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button"; cancelBtn.className = "docxedit-menu-item"; cancelBtn.textContent = t("cancel");
  const applyBtn = document.createElement("button");
  applyBtn.type = "button"; applyBtn.className = "docxedit-menu-item docxedit-dialog-primary"; applyBtn.textContent = t("insert");
  actions.append(cancelBtn, applyBtn);
  panel.append(title, input, preview, actions);
  overlay.appendChild(panel);
  wrap.appendChild(overlay);

  let editing: HTMLElement | null = null; // the equation being edited, or null for a fresh insert
  const close = (): void => { overlay.hidden = true; editing = null; };

  // Render the current LaTeX into the preview (a temml error renders inline in red).
  let previewTimer = 0;
  const renderPreview = async (): Promise<void> => {
    const latex = input.value.trim();
    if (!latex) { preview.replaceChildren(); return; }
    try { preview.innerHTML = toMathml(await loadTemml(), latex); } catch { preview.textContent = "…"; }
  };
  input.addEventListener("input", () => { window.clearTimeout(previewTimer); previewTimer = window.setTimeout(renderPreview, 150); });

  const open = (target?: HTMLElement): void => {
    captureSel();
    editing = target ?? null;
    input.value = target?.getAttribute("data-latex") ?? "";
    void renderPreview();
    overlay.hidden = false;
    input.focus();
  };

  const mkSpan = (latex: string, mathml: string): HTMLElement => {
    const span = document.createElement("span");
    span.className = "docx-eq";
    span.setAttribute("data-rdoc-eq", "");
    span.setAttribute("data-latex", latex);
    span.setAttribute("contenteditable", "false");
    span.innerHTML = mathml; // a fresh equation has no original OMML, so the writer rebuilds it from MathML
    return span;
  };
  const apply = async (): Promise<void> => {
    const latex = input.value.trim();
    const target = editing;
    close();
    if (target && !latex) { target.remove(); mark(); return; } // cleared: remove the equation
    if (!latex) return;
    let mathml: string;
    try { mathml = toMathml(await loadTemml(), latex); } catch { return; }
    const span = mkSpan(latex, mathml);
    if (target) { target.replaceWith(span); mark(); return; }
    // Insert at the saved caret via a Range (execCommand misplaces a non-editable inline span,
    // dropping it outside the paragraph); this matches how bookmarks / captions are inserted.
    restoreSel();
    const sel = window.getSelection();
    const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    const cont = range && (range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element));
    if (range && cont && doc.contains(cont)) { range.collapse(false); range.insertNode(span); range.setStartAfter(span); range.collapse(true); }
    else { (doc.querySelector("p,h1,h2,h3,li") ?? doc).appendChild(span); } // no caret: append to the first block
    mark();
  };

  cancelBtn.addEventListener("click", close);
  applyBtn.addEventListener("click", () => void apply());
  overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) close(); });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.preventDefault(); close(); }
    else if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) { e.preventDefault(); void apply(); }
  });
  // Click an existing equation to edit it.
  const onClick = (e: MouseEvent) => {
    const eq = (e.target as HTMLElement).closest?.(".docx-eq") as HTMLElement | null;
    if (eq && doc.contains(eq)) { e.preventDefault(); open(eq); }
  };
  doc.addEventListener("click", onClick);

  return {
    openDialog: () => open(),
    teardown(): void { doc.removeEventListener("click", onClick); overlay.remove(); },
  };
}
