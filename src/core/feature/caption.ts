// Captions: a numbered paragraph attached under a figure (image) or table. The number is a
// self-updating "seq" field (Word SEQ / odf text:sequence); the engine's decorateFields renumbers
// each type in document order. A caption paragraph carries data-rdoc-caption="figure|table", which
// is how the cross-reference dialog lists it and how the adapters tag it on a round-trip.
import { t } from "../i18n";

export type CaptionKind = "figure" | "table";

// The sequence identifier stored on the field (the Word SEQ name / odf sequence name). Kept distinct
// from the visible label, which is localised.
export const SEQ_ID: Record<CaptionKind, string> = { figure: "Figure", table: "Table" };

export function captionLabel(kind: CaptionKind): string {
  return kind === "table" ? t("captionTable") : t("captionFigure");
}

/** The auto-numbering field shown inside a caption (non-editable; decorateFields sets its number). */
export function makeSeqField(kind: CaptionKind): HTMLSpanElement {
  const s = document.createElement("span");
  s.className = "docx-field docx-field-seq";
  s.setAttribute("data-field", "seq");
  s.setAttribute("data-seq", SEQ_ID[kind]);
  s.contentEditable = "false";
  s.textContent = "1";
  return s;
}

/** Build a caption paragraph: "<Label> <n>: <text>" (the ": text" only when there is text). */
export function buildCaption(kind: CaptionKind, text: string): HTMLParagraphElement {
  const p = document.createElement("p");
  p.setAttribute("data-rdoc-caption", kind);
  p.append(document.createTextNode(captionLabel(kind) + " "), makeSeqField(kind));
  const trimmed = text.trim();
  if (trimmed) p.append(document.createTextNode(": " + trimmed));
  return p;
}

/** The user-facing caption text: everything after the number field, minus a leading ": ". */
export function captionText(p: Element): string {
  const seq = p.querySelector('[data-field="seq"]');
  if (!seq) return (p.textContent || "").trim();
  let s = "";
  for (let n = seq.nextSibling; n; n = n.nextSibling) s += n.textContent || "";
  return s.replace(/^\s*:\s*/, "").trim();
}

/** Replace the trailing text of an existing caption, keeping its label and number field. */
export function setCaptionText(p: Element, text: string): void {
  const seq = p.querySelector('[data-field="seq"]');
  if (!seq) return;
  while (seq.nextSibling) seq.nextSibling.remove();
  const trimmed = text.trim();
  if (trimmed) p.append(document.createTextNode(": " + trimmed));
}

/** A caption paragraph already sitting right after `block`, of the given kind, if any. */
export function captionAfter(block: Element, kind: CaptionKind): HTMLElement | null {
  const next = block.nextElementSibling as HTMLElement | null;
  return next && next.getAttribute("data-rdoc-caption") === kind ? next : null;
}

/** The top-level block (direct child of the editable body) containing `el`. */
export function topBlock(el: Element, doc: Element): HTMLElement | null {
  let n: Element = el;
  while (n.parentElement && n.parentElement !== doc) n = n.parentElement;
  return n.parentElement === doc ? (n as HTMLElement) : null;
}

// Add, update, or (when text is empty and one exists) remove the caption attached to `block`.
export function applyCaption(block: HTMLElement, kind: CaptionKind, text: string): void {
  const existing = captionAfter(block, kind);
  if (!text.trim()) { existing?.remove(); return; }
  if (existing) { setCaptionText(existing, text); return; }
  block.after(buildCaption(kind, text));
}
