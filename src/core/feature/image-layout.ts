// Image layout toolbar: a small bar shown above a selected image to set how it sits relative
// to the text (in line / wrap / break / behind / in front), its alignment, and its alt text.
// It only writes the format-agnostic data-rdoc-* attributes (the adapters map them to docx
// wp:anchor / odt draw:frame); CSS renders them. Behind/front images can be dragged to position.
import { t } from "../i18n";
import { makeDialogAccessible } from "./dialog-a11y";
import type { ImageWrap } from "../types";
import { applyCaption, captionAfter, captionText, topBlock } from "./caption";

export interface ImageLayoutDeps {
  wrap: HTMLElement;
  doc: HTMLElement; // the editable body; the offset parent for behind/front images
  scroll: HTMLElement;
  mark: () => void;
  getZoom: () => number;
  reposition: () => void; // re-place the resize handle after a layout change moves the image
}

const WRAP_ICONS: Record<"inline" | "square" | "tight" | "topbottom" | "behind" | "front", string> = {
  inline:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="5" width="5" height="6" rx="1" fill="currentColor"/><rect x="8" y="5.4" width="6.5" height="1.4" rx=".6" fill="currentColor"/><rect x="8" y="9.2" width="6.5" height="1.4" rx=".6" fill="currentColor"/></svg>',
  square:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="2.5" width="6" height="6" rx="1" fill="currentColor"/><rect x="9" y="2.8" width="5.5" height="1.3" rx=".5" fill="currentColor"/><rect x="9" y="6" width="5.5" height="1.3" rx=".5" fill="currentColor"/><rect x="1.5" y="10.6" width="13" height="1.3" rx=".5" fill="currentColor"/><rect x="1.5" y="13.2" width="13" height="1.3" rx=".5" fill="currentColor"/></svg>',
  tight:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><path d="M8 2.2 11 6 8 9.8 5 6z" fill="currentColor"/><rect x="11.5" y="3.2" width="3" height="1.2" rx=".5" fill="currentColor"/><rect x="12" y="6" width="2.5" height="1.2" rx=".5" fill="currentColor"/><rect x="1.5" y="3.2" width="3" height="1.2" rx=".5" fill="currentColor"/><rect x="1.5" y="6" width="2.5" height="1.2" rx=".5" fill="currentColor"/><rect x="1.5" y="11" width="13" height="1.3" rx=".5" fill="currentColor"/><rect x="1.5" y="13.4" width="13" height="1.3" rx=".5" fill="currentColor"/></svg>',
  topbottom:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="1.5" width="13" height="1.3" rx=".5" fill="currentColor"/><rect x="1.5" y="3.9" width="13" height="1.3" rx=".5" fill="currentColor"/><rect x="4.5" y="6.4" width="7" height="3.2" rx="1" fill="currentColor"/><rect x="1.5" y="10.8" width="13" height="1.3" rx=".5" fill="currentColor"/><rect x="1.5" y="13.2" width="13" height="1.3" rx=".5" fill="currentColor"/></svg>',
  behind:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><rect x="3.5" y="3.5" width="9" height="9" rx="1" fill="currentColor" opacity=".35"/><rect x="1.5" y="5.2" width="13" height="1.5" rx=".6" fill="currentColor"/><rect x="1.5" y="8.6" width="13" height="1.5" rx=".6" fill="currentColor"/></svg>',
  front:
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><rect x="1.5" y="5.2" width="13" height="1.5" rx=".6" fill="currentColor" opacity=".35"/><rect x="1.5" y="8.6" width="13" height="1.5" rx=".6" fill="currentColor" opacity=".35"/><rect x="3.5" y="3.5" width="9" height="9" rx="1" fill="currentColor"/></svg>',
};
const ALIGN_ICONS: Record<"left" | "center" | "right", string> = {
  left: '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><rect x="2" y="3" width="8" height="1.6" rx=".6"/><rect x="2" y="7" width="12" height="1.6" rx=".6"/><rect x="2" y="11" width="8" height="1.6" rx=".6"/></svg>',
  center: '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><rect x="4" y="3" width="8" height="1.6" rx=".6"/><rect x="2" y="7" width="12" height="1.6" rx=".6"/><rect x="4" y="11" width="8" height="1.6" rx=".6"/></svg>',
  right: '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><rect x="6" y="3" width="8" height="1.6" rx=".6"/><rect x="2" y="7" width="12" height="1.6" rx=".6"/><rect x="6" y="11" width="8" height="1.6" rx=".6"/></svg>',
};

export function setupImageLayout(deps: ImageLayoutDeps) {
  const { wrap, doc, scroll, mark, getZoom, reposition } = deps;
  let img: HTMLImageElement | null = null;

  const bar = document.createElement("div");
  bar.className = "docxedit-imgbar";
  bar.hidden = true;
  const mkBtn = (svg: string, title: string, fn: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-imgbar-btn";
    b.innerHTML = svg;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the image selected
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      fn();
    });
    return b;
  };
  // Apply a wrap mode. inline = remove the layout attrs (back in line with text). Switching to
  // behind/front seeds the offset from where the image currently sits so it does not jump.
  const setWrap = (mode: "inline" | ImageWrap): void => {
    if (!img) return;
    if (mode === "inline") {
      for (const a of ["data-rdoc-wrap", "data-rdoc-align", "data-rdoc-x", "data-rdoc-y", "data-rdoc-wrapdist"]) img.removeAttribute(a);
      img.style.cssText = ""; // drop the left/top/margin the floating layout set
    } else {
      if (mode === "behind" || mode === "front") {
        if (!img.getAttribute("data-rdoc-wrap")) {
          const x = img.offsetLeft;
          const y = img.offsetTop;
          img.setAttribute("data-rdoc-x", String(x));
          img.setAttribute("data-rdoc-y", String(y));
          img.style.left = `${x}px`;
          img.style.top = `${y}px`;
        }
        for (const at of ["data-rdoc-absx", "data-rdoc-absy"]) img.removeAttribute(at); // behind/front carry their own x/y
      } else {
        // square / tight / top-and-bottom: the editor positions these by alignment, so drop any
        // imported absolute offset (it would otherwise re-emit a stale position the user can't see).
        for (const at of ["data-rdoc-absx", "data-rdoc-absy", "data-rdoc-x", "data-rdoc-y"]) img.removeAttribute(at);
      }
      img.setAttribute("data-rdoc-wrap", mode);
      if (!img.getAttribute("data-rdoc-align")) img.setAttribute("data-rdoc-align", "left");
    }
    mark();
    reposition();
    sync();
    place();
  };
  const setAlign = (a: "left" | "center" | "right"): void => {
    if (!img || !img.getAttribute("data-rdoc-wrap")) return;
    img.setAttribute("data-rdoc-align", a);
    // Choosing an alignment replaces an imported absolute offset with the alignment model.
    for (const at of ["data-rdoc-absx", "data-rdoc-absy", "data-rdoc-x", "data-rdoc-y"]) img.removeAttribute(at);
    mark();
    reposition();
    sync();
    place();
  };
  // A small dialog to set the image's alt text and an (optional) figure caption together.
  const overlay = document.createElement("div");
  overlay.className = "docxedit-dialog-overlay";
  overlay.hidden = true;
  const panel = document.createElement("div");
  panel.className = "docxedit-dialog docxedit-imgdialog";
  const dlgTitle = document.createElement("div");
  dlgTitle.className = "docxedit-dialog-title";
  dlgTitle.textContent = t("imageOptions");
  const mkField = (label: string): { row: HTMLElement; input: HTMLInputElement } => {
    const row = document.createElement("label");
    row.className = "docxedit-dialog-row docxedit-imgdialog-field";
    const span = document.createElement("span");
    span.textContent = label;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "docxedit-dialog-font";
    row.append(span, input);
    return { row, input };
  };
  const altField = mkField(t("altText"));
  const capField = mkField(t("caption"));
  const dlgActions = document.createElement("div");
  dlgActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const dlgCancel = document.createElement("button");
  dlgCancel.type = "button"; dlgCancel.className = "docxedit-menu-item"; dlgCancel.textContent = t("cancel");
  const dlgApply = document.createElement("button");
  dlgApply.type = "button"; dlgApply.className = "docxedit-menu-item docxedit-dialog-primary"; dlgApply.textContent = t("apply");
  dlgActions.append(dlgCancel, dlgApply);
  panel.append(dlgTitle, altField.row, capField.row, dlgActions);
  overlay.appendChild(panel);
  wrap.appendChild(overlay);
  makeDialogAccessible(overlay);
  // The dialog keeps its own image reference: opening it moves focus, which deselects the image
  // (img -> null), so apply must act on the image captured when the dialog opened.
  let dlgImg: HTMLImageElement | null = null;
  const closeDlg = (): void => { overlay.hidden = true; dlgImg = null; };
  const editAlt = (): void => {
    if (!img) return;
    dlgImg = img;
    altField.input.value = img.getAttribute("alt") || "";
    const block = topBlock(img, doc);
    const cap = block ? captionAfter(block, "figure") : null;
    capField.input.value = cap ? captionText(cap) : "";
    overlay.hidden = false;
    altField.input.focus();
  };
  dlgApply.addEventListener("click", () => {
    if (!dlgImg) { closeDlg(); return; }
    dlgImg.setAttribute("alt", altField.input.value);
    const block = topBlock(dlgImg, doc);
    if (block) applyCaption(block, "figure", capField.input.value);
    closeDlg();
    mark();
  });
  dlgCancel.addEventListener("click", closeDlg);
  overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) closeDlg(); });

  const wrapBtns: Record<string, HTMLButtonElement> = {
    inline: mkBtn(WRAP_ICONS.inline, t("wrapInline"), () => setWrap("inline")),
    square: mkBtn(WRAP_ICONS.square, t("wrapSquare"), () => setWrap("square")),
    tight: mkBtn(WRAP_ICONS.tight, t("wrapTight"), () => setWrap("tight")),
    topbottom: mkBtn(WRAP_ICONS.topbottom, t("wrapTopBottom"), () => setWrap("topbottom")),
    behind: mkBtn(WRAP_ICONS.behind, t("wrapBehind"), () => setWrap("behind")),
    front: mkBtn(WRAP_ICONS.front, t("wrapFront"), () => setWrap("front")),
  };
  const sep = document.createElement("span");
  sep.className = "docxedit-imgbar-sep";
  const alignBtns: Record<string, HTMLButtonElement> = {
    left: mkBtn(ALIGN_ICONS.left, t("alignLeft"), () => setAlign("left")),
    center: mkBtn(ALIGN_ICONS.center, t("alignCenter"), () => setAlign("center")),
    right: mkBtn(ALIGN_ICONS.right, t("alignRight"), () => setAlign("right")),
  };
  const sep2 = document.createElement("span");
  sep2.className = "docxedit-imgbar-sep";
  const altBtn = document.createElement("button");
  altBtn.type = "button";
  altBtn.className = "docxedit-imgbar-btn docxedit-imgbar-alt";
  altBtn.textContent = t("altText");
  altBtn.title = t("imageOptions");
  altBtn.addEventListener("mousedown", (e) => e.preventDefault());
  altBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    editAlt();
  });
  bar.append(wrapBtns.inline!, wrapBtns.square!, wrapBtns.tight!, wrapBtns.topbottom!, wrapBtns.behind!, wrapBtns.front!, sep, alignBtns.left!, alignBtns.center!, alignBtns.right!, sep2, altBtn);
  wrap.appendChild(bar);

  // Reflect the image's current wrap/align on the buttons; hide alignment for behind/front.
  const sync = (): void => {
    const w = img?.getAttribute("data-rdoc-wrap") ?? "inline";
    for (const [k, b] of Object.entries(wrapBtns)) b.classList.toggle("is-on", k === w);
    const aligned = w === "square" || w === "topbottom" || w === "tight";
    sep.hidden = !aligned;
    for (const b of Object.values(alignBtns)) b.hidden = !aligned;
    const a = img?.getAttribute("data-rdoc-align") ?? "left";
    for (const [k, b] of Object.entries(alignBtns)) b.classList.toggle("is-on", k === a);
  };
  // Position the bar just above the image (or below if there is no room), clamped to the view.
  const place = (): void => {
    if (!img || !wrap.contains(img)) {
      bar.hidden = true;
      return;
    }
    bar.hidden = false;
    const ir = img.getBoundingClientRect();
    const bw = bar.offsetWidth || 220;
    const bh = bar.offsetHeight || 32;
    let left = Math.max(8, Math.min(ir.left, window.innerWidth - bw - 8));
    let top = ir.top - bh - 6;
    if (top < 8) top = ir.bottom + 6;
    bar.style.left = `${left}px`;
    bar.style.top = `${top}px`;
  };

  // Drag a behind/front image to reposition it (offsets are in the doc's own px, so divide the
  // screen delta by the zoom). Other wrap modes flow with the text and are not draggable.
  let drag: { startX: number; startY: number; baseX: number; baseY: number } | null = null;
  const onDown = (e: PointerEvent): void => {
    if (!img || img !== e.target) return;
    const mode = img.getAttribute("data-rdoc-wrap");
    if (mode !== "behind" && mode !== "front") return;
    e.preventDefault();
    drag = { startX: e.clientX, startY: e.clientY, baseX: Number(img.getAttribute("data-rdoc-x")) || 0, baseY: Number(img.getAttribute("data-rdoc-y")) || 0 };
    try { img.setPointerCapture(e.pointerId); } catch { /* no pointer */ }
  };
  const onMove = (e: PointerEvent): void => {
    if (!drag || !img) return;
    const z = getZoom() || 1;
    const x = Math.round(drag.baseX + (e.clientX - drag.startX) / z);
    const y = Math.round(drag.baseY + (e.clientY - drag.startY) / z);
    img.setAttribute("data-rdoc-x", String(x));
    img.setAttribute("data-rdoc-y", String(y));
    img.style.left = `${x}px`;
    img.style.top = `${y}px`;
    place();
  };
  const onUp = (e: PointerEvent): void => {
    if (!drag) return;
    drag = null;
    try { img?.releasePointerCapture(e.pointerId); } catch { /* not captured */ }
    reposition();
    mark();
  };
  doc.addEventListener("pointerdown", onDown);
  doc.addEventListener("pointermove", onMove);
  doc.addEventListener("pointerup", onUp);
  scroll.addEventListener("scroll", place);

  const onSelect = (next: HTMLImageElement | null): void => {
    img = next;
    if (!img) {
      bar.hidden = true;
      return;
    }
    sync();
    place();
  };
  const teardown = (): void => {
    doc.removeEventListener("pointerdown", onDown);
    doc.removeEventListener("pointermove", onMove);
    doc.removeEventListener("pointerup", onUp);
    scroll.removeEventListener("scroll", place);
    bar.remove();
    overlay.remove();
  };
  return { onSelect, reposition: place, teardown };
}
