// Floating formatting bar (desktop only): a compact bold/italic/.../colour bar that appears
// near the caret or selection when the mouse is close to it, for quick formatting without
// reaching for the top toolbar. Hidden on coarse-pointer (touch) devices.
import { t } from "../../i18n";

export interface FloatBarDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  getActiveEl: () => HTMLElement;
  beginFormatChange: () => void;
  exec: (cmd: string, val?: string) => void;
  queryState: (cmd: string) => boolean;
  withSc: (title: string, key: string, opts?: { shift?: boolean; alt?: boolean }) => string;
  /** Vertical (tategaki) writing: stack the bar top-to-bottom beside the caret. */
  vertical: boolean;
}

export function setupFloatBar(deps: FloatBarDeps) {
  const { wrap, regions, getActiveEl, beginFormatChange, exec, queryState, withSc, vertical } = deps;
  const coarse = typeof window.matchMedia === "function" && window.matchMedia("(hover: none), (pointer: coarse)").matches;
  const floatBar = document.createElement("div");
  floatBar.className = `docxedit-floatbar${vertical ? " is-vertical" : ""}`;
  floatBar.hidden = true;
  let floatHideTimer = 0;
  let floatHovered = false;
  const fbtn = (label: string, title: string, cmd: string, cls: string): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `docxedit-floatbar-btn ${cls}`;
    b.textContent = label;
    b.title = title;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the selection
    b.addEventListener("click", () => {
      beginFormatChange();
      exec(cmd);
      updateFloatStates();
    });
    return b;
  };
  // A colour input applies foreColor / hiliteColor; keep the bar open while its picker is up.
  const fcolor = (title: string, cmd: string, value: string): HTMLInputElement => {
    const c = document.createElement("input");
    c.type = "color";
    c.value = value;
    c.title = title;
    c.className = "docxedit-floatbar-color";
    c.addEventListener("mousedown", () => {
      floatHovered = true;
      window.clearTimeout(floatHideTimer);
      getActiveEl().focus();
    });
    c.addEventListener("input", () => {
      beginFormatChange();
      exec(cmd, c.value);
    });
    c.addEventListener("change", () => {
      floatHovered = false;
    });
    return c;
  };
  const fBold = fbtn("B", withSc(t("bold"), "B"), "bold", "docxedit-tb-bold");
  const fItalic = fbtn("I", withSc(t("italic"), "I"), "italic", "docxedit-tb-italic");
  const fUnderline = fbtn("U", withSc(t("underline"), "U"), "underline", "docxedit-tb-underline");
  const fStrike = fbtn("S", t("strikethrough"), "strikeThrough", "docxedit-tb-strike");
  const fSup = fbtn("x²", t("superscript"), "superscript", "");
  const fSub = fbtn("x₂", t("subscript"), "subscript", "");
  const fColor = fcolor(t("textColor"), "foreColor", "#000000");
  const fBg = fcolor(t("highlight"), "hiliteColor", "#ffff00");
  floatBar.append(fBold, fItalic, fUnderline, fStrike, fSup, fSub, fColor, fBg);
  wrap.appendChild(floatBar);
  const updateFloatStates = () => {
    fBold.classList.toggle("is-on", queryState("bold"));
    fItalic.classList.toggle("is-on", queryState("italic"));
    fUnderline.classList.toggle("is-on", queryState("underline"));
    fStrike.classList.toggle("is-on", queryState("strikeThrough"));
    fSup.classList.toggle("is-on", queryState("superscript"));
    fSub.classList.toggle("is-on", queryState("subscript"));
  };
  const hideFloat = () => {
    floatBar.hidden = true;
  };
  const scheduleFloatHide = () => {
    window.clearTimeout(floatHideTimer);
    floatHideTimer = window.setTimeout(() => {
      if (!floatHovered) hideFloat();
    }, 350);
  };
  const showFloatAt = (rect: DOMRect) => {
    floatBar.hidden = false;
    const bw = floatBar.offsetWidth || (vertical ? 30 : 180);
    const bh = floatBar.offsetHeight || (vertical ? 180 : 32);
    if (vertical) {
      // Tategaki: a vertical stack to the right of the caret (flip left if there is no room),
      // centred along the caret's height.
      let left = rect.right + 8;
      if (left + bw > window.innerWidth - 8) left = rect.left - bw - 8;
      left = Math.max(8, left);
      let top = rect.top + rect.height / 2 - bh / 2;
      top = Math.max(8, Math.min(top, window.innerHeight - bh - 8));
      floatBar.style.left = `${left}px`;
      floatBar.style.top = `${top}px`;
      return;
    }
    let left = rect.left + rect.width / 2 - bw / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - bw - 8));
    let top = rect.top - bh - 8;
    if (top < 8) top = rect.bottom + 8; // not enough room above -> below
    floatBar.style.left = `${left}px`;
    floatBar.style.top = `${top}px`;
  };
  const selectionRect = (): DOMRect | null => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    const n = range.startContainer;
    const el = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
    if (!el || !regions.some((r) => r.contains(el))) return null;
    const rect = range.getBoundingClientRect();
    return rect && (rect.width > 0 || rect.height > 0) ? rect : null;
  };
  const onFloatMouseMove = (e: MouseEvent) => {
    if (floatHovered) return;
    // While an image is selected its own layout toolbar is shown; the text formatting bar must
    // stay out of the way (the selected image sits next to the caret, so proximity would trip it).
    if (wrap.querySelector("img.sel")) {
      hideFloat();
      return;
    }
    const rect = selectionRect();
    if (!rect) {
      scheduleFloatHide();
      return;
    }
    const dx = Math.max(rect.left - e.clientX, 0, e.clientX - rect.right);
    const dy = Math.max(rect.top - e.clientY, 0, e.clientY - rect.bottom);
    if (dx < 110 && dy < 90) {
      window.clearTimeout(floatHideTimer);
      updateFloatStates();
      showFloatAt(rect);
    } else {
      scheduleFloatHide();
    }
  };
  if (!coarse) {
    floatBar.addEventListener("mouseenter", () => {
      floatHovered = true;
      window.clearTimeout(floatHideTimer);
    });
    floatBar.addEventListener("mouseleave", () => {
      floatHovered = false;
      scheduleFloatHide();
    });
    document.addEventListener("mousemove", onFloatMouseMove);
    for (const r of regions) {
      r.addEventListener("keydown", hideFloat); // typing dismisses it
      r.addEventListener("scroll", hideFloat);
      r.addEventListener("mousedown", (e) => { if ((e.target as HTMLElement)?.closest?.("img")) hideFloat(); }); // selecting an image dismisses it
    }
    wrap.addEventListener("scroll", hideFloat, true);
  }

  const teardown = () => {
    document.removeEventListener("mousemove", onFloatMouseMove);
    window.clearTimeout(floatHideTimer);
  };
  return { teardown };
}
