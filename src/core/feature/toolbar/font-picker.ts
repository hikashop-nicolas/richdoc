// The toolbar font picker: fontshelf's grouped <select> wired to the editor's selection and
// tip. Detection, the catalog, consent and the dialogs live in fontshelf.
import { createFontSelect } from "fontshelf";
import type { FontOptions } from "../../types";

export interface FontPickerDeps {
  wrap: HTMLElement; // chrome the dialogs overlay
  docFonts: readonly string[]; // fonts the open document uses
  fontOptions: FontOptions | undefined;
  getActiveEl: () => HTMLElement;
  captureSel: () => void;
  restoreSel: () => void;
  apply: (family: string) => void; // set the font on the selection
}

export function setupFontPicker(deps: FontPickerDeps) {
  const { wrap, fontOptions } = deps;
  const tip = (text: string): void => {
    const el = document.createElement("div");
    el.className = "docxedit-cmt-pop docxedit-cmt-tip";
    el.setAttribute("role", "status");
    el.textContent = text;
    wrap.appendChild(el);
    setTimeout(() => el.remove(), 2500);
  };
  const picker = createFontSelect({
    docFonts: deps.docFonts,
    local: fontOptions?.local,
    remote: fontOptions?.remote ?? false,
    container: wrap,
    apply: deps.apply,
    onAction: deps.captureSel,
    onPicked: deps.restoreSel,
    focusTarget: deps.getActiveEl,
    notify: tip,
  });
  return {
    select: picker.select,
    /** Every font in the picker (document fonts first). */
    fonts: (): string[] => picker.fonts(),
    has: (family: string): boolean => picker.has(family),
    /** Show a value (caret sync); it is also what an action entry restores. */
    set: (value: string): void => picker.set(value),
    teardown: (): void => picker.teardown(),
  };
}
