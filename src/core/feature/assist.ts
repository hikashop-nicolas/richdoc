// On-device writing assist: a toolbar control that runs translate / elaborate / shorten /
// write over the current selection (or, for "write", a typed instruction) using localml,
// entirely in the browser. localml is imported LAZILY (dynamic import) the first time an
// action runs, so the editor's base bundle never pulls in transformers.js until used, and a
// host that sets `assist:false` never loads it at all.
//
// The result streams into a preview popover; nothing touches the document until the user hits
// Accept, which applies it through execCommand so it joins the native undo stack and the
// editor's dirty/reflow tracking (like find-replace).
import { t } from "../i18n";

export interface AssistDeps {
  doc: HTMLElement; // the editable body
  wrap: HTMLElement; // the editor chrome the popover overlays
  regions: HTMLElement[]; // editable regions (body + header/footer); a selection must be inside one
  locale: string; // the editor locale, used as the default translate source language
}

type Task = "translate" | "elaborate" | "shorten" | "write";
const SELECTION_TASKS: Task[] = ["translate", "elaborate", "shorten"]; // need a non-empty selection

const ICON = (paths: string): string =>
  `<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${paths}</svg>`;
// A small "sparkle" mark for the assist control.
const SPARKLE = ICON('<path d="M8 2.5l1.1 2.9L12 6.5 9.1 7.6 8 10.5 6.9 7.6 4 6.5l2.9-1.1z"/><path d="M12.5 11l.5 1.3 1.3.5-1.3.5-.5 1.3-.5-1.3-1.3-.5 1.3-.5z"/>');

export function setupAssist(deps: AssistDeps) {
  const { wrap, regions } = deps;

  // --- toolbar control -------------------------------------------------------
  const button = document.createElement("button");
  button.type = "button";
  button.className = "docxedit-assist-btn";
  button.innerHTML = `${SPARKLE}<span>${t("assist")}</span>`;
  button.title = t("assist");
  button.setAttribute("aria-label", t("assist"));
  button.setAttribute("aria-haspopup", "menu");
  button.addEventListener("mousedown", (e) => e.preventDefault()); // keep the document selection

  // --- dropdown menu ---------------------------------------------------------
  const menu = document.createElement("div");
  menu.className = "docxedit-assist-menu";
  menu.setAttribute("role", "menu");
  menu.hidden = true;
  const TASK_LABEL: Record<Task, string> = {
    translate: "assistTranslate",
    elaborate: "assistElaborate",
    shorten: "assistShorten",
    write: "assistWrite",
  };
  const items: Record<Task, HTMLButtonElement> = {} as Record<Task, HTMLButtonElement>;
  for (const task of ["translate", "elaborate", "shorten", "write"] as Task[]) {
    const it = document.createElement("button");
    it.type = "button";
    it.className = "docxedit-assist-item";
    it.setAttribute("role", "menuitem");
    it.textContent = t(TASK_LABEL[task]);
    it.addEventListener("mousedown", (e) => e.preventDefault());
    it.addEventListener("click", () => {
      closeMenu();
      openPopover(task);
    });
    items[task] = it;
    menu.appendChild(it);
  }

  const closeMenu = (): void => {
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
  };
  const openMenu = (): void => {
    // Selection-based tasks need a non-empty selection inside the document.
    const sel = selectionText();
    for (const task of SELECTION_TASKS) items[task].disabled = !sel;
    menu.hidden = false;
    button.setAttribute("aria-expanded", "true");
  };
  button.addEventListener("click", () => (menu.hidden ? openMenu() : closeMenu()));
  document.addEventListener("mousedown", (e) => {
    if (!button.isConnected) return; // this editor was destroyed; ignore
    if (!menu.hidden && !menu.contains(e.target as Node) && e.target !== button) closeMenu();
  });

  // --- preview popover -------------------------------------------------------
  const pop = document.createElement("div");
  pop.className = "docxedit-assist-pop";
  pop.hidden = true;
  const popTitle = document.createElement("div");
  popTitle.className = "docxedit-assist-title";
  const controls = document.createElement("div"); // task-specific inputs (langs / instruction)
  controls.className = "docxedit-assist-controls";
  const output = document.createElement("div");
  output.className = "docxedit-assist-output";
  output.setAttribute("aria-live", "polite");
  const progress = document.createElement("div");
  progress.className = "docxedit-assist-progress";
  const actions = document.createElement("div");
  actions.className = "docxedit-assist-actions";
  const genBtn = mkBtn(t("assistGenerate"), () => void run());
  const acceptBtn = mkBtn(t("assistAccept"), () => apply());
  const cancelBtn = mkBtn(t("assistCancel"), () => closePopover());
  acceptBtn.classList.add("is-primary");
  actions.append(genBtn, acceptBtn, cancelBtn);
  pop.append(popTitle, controls, progress, output, actions);
  wrap.appendChild(pop);

  // --- state -----------------------------------------------------------------
  let task: Task = "elaborate";
  let savedRange: Range | null = null; // where the result will be applied
  let savedText = ""; // the selected text captured at open (the selection may change after)
  let result = "";
  let running: { cancel(): void } | null = null;
  let fromSel: HTMLSelectElement | null = null;
  let toSel: HTMLSelectElement | null = null;
  let instr: HTMLTextAreaElement | null = null;

  function openPopover(next: Task): void {
    task = next;
    savedRange = currentRange();
    savedText = savedRange ? savedRange.toString() : "";
    result = "";
    output.textContent = "";
    progress.textContent = "";
    acceptBtn.disabled = true;
    popTitle.textContent = t(TASK_LABEL[next]);
    buildControls(next);
    pop.hidden = false;
    (instr ?? genBtn).focus();
  }

  function buildControls(next: Task): void {
    controls.textContent = "";
    fromSel = toSel = null;
    instr = null;
    if (next === "translate") {
      fromSel = document.createElement("select");
      toSel = document.createElement("select");
      controls.append(labelled(t("assistFrom"), fromSel), labelled(t("assistTo"), toSel));
      void populateLangs(); // lazy: pulls the language list from localml
    } else if (next === "write") {
      instr = document.createElement("textarea");
      instr.className = "docxedit-assist-instr";
      instr.placeholder = t("assistWritePlaceholder");
      controls.append(instr);
    } else {
      const preview = document.createElement("div");
      preview.className = "docxedit-assist-src";
      preview.textContent = truncate(savedText, 240);
      controls.append(preview);
    }
  }

  async function populateLangs(): Promise<void> {
    try {
      const { TRANSLATE_LANGS } = await import("localml/translate");
      for (const l of TRANSLATE_LANGS) {
        fromSel!.add(new Option(l.label, l.code));
        toSel!.add(new Option(l.label, l.code));
      }
      fromSel!.value = TRANSLATE_LANGS.some((l) => l.code === deps.locale) ? deps.locale : "en";
      toSel!.value = fromSel!.value === "en" ? "fr" : "en";
    } catch {
      progress.textContent = t("assistUnavailable");
    }
  }

  async function run(): Promise<void> {
    running?.cancel();
    result = "";
    output.textContent = "";
    acceptBtn.disabled = true;
    genBtn.disabled = true;
    progress.textContent = t("assistLoading");
    const onProgress = (stage: string, ratio: number): void => {
      progress.textContent = stage === "download" ? t("assistDownloading", { pct: Math.round(ratio * 100) }) : t("assistGenerating");
    };
    try {
      if (task === "translate") {
        const { runTranslate, DEFAULT_TRANSLATE_MODEL } = await import("localml/translate");
        const lines = savedText.split("\n");
        const parts: string[] = [];
        const r = runTranslate(lines, { model: DEFAULT_TRANSLATE_MODEL, srcLang: fromSel!.value, tgtLang: toSel!.value }, {
          onProgress: (p) => onProgress(p.stage, p.ratio),
          onPartial: (start, texts) => {
            texts.forEach((tx, k) => (parts[start + k] = tx));
            result = parts.join("\n");
            output.textContent = result;
          },
          onDevice: () => (progress.textContent = t("assistGenerating")),
        });
        running = r;
        await r.done;
      } else {
        const { runGenerate } = await import("localml/generate");
        const input = task === "write" ? (instr?.value ?? "") : savedText;
        if (!input.trim()) {
          progress.textContent = t("assistNeedInput");
          genBtn.disabled = false;
          return;
        }
        const r = runGenerate(input, { task }, {
          onProgress: (p) => onProgress(p.stage, p.ratio),
          onPartial: (text) => {
            result = text;
            output.textContent = text;
          },
          onDevice: () => (progress.textContent = t("assistGenerating")),
        });
        running = { cancel: r.cancel };
        const out = await r.done;
        result = out.text || result;
        output.textContent = result;
      }
      progress.textContent = t("assistDone");
      acceptBtn.disabled = !result.trim();
    } catch (e) {
      progress.textContent = t("assistError", { msg: (e as Error).message });
    } finally {
      running = null;
      genBtn.disabled = false;
    }
  }

  function apply(): void {
    if (!result.trim() || !savedRange) return closePopover();
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
    deps.doc.focus();
    // execCommand keeps the edit on the native undo stack and fires the input listeners the
    // editor uses for dirty/reflow/word-count.
    document.execCommand("insertText", false, result);
    closePopover();
  }

  function closePopover(): void {
    running?.cancel();
    running = null;
    pop.hidden = true;
    savedRange = null;
  }

  document.addEventListener("keydown", (e) => {
    if (!button.isConnected || e.key !== "Escape") return; // ignore once the editor is gone
    if (!pop.hidden) closePopover();
    else if (!menu.hidden) closeMenu();
  });

  // Button + its dropdown share a positioned container so the menu drops under the button.
  const control = document.createElement("span");
  control.className = "docxedit-assist";
  control.append(button, menu);

  // --- helpers ---------------------------------------------------------------
  // The selection Range if it sits inside an editable region, else null.
  function currentRange(): Range | null {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const r = sel.getRangeAt(0);
    return regions.some((reg) => reg.contains(r.commonAncestorContainer)) ? r.cloneRange() : null;
  }
  function selectionText(): string {
    const r = currentRange();
    return r ? r.toString() : "";
  }

  return { control };
}

function mkBtn(label: string, fn: () => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "docxedit-assist-action";
  b.textContent = label;
  b.addEventListener("mousedown", (e) => e.preventDefault());
  b.addEventListener("click", (e) => {
    e.preventDefault();
    fn();
  });
  return b;
}

function labelled(label: string, control: HTMLElement): HTMLElement {
  const l = document.createElement("label");
  l.className = "docxedit-assist-field";
  const s = document.createElement("span");
  s.textContent = label;
  l.append(s, control);
  return l;
}

const truncate = (s: string, n: number): string => (s.length > n ? `${s.slice(0, n)}…` : s);
