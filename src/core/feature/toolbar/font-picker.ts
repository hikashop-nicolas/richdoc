// The font picker: a native <select> (robust on touch and screen readers) grouped as
// "In this document" / "Fonts" / "More fonts". The last group holds action entries
// ("Show my fonts", "Get more fonts") that run and then restore the previous choice.
import { t } from "../../i18n";
import { makeDialogAccessible } from "../dialog-a11y";
import {
  BASE_FONTS, LOCAL_CANDIDATES, WEB_FONTS, GROUP_SAMPLE, buildFontGroups, canDraw, detectInstalled,
  hasLocalFontAccess, localFontsPermission, queryLocalFamilies, readFontPrefs, writeFontPrefs,
  loadWebFont, loadWebFontPreviews, findWebFont, mergeNames, previewFamily, cssFamily, type WebFont,
} from "../../fonts";
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

const ACTION = "\u0001"; // value prefix of the action entries
const SHOW_MINE = `${ACTION}local`;
const GET_MORE = `${ACTION}remote`;

// Installed fonts found this session, shared by every editor.
let sessionLocal: { names: string[]; full: boolean } | null = null;
// Web font families loaded this session.
const sessionAdded = new Set<string>();

export function setupFontPicker(deps: FontPickerDeps) {
  const { wrap, fontOptions, getActiveEl, captureSel, restoreSel, apply } = deps;
  const remote = !!fontOptions?.remote && fontOptions.remote.provider === "google";
  const localOn = fontOptions?.local !== false;
  const docFonts = [...deps.docFonts];
  let local: string[] = sessionLocal?.names ?? [];
  let localFull = sessionLocal?.full ?? false; // the full queryLocalFonts list, not just detection
  let added: string[] = [];
  let all: string[] = [];
  let disposed = false;

  const select = document.createElement("select");
  select.title = t("font");
  select.setAttribute("aria-label", t("font"));

  const render = (): void => {
    const keep = select.value;
    const groups = buildFontGroups({ docFonts, base: BASE_FONTS, extra: [...local, ...added], canDraw: (n) => sessionAdded.has(n.toLowerCase()) || canDraw(n) });
    all = [...groups.doc.map((d) => d.name), ...groups.fonts];
    const head = new Option(t("font"), "");
    head.disabled = true;
    const parts: (HTMLOptionElement | HTMLOptGroupElement)[] = [head];
    const group = (label: string, opts: HTMLOptionElement[]): void => {
      if (!opts.length) return;
      const g = document.createElement("optgroup");
      g.label = label;
      g.append(...opts);
      parts.push(g);
    };
    group(t("fontsInDocument"), groups.doc.map((d) => new Option(d.missing ? t("fontMissing", { name: d.name }) : d.name, d.name)));
    group(t("fontsAll"), groups.fonts.map((f) => new Option(f, f)));
    const actions: HTMLOptionElement[] = [];
    if (localOn && !localFull) actions.push(new Option(t("fontsShowMine"), SHOW_MINE));
    if (remote) actions.push(new Option(t("fontsGetMore"), GET_MORE));
    group(t("fontsMore"), actions);
    select.replaceChildren(...parts);
    select.value = all.includes(keep) ? keep : "";
  };

  // The value before the current interaction, restored after an action entry.
  let prev = "";
  let navKey = false;
  const remember = (): void => { if (!select.value.startsWith(ACTION)) prev = select.value; };
  select.addEventListener("pointerdown", remember);
  select.addEventListener("focus", remember);
  select.addEventListener("mousedown", () => getActiveEl().focus());
  // Arrow keys on a closed select fire change per step: never run an action from those.
  select.addEventListener("keydown", (e) => {
    remember();
    navKey = /^(Arrow|Page|Home|End)/.test(e.key);
  });
  select.addEventListener("change", () => {
    const v = select.value;
    const fromNav = navKey;
    navKey = false;
    if (!v) return;
    if (!v.startsWith(ACTION)) {
      prev = v;
      apply(v);
      return;
    }
    select.value = prev;
    if (fromNav) return;
    captureSel();
    if (v === SHOW_MINE) void showMine(true);
    else if (v === GET_MORE) openRemote();
  });

  const tip = (text: string): void => {
    const el = document.createElement("div");
    el.className = "docxedit-cmt-pop docxedit-cmt-tip";
    el.setAttribute("role", "status");
    el.textContent = text;
    wrap.appendChild(el);
    setTimeout(() => el.remove(), 2500);
  };

  // Installed fonts: the Local Font Access API when allowed, else probe the built-in list.
  // `gesture` = the user asked now, so the permission prompt may show.
  const showMine = async (gesture: boolean): Promise<void> => {
    const before = all.length;
    let names: string[] | null = null;
    if (hasLocalFontAccess() && (gesture || (await localFontsPermission()) === "granted")) names = await queryLocalFamilies();
    const full = !!names;
    if (!names) names = await detectInstalled(LOCAL_CANDIDATES);
    if (disposed) return;
    local = mergeNames(local, names);
    localFull = localFull || full;
    sessionLocal = { names: local, full: localFull };
    if (gesture) writeFontPrefs({ local: true });
    render();
    if (gesture) {
      const n = all.length - before;
      tip(n > 0 ? t("fontsFound", { n }) : t("fontsNoneFound"));
    }
  };

  // --- Get more fonts ----------------------------------------------------------------------
  let consentOverlay: HTMLElement | null = null;
  let catalogOverlay: HTMLElement | null = null;
  let catalogSearch: HTMLInputElement | null = null;
  const openRemote = (): void => {
    if (readFontPrefs().consent) openCatalog();
    else openConsent();
  };

  const dialog = (cls: string, title: string): { overlay: HTMLElement; panel: HTMLElement; actions: HTMLElement } => {
    const overlay = document.createElement("div");
    overlay.className = "docxedit-dialog-overlay";
    overlay.hidden = true;
    const panel = document.createElement("div");
    panel.className = `docxedit-dialog ${cls}`;
    const h = document.createElement("div");
    h.className = "docxedit-dialog-title";
    h.textContent = title;
    const actions = document.createElement("div");
    actions.className = "docxedit-dialog-row docxedit-dialog-actions";
    panel.append(h, actions);
    overlay.appendChild(panel);
    wrap.appendChild(overlay);
    overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) overlay.hidden = true; });
    overlay.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); overlay.hidden = true; } });
    return { overlay, panel, actions };
  };
  const button = (label: string, primary = false): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = primary ? "docxedit-menu-item docxedit-dialog-primary" : "docxedit-menu-item";
    b.textContent = label;
    return b;
  };

  const openConsent = (): void => {
    if (!consentOverlay) {
      const { overlay, panel, actions } = dialog("docxedit-fonts-consent", t("fontsGetTitle"));
      const p = document.createElement("p");
      p.className = "docxedit-fonts-note";
      p.textContent = t("fontsConsent");
      panel.insertBefore(p, actions);
      const cancel = button(t("cancel"));
      const ok = button(t("fontsContinue"), true);
      actions.append(cancel, ok);
      cancel.addEventListener("click", () => (overlay.hidden = true));
      ok.addEventListener("click", () => {
        overlay.hidden = true;
        writeFontPrefs({ consent: true });
        openCatalog();
      });
      makeDialogAccessible(overlay, ok);
      consentOverlay = overlay;
    }
    consentOverlay.hidden = false;
  };

  const openCatalog = (): void => {
    if (!catalogOverlay) buildCatalog();
    catalogOverlay!.hidden = false;
    void loadWebFontPreviews(WEB_FONTS);
    setTimeout(() => catalogSearch?.focus(), 0);
  };

  const buildCatalog = (): void => {
    const { overlay, panel, actions } = dialog("docxedit-fonts-catalog", t("fontsGetTitle"));
    const note = document.createElement("p");
    note.className = "docxedit-fonts-note";
    note.textContent = t("fontsCatalogNote");
    const search = document.createElement("input");
    search.type = "search";
    search.className = "docxedit-dialog-textarea docxedit-fonts-search";
    search.placeholder = t("fontsSearch");
    search.setAttribute("aria-label", t("fontsSearch"));
    const list = document.createElement("div");
    list.className = "docxedit-fonts-list";
    const status = document.createElement("div");
    status.className = "docxedit-fonts-status";
    status.setAttribute("role", "status");
    status.setAttribute("aria-live", "polite");
    const rows: { el: HTMLElement; key: string }[] = [];
    let heading: HTMLElement | null = null;
    let lastGroup = "";
    for (const wf of WEB_FONTS) {
      const label = t(`fontGroup_${wf.group}`);
      if (wf.group !== lastGroup) {
        heading = document.createElement("div");
        heading.className = "docxedit-fonts-group";
        heading.textContent = label;
        list.appendChild(heading);
        lastGroup = wf.group;
      }
      const b = document.createElement("button");
      b.type = "button";
      b.className = "docxedit-fonts-item";
      b.style.fontFamily = `${cssFamily(previewFamily(wf.family))}, ${cssFamily(wf.family)}, system-ui, sans-serif`;
      b.textContent = wf.family;
      const sample = GROUP_SAMPLE[wf.group];
      if (sample) {
        const s = document.createElement("span");
        s.className = "docxedit-fonts-sample";
        s.textContent = sample;
        s.setAttribute("aria-hidden", "true");
        b.appendChild(s);
      }
      b.setAttribute("aria-label", `${wf.family}, ${label}`);
      b.addEventListener("click", () => void pickWebFont(wf, status, overlay));
      list.appendChild(b);
      rows.push({ el: b, key: `${wf.family} ${label}`.toLowerCase() });
      (b as HTMLElement & { _head?: HTMLElement })._head = heading!;
    }
    search.addEventListener("input", () => {
      const q = search.value.trim().toLowerCase();
      const shownHeads = new Set<HTMLElement>();
      for (const r of rows) {
        const show = !q || r.key.includes(q);
        r.el.hidden = !show;
        if (show) shownHeads.add((r.el as HTMLElement & { _head?: HTMLElement })._head!);
      }
      for (const h of Array.from(list.querySelectorAll<HTMLElement>(".docxedit-fonts-group"))) h.hidden = !shownHeads.has(h);
    });
    search.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        rows.find((r) => !r.el.hidden)?.el.click();
      }
    });
    const close = button(t("findClose"));
    close.addEventListener("click", () => (overlay.hidden = true));
    actions.append(close);
    panel.insertBefore(note, actions);
    panel.insertBefore(search, actions);
    panel.insertBefore(list, actions);
    panel.insertBefore(status, actions);
    makeDialogAccessible(overlay, search);
    catalogOverlay = overlay;
    catalogSearch = search;
  };

  const pickWebFont = async (wf: WebFont, status: HTMLElement, overlay: HTMLElement): Promise<void> => {
    status.textContent = t("fontsLoading", { name: wf.family });
    const ok = await loadWebFont(wf.family);
    if (disposed) return;
    if (!ok) {
      status.textContent = t("fontsLoadFailed", { name: wf.family });
      return;
    }
    status.textContent = "";
    sessionAdded.add(wf.family.toLowerCase());
    const prefs = readFontPrefs();
    writeFontPrefs({ added: mergeNames(prefs.added ?? [], [wf.family]) });
    added = mergeNames(added, [wf.family]);
    render();
    overlay.hidden = true;
    restoreSel();
    select.value = wf.family;
    prev = wf.family;
    apply(wf.family);
  };

  // Start-up: base + document fonts now, then what the user opted into before.
  render();
  const prefs = readFontPrefs();
  if (localOn && (prefs.local || sessionLocal)) void showMine(false);
  if (remote && prefs.consent) {
    // Families the user added earlier, plus document fonts from the catalog that are missing.
    const want = mergeNames(prefs.added ?? [], docFonts.filter((f) => findWebFont(f) && !canDraw(f)));
    for (const f of want) {
      void loadWebFont(f).then((ok) => {
        if (!ok || disposed) return;
        sessionAdded.add(f.toLowerCase());
        added = mergeNames(added, [findWebFont(f)!.family]);
        render();
      });
    }
  }

  return {
    select,
    /** Every font in the picker (document fonts first). */
    fonts: (): string[] => all,
    has: (family: string): boolean => all.includes(family),
    /** Show a value (caret sync); it is also what an action entry restores. */
    set: (value: string): void => {
      select.value = value;
      prev = select.value;
    },
    teardown: (): void => {
      disposed = true;
      consentOverlay?.remove();
      catalogOverlay?.remove();
    },
  };
}
