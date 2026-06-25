// Toolbar: the formatting controls and their wiring. Builds the button/select factories,
// the bold/italic/underline + list + alignment commands (execCommand), the paragraph-style,
// font, size and colour pickers, link / page-break / image / comment actions, hosts the
// track-changes setup (which needs the button factories), and lays the single row out with
// an overflow "more" popover. The feature behaviours it drives come in as deps.
import { t } from "../i18n";
import { firstFontFamily, fontSizeToHalfPt } from "../util";
import { setupTrackChanges } from "./track-changes";
import type { Adapter, Capabilities, CommentThread, EditorOptions, NewStyle, RichDoc } from "../types";

export interface ToolbarDeps {
  toolbar: HTMLElement;
  wrap: HTMLElement;
  doc: HTMLElement;
  regions: HTMLElement[];
  caps: Capabilities;
  options: EditorOptions;
  parts: RichDoc;
  adapter: Adapter;
  getActiveEl: () => HTMLElement;
  mark: () => void;
  positionCards: () => void;
  addThreadCard: (thread: CommentThread) => HTMLElement;
  setActiveComment: (id: string | null) => void;
  allocId: () => string;
  freshParaId: () => string;
  insertImage: () => void;
  styleBar: HTMLElement; // the bottom-bar left slot for the style pickers
  newStyles: NewStyle[]; // styles authored in-session, collected for save
  newStyleCss: HTMLStyleElement; // live <style> for the appearance of in-session styles
}

export function setupToolbar(deps: ToolbarDeps) {
  const {
    toolbar, wrap, doc, regions, caps, options, parts, adapter, getActiveEl, mark, positionCards,
    addThreadCard, setActiveComment, allocId, freshParaId, insertImage, styleBar,
    newStyles, newStyleCss,
  } = deps;

  // Clicking a toolbar <select> can drop the editor's selection (especially a non-collapsed
  // one). Capture it on mousedown, restore it before a style action reads it.
  let savedSel: Range | null = null;
  const captureSel = (): void => {
    const s = window.getSelection();
    if (!s || !s.rangeCount) return;
    const r = s.getRangeAt(0);
    const n = r.startContainer;
    const el = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
    if (el && regions.some((reg) => reg.contains(el))) savedSel = r.cloneRange();
  };
  const restoreSel = (): void => {
    if (!savedSel) return;
    const s = window.getSelection();
    if (s) {
      s.removeAllRanges();
      s.addRange(savedSel);
    }
  };

  const exec = (cmd: string, val?: string) => {
    getActiveEl().focus();
    document.execCommand(cmd, false, val);
    mark();
    syncToolbarState();
  };
  // Wrap the current selection in a span carrying one CSS property (for font size, which
  // has no execCommand equivalent in CSS mode). No-op on a collapsed selection.
  const styleSel = (prop: string, val: string) => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (range.collapsed) return;
    const span = document.createElement("span");
    (span.style as unknown as Record<string, string>)[prop] = val;
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
    } catch {
      return;
    }
    sel.removeAllRanges();
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.addRange(r2);
    mark();
    syncToolbarState();
  };
  const btn = (label: string, title: string, fn: () => void, cls = "") => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    if (cls) b.className = cls;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", fn);
    return b;
  };
  const sep = () => {
    const s = document.createElement("span");
    s.className = "sep";
    return s;
  };
  const alignIcon = (rows: [number, number][]): string =>
    `<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">${rows
      .map(([x, w], k) => `<rect x="${x}" y="${3 + k * 4}" width="${w}" height="1.6" rx=".6"/>`)
      .join("")}</svg>`;
  const iconBtn = (svg: string, title: string, fn: () => void) => {
    const b = btn("", title, fn);
    b.innerHTML = svg;
    return b;
  };
  const bulletIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<circle cx="2.3" cy="4" r="1.3"/><circle cx="2.3" cy="11" r="1.3"/>' +
    '<rect x="6" y="3.2" width="9" height="1.6" rx=".6"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6"/></svg>';
  const numberIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">' +
    '<text x="0.3" y="6.2" font-size="6" font-family="sans-serif" fill="currentColor">1</text>' +
    '<text x="0.3" y="13.4" font-size="6" font-family="sans-serif" fill="currentColor">2</text>' +
    '<rect x="6" y="3.2" width="9" height="1.6" rx=".6" fill="currentColor"/><rect x="6" y="10.2" width="9" height="1.6" rx=".6" fill="currentColor"/></svg>';
  const linkIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true">' +
    '<path d="M6.6 9.4l2.8-2.8"/><path d="M7.2 4.6l1-1a2.4 2.4 0 0 1 3.4 3.4l-1 1"/><path d="M8.8 11.4l-1 1a2.4 2.4 0 0 1-3.4-3.4l1-1"/></svg>';

  const block = document.createElement("select");
  block.title = t("paragraphStyle");
  block.setAttribute("aria-label", t("paragraphStyle"));
  const BUILTIN_BLOCKS = new Set(["P", "H1", "H2", "H3"]);
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(t(key), v));
  }
  // The document's own named paragraph styles (Title, Quote, ...), value = style id with a
  // "s:" prefix so it never collides with the built-in block tags. The list is mutable so a
  // style authored in-session can be appended.
  const namedStyles = [...(parts.paragraphStyles ?? [])];
  const paraGroup = document.createElement("optgroup");
  paraGroup.label = t("documentStyles");
  for (const s of namedStyles) paraGroup.appendChild(new Option(s.name, `s:${s.id}`));
  if (namedStyles.length) block.appendChild(paraGroup);
  block.add(new Option(t("newParagraphStyle"), "__newpara__"));
  const paraEditOpt = new Option(t("editCurrentStyle"), "__editpara__");
  paraEditOpt.hidden = true; // shown only when the caret is in a styled paragraph
  block.add(paraEditOpt);
  block.addEventListener("mousedown", () => {
    captureSel();
    getActiveEl().focus();
  });
  block.addEventListener("change", () => {
    const v = block.value;
    if (v === "__newpara__") {
      createParagraphStyle();
    } else if (v === "__editpara__") {
      editCurrentParagraphStyle();
    } else if (BUILTIN_BLOCKS.has(v)) {
      // a built-in block clears any named style on the affected paragraphs
      exec("formatBlock", v);
      for (const b of selectedBlocks()) b.removeAttribute("data-rdoc-style");
      mark();
    } else if (v.startsWith("s:")) {
      // a named paragraph style: drop heading state, then tag the blocks with the style id
      getActiveEl().focus();
      document.execCommand("formatBlock", false, "P");
      for (const b of selectedBlocks()) b.setAttribute("data-rdoc-style", v.slice(2));
      mark();
      syncToolbarState();
    }
  });

  // Character styles: wrap the selection in a span carrying the style id (or strip it). The
  // appearance comes from the injected style CSS; the id round-trips to w:rStyle / text:style-name.
  const charStyles = [...(parts.characterStyles ?? [])];
  const applyCStyle = (id: string | null): void => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (id === null) {
      // clear: strip the attribute from styled spans intersecting the selection
      for (const s of Array.from(getActiveEl().querySelectorAll<HTMLElement>("[data-rdoc-cstyle]"))) {
        if (range.intersectsNode(s)) s.removeAttribute("data-rdoc-cstyle");
      }
      mark();
      syncToolbarState();
      return;
    }
    if (range.collapsed) return; // a character style needs a selection to wrap
    const span = document.createElement("span");
    span.setAttribute("data-rdoc-cstyle", id);
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
    } catch {
      return;
    }
    sel.removeAllRanges();
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.addRange(r2);
    mark();
    syncToolbarState();
  };
  const cstyleSel = document.createElement("select");
  cstyleSel.title = t("characterStyle");
  cstyleSel.setAttribute("aria-label", t("characterStyle"));
  const cstyleHead = new Option(t("characterStyle"), "");
  cstyleHead.disabled = true;
  cstyleSel.add(cstyleHead);
  cstyleSel.add(new Option(t("noCharStyle"), "none"));
  for (const s of charStyles) cstyleSel.add(new Option(s.name, `c:${s.id}`));
  const cstyleNewOpt = new Option(t("newCharacterStyle"), "__newchar__");
  cstyleSel.add(cstyleNewOpt);
  const cstyleEditOpt = new Option(t("editCurrentStyle"), "__editchar__");
  cstyleEditOpt.hidden = true; // shown only when the caret is in a styled run
  cstyleSel.add(cstyleEditOpt);
  cstyleSel.addEventListener("mousedown", () => {
    captureSel();
    getActiveEl().focus();
  });
  cstyleSel.addEventListener("change", () => {
    const v = cstyleSel.value;
    if (v === "__newchar__") createCharacterStyle();
    else if (v === "__editchar__") editCurrentCharacterStyle();
    else if (v === "none") applyCStyle(null);
    else if (v.startsWith("c:")) applyCStyle(v.slice(2));
  });

  // --- Authoring and editing styles -------------------------------------------------------
  const styleName = (id: string): string => [...namedStyles, ...charStyles].find((s) => s.id === id)?.name ?? id;
  // Each known style's current definition (CSS), so the edit dialog can be prefilled.
  const styleDefMap = new Map<string, { kind: "paragraph" | "character"; css: Record<string, string> }>();
  for (const d of parts.styleDefs ?? []) styleDefMap.set(d.id, { kind: d.kind, css: { ...d.css } });
  const existingIds = new Set<string>([...namedStyles, ...charStyles].map((s) => s.id));
  const makeStyleId = (name: string): string => {
    const base = name.replace(/[^A-Za-z0-9]/g, "") || "Style";
    let id = base;
    let n = 1;
    while (existingIds.has(id)) id = `${base}${++n}`;
    existingIds.add(id);
    return id;
  };
  const cssBody = (css: Record<string, string>): string => Object.entries(css).map(([k, v]) => `${k}:${v}`).join(";");
  // In-session CSS for created/edited styles, rebuilt whole so an edit replaces (not stacks).
  const overrideCss = new Map<string, string>();
  const rebuildNewStyleCss = (): void => {
    newStyleCss.textContent = [...overrideCss.values()].join("\n");
  };
  // Create or update a style: refresh its definition, its live CSS, and its save record.
  const upsertStyle = (kind: "paragraph" | "character", id: string, name: string, css: Record<string, string>): void => {
    styleDefMap.set(id, { kind, css });
    const attr = kind === "paragraph" ? "data-rdoc-style" : "data-rdoc-cstyle";
    const body = cssBody(css);
    if (body) overrideCss.set(id, `.docxedit-doc [${attr}="${id}"]{${body}}`);
    else overrideCss.delete(id);
    rebuildNewStyleCss();
    const i = newStyles.findIndex((s) => s.id === id);
    if (i >= 0) newStyles[i] = { id, name, kind, css };
    else newStyles.push({ id, name, kind, css });
  };
  // Register a brand-new style and add it to the matching dropdown.
  const registerStyle = (kind: "paragraph" | "character", name: string, css: Record<string, string>): string => {
    const id = makeStyleId(name);
    upsertStyle(kind, id, name, css);
    if (kind === "paragraph") {
      namedStyles.push({ id, name });
      if (!paraGroup.parentNode) block.insertBefore(paraGroup, block.querySelector('option[value="__newpara__"]'));
      paraGroup.appendChild(new Option(name, `s:${id}`));
    } else {
      charStyles.push({ id, name });
      cstyleSel.add(new Option(name, `c:${id}`), cstyleNewOpt);
    }
    return id;
  };
  // Capture a block's direct paragraph formatting as style CSS.
  const captureParaCss = (b: HTMLElement): Record<string, string> => {
    const css: Record<string, string> = {};
    if (b.style.textAlign && b.style.textAlign !== "left") css["text-align"] = b.style.textAlign;
    if (b.style.marginLeft) css["margin-left"] = b.style.marginLeft;
    if (b.style.lineHeight) css["line-height"] = b.style.lineHeight;
    if (b.style.marginTop) css["margin-top"] = b.style.marginTop;
    if (b.style.marginBottom) css["margin-bottom"] = b.style.marginBottom;
    if (b.style.backgroundColor) css["background-color"] = rgbToHex(b.style.backgroundColor) ?? "";
    return css;
  };
  // Capture the selection's run formatting (toggles via execCommand state, the rest computed).
  const rgbToHex = (rgb: string): string | null => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(rgb);
    if (!m) return null;
    return "#" + [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("");
  };
  const captureCharCss = (): Record<string, string> => {
    const css: Record<string, string> = {};
    if (queryState("bold")) css["font-weight"] = "bold";
    if (queryState("italic")) css["font-style"] = "italic";
    const deco: string[] = [];
    if (queryState("underline")) deco.push("underline");
    if (queryState("strikeThrough")) deco.push("line-through");
    if (deco.length) css["text-decoration"] = deco.join(" ");
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as HTMLElement)) : null;
    if (el) {
      const cs = getComputedStyle(el);
      const hex = rgbToHex(cs.color);
      if (hex && hex !== "#000000") css["color"] = hex;
      const px = parseFloat(cs.fontSize);
      if (px) css["font-size"] = `${Math.round((px * 72) / 96)}pt`;
      const fam = firstFontFamily(cs.fontFamily);
      if (fam) css["font-family"] = `'${fam.replace(/'/g, "")}'`;
    }
    return css;
  };
  // --- New-style dialog: a name plus formatting controls -----------------------------------
  let dlgKind: "paragraph" | "character" = "paragraph";
  let dlgBlocks: HTMLElement[] = []; // target paragraphs (paragraph kind)
  let dlgRange: Range | null = null; // target selection (character kind)
  const overlay = document.createElement("div");
  overlay.className = "docxedit-dialog-overlay";
  overlay.hidden = true;
  const panel = document.createElement("div");
  panel.className = "docxedit-dialog";
  overlay.appendChild(panel);
  const dlgTitle = document.createElement("div");
  dlgTitle.className = "docxedit-dialog-title";
  const dlgName = document.createElement("input");
  dlgName.type = "text";
  dlgName.className = "docxedit-name-input";
  dlgName.placeholder = t("styleName");
  // Same look as the main toolbar: styled letters for B/I/U/S, the shared SVG for alignment.
  const toggleBtn = (label: string, title: string, styleClass: string): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `docxedit-dialog-toggle ${styleClass}`;
    b.textContent = label;
    b.title = title;
    b.addEventListener("click", () => b.classList.toggle("is-on"));
    return b;
  };
  const dB = toggleBtn("B", t("bold"), "docxedit-dlg-bold");
  const dI = toggleBtn("I", t("italic"), "docxedit-dlg-italic");
  const dU = toggleBtn("U", t("underline"), "docxedit-dlg-underline");
  const dS = toggleBtn("S", t("strikethrough"), "docxedit-dlg-strike");
  const fmtRow = document.createElement("div");
  fmtRow.className = "docxedit-dialog-row";
  fmtRow.append(dB, dI, dU, dS);
  const alignBtns: Record<string, HTMLButtonElement> = {};
  const alignRow = document.createElement("div");
  alignRow.className = "docxedit-dialog-row";
  const alignSvg: Record<string, [number, number][]> = {
    left: [[2, 12], [2, 8], [2, 11]], center: [[2, 12], [4, 8], [3, 10]], right: [[2, 12], [6, 8], [3, 11]], justify: [[2, 12], [2, 12], [2, 12]],
  };
  for (const [val, key] of [["left", "alignLeft"], ["center", "alignCenter"], ["right", "alignRight"], ["justify", "alignJustify"]] as const) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-dialog-toggle";
    b.innerHTML = alignIcon(alignSvg[val]!);
    b.title = t(key);
    b.addEventListener("click", () => {
      const on = b.classList.contains("is-on");
      for (const k of Object.keys(alignBtns)) alignBtns[k]!.classList.remove("is-on");
      if (!on) b.classList.add("is-on");
    });
    alignBtns[val] = b;
    alignRow.appendChild(b);
  }
  const sizeRow = document.createElement("div");
  sizeRow.className = "docxedit-dialog-row";
  const dSize = document.createElement("input");
  dSize.type = "number";
  dSize.min = "1";
  dSize.className = "docxedit-dialog-size";
  dSize.placeholder = t("size");
  const dColorChk = document.createElement("input");
  dColorChk.type = "checkbox";
  const dColor = document.createElement("input");
  dColor.type = "color";
  dColor.value = "#000000";
  dColor.title = t("textColor");
  dColor.addEventListener("input", () => (dColorChk.checked = true));
  const colorLabel = document.createElement("label");
  colorLabel.className = "docxedit-dialog-color";
  colorLabel.append(dColorChk, document.createTextNode(t("textColor")), dColor);
  sizeRow.append(dSize, colorLabel);
  // Background colour (checkbox + picker) and font family (populated on open from FONTS).
  const bgRow = document.createElement("div");
  bgRow.className = "docxedit-dialog-row";
  const dBgChk = document.createElement("input");
  dBgChk.type = "checkbox";
  const dBg = document.createElement("input");
  dBg.type = "color";
  dBg.value = "#ffff00";
  dBg.title = t("highlight");
  dBg.addEventListener("input", () => (dBgChk.checked = true));
  const bgLabel = document.createElement("label");
  bgLabel.className = "docxedit-dialog-color";
  bgLabel.append(dBgChk, document.createTextNode(t("background")), dBg);
  const dFont = document.createElement("select");
  dFont.className = "docxedit-dialog-font";
  dFont.title = t("font");
  let dFontFilled = false; // populated lazily from FONTS (defined later)
  bgRow.append(dFont, bgLabel);
  const dCreate = document.createElement("button");
  dCreate.type = "button";
  dCreate.className = "docxedit-menu-item docxedit-dialog-primary";
  dCreate.textContent = t("createStyle");
  const dCancel = document.createElement("button");
  dCancel.type = "button";
  dCancel.className = "docxedit-menu-item";
  dCancel.textContent = t("cancel");
  const btnRow = document.createElement("div");
  btnRow.className = "docxedit-dialog-row docxedit-dialog-actions";
  btnRow.append(dCancel, dCreate);
  panel.append(dlgTitle, dlgName, fmtRow, alignRow, sizeRow, bgRow, btnRow);
  wrap.appendChild(overlay);

  let dlgEditId: string | null = null; // set when editing an existing style
  let dlgPassthrough: Record<string, string> = {}; // captured props with no dialog control
  const closeStyleDialog = (): void => {
    overlay.hidden = true;
    dlgBlocks = [];
    dlgRange = null;
    dlgEditId = null;
    dlgPassthrough = {};
    dlgName.disabled = false;
    syncToolbarState(); // reset the dropdowns off the "New..." entry
  };
  // Build the style CSS from the dialog controls, carrying through props with no control
  // (indent / spacing captured at open) so they are not lost.
  const dialogCss = (): Record<string, string> => {
    const css: Record<string, string> = { ...dlgPassthrough };
    css["font-weight"] = dB.classList.contains("is-on") ? "bold" : "normal";
    css["font-style"] = dI.classList.contains("is-on") ? "italic" : "normal";
    const deco: string[] = [];
    if (dU.classList.contains("is-on")) deco.push("underline");
    if (dS.classList.contains("is-on")) deco.push("line-through");
    if (deco.length) css["text-decoration"] = deco.join(" ");
    const sz = parseInt(dSize.value, 10);
    if (sz > 0) css["font-size"] = `${sz}pt`;
    else delete css["font-size"];
    if (dColorChk.checked) css["color"] = dColor.value;
    if (dBgChk.checked) css["background-color"] = dBg.value;
    if (dFont.value) css["font-family"] = `'${dFont.value}'`;
    if (dlgKind === "paragraph") {
      delete css["text-align"];
      for (const k of Object.keys(alignBtns)) if (alignBtns[k]!.classList.contains("is-on")) css["text-align"] = k;
    }
    return css;
  };
  // Fill the dialog controls from a CSS map; stash the un-controlled props for passthrough.
  const fillDialog = (pre: Record<string, string>): void => {
    if (!dFontFilled) {
      dFont.add(new Option(t("defaultFont"), ""));
      for (const f of FONTS) dFont.add(new Option(f, f));
      dFontFilled = true;
    }
    dB.classList.toggle("is-on", pre["font-weight"] === "bold");
    dI.classList.toggle("is-on", pre["font-style"] === "italic");
    dU.classList.toggle("is-on", /underline/.test(pre["text-decoration"] ?? ""));
    dS.classList.toggle("is-on", /line-through/.test(pre["text-decoration"] ?? ""));
    dSize.value = pre["font-size"] ? String(parseInt(pre["font-size"], 10)) : "";
    dColorChk.checked = !!pre["color"];
    dColor.value = pre["color"] ?? "#000000";
    dBgChk.checked = !!pre["background-color"];
    dBg.value = pre["background-color"] ?? "#ffff00";
    const fam = (pre["font-family"] ?? "").replace(/['"]/g, "").split(",")[0]?.trim() ?? "";
    dFont.value = Array.from(dFont.options).some((o) => o.value === fam) ? fam : "";
    for (const k of Object.keys(alignBtns)) alignBtns[k]!.classList.toggle("is-on", pre["text-align"] === k);
    dlgPassthrough = {};
    for (const p of ["margin-left", "margin-top", "margin-bottom", "line-height"]) if (pre[p]) dlgPassthrough[p] = pre[p]!;
  };
  const openStyleDialog = (kind: "paragraph" | "character"): void => {
    restoreSel(); // the <select> click may have dropped the editor selection
    dlgKind = kind;
    dlgEditId = null;
    if (kind === "paragraph") {
      dlgBlocks = selectedBlocks();
      if (!dlgBlocks.length) {
        syncToolbarState();
        return; // no paragraph to apply the style to
      }
      dlgRange = null;
    } else {
      // A character style applies to a selection; with none, still open to define one for later.
      const sel = window.getSelection();
      dlgRange = sel && sel.rangeCount && !sel.getRangeAt(0).collapsed ? sel.getRangeAt(0).cloneRange() : null;
      dlgBlocks = [];
    }
    const pre = kind === "paragraph" ? { ...captureCharCss(), ...captureParaCss(dlgBlocks[0]!) } : captureCharCss();
    dlgTitle.textContent = (kind === "paragraph" ? t("newParagraphStyle") : t("newCharacterStyle")).replace(/…$/, "");
    dlgName.value = "";
    dlgName.disabled = false;
    fillDialog(pre);
    alignRow.hidden = kind !== "paragraph";
    overlay.hidden = false;
    dlgName.focus();
  };
  // Edit an existing style's definition; on save it updates every block/run using it.
  const editStyle = (kind: "paragraph" | "character", id: string): void => {
    const def = styleDefMap.get(id);
    if (!def) {
      syncToolbarState();
      return;
    }
    dlgKind = kind;
    dlgEditId = id;
    dlgBlocks = [];
    dlgRange = null;
    dlgTitle.textContent = t("editStyle") + ": " + styleName(id);
    dlgName.value = styleName(id);
    dlgName.disabled = true; // renaming a style id is not supported
    fillDialog(def.css);
    alignRow.hidden = kind !== "paragraph";
    overlay.hidden = false;
    dCreate.focus();
  };
  const submitStyleDialog = (): void => {
    const name = dlgName.value.trim();
    if (!name) {
      dlgName.focus();
      return;
    }
    const css = dialogCss();
    if (dlgEditId) {
      upsertStyle(dlgKind, dlgEditId, name, css); // update the definition in place
      mark();
    } else if (dlgKind === "paragraph") {
      const id = registerStyle("paragraph", name, css);
      for (const b of dlgBlocks) {
        b.setAttribute("data-rdoc-style", id);
        // captured direct formatting now lives in the style; clear it so it is not written twice
        for (const p of ["textAlign", "marginLeft", "lineHeight", "marginTop", "marginBottom"] as const) b.style[p] = "";
      }
      mark();
    } else {
      const id = registerStyle("character", name, css);
      if (dlgRange) {
        const s = window.getSelection();
        if (s) {
          s.removeAllRanges();
          s.addRange(dlgRange);
        }
        applyCStyle(id); // wrap the selection in the new style
      } else {
        mark(); // defined but not applied (no selection); available in the dropdown to apply later
      }
    }
    closeStyleDialog();
  };
  dCreate.addEventListener("click", submitStyleDialog);
  dCancel.addEventListener("click", closeStyleDialog);
  overlay.addEventListener("mousedown", (e) => {
    if (e.target === overlay) closeStyleDialog();
  });
  dlgName.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submitStyleDialog();
    } else if (e.key === "Escape") {
      closeStyleDialog();
    }
  });
  const createParagraphStyle = (): void => openStyleDialog("paragraph");
  const createCharacterStyle = (): void => openStyleDialog("character");
  // Edit the style applied at the caret (paragraph) / around the selection (character).
  const styledAncestor = (attr: string): HTMLElement | null => {
    const node = window.getSelection()?.anchorNode as Node | null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as HTMLElement)) : null;
    return (el?.closest?.(`[${attr}]`) as HTMLElement | null) ?? null;
  };
  const editCurrentParagraphStyle = (): void => {
    restoreSel();
    const id = styledAncestor("data-rdoc-style")?.getAttribute("data-rdoc-style");
    if (id && styleDefMap.get(id)?.kind === "paragraph") editStyle("paragraph", id);
    else syncToolbarState();
  };
  const editCurrentCharacterStyle = (): void => {
    restoreSel();
    const id = styledAncestor("data-rdoc-cstyle")?.getAttribute("data-rdoc-cstyle");
    if (id && styleDefMap.get(id)?.kind === "character") editStyle("character", id);
    else syncToolbarState();
  };

  // A select whose first option is a non-selectable title; firing fn(value) on change.
  const pickerSelect = (title: string, opts: [string, string][], fn: (v: string) => void): HTMLSelectElement => {
    const s = document.createElement("select");
    s.title = title;
    s.setAttribute("aria-label", title);
    const head = new Option(title, "");
    head.disabled = true;
    head.selected = true;
    s.add(head);
    for (const [v, label] of opts) s.add(new Option(label, v));
    s.addEventListener("mousedown", () => getActiveEl().focus());
    s.addEventListener("change", () => {
      if (s.value) fn(s.value); // keep the chosen value shown (do not reset to the placeholder)
    });
    return s;
  };

  // Text colour: a native colour input that applies w:color via foreColor (CSS mode).
  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.value = "#000000";
  colorInput.title = t("textColor");
  colorInput.setAttribute("aria-label", t("textColor"));
  colorInput.className = "docxedit-color";
  colorInput.addEventListener("mousedown", () => getActiveEl().focus());
  colorInput.addEventListener("input", () => {
    beginFormatChange();
    exec("foreColor", colorInput.value);
  });

  // Background colour: a free colour picker (maps to w:highlight when it matches a named
  // highlight exactly, otherwise to arbitrary w:shd shading). A button clears it.
  const bgWrap = document.createElement("span");
  bgWrap.className = "docxedit-bg";
  const bgInput = document.createElement("input");
  bgInput.type = "color";
  bgInput.value = "#ffff00";
  bgInput.title = t("highlight");
  bgInput.setAttribute("aria-label", t("highlight"));
  bgInput.className = "docxedit-color";
  bgInput.addEventListener("mousedown", () => getActiveEl().focus());
  bgInput.addEventListener("input", () => {
    beginFormatChange();
    exec("hiliteColor", bgInput.value);
  });
  const bgClear = btn("⌫", t("none"), () => exec("hiliteColor", "transparent"), "docxedit-bg-clear");
  bgWrap.append(bgInput, bgClear);

  // Fonts/sizes the document actually uses but that are not in the defaults are added to
  // the pickers, so the caret-sync can show them instead of falling back to the placeholder.
  const BASE_FONTS = ["Arial", "Calibri", "Century", "Courier New", "Georgia", "Times New Roman", "Verdana"];
  const BASE_SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48"];
  const docFonts = new Set<string>();
  const docSizes = new Set<string>();
  for (const el of Array.from(doc.querySelectorAll<HTMLElement>("[style]"))) {
    const fam = firstFontFamily(el.style.fontFamily);
    if (fam) docFonts.add(fam);
    const hp = fontSizeToHalfPt(el.style.fontSize);
    if (hp) docSizes.add(String(Math.round(hp / 2)));
  }
  if (parts.defaultFont) docFonts.add(parts.defaultFont);
  const FONTS = [...BASE_FONTS, ...[...docFonts].filter((f) => !BASE_FONTS.includes(f)).sort()];
  const SIZES = [...new Set([...BASE_SIZES, ...docSizes])].sort((a, b) => Number(a) - Number(b));
  const fontSel = pickerSelect(t("font"), FONTS.map((f) => [f, f] as [string, string]), (v) => {
    beginFormatChange();
    exec("fontName", v);
  });

  const sizeSel = pickerSelect(t("size"), SIZES.map((s) => [s, s] as [string, string]), (v) => {
    beginFormatChange();
    styleSel("fontSize", `${v}pt`);
  });

  // A new (empty) document gets a default font + size, shown in the pickers and applied so
  // typing starts in them. Existing documents keep their own fonts; the pickers stay blank.
  if (caps.fontControls && !doc.textContent?.trim()) {
    const defFont = parts.defaultFont && FONTS.includes(parts.defaultFont) ? parts.defaultFont : "Arial";
    const defSize = "11";
    fontSel.value = defFont;
    sizeSel.value = defSize;
    doc.style.fontFamily = `'${defFont}', sans-serif`;
    doc.style.fontSize = `${defSize}pt`;
  }

  const insertPageBreak = () => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const el = document.createElement("span");
    el.className = "docx-pagebreak";
    el.contentEditable = "false";
    el.setAttribute("data-docx-pagebreak", "manual");
    el.setAttribute("data-label", t("pageBreak"));
    range.insertNode(el);
    range.setStartAfter(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    mark();
  };
  const pbIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="3" y="1.5" width="10" height="4" rx=".5"/><rect x="3" y="10.5" width="10" height="4" rx=".5"/>' +
    '<line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="2 1.6"/></svg>';

  // Insert an image: read a file, show it via a data URL, and let the serializer embed it.
  const imgIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
    '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

  // Add a comment over the current selection: wrap it in comment-range markers and a
  // reference marker that carries the text, so the serializer can build comments.xml.
  const addComment = () => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed) {
      const tip = document.createElement("div");
      tip.className = "docxedit-cmt-pop docxedit-cmt-tip";
      tip.textContent = t("commentSelect");
      wrap.appendChild(tip);
      setTimeout(() => tip.remove(), 1800);
      return;
    }
    const text = prompt(t("commentPrompt"));
    if (!text) return;
    const id = allocId();
    const author = options.author || "Author";
    const date = options.now || new Date().toISOString();
    const paraId = freshParaId();
    const { start, end, ref } = adapter.newCommentMarkers({ id, author, date, text, paraId });
    const range = sel.getRangeAt(0);
    const visual = document.createElement("span");
    visual.className = "docx-comment";
    visual.setAttribute("data-comment-id", id);
    visual.appendChild(range.extractContents());
    range.insertNode(visual);
    const parent = visual.parentNode;
    if (parent) {
      parent.insertBefore(start, visual);
      parent.insertBefore(end, visual.nextSibling);
      parent.insertBefore(ref, end.nextSibling);
    }
    addThreadCard({ id, author, date, text, reactions: [], replies: [], paraId, resolved: false });
    setActiveComment(id);
    positionCards();
    mark();
  };
  const cmtIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<path d="M2 3.5h12v8H6l-3 2.5V11.5H2z"/><line x1="5" y1="6.2" x2="11" y2="6.2"/><line x1="5" y1="8.6" x2="9" y2="8.6"/></svg>';

  // Track changes (suggestion mode) lives in its own feature module; it builds its own
  // toolbar buttons and exposes beginFormatChange for the bold/italic/underline buttons.
  const { beginFormatChange, suggestBtn, acceptAllBtn, rejectAllBtn, updateChangeButtons } = setupTrackChanges({
    doc, wrap, regions, options, mark, positionCards, getActiveEl, iconBtn, btn,
  });

  // Toolbar: shared controls always shown; image/comment/page-break/track-changes are
  // gated by the adapter's capabilities so a format can hide what it cannot serialize.
  const insertLink = () => {
    const url = prompt(t("linkPrompt"), "https://");
    if (url === null) return;
    if (url === "") exec("unlink");
    else exec("createLink", url);
  };
  const linkBtn = iconBtn(linkIcon, t("linkAria"), insertLink);
  const supIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<text x="0" y="14" font-size="11" font-family="serif">x</text><text x="8" y="7" font-size="7" font-family="serif">2</text></svg>';
  const subIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<text x="0" y="11" font-size="11" font-family="serif">x</text><text x="8" y="16" font-size="7" font-family="serif">2</text></svg>';
  // Named so their pressed state can be reflected from the caret (see syncToolbarState).
  const boldBtn = btn("B", t("bold"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold");
  const italicBtn = btn("I", t("italic"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic");
  const underlineBtn = btn("U", t("underline"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline");
  const strikeBtn = btn("S", t("strikethrough"), () => { beginFormatChange(); exec("strikeThrough"); }, "docxedit-tb-strike");
  const supBtn = iconBtn(supIcon, t("superscript"), () => { beginFormatChange(); exec("superscript"); });
  const subBtn = iconBtn(subIcon, t("subscript"), () => { beginFormatChange(); exec("subscript"); });
  const alignLeftBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), t("alignLeft"), () => exec("justifyLeft")) : null;
  const alignCenterBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), t("alignCenter"), () => exec("justifyCenter")) : null;
  const alignRightBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), t("alignRight"), () => exec("justifyRight")) : null;
  const alignJustifyBtn = caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), t("alignJustify"), () => exec("justifyFull")) : null;

  // Paragraph indent + line spacing operate on the block(s) intersecting the selection.
  const BLOCK_SEL = "p,h1,h2,h3,h4,h5,h6,li,blockquote,div";
  const selectedBlocks = (): HTMLElement[] => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return [];
    const region = getActiveEl();
    const blockOf = (n: Node | null): HTMLElement | null => {
      const el = n ? (n.nodeType === 3 ? n.parentElement : (n as Element)) : null;
      const b = el?.closest?.(BLOCK_SEL) as HTMLElement | null;
      return b && region.contains(b) && !b.classList.contains("docxedit-doc") ? b : null;
    };
    const range = sel.getRangeAt(0);
    const start = blockOf(range.startContainer);
    const end = blockOf(range.endContainer);
    const out: HTMLElement[] = start ? [start] : [];
    let n = start;
    while (n && n !== end) {
      n = n.nextElementSibling as HTMLElement | null;
      if (n && n.matches(BLOCK_SEL)) out.push(n);
    }
    if (end && !out.includes(end)) out.push(end);
    return out;
  };
  const indentIcon = (dir: number): string =>
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">' +
    '<rect x="2" y="2.5" width="12" height="1.6" rx=".6"/><rect x="7" y="6" width="7" height="1.6" rx=".6"/>' +
    '<rect x="7" y="9.4" width="7" height="1.6" rx=".6"/><rect x="2" y="12.9" width="12" height="1.6" rx=".6"/>' +
    (dir > 0 ? '<path d="M2 6.2l2.6 1.9L2 10z"/>' : '<path d="M4.6 6.2L2 8.1l2.6 1.9z"/>') + "</svg>";
  const STEP = 48; // 0.5in
  const marginAdjust = (b: HTMLElement, dir: 1 | -1) => {
    const next = Math.max(0, (parseFloat(b.style.marginLeft) || 0) + dir * STEP);
    b.style.marginLeft = next ? `${next}px` : "";
  };
  // Indent a list item by sinking it into a sublist on its previous sibling (real nesting,
  // so it round-trips as w:ilvl / a deeper text:list rather than a margin).
  const indentLi = (li: HTMLElement) => {
    const prev = li.previousElementSibling as HTMLElement | null;
    if (!prev || prev.tagName !== "LI") return marginAdjust(li, 1); // first item can't nest
    const parentTag = (li.parentElement?.tagName ?? "UL").toLowerCase();
    let sub = prev.lastElementChild as HTMLElement | null;
    if (!sub || (sub.tagName !== "UL" && sub.tagName !== "OL")) {
      sub = li.ownerDocument.createElement(parentTag);
      prev.appendChild(sub);
    }
    sub.appendChild(li);
  };
  const outdentLi = (li: HTMLElement) => {
    const parentList = li.parentElement as HTMLElement | null;
    const grandLi = parentList?.parentElement as HTMLElement | null;
    if (!parentList || !grandLi || grandLi.tagName !== "LI") return marginAdjust(li, -1); // already top level
    const outerList = grandLi.parentElement as HTMLElement;
    // following siblings stay nested under the outdented item, preserving their depth
    let next: Element | null = li.nextElementSibling;
    if (next) {
      const newSub = li.ownerDocument.createElement(parentList.tagName.toLowerCase());
      while (next) {
        const after: Element | null = next.nextElementSibling;
        newSub.appendChild(next);
        next = after;
      }
      li.appendChild(newSub);
    }
    outerList.insertBefore(li, grandLi.nextElementSibling);
    if (!parentList.children.length) parentList.remove();
  };
  const adjustIndent = (dir: 1 | -1) => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) {
      if (b.tagName === "LI") dir === 1 ? indentLi(b) : outdentLi(b);
      else marginAdjust(b, dir);
    }
    mark();
  };
  const indentBtn = iconBtn(indentIcon(1), t("indent"), () => adjustIndent(1));
  const outdentBtn = iconBtn(indentIcon(-1), t("outdent"), () => adjustIndent(-1));

  // Line spacing: an icon button (like the others) opening a small menu of presets.
  const lineSpacingIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M6 4h8M6 8h8M6 12h8"/><path d="M2.6 3.6v8.8M1.4 4.8 2.6 3.6 3.8 4.8M1.4 11.2 2.6 12.4 3.8 11.2"/></svg>';
  const lineSpacingMenu = document.createElement("div");
  lineSpacingMenu.className = "docxedit-menu";
  lineSpacingMenu.hidden = true;
  const setLineHeight = (v: string): void => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) b.style.lineHeight = v;
    mark();
    lineSpacingMenu.hidden = true;
  };
  const lsButtons = ([["1", "1.0"], ["1.15", "1.15"], ["1.5", "1.5"], ["2", "2.0"]] as const).map(([v, label]) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.textContent = label;
    b.dataset.v = v;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      setLineHeight(v);
    });
    return b;
  });
  lineSpacingMenu.append(...lsButtons);
  // Paragraph space before/after: a default add, or an explicit 0 to remove inherited spacing.
  const SPACE_ADD = 14; // px, ~10.5pt, a sensible "add space" default
  const divider = document.createElement("div");
  divider.className = "docxedit-menu-sep";
  const setSpace = (side: "marginTop" | "marginBottom", on: boolean): void => {
    getActiveEl().focus();
    for (const b of selectedBlocks()) b.style[side] = on ? `${SPACE_ADD}px` : "0px";
    mark();
    lineSpacingMenu.hidden = true;
  };
  const spaceItem = (side: "marginTop" | "marginBottom"): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.dataset.side = side;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      setSpace(side, b.dataset.on !== "1"); // toggle off when currently on
    });
    return b;
  };
  const spaceBeforeItem = spaceItem("marginTop");
  const spaceAfterItem = spaceItem("marginBottom");
  lineSpacingMenu.append(divider, spaceBeforeItem, spaceAfterItem);
  // Reflect the current block's state on the two toggle items (add vs remove).
  const refreshSpaceItems = (): void => {
    const b = selectedBlocks()[0];
    const on = (prop: "marginTop" | "marginBottom"): boolean => (parseFloat(b?.style[prop] ?? "") || 0) > 0;
    const beforeOn = on("marginTop");
    const afterOn = on("marginBottom");
    spaceBeforeItem.dataset.on = beforeOn ? "1" : "0";
    spaceAfterItem.dataset.on = afterOn ? "1" : "0";
    spaceBeforeItem.textContent = t(beforeOn ? "removeSpaceBefore" : "addSpaceBefore");
    spaceAfterItem.textContent = t(afterOn ? "removeSpaceAfter" : "addSpaceAfter");
  };
  const lineSpacingBtn = iconBtn(lineSpacingIcon, t("lineSpacing"), () => {});
  lineSpacingBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = lineSpacingBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    lineSpacingMenu.style.left = `${r.left - wr.left}px`;
    lineSpacingMenu.style.top = `${r.bottom - wr.top + 2}px`;
    if (lineSpacingMenu.hidden) refreshSpaceItems();
    lineSpacingMenu.hidden = !lineSpacingMenu.hidden;
  });

  // --- Insert table: a button opening a grid picker (drag/hover to size, click to insert) ---
  const tableIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><line x1="1.5" y1="6.2" x2="14.5" y2="6.2"/>' +
    '<line x1="1.5" y1="9.9" x2="14.5" y2="9.9"/><line x1="6" y1="2.5" x2="6" y2="13.5"/><line x1="10.3" y1="2.5" x2="10.3" y2="13.5"/></svg>';
  const insertTable = (rows: number, cols: number) => {
    getActiveEl().focus();
    const table = document.createElement("table");
    table.className = "docx-table";
    table.contentEditable = "false";
    // Fill the available content width with equal columns (a <colgroup> the adapters read).
    const cs = getComputedStyle(doc);
    const avail = Math.max(120, doc.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0));
    const colW = Math.floor(avail / cols);
    const colgroup = document.createElement("colgroup");
    for (let c = 0; c < cols; c++) {
      const col = document.createElement("col");
      col.style.width = `${colW}px`;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
    for (let r = 0; r < rows; r++) {
      const tr = table.insertRow();
      for (let c = 0; c < cols; c++) {
        const td = tr.insertCell();
        const div = document.createElement("div");
        div.className = "docx-cell";
        div.contentEditable = "true";
        div.innerHTML = "<br>";
        td.appendChild(div);
      }
    }
    // Insert as a top-level block after the block holding the caret, then a paragraph after.
    const sel = window.getSelection();
    const trailing = document.createElement("p");
    trailing.innerHTML = "<br>";
    let anchor: HTMLElement | null = null;
    if (sel && sel.rangeCount) {
      let el: Element | null = ((n) => (n.nodeType === 3 ? n.parentElement : (n as Element)))(sel.getRangeAt(0).startContainer);
      while (el && el.parentElement !== doc) el = el.parentElement;
      anchor = el as HTMLElement | null;
    }
    if (anchor && anchor.parentElement === doc) {
      doc.insertBefore(table, anchor.nextSibling);
      doc.insertBefore(trailing, table.nextSibling);
    } else {
      doc.append(table, trailing);
    }
    mark();
    const firstCell = table.querySelector(".docx-cell") as HTMLElement | null;
    if (firstCell) {
      const r = document.createRange();
      r.selectNodeContents(firstCell);
      r.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(r);
      firstCell.focus();
    }
  };
  const GRID_ROWS = 8;
  const GRID_COLS = 10;
  const tablePicker = document.createElement("div");
  tablePicker.className = "docxedit-table-picker";
  tablePicker.hidden = true;
  const tableGrid = document.createElement("div");
  tableGrid.className = "docxedit-table-grid";
  const tableLabel = document.createElement("div");
  tableLabel.className = "docxedit-table-label";
  tableLabel.textContent = t("insertTable");
  const squares: HTMLElement[] = [];
  const highlight = (rr: number, cc: number) => {
    for (let r = 0; r < GRID_ROWS; r++) for (let c = 0; c < GRID_COLS; c++) squares[r * GRID_COLS + c].classList.toggle("on", r <= rr && c <= cc);
    tableLabel.textContent = `${cc + 1} × ${rr + 1}`;
  };
  for (let r = 0; r < GRID_ROWS; r++) {
    for (let c = 0; c < GRID_COLS; c++) {
      const sq = document.createElement("div");
      sq.className = "docxedit-table-sq";
      sq.addEventListener("mouseenter", () => highlight(r, c));
      sq.addEventListener("mousedown", (e) => {
        e.preventDefault();
        insertTable(r + 1, c + 1);
        tablePicker.hidden = true;
      });
      tableGrid.appendChild(sq);
      squares.push(sq);
    }
  }
  tablePicker.append(tableGrid, tableLabel);
  const tableBtn = iconBtn(tableIcon, t("insertTable"), () => {});
  tableBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = tableBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    tablePicker.style.left = `${r.left - wr.left}px`;
    tablePicker.style.top = `${r.bottom - wr.top + 2}px`;
    tablePicker.hidden = !tablePicker.hidden;
    if (!tablePicker.hidden) highlight(-1, -1);
  });

  // --- Insert field: page number / count / table of contents (docx fields) ---------------
  const fieldIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<path d="M5 2.5C3.5 2.5 3.5 5 3.5 8s0 5.5-1.5 5.5"/><path d="M11 2.5c1.5 0 1.5 2.5 1.5 5.5s0 5.5 1.5 5.5"/></svg>';
  const insertAtCaret = (node: Node): void => {
    getActiveEl().focus();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    getActiveEl().dispatchEvent(new Event("input", { bubbles: true })); // sync header/footer clone + reflow
  };
  const fieldSpan = (instr: string, text: string): HTMLElement => {
    const s = document.createElement("span");
    s.className = "docx-field";
    s.contentEditable = "false";
    s.setAttribute("data-field", instr);
    s.textContent = text;
    return s;
  };
  const insertField = (instr: string) => insertAtCaret(fieldSpan(instr, "1"));
  const insertPageXofY = () => {
    const frag = document.createDocumentFragment();
    frag.append(fieldSpan("PAGE", "1"), document.createTextNode(" / "), fieldSpan("NUMPAGES", "1"));
    insertAtCaret(frag);
  };
  const insertTOC = () => {
    const toc = document.createElement("div");
    toc.className = "docx-field-toc";
    toc.contentEditable = "false";
    toc.setAttribute("data-field", "TOC");
    toc.innerHTML = `<div class="docx-field-toc-empty">${t("tocEmpty")}</div>`;
    const trailing = document.createElement("p");
    trailing.innerHTML = "<br>";
    // The TOC belongs in the body: insert after the caret's top-level block, else at the top.
    const sel = window.getSelection();
    let anchor: HTMLElement | null = null;
    if (sel && sel.rangeCount && doc.contains(sel.getRangeAt(0).startContainer)) {
      let el: Element | null = ((n) => (n.nodeType === 3 ? n.parentElement : (n as Element)))(sel.getRangeAt(0).startContainer);
      while (el && el.parentElement !== doc) el = el.parentElement;
      anchor = el as HTMLElement | null;
    }
    if (anchor && anchor.parentElement === doc) {
      doc.insertBefore(toc, anchor.nextSibling);
      doc.insertBefore(trailing, toc.nextSibling);
    } else {
      doc.insertBefore(trailing, doc.firstChild);
      doc.insertBefore(toc, doc.firstChild);
    }
    doc.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const fieldsMenu = document.createElement("div");
  fieldsMenu.className = "docxedit-menu";
  fieldsMenu.hidden = true;
  for (const it of [
    { label: t("fieldPageNumber"), fn: () => insertField("PAGE") },
    { label: t("fieldPageCount"), fn: () => insertField("NUMPAGES") },
    { label: t("fieldPageXofY"), fn: insertPageXofY },
    { label: t("fieldToc"), fn: insertTOC },
  ]) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-menu-item";
    b.textContent = it.label;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      it.fn();
      fieldsMenu.hidden = true;
    });
    fieldsMenu.appendChild(b);
  }
  const fieldsBtn = iconBtn(fieldIcon, t("insertField"), () => {});
  fieldsBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = fieldsBtn.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    fieldsMenu.style.left = `${r.left - wr.left}px`;
    fieldsMenu.style.top = `${r.bottom - wr.top + 2}px`;
    fieldsMenu.hidden = !fieldsMenu.hidden;
  });

  // Reflect the caret's formatting in the controls: the font / size / paragraph-style
  // pickers track the text under the caret, and bold/italic/underline/alignment show their
  // pressed state. Runs (rAF-coalesced) on every selection change and after a command.
  const queryState = (cmd: string): boolean => {
    try { return document.queryCommandState(cmd); } catch { return false; }
  };
  const syncToolbarState = () => {
    const sel = window.getSelection();
    const node = sel && sel.rangeCount ? sel.anchorNode : null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as Element)) : null;
    if (!el || !regions.some((r) => r.contains(el))) return; // only while editing a region
    const cs = getComputedStyle(el);
    if (caps.fontControls) {
      const fam = firstFontFamily(cs.fontFamily);
      fontSel.value = fam && FONTS.includes(fam) ? fam : "";
      const hp = fontSizeToHalfPt(cs.fontSize);
      const pt = hp ? String(Math.round(hp / 2)) : "";
      sizeSel.value = SIZES.includes(pt) ? pt : "";
    }
    const styled = el.closest("[data-rdoc-style]") as HTMLElement | null;
    const styleId = styled && regions.some((r) => r.contains(styled)) ? styled.getAttribute("data-rdoc-style") : null;
    const editablePara = !!(styleId && namedStyles.some((s) => s.id === styleId));
    const tag = el.closest("h1,h2,h3,p")?.tagName ?? "";
    if (editablePara) block.value = `s:${styleId}`;
    else block.value = tag === "H1" || tag === "H2" || tag === "H3" ? tag : "P";
    paraEditOpt.hidden = !editablePara; // only offer "Edit current style" when one is applied
    // reflect the character style enclosing the caret (or none)
    const cstyled = el.closest("[data-rdoc-cstyle]") as HTMLElement | null;
    const cId = cstyled && regions.some((r) => r.contains(cstyled)) ? cstyled.getAttribute("data-rdoc-cstyle") : null;
    const editableChar = !!(cId && charStyles.some((s) => s.id === cId));
    if (editableChar) cstyleSel.value = `c:${cId}`;
    else cstyleSel.selectedIndex = 0;
    cstyleEditOpt.hidden = !editableChar;
    const lh = (el.closest(BLOCK_SEL) as HTMLElement | null)?.style.lineHeight ?? "";
    for (const b of lsButtons) b.classList.toggle("is-on", b.dataset.v === lh);
    const setOn = (b: HTMLElement | null, on: boolean) => b?.classList.toggle("is-on", on);
    setOn(boldBtn, queryState("bold"));
    setOn(italicBtn, queryState("italic"));
    setOn(underlineBtn, queryState("underline"));
    setOn(strikeBtn, queryState("strikeThrough"));
    setOn(supBtn, queryState("superscript"));
    setOn(subBtn, queryState("subscript"));
    setOn(alignLeftBtn, queryState("justifyLeft"));
    setOn(alignCenterBtn, queryState("justifyCenter"));
    setOn(alignRightBtn, queryState("justifyRight"));
    setOn(alignJustifyBtn, queryState("justifyFull"));
  };
  // Coalesce bursts of selectionchange (e.g. during a drag-select) with a short timer.
  // setTimeout (not rAF) so it still fires when the tab is in the background.
  let syncTimer = 0;
  const scheduleSync = () => {
    window.clearTimeout(syncTimer);
    syncTimer = window.setTimeout(syncToolbarState, 16);
  };
  document.addEventListener("selectionchange", scheduleSync);

  // The paragraph + character style pickers live in the bottom bar's left slot.
  styleBar.append(block, cstyleSel);

  // Two clusters collapse into a single dropdown button when the toolbar runs out of room:
  // the character-formatting controls ("style"), and the insert controls.
  const styleSrc: (HTMLElement | null)[] = [boldBtn, italicBtn, underlineBtn, strikeBtn, supBtn, subBtn, caps.textColor ? colorInput : null, caps.textColor ? bgWrap : null, caps.fontControls ? fontSel : null, caps.fontControls ? sizeSel : null];
  const insertSrc: (HTMLElement | null)[] = [caps.images ? iconBtn(imgIcon, t("insertImage"), insertImage) : null, caps.tables ? tableBtn : null, caps.fields ? fieldsBtn : null, caps.comments ? iconBtn(cmtIcon, t("addComment"), addComment) : null, caps.pageBreak ? iconBtn(pbIcon, t("insertPageBreak"), insertPageBreak) : null, linkBtn];
  const styleControls = styleSrc.filter((n): n is HTMLElement => n != null);
  const insertControls = insertSrc.filter((n): n is HTMLElement => n != null);
  const caret = '<svg width="8" height="8" viewBox="0 0 8 8" fill="currentColor" aria-hidden="true"><path d="M1 2.5 4 5.5 7 2.5z"/></svg>';
  const styleGroupSvg = '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 13 7 3h2l3 10h-2l-.66-2.4H6.66L6 13H4zm3.1-4.2h1.8L8 5.4 7.1 8.8z"/></svg>';
  const insertGroupSvg = '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M7 2h2v5h5v2H9v5H7V9H2V7h5z"/></svg>';
  // A collapsible cluster: a slot that holds the controls inline or a group button + a popover.
  const makeGroup = (controls: HTMLElement[], svg: string, title: string) => {
    const slot = document.createElement("span");
    slot.className = "docxedit-tb-slot";
    const btn = iconBtn(svg + caret, title, () => {});
    btn.classList.add("docxedit-tb-group");
    const menu = document.createElement("div");
    menu.className = "docxedit-toolbar docxedit-tb-overflow docxedit-tb-groupmenu";
    menu.hidden = true;
    let collapsed = false;
    btn.addEventListener("mousedown", (e) => e.preventDefault());
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const r = btn.getBoundingClientRect();
      const wr = wrap.getBoundingClientRect();
      menu.style.top = `${r.bottom - wr.top + 2}px`;
      menu.style.left = `${r.left - wr.left}px`;
      menu.style.right = "auto";
      menu.hidden = !menu.hidden;
    });
    slot.replaceChildren(...controls); // start expanded
    const expand = () => {
      if (!collapsed) return;
      slot.replaceChildren(...controls);
      menu.hidden = true;
      collapsed = false;
    };
    const collapse = () => {
      if (collapsed || !controls.length) return;
      menu.replaceChildren(...controls);
      slot.replaceChildren(btn);
      collapsed = true;
    };
    return { slot, btn, menu, expand, collapse, has: controls.length > 0 };
  };
  const styleGroup = makeGroup(styleControls, styleGroupSvg, t("formatting"));
  const insertGroup = makeGroup(insertControls, insertGroupSvg, t("insertMenu"));

  const items: (Node | null)[] = [
    styleGroup.has ? styleGroup.slot : null,
    styleGroup.has ? sep() : null,
    iconBtn(bulletIcon, t("bulleted"), () => exec("insertUnorderedList")),
    iconBtn(numberIcon, t("numbered"), () => exec("insertOrderedList")),
    caps.alignment ? sep() : null,
    alignLeftBtn,
    alignCenterBtn,
    alignRightBtn,
    alignJustifyBtn,
    outdentBtn,
    indentBtn,
    lineSpacingBtn,
    insertGroup.has ? sep() : null,
    insertGroup.has ? insertGroup.slot : null,
    caps.trackChanges ? sep() : null,
    caps.trackChanges ? suggestBtn : null,
    caps.trackChanges ? acceptAllBtn : null,
    caps.trackChanges ? rejectAllBtn : null,
  ];
  // Overflow menu: anything that still does not fit after the groups collapse moves into a "…"
  // popover so nothing is lost on very narrow widths.
  const toolbarItems = items.filter((n): n is HTMLElement => n != null);
  const moreBtn = document.createElement("button");
  moreBtn.type = "button";
  moreBtn.className = "docxedit-tb-more";
  moreBtn.textContent = "⋯";
  moreBtn.title = t("moreTools");
  moreBtn.setAttribute("aria-label", t("moreTools"));
  const overflow = document.createElement("div");
  overflow.className = "docxedit-toolbar docxedit-tb-overflow";
  overflow.hidden = true;
  moreBtn.addEventListener("mousedown", (e) => e.preventDefault());
  moreBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    overflow.style.top = `${toolbar.offsetHeight}px`;
    overflow.hidden = !overflow.hidden;
  });
  const closeOverflow = (e: MouseEvent) => {
    if (!overflow.hidden && !overflow.contains(e.target as Node) && e.target !== moreBtn) overflow.hidden = true;
    if (!styleGroup.menu.hidden && !styleGroup.menu.contains(e.target as Node) && !styleGroup.btn.contains(e.target as Node)) styleGroup.menu.hidden = true;
    if (!insertGroup.menu.hidden && !insertGroup.menu.contains(e.target as Node) && !insertGroup.btn.contains(e.target as Node)) insertGroup.menu.hidden = true;
    if (!tablePicker.hidden && !tablePicker.contains(e.target as Node) && !tableBtn.contains(e.target as Node)) tablePicker.hidden = true;
    if (!lineSpacingMenu.hidden && !lineSpacingMenu.contains(e.target as Node) && !lineSpacingBtn.contains(e.target as Node)) lineSpacingMenu.hidden = true;
    if (!fieldsMenu.hidden && !fieldsMenu.contains(e.target as Node) && !fieldsBtn.contains(e.target as Node)) fieldsMenu.hidden = true;
  };
  document.addEventListener("click", closeOverflow);
  toolbar.append(...toolbarItems, moreBtn);
  wrap.append(overflow, styleGroup.menu, insertGroup.menu, tablePicker, lineSpacingMenu, fieldsMenu);

  const fits = () => toolbar.scrollWidth <= toolbar.clientWidth + 1;
  const layoutToolbar = () => {
    overflow.hidden = true;
    for (const it of toolbarItems) toolbar.insertBefore(it, moreBtn); // pull everything back in
    moreBtn.style.display = "none";
    insertGroup.expand();
    styleGroup.expand();
    if (fits()) return;
    insertGroup.collapse(); // first collapse the insert cluster
    if (fits()) return;
    styleGroup.collapse(); // then the formatting cluster
    if (fits()) return;
    moreBtn.style.display = ""; // final fallback: pocket trailing items
    for (let i = toolbarItems.length - 1; i >= 0; i--) {
      if (fits()) break;
      overflow.insertBefore(toolbarItems[i], overflow.firstChild);
    }
  };
  layoutToolbar();
  requestAnimationFrame(layoutToolbar);
  setTimeout(layoutToolbar, 150);
  const toolbarObserver = new ResizeObserver(() => layoutToolbar());
  toolbarObserver.observe(toolbar);

  scheduleSync(); // reflect the initial caret position once mounted

  // --- Floating formatting bar (desktop): quick formatting near the caret on mouse proximity --
  const coarse = typeof window.matchMedia === "function" && window.matchMedia("(hover: none), (pointer: coarse)").matches;
  const floatBar = document.createElement("div");
  floatBar.className = "docxedit-floatbar";
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
  const fBold = fbtn("B", t("bold"), "bold", "docxedit-tb-bold");
  const fItalic = fbtn("I", t("italic"), "italic", "docxedit-tb-italic");
  const fUnderline = fbtn("U", t("underline"), "underline", "docxedit-tb-underline");
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
    const bw = floatBar.offsetWidth || 180;
    const bh = floatBar.offsetHeight || 32;
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
    }
    wrap.addEventListener("scroll", hideFloat, true);
  }

  // --- Keyboard shortcuts (Word-like) -------------------------------------------------------
  // Bound to the editable regions so they fire only while editing (single-key bindings such as
  // Ctrl+R for right-align therefore leave the browser's own shortcuts alone everywhere else).
  const applyBlockTag = (tag: string) => {
    getActiveEl().focus();
    document.execCommand("formatBlock", false, tag);
    for (const b of selectedBlocks()) b.removeAttribute("data-rdoc-style");
    mark();
    syncToolbarState();
  };
  const changeFontSize = (delta: number) => {
    const node = window.getSelection()?.anchorNode as Node | null;
    const el = node ? (node.nodeType === 3 ? node.parentElement : (node as HTMLElement)) : null;
    let pt = 11;
    if (el) {
      const px = parseFloat(getComputedStyle(el).fontSize);
      if (px) pt = Math.round((px * 72) / 96);
    }
    beginFormatChange();
    styleSel("fontSize", `${Math.max(1, pt + delta)}pt`);
  };
  const onShortcut = (e: KeyboardEvent) => {
    if (!(e.metaKey || e.ctrlKey)) return;
    const k = e.key.toLowerCase();
    const run = (fn: () => void) => {
      e.preventDefault();
      fn();
    };
    if (e.altKey && !e.shiftKey) {
      if (k === "1") return run(() => applyBlockTag("H1"));
      if (k === "2") return run(() => applyBlockTag("H2"));
      if (k === "3") return run(() => applyBlockTag("H3"));
      if (k === "0") return run(() => applyBlockTag("P"));
      return;
    }
    if (e.shiftKey) {
      if (k === "l") return run(() => exec("insertUnorderedList")); // bulleted list
      if (k === "7") return run(() => exec("insertOrderedList")); // numbered list
      return;
    }
    switch (k) {
      case "b": return run(() => { beginFormatChange(); exec("bold"); });
      case "i": return run(() => { beginFormatChange(); exec("italic"); });
      case "u": return run(() => { beginFormatChange(); exec("underline"); });
      case "k": return run(insertLink);
      case "l": if (caps.alignment) return run(() => exec("justifyLeft")); return;
      case "e": if (caps.alignment) return run(() => exec("justifyCenter")); return;
      case "r": if (caps.alignment) return run(() => exec("justifyRight")); return;
      case "j": if (caps.alignment) return run(() => exec("justifyFull")); return;
      case "]": if (caps.fontControls) return run(() => changeFontSize(1)); return;
      case "[": if (caps.fontControls) return run(() => changeFontSize(-1)); return;
    }
  };
  for (const r of regions) r.addEventListener("keydown", onShortcut);

  const teardown = () => {
    toolbarObserver.disconnect();
    document.removeEventListener("click", closeOverflow);
    document.removeEventListener("selectionchange", scheduleSync);
    document.removeEventListener("mousemove", onFloatMouseMove);
    for (const r of regions) r.removeEventListener("keydown", onShortcut);
    window.clearTimeout(syncTimer);
    window.clearTimeout(floatHideTimer);
  };
  return { updateChangeButtons, teardown };
}
