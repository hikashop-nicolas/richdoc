// Named styles: the paragraph-style and character-style dropdowns (which live in the bottom
// bar), plus authoring and editing styles through a small dialog. Applying a style tags the
// block/run with data-rdoc-style / data-rdoc-cstyle; the appearance comes from injected CSS and
// the id round-trips to w:pStyle/w:rStyle (docx) or text:style-name (odt). The behaviours it
// needs (selection capture, block selection, the caret-state resync) come in as deps.
import { t } from "../../i18n";
import { firstFontFamily, blockBorders, parseCssBorder } from "../../util";
import { alignIcon } from "./icons";
import type { NewStyle, RichDoc } from "../../types";

export interface StylesDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  parts: RichDoc;
  styleBar: HTMLElement; // the bottom-bar slot the two dropdowns live in
  newStyles: NewStyle[];
  newStyleCss: HTMLStyleElement;
  fonts: string[]; // the document's font list, for the dialog's font picker
  getActiveEl: () => HTMLElement;
  mark: () => void;
  exec: (cmd: string, val?: string) => void;
  selectedBlocks: () => HTMLElement[];
  queryState: (cmd: string) => boolean;
  captureSel: () => void;
  restoreSel: () => void;
  resync: () => void; // the toolbar's syncToolbarState, to refresh the caret-state after a change
  sc: (key: string, opts?: { shift?: boolean; alt?: boolean }) => string; // shortcut label formatter
}

export function setupStyles(deps: StylesDeps) {
  const { wrap, regions, parts, styleBar, newStyles, newStyleCss, fonts: FONTS, getActiveEl, mark, exec, selectedBlocks, queryState, captureSel, restoreSel, resync, sc } = deps;

  const block = document.createElement("select");
  block.title = t("paragraphStyle");
  block.setAttribute("aria-label", t("paragraphStyle"));
  const BUILTIN_BLOCKS = new Set(["P", "H1", "H2", "H3"]);
  const blockSc: Record<string, string> = { P: sc("0", { alt: true }), H1: sc("1", { alt: true }), H2: sc("2", { alt: true }), H3: sc("3", { alt: true }) };
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(`${t(key)}  (${blockSc[v]})`, v));
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
      resync();
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
      resync();
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
    resync();
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
    const bs = blockBorders(b);
    for (const x of bs) css[`border-${x.side}`] = `${x.px}px ${x.style} #${x.hex.toLowerCase()}`;
    if (bs.length) css["padding"] = "2px 6px";
    const tabs = b.getAttribute("data-rdoc-tabstops");
    if (tabs) css["--rdoc-tabstops"] = tabs;
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
  let dFontFilled = false; // populated lazily from FONTS
  bgRow.append(dFont, bgLabel);
  // Paragraph border (paragraph styles only): a side preset + colour, line style and width.
  const borderRow = document.createElement("div");
  borderRow.className = "docxedit-dialog-row";
  const dBorder = document.createElement("select");
  dBorder.className = "docxedit-dialog-font";
  dBorder.title = t("paraBorder");
  for (const [v, key] of [["", "borderNone"], ["box", "borderAll"], ["top", "borderTop"], ["bottom", "borderBottom"], ["left", "borderLeft"], ["right", "borderRight"], ["topbottom", "borderTopBottom"], ["leftright", "borderLeftRight"]] as const) dBorder.add(new Option(t(key), v));
  const dBorderColor = document.createElement("input");
  dBorderColor.type = "color";
  dBorderColor.value = "#000000";
  dBorderColor.title = t("borderColor");
  const dBorderStyle = document.createElement("select");
  dBorderStyle.title = t("borderStyle");
  for (const [v, key] of [["solid", "bsSolid"], ["dashed", "bsDashed"], ["dotted", "bsDotted"], ["double", "bsDouble"]] as const) dBorderStyle.add(new Option(t(key), v));
  const dBorderWidth = document.createElement("select");
  dBorderWidth.title = t("borderWidth");
  for (const w of [1, 2, 3, 4]) dBorderWidth.add(new Option(`${w} px`, String(w)));
  borderRow.append(dBorder, dBorderColor, dBorderStyle, dBorderWidth);
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
  panel.append(dlgTitle, dlgName, fmtRow, alignRow, sizeRow, bgRow, borderRow, btnRow);
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
    resync(); // reset the dropdowns off the "New..." entry
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
      for (const s of ["top", "right", "bottom", "left"]) delete css[`border-${s}`];
      delete css["padding"];
      const preset = dBorder.value;
      if (preset) {
        const sides = preset === "box" ? ["top", "right", "bottom", "left"] : preset === "topbottom" ? ["top", "bottom"] : preset === "leftright" ? ["left", "right"] : [preset];
        const spec = `${dBorderWidth.value}px ${dBorderStyle.value} ${dBorderColor.value}`;
        for (const s of sides) css[`border-${s}`] = spec;
        css["padding"] = "2px 6px";
      }
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
    // Border: pick the preset matching which sides are present, and its colour.
    const bsides = ["top", "right", "bottom", "left"].filter((s) => parseCssBorder(pre[`border-${s}`]));
    const bset = new Set(bsides);
    dBorder.value = bsides.length === 4 ? "box"
      : bsides.length === 1 ? bsides[0]! // top / bottom / left / right
      : bsides.length === 2 && bset.has("top") && bset.has("bottom") ? "topbottom"
      : bsides.length === 2 && bset.has("left") && bset.has("right") ? "leftright" : "";
    const firstB = bsides.length ? parseCssBorder(pre[`border-${bsides[0]}`]) : null;
    dBorderColor.value = firstB ? `#${firstB.hex.toLowerCase()}` : "#000000";
    dBorderStyle.value = firstB && ["solid", "dashed", "dotted", "double"].includes(firstB.style) ? firstB.style : "solid";
    dBorderWidth.value = String(firstB ? Math.min(4, Math.max(1, firstB.px)) : 1);
    dlgPassthrough = {};
    for (const p of ["margin-left", "margin-top", "margin-bottom", "line-height", "--rdoc-tabstops"]) if (pre[p]) dlgPassthrough[p] = pre[p]!;
  };
  const openStyleDialog = (kind: "paragraph" | "character"): void => {
    restoreSel(); // the <select> click may have dropped the editor selection
    dlgKind = kind;
    dlgEditId = null;
    if (kind === "paragraph") {
      dlgBlocks = selectedBlocks();
      if (!dlgBlocks.length) {
        resync();
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
    borderRow.hidden = kind !== "paragraph";
    overlay.hidden = false;
    dlgName.focus();
  };
  // Edit an existing style's definition; on save it updates every block/run using it.
  const editStyle = (kind: "paragraph" | "character", id: string): void => {
    const def = styleDefMap.get(id);
    if (!def) {
      resync();
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
    borderRow.hidden = kind !== "paragraph";
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
        for (const p of ["textAlign", "marginLeft", "lineHeight", "marginTop", "marginBottom", "backgroundColor", "borderTop", "borderRight", "borderBottom", "borderLeft", "padding"] as const) b.style[p] = "";
        b.removeAttribute("data-rdoc-tabstops"); // tab stops now live in the style
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
    else resync();
  };
  const editCurrentCharacterStyle = (): void => {
    restoreSel();
    const id = styledAncestor("data-rdoc-cstyle")?.getAttribute("data-rdoc-cstyle");
    if (id && styleDefMap.get(id)?.kind === "character") editStyle("character", id);
    else resync();
  };

  // The paragraph + character style pickers live in the bottom bar's left slot.
  styleBar.append(block, cstyleSel);

  // Reflect the caret position in the two dropdowns (and whether "Edit current style" shows).
  const syncStyleState = (el: Element): void => {
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
  };

  return { block, cstyleSel, syncStyleState };
}
