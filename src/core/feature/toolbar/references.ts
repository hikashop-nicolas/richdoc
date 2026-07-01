// Bookmarks, cross-references, captions and (internal / web) links: the four insert-menu features
// that revolve around referenceable targets. They share a body of helpers (bookmark id minting,
// wrapping a heading/caption in a bookmark, the target lists) so they live together, apart from the
// main toolbar wiring. setupReferences builds the dialogs and returns the four open/insert handlers;
// the toolbar creates the buttons that call them.
import { t } from "../../i18n";
import { makeDialogAccessible } from "../dialog-a11y";
import { applyCaption, topBlock, captionText, captionAfter, type CaptionKind } from "../caption";

export interface ReferencesDeps {
  doc: HTMLElement; // the editable body
  wrap: HTMLElement; // the editor chrome the dialog overlays attach to
  mark: () => void; // flag the document dirty + schedule reflow
  exec: (cmd: string, val?: string) => void; // execCommand wrapper (createLink / unlink)
  captureSel: () => void; // save the current selection before a dialog steals focus
  restoreSel: () => void; // restore it before acting
  getActiveEl: () => HTMLElement; // the region (body / header / footer) that last had focus
}

export function setupReferences(deps: ReferencesDeps) {
  const { doc, wrap, mark, exec, captureSel, restoreSel, getActiveEl } = deps;

  // --- Bookmarks + cross-references ---------------------------------------------------------------
  const nextBmId = (): number => {
    let max = 0;
    for (const e of Array.from(doc.querySelectorAll("[data-rdoc-bm-id]"))) max = Math.max(max, parseInt(e.getAttribute("data-rdoc-bm-id") || "0", 10) || 0);
    return max + 1;
  };
  const mkBmEl = (cls: string, name: string | null, id: string): HTMLElement => {
    const a = document.createElement("a");
    a.className = cls;
    if (name) a.setAttribute("data-rdoc-bm", name);
    a.setAttribute("data-rdoc-bm-id", id);
    a.contentEditable = "false";
    return a;
  };
  // Insert a bookmark: a point at the caret, or a range wrapping the selection (a start marker at the
  // selection start + an end marker at its end, so the spanned text stays referenceable).
  const insertBookmark = (): void => {
    captureSel();
    const name = (window.prompt(t("bookmarkPrompt")) || "").trim();
    if (!name) return;
    restoreSel();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const cont = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element);
    if (!cont || !doc.contains(cont)) return;
    const id = String(nextBmId());
    const start = mkBmEl("docx-bookmark", name, id);
    const end = mkBmEl("docx-bookmark-end", null, id);
    end.setAttribute("data-rdoc-bm-end", name); // odt pairs by name; docx by id
    if (range.collapsed) { range.insertNode(end); range.insertNode(start); }
    else { const r2 = range.cloneRange(); r2.collapse(false); r2.insertNode(end); range.collapse(true); range.insertNode(start); }
    mark();
  };
  // A block's referenceable bookmark name (heading or caption): reuse one already wrapping it, else
  // wrap its content in a fresh _Ref bookmark so the cross-ref target survives a save (neither odf nor
  // ooxml treats a heading/caption as a bookmark, so a wrapping bookmark is what makes it referenceable).
  const blockBmName = (block: HTMLElement): string => {
    const existing = block.querySelector<HTMLElement>(":scope > .docx-bookmark");
    if (existing) return existing.getAttribute("data-rdoc-bm") || "";
    const id = String(nextBmId());
    const name = `_Ref${id}`;
    const start = mkBmEl("docx-bookmark", name, id);
    const end = mkBmEl("docx-bookmark-end", null, id);
    end.setAttribute("data-rdoc-bm-end", name);
    block.insertBefore(start, block.firstChild);
    block.appendChild(end);
    return name;
  };
  // A bookmark over only part of a caption, so "label and number" / "caption text" references round-trip
  // as ordinary bookmarks (a narrower range) rather than needing a non-standard field switch.
  const captionRefName = (cap: HTMLElement, gran: "label" | "captext"): string => {
    const reuse = Array.from(cap.querySelectorAll<HTMLElement>(":scope > .docx-bookmark")).find((b) => b.getAttribute("data-rdoc-bm-gran") === gran);
    if (reuse) return reuse.getAttribute("data-rdoc-bm") || "";
    const seq = cap.querySelector(":scope > [data-field='seq']");
    if (!seq) return blockBmName(cap); // no number to split on: fall back to the whole caption
    const id = String(nextBmId());
    const name = `_Ref${id}`;
    const start = mkBmEl("docx-bookmark", name, id);
    start.setAttribute("data-rdoc-bm-gran", gran);
    const end = mkBmEl("docx-bookmark-end", null, id);
    end.setAttribute("data-rdoc-bm-end", name);
    if (gran === "label") {
      cap.insertBefore(start, cap.firstChild); // ... start of caption
      (seq as ChildNode).after(end); //          ... through the number = "Figure 1"
    } else {
      const after = seq.nextSibling; // the ": text" node following the number
      if (after && after.nodeType === 3) {
        const m = /^\s*:\s*/.exec((after as Text).textContent || "");
        const textNode = m ? (after as Text).splitText(m[0].length) : after; // drop the leading ": "
        (textNode as ChildNode).before(start);
      } else {
        (seq as ChildNode).after(start);
      }
      cap.appendChild(end); // ... to the caption end = the description only
    }
    return name;
  };
  type XrefKind = "heading" | "bookmark" | "figure" | "table" | "equation";
  // The cross-reference targets of a kind, in document order: headings / captioned figures / captioned
  // tables (each referenced via a wrapping bookmark), or named bookmarks.
  const xrefTargets = (kind: XrefKind): { label: string; el: HTMLElement }[] => {
    if (kind === "bookmark") return Array.from(doc.querySelectorAll<HTMLElement>(".docx-bookmark")).map((b) => ({ label: b.getAttribute("data-rdoc-bm") || "", el: b }));
    const sel = kind === "heading" ? "h1,h2,h3" : `[data-rdoc-caption="${kind}"]`;
    return Array.from(doc.querySelectorAll<HTMLElement>(sel)).filter((h) => !h.closest(".docx-field-toc")).map((h) => ({ label: (h.textContent || "").trim() || t("untitled"), el: h }));
  };
  const mkXrefRadio = (group: string, value: string, label: string, checked: boolean) => {
    const lab = document.createElement("label");
    lab.className = "docxedit-noteinsert-opt";
    const input = document.createElement("input");
    input.type = "radio"; input.name = group; input.value = value; input.checked = checked;
    lab.append(input, document.createTextNode(` ${label}`));
    return { lab, input };
  };

  const xrefOverlay = document.createElement("div");
  xrefOverlay.className = "docxedit-dialog-overlay";
  xrefOverlay.hidden = true;
  const xrefPanel = document.createElement("div");
  xrefPanel.className = "docxedit-dialog docxedit-xref";
  const xrefTitle = document.createElement("div");
  xrefTitle.className = "docxedit-dialog-title";
  xrefTitle.textContent = t("insertCrossRef");
  const typeRow = document.createElement("div");
  typeRow.className = "docxedit-dialog-row docxedit-xref-kind";
  const tHeading = mkXrefRadio("rdoc-xref-kind", "heading", t("refHeadings"), true);
  const tBookmark = mkXrefRadio("rdoc-xref-kind", "bookmark", t("refBookmarks"), false);
  const tFigure = mkXrefRadio("rdoc-xref-kind", "figure", t("refFigures"), false);
  const tTable = mkXrefRadio("rdoc-xref-kind", "table", t("refTables"), false);
  const tEquation = mkXrefRadio("rdoc-xref-kind", "equation", t("refEquations"), false);
  const xrefKinds: { r: { lab: HTMLElement; input: HTMLInputElement }; kind: XrefKind }[] = [
    { r: tHeading, kind: "heading" }, { r: tBookmark, kind: "bookmark" },
    { r: tFigure, kind: "figure" }, { r: tTable, kind: "table" }, { r: tEquation, kind: "equation" },
  ];
  typeRow.append(tHeading.lab, tBookmark.lab, tFigure.lab, tTable.lab, tEquation.lab);
  const targetSel = document.createElement("select");
  targetSel.className = "docxedit-dialog-font";
  let xrefEls: HTMLElement[] = [];
  const fillTargets = () => {
    const kind = xrefKinds.find((k) => k.r.input.checked)?.kind ?? "heading";
    // "Label and number" / "Caption text" only make sense for captioned figures/tables/equations.
    const caption = kind === "figure" || kind === "table" || kind === "equation";
    fmtLabel.lab.hidden = !caption;
    fmtCaptionText.lab.hidden = !caption;
    if (!caption && (fmtLabel.input.checked || fmtCaptionText.input.checked)) fmtText.input.checked = true;
    const list = xrefTargets(kind);
    xrefEls = list.map((x) => x.el);
    targetSel.replaceChildren(...list.map((x, i) => { const o = document.createElement("option"); o.value = String(i); o.textContent = x.label; return o; }));
  };
  for (const k of xrefKinds) k.r.input.addEventListener("change", fillTargets);
  const fmtRow = document.createElement("div");
  fmtRow.className = "docxedit-dialog-row docxedit-xref-fmt";
  const fmtText = mkXrefRadio("rdoc-xref-fmt", "text", t("refFormatText"), true);
  const fmtLabel = mkXrefRadio("rdoc-xref-fmt", "label", t("refFormatLabel"), false);
  const fmtCaptionText = mkXrefRadio("rdoc-xref-fmt", "captext", t("refFormatCaptionText"), false);
  const fmtPage = mkXrefRadio("rdoc-xref-fmt", "page", t("refFormatPage"), false);
  const fmtDirection = mkXrefRadio("rdoc-xref-fmt", "direction", t("refFormatDirection"), false);
  fmtRow.append(fmtText.lab, fmtLabel.lab, fmtCaptionText.lab, fmtPage.lab, fmtDirection.lab);
  const xrefActions = document.createElement("div");
  xrefActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const xrefCancel = document.createElement("button");
  xrefCancel.type = "button"; xrefCancel.className = "docxedit-menu-item"; xrefCancel.textContent = t("cancel");
  const xrefInsertBtn = document.createElement("button");
  xrefInsertBtn.type = "button"; xrefInsertBtn.className = "docxedit-menu-item docxedit-dialog-primary"; xrefInsertBtn.textContent = t("insert");
  xrefActions.append(xrefCancel, xrefInsertBtn);
  xrefPanel.append(xrefTitle, typeRow, targetSel, fmtRow, xrefActions);
  xrefOverlay.appendChild(xrefPanel);
  wrap.appendChild(xrefOverlay);
  makeDialogAccessible(xrefOverlay);
  const closeXref = () => { xrefOverlay.hidden = true; };
  const openXrefDialog = () => {
    captureSel();
    // Declutter: show only the target kinds that have something to reference, and select the first.
    let first: { lab: HTMLElement; input: HTMLInputElement } | null = null;
    for (const k of xrefKinds) {
      const has = xrefTargets(k.kind).length > 0;
      k.r.lab.hidden = !has;
      if (has && !first) first = k.r;
    }
    (first ?? tHeading).input.checked = true;
    fmtText.input.checked = true;
    fillTargets();
    xrefOverlay.hidden = false;
  };
  xrefOverlay.addEventListener("mousedown", (e) => { if (e.target === xrefOverlay) closeXref(); });
  xrefCancel.addEventListener("click", closeXref);
  xrefInsertBtn.addEventListener("click", () => {
    const target = xrefEls[Number(targetSel.value)];
    if (!target) { closeXref(); return; }
    // "Label and number" / "Caption text" target a narrower bookmark inside the caption (fmt stays text);
    // the others reference a bookmark over the whole heading / caption / bookmark.
    const gran = fmtLabel.input.checked ? "label" : fmtCaptionText.input.checked ? "captext" : null;
    const fmt = gran ? "text" : fmtPage.input.checked ? "page" : fmtDirection.input.checked ? "direction" : "text";
    const name = gran ? captionRefName(target, gran)
      : target.hasAttribute("data-rdoc-bm") ? target.getAttribute("data-rdoc-bm") || ""
      : blockBmName(target);
    if (!name) { closeXref(); return; }
    closeXref();
    restoreSel();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const cont = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element);
    if (!cont || !doc.contains(cont)) return;
    const xr = document.createElement("a");
    xr.className = "docx-xref";
    xr.setAttribute("data-rdoc-xref", name);
    xr.setAttribute("data-rdoc-xref-fmt", fmt);
    xr.contentEditable = "false";
    // A placeholder until the next reflow recomputes it (page number / above-below / target text).
    xr.textContent = fmt === "page" ? "1" : fmt === "direction" ? t("refBelow") : (target.textContent || name).trim().slice(0, 80) || name;
    range.collapse(false);
    range.insertNode(xr);
    range.setStartAfter(xr);
    mark();
  });

  // Insert caption dialog: pick a label (Figure / Table / Equation) and optional text, captioning the
  // block at the caret. This is the entry point for equation captions (which have no image/table to hang off).
  const capOverlay = document.createElement("div");
  capOverlay.className = "docxedit-dialog-overlay";
  capOverlay.hidden = true;
  const capPanel = document.createElement("div");
  capPanel.className = "docxedit-dialog docxedit-caption";
  const capTitle = document.createElement("div");
  capTitle.className = "docxedit-dialog-title";
  capTitle.textContent = t("insertCaption");
  const mkCapField = (label: string, control: HTMLElement): HTMLElement => {
    const row = document.createElement("label");
    row.className = "docxedit-dialog-row docxedit-imgdialog-field";
    const span = document.createElement("span");
    span.textContent = label;
    row.append(span, control);
    return row;
  };
  const capTypeSel = document.createElement("select");
  capTypeSel.className = "docxedit-dialog-font";
  for (const [val, key] of [["figure", "captionFigure"], ["table", "captionTable"], ["equation", "captionEquation"]] as const) capTypeSel.add(new Option(t(key), val));
  capTypeSel.value = "equation";
  const capTextInput = document.createElement("input");
  capTextInput.type = "text";
  capTextInput.className = "docxedit-dialog-font";
  const capActions = document.createElement("div");
  capActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const capCancel = document.createElement("button");
  capCancel.type = "button"; capCancel.className = "docxedit-menu-item"; capCancel.textContent = t("cancel");
  const capInsertBtn = document.createElement("button");
  capInsertBtn.type = "button"; capInsertBtn.className = "docxedit-menu-item docxedit-dialog-primary"; capInsertBtn.textContent = t("insert");
  capActions.append(capCancel, capInsertBtn);
  capPanel.append(capTitle, mkCapField(t("captionType"), capTypeSel), mkCapField(t("caption"), capTextInput), capActions);
  capOverlay.appendChild(capPanel);
  wrap.appendChild(capOverlay);
  makeDialogAccessible(capOverlay);
  const closeCap = () => { capOverlay.hidden = true; };
  const captionBlock = (): HTMLElement | null => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const c = sel.getRangeAt(0).startContainer;
    const el = c.nodeType === 3 ? c.parentElement : (c as Element);
    return el && doc.contains(el) ? topBlock(el, doc) : null;
  };
  const prefillCaption = () => {
    const block = captionBlock();
    const cap = block ? captionAfter(block, capTypeSel.value as CaptionKind) : null;
    capTextInput.value = cap ? captionText(cap) : "";
  };
  const openCaptionDialog = () => {
    captureSel();
    prefillCaption();
    capOverlay.hidden = false;
    capTextInput.focus();
  };
  capTypeSel.addEventListener("change", prefillCaption);
  capOverlay.addEventListener("mousedown", (e) => { if (e.target === capOverlay) closeCap(); });
  capCancel.addEventListener("click", closeCap);
  capInsertBtn.addEventListener("click", () => {
    closeCap();
    restoreSel();
    const block = captionBlock();
    if (!block) return;
    applyCaption(block, capTypeSel.value as CaptionKind, capTextInput.value);
    mark();
  });

  // Link dialog: a web address, or a place in this document (a heading / bookmark / caption, referenced
  // by an href="#name" anchor). An in-document target without a name yet gets a wrapping bookmark minted.
  const linkOverlay = document.createElement("div");
  linkOverlay.className = "docxedit-dialog-overlay";
  linkOverlay.hidden = true;
  const linkPanel = document.createElement("div");
  linkPanel.className = "docxedit-dialog docxedit-link";
  const linkTitle = document.createElement("div");
  linkTitle.className = "docxedit-dialog-title";
  linkTitle.textContent = t("link");
  const linkModeRow = document.createElement("div");
  linkModeRow.className = "docxedit-dialog-row docxedit-noteinsert-kind";
  const mWeb = mkXrefRadio("rdoc-link-mode", "web", t("linkWeb"), true);
  const mDoc = mkXrefRadio("rdoc-link-mode", "doc", t("linkInDoc"), false);
  linkModeRow.append(mWeb.lab, mDoc.lab);
  const linkUrlInput = document.createElement("input");
  linkUrlInput.type = "text";
  linkUrlInput.className = "docxedit-dialog-font";
  linkUrlInput.placeholder = "https://";
  const linkUrlRow = document.createElement("div");
  linkUrlRow.className = "docxedit-dialog-row";
  linkUrlRow.append(linkUrlInput);
  const linkTargetSel = document.createElement("select");
  linkTargetSel.className = "docxedit-dialog-font";
  const linkTargetRow = document.createElement("div");
  linkTargetRow.className = "docxedit-dialog-row";
  linkTargetRow.append(linkTargetSel);
  let linkTargets: { label: string; el: HTMLElement }[] = [];
  const LINK_GROUPS: [XrefKind, () => string][] = [
    ["heading", () => t("refHeadings")], ["bookmark", () => t("refBookmarks")],
    ["figure", () => t("refFigures")], ["table", () => t("refTables")], ["equation", () => t("refEquations")],
  ];
  const fillLinkTargets = () => {
    linkTargets = [];
    linkTargetSel.replaceChildren();
    for (const [kind, label] of LINK_GROUPS) {
      const items = xrefTargets(kind);
      if (!items.length) continue;
      const og = document.createElement("optgroup");
      og.label = label();
      for (const it of items) {
        const o = document.createElement("option");
        o.value = String(linkTargets.length);
        o.textContent = it.label;
        og.appendChild(o);
        linkTargets.push(it);
      }
      linkTargetSel.appendChild(og);
    }
  };
  const syncLinkMode = () => { linkUrlRow.hidden = !mWeb.input.checked; linkTargetRow.hidden = mWeb.input.checked; };
  mWeb.input.addEventListener("change", syncLinkMode);
  mDoc.input.addEventListener("change", syncLinkMode);
  const linkActions = document.createElement("div");
  linkActions.className = "docxedit-dialog-row docxedit-dialog-actions";
  const linkCancel = document.createElement("button");
  linkCancel.type = "button"; linkCancel.className = "docxedit-menu-item"; linkCancel.textContent = t("cancel");
  const linkApply = document.createElement("button");
  linkApply.type = "button"; linkApply.className = "docxedit-menu-item docxedit-dialog-primary"; linkApply.textContent = t("insert");
  linkActions.append(linkCancel, linkApply);
  linkPanel.append(linkTitle, linkModeRow, linkUrlRow, linkTargetRow, linkActions);
  linkOverlay.appendChild(linkPanel);
  wrap.appendChild(linkOverlay);
  makeDialogAccessible(linkOverlay);
  const closeLink = () => { linkOverlay.hidden = true; };
  const openLinkDialog = () => {
    captureSel();
    // Prefill the web field from an existing link on the selection.
    const a = getActiveEl()?.closest?.("a[href]") as HTMLAnchorElement | null;
    const existing = a?.getAttribute("href") ?? "";
    const inDoc = existing.startsWith("#");
    mWeb.input.checked = !inDoc;
    mDoc.input.checked = inDoc;
    linkUrlInput.value = inDoc ? "" : existing;
    fillLinkTargets();
    syncLinkMode();
    linkOverlay.hidden = false;
    (mWeb.input.checked ? linkUrlInput : linkTargetSel).focus();
  };
  linkOverlay.addEventListener("mousedown", (e) => { if (e.target === linkOverlay) closeLink(); });
  linkCancel.addEventListener("click", closeLink);
  linkApply.addEventListener("click", () => {
    closeLink();
    restoreSel();
    if (mWeb.input.checked) {
      const url = linkUrlInput.value.trim();
      if (url) exec("createLink", url); else exec("unlink");
      return;
    }
    const target = linkTargets[Number(linkTargetSel.value)];
    if (!target) return;
    const name = target.el.hasAttribute("data-rdoc-bm") ? target.el.getAttribute("data-rdoc-bm") || "" : blockBmName(target.el);
    if (!name) return;
    const tagInternalLinks = () => doc.querySelectorAll<HTMLElement>('a[href^="#"]:not([title])').forEach((a) => a.setAttribute("title", t("linkFollow")));
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed) { exec("createLink", "#" + name); tagInternalLinks(); return; }
    // Nothing selected: drop in the target's label as the link text.
    if (sel && sel.rangeCount) {
      const a2 = document.createElement("a");
      a2.setAttribute("href", "#" + name);
      a2.setAttribute("title", t("linkFollow"));
      a2.textContent = target.label;
      const r = sel.getRangeAt(0); r.insertNode(a2); r.setStartAfter(a2); r.collapse(true);
      mark();
    }
  });

  return { insertBookmark, openCaptionDialog, openXrefDialog, openLinkDialog };
}
