// Track-changes (suggestion mode): intercepts editing so typing/deleting in suggest mode
// produces ins/del marks and paragraph-mark and rPr-change records, plus accept/reject
// (per change via a click popover, or all at once) and the toolbar buttons. Extracted from
// the engine; the toolbar button factories and the active-region getter come in as deps.
import { t } from "../i18n";
import { toHex6, fontSizeToHalfPt, firstFontFamily } from "../util";
import type { EditorOptions } from "../types";

export interface TrackChangesDeps {
  doc: HTMLElement;
  wrap: HTMLElement;
  regions: HTMLElement[];
  options: EditorOptions;
  mark: (coalesce?: string) => void;
  positionCards: () => void;
  getActiveEl: () => HTMLElement;
  iconBtn: (svg: string, title: string, fn: () => void) => HTMLElement;
  btn: (label: string, title: string, fn: () => void, cls?: string) => HTMLElement;
}

export function setupTrackChanges(deps: TrackChangesDeps) {
  const { doc, wrap, regions, options, mark, positionCards, getActiveEl, iconBtn, btn } = deps;

  // --- Suggestion mode (track changes) ---------------------------------------
  let suggesting = false;
  const sugAuthor = () => options.author || "Author";
  const sugDate = () => options.now || new Date().toISOString();
  const blockOf = (n: Node | null): HTMLElement | null => {
    const start = n && (n.nodeType === 3 ? n.parentElement : (n as Element));
    const el = start?.closest?.("p,h1,h2,h3,h4,h5,h6,li,div") as HTMLElement | null;
    // never return a region container (the doc/header/footer itself)
    if (!el || el.classList.contains("docxedit-doc") || el.classList.contains("docxedit-header") || el.classList.contains("docxedit-footer")) return null;
    return el;
  };
  const markPara = (el: HTMLElement, kind: "ins" | "del") => {
    el.setAttribute("data-rev-para", kind);
    el.setAttribute("data-rev-author", sugAuthor());
    el.setAttribute("data-rev-date", sugDate());
    el.classList.add(`docx-para-${kind}`);
  };
  // Capture current run formatting from an element's computed style (for rPrChange "old").
  const captureFmt = (el: Element | null): Record<string, unknown> => {
    if (!el) return {};
    const cs = getComputedStyle(el);
    const deco = cs.textDecorationLine || "";
    const wght = Number(cs.fontWeight);
    return {
      b: cs.fontWeight === "bold" || wght >= 600,
      i: cs.fontStyle === "italic",
      u: /underline/.test(deco),
      strike: /line-through/.test(deco),
      color: toHex6(cs.color),
      sizeHalfPt: fontSizeToHalfPt(cs.fontSize),
      font: firstFontFamily(cs.fontFamily),
    };
  };
  const fmtToStyle = (f: Record<string, unknown>): string => {
    const p: string[] = [];
    if (f.b) p.push("font-weight:bold");
    if (f.i) p.push("font-style:italic");
    const dec = [f.u ? "underline" : "", f.strike ? "line-through" : ""].filter(Boolean).join(" ");
    if (dec) p.push(`text-decoration:${dec}`);
    if (f.color) p.push(`color:#${f.color}`);
    if (f.sizeHalfPt) p.push(`font-size:${(f.sizeHalfPt as number) / 2}pt`);
    if (f.font) p.push(`font-family:'${String(f.font).replace(/'/g, "")}', serif`);
    return p.join(";");
  };
  // Record a formatting change: wrap the selection so its old props become an rPrChange.
  const beginFormatChange = () => {
    if (!suggesting) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || sel.isCollapsed) return;
    const range = sel.getRangeAt(0);
    const anc = range.commonAncestorContainer;
    const el = anc.nodeType === 3 ? anc.parentElement : (anc as Element);
    if (el?.closest?.(".docx-rpr-change")) return; // already inside a change
    const span = document.createElement("span");
    span.className = "docx-rpr-change";
    span.setAttribute("data-old", JSON.stringify(captureFmt(el)));
    span.setAttribute("data-rev-author", sugAuthor());
    span.setAttribute("data-rev-date", sugDate());
    span.appendChild(range.extractContents());
    range.insertNode(span);
    const r2 = document.createRange();
    r2.selectNodeContents(span);
    sel.removeAllRanges();
    sel.addRange(r2);
  };

  const insertSuggestText = (data: string) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    // Merge into an adjacent same-author insertion: the caret's own .docx-ins, or one
    // ending just before the caret.
    const sc = range.startContainer;
    let host = (sc.nodeType === 3 ? sc.parentElement : (sc as Element))?.closest?.(".docx-ins") as HTMLElement | null;
    if (!host && sc.nodeType === 1) {
      const prev = (sc as Element).childNodes[range.startOffset - 1];
      if (prev && prev.nodeType === 1) host = (prev as Element).closest?.(".docx-ins") as HTMLElement | null;
    }
    let after: Node;
    if (host && host.getAttribute("data-author") === sugAuthor()) {
      const tn = document.createTextNode(data);
      if (host.contains(sc)) range.insertNode(tn);
      else host.appendChild(tn);
      after = tn;
    } else {
      const ins = document.createElement("ins");
      ins.className = "docx-ins";
      ins.setAttribute("data-author", sugAuthor());
      ins.setAttribute("data-date", sugDate());
      ins.textContent = data;
      range.insertNode(ins);
      after = ins.firstChild ?? ins; // caret inside the ins, so the next keystroke merges
    }
    const r2 = document.createRange();
    r2.setStartAfter(after);
    r2.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r2);
  };

  const suggestDelete = (range: Range) => {
    if (range.collapsed) return;
    const anc = range.commonAncestorContainer;
    const ancEl = anc.nodeType === 3 ? anc.parentElement : (anc as Element);
    const ownIns = ancEl?.closest?.(".docx-ins");
    const frag = range.extractContents();
    if (ownIns && ownIns.getAttribute("data-author") === sugAuthor()) {
      range.collapse(true); // deleting your own pending insertion -> just drop it
    } else {
      const del = document.createElement("del");
      del.className = "docx-del";
      del.setAttribute("data-author", sugAuthor());
      del.setAttribute("data-date", sugDate());
      del.appendChild(frag);
      range.insertNode(del);
      range.setStartBefore(del);
      range.collapse(true);
    }
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  };

  const onBeforeInput = (e: Event) => {
    if (!suggesting) return;
    const ie = e as InputEvent;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const type = ie.inputType;
    if (type === "insertText" || type === "insertReplacementText" || type === "insertFromPaste") {
      ie.preventDefault();
      const data = ie.data ?? ie.dataTransfer?.getData("text/plain") ?? "";
      if (!data) return;
      const range = sel.getRangeAt(0);
      if (!range.collapsed) suggestDelete(range);
      insertSuggestText(data);
      mark("suggest");
      positionCards();
    } else if (type.startsWith("delete")) {
      ie.preventDefault();
      const range = sel.getRangeAt(0);
      let r = range;
      if (range.collapsed) {
        const tr = ie.getTargetRanges?.()[0];
        if (tr) {
          r = document.createRange();
          r.setStart(tr.startContainer, tr.startOffset);
          r.setEnd(tr.endContainer, tr.endOffset);
        }
      }
      if (r.collapsed) return;
      // A deletion that spans a paragraph boundary is a paragraph-mark deletion (merge):
      // mark the first block's mark as deleted instead of merging.
      const sb = blockOf(r.startContainer);
      const eb = blockOf(r.endContainer);
      if (sb && eb && sb !== eb) {
        if (!sb.hasAttribute("data-rev-para")) markPara(sb, "del");
      } else {
        suggestDelete(r);
      }
      mark("suggest");
      positionCards();
    }
  };
  for (const r of regions) r.addEventListener("beforeinput", onBeforeInput);
  // Paragraph split (Enter) in suggesting mode -> mark the first half's paragraph mark.
  for (const region of regions)
    region.addEventListener("input", (e) => {
      if (!suggesting || (e as InputEvent).inputType !== "insertParagraph") return;
      const sel = window.getSelection();
      const block = sel && sel.rangeCount ? blockOf(sel.getRangeAt(0).startContainer) : null;
      const first = block?.previousElementSibling as HTMLElement | null;
      if (first && !first.hasAttribute("data-rev-para")) markPara(first, "ins");
      positionCards();
    });

  const unwrap = (el: Element) => {
    while (el.firstChild) el.parentNode?.insertBefore(el.firstChild, el);
    el.remove();
  };
  const mergeWithNext = (el: HTMLElement) => {
    const next = el.nextElementSibling;
    if (next) {
      while (next.firstChild) el.appendChild(next.firstChild);
      next.remove();
    }
  };
  const clearPara = (el: HTMLElement) => {
    el.removeAttribute("data-rev-para");
    el.removeAttribute("data-rev-author");
    el.removeAttribute("data-rev-date");
    el.classList.remove("docx-para-ins", "docx-para-del");
  };
  const resolveChange = (el: Element, accept: boolean) => {
    if (el.classList.contains("docx-rpr-change")) {
      // formatting change: accept keeps the new look; reject restores the old props
      if (accept) unwrap(el);
      else {
        let old: Record<string, unknown> = {};
        try {
          old = JSON.parse(el.getAttribute("data-old") || "{}");
        } catch {
          /* keep default */
        }
        const span = document.createElement("span");
        const style = fmtToStyle(old);
        if (style) span.setAttribute("style", style);
        span.textContent = el.textContent;
        el.replaceWith(span);
      }
    } else if (el.hasAttribute("data-rev-para")) {
      const kind = el.getAttribute("data-rev-para");
      const merge = kind === "ins" ? !accept : accept; // ins-reject and del-accept both merge
      if (merge) mergeWithNext(el as HTMLElement);
      clearPara(el as HTMLElement);
    } else {
      const isDel = el.classList.contains("docx-del");
      if (accept ? isDel : !isDel) el.remove();
      else unwrap(el);
    }
    mark();
    positionCards();
  };
  const resolveAll = (accept: boolean) => {
    for (const el of Array.from(wrap.querySelectorAll(".docx-ins, .docx-del, .docx-rpr-change, [data-rev-para]"))) resolveChange(el, accept);
  };

  // Accept/reject popover when a tracked change is clicked.
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement).closest?.(".docxedit-change-pop")) return;
    document.querySelector(".docxedit-change-pop")?.remove();
    const ch = (e.target as HTMLElement).closest?.(".docx-ins, .docx-del, .docx-rpr-change") as HTMLElement | null;
    if (!ch) return;
    const pop = document.createElement("div");
    pop.className = "docxedit-change-pop";
    const mk = (label: string, title: string, accept: boolean) => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.title = title;
      b.addEventListener("click", (ev) => {
        ev.stopPropagation();
        resolveChange(ch, accept);
        pop.remove();
      });
      return b;
    };
    pop.append(mk("✓", t("accept"), true), mk("✕", t("reject"), false));
    wrap.appendChild(pop);
    const cr = ch.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    pop.style.left = `${Math.min(cr.left - wr.left, wrap.clientWidth - 80)}px`;
    pop.style.top = `${cr.bottom - wr.top + 3}px`;
  });

  const suggestIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3" aria-hidden="true">' +
    '<path d="M2 11.5 10.5 3l2.5 2.5L4.5 14H2z"/><line x1="9" y1="4.5" x2="11.5" y2="7"/></svg>';
  const suggestBtn = iconBtn(suggestIcon, t("suggesting"), () => {
    suggesting = !suggesting;
    suggestBtn.classList.toggle("is-on", suggesting);
    getActiveEl().focus();
  });

  const acceptAllBtn = btn("✓", t("acceptAll"), () => resolveAll(true));
  const rejectAllBtn = btn("✕", t("rejectAll"), () => resolveAll(false));
  const updateChangeButtons = () => {
    const has = !!doc.querySelector(".docx-ins, .docx-del");
    acceptAllBtn.style.display = has ? "" : "none";
    rejectAllBtn.style.display = has ? "" : "none";
  };

  return { beginFormatChange, suggestBtn, acceptAllBtn, rejectAllBtn, updateChangeButtons };
}
