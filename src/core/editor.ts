// Shared rich-document editor engine. Renders the editable surface, toolbar, comments
// panel, track changes, image/page chrome and passthrough; the format-specific parse,
// serialize and comment markers come from an Adapter (see core/types.ts). docx is the
// reference adapter; odt reuses this same engine.

import { t } from "./i18n";
import { defaultPageGeometry, paginate } from "./page";
import { bytesToBase64, toHex6, fontSizeToHalfPt, firstFontFamily } from "./util";
import type { Adapter, EditorOptions, RichEditor, RichDoc, CommentEntry, CommentThread } from "./types";
import "../adapters/docx/docxedit.css";

export function createRichEditor(container: HTMLElement, adapter: Adapter, options: EditorOptions = {}): RichEditor {
  const original = adapter.original;
  const caps = adapter.capabilities;
  let dirty = false;

  const wrap = document.createElement("div");
  wrap.className = "docxedit-wrap";
  const toolbar = document.createElement("div");
  toolbar.className = "docxedit-toolbar";
  toolbar.setAttribute("role", "toolbar");
  toolbar.setAttribute("aria-label", t("toolbar"));
  const scroll = document.createElement("div");
  scroll.className = "docxedit-scroll";
  const page = document.createElement("div");
  page.className = "docxedit-page";
  const doc = document.createElement("div");
  doc.className = "docxedit-doc";
  doc.contentEditable = "true";
  doc.spellcheck = false;
  doc.setAttribute("role", "textbox");
  doc.setAttribute("aria-multiline", "true");
  doc.setAttribute("aria-label", t("documentText"));

  let parts: RichDoc = { body: "<p><br></p>", header: "", footer: "", comments: [] };
  try {
    parts = adapter.read();
  } catch (e) {
    console.warn("richdoc: failed to parse document", e);
  }
  doc.innerHTML = parts.body || "<p><br></p>";

  // Render the document in its own embedded typefaces, if the adapter supplied any.
  const fontUrls: string[] = parts.fontUrls ?? [];
  if (parts.fontCss) {
    const fs = document.createElement("style");
    fs.textContent = parts.fontCss;
    wrap.appendChild(fs);
  }
  if (parts.defaultFont) page.style.setProperty("--docxedit-doc-font", `"${parts.defaultFont.replace(/"/g, "")}"`);

  // Page geometry: render at the document's real size and margins, or the default size
  // (A4 unless options override) when the file declares none. Height is unused until
  // pagination (Phase 1); width and margins apply now.
  const geometry = parts.page ?? defaultPageGeometry(options.defaultPageSize ?? "a4");
  page.style.setProperty("--rdoc-page-width", `${geometry.widthPx}px`);
  page.style.setProperty("--rdoc-margin-top", `${geometry.margin.top}px`);
  page.style.setProperty("--rdoc-margin-right", `${geometry.margin.right}px`);
  page.style.setProperty("--rdoc-margin-bottom", `${geometry.margin.bottom}px`);
  page.style.setProperty("--rdoc-margin-left", `${geometry.margin.left}px`);

  const band = (cls: string, label: string, html: string): HTMLElement | null => {
    if (!html) return null;
    const el = document.createElement("div");
    el.className = cls;
    el.contentEditable = "true";
    el.spellcheck = false;
    el.setAttribute("role", "textbox");
    el.setAttribute("aria-multiline", "true");
    el.setAttribute("aria-label", label);
    el.innerHTML = html;
    return el;
  };
  const header = band("docxedit-header", t("header"), parts.header);
  const footer = band("docxedit-footer", t("footer"), parts.footer);

  // Paginated view: one continuous editable body (doc) on top of a layer of page-card
  // decorations, with inert spacer gaps inserted at page boundaries. Pageless view keeps
  // the body and header/footer stacked in one card (the previous behaviour).
  const paginated = options.paginated ?? true;
  const pagelayer = document.createElement("div"); // cards + page numbers, behind the body
  pagelayer.className = "docxedit-pagelayer";
  pagelayer.setAttribute("aria-hidden", "true");
  const hflayer = document.createElement("div"); // header/footer clones, above the body (clickable)
  hflayer.className = "docxedit-hflayer";
  // Off-screen holder so header/footer can be measured (and kept as the save source)
  // without showing as in-flow bands in paginated mode.
  const measure = document.createElement("div");
  measure.className = "docxedit-measure";
  if (paginated) {
    page.classList.add("is-paginated");
    page.append(pagelayer, doc, hflayer);
    if (header) measure.appendChild(header);
    if (footer) measure.appendChild(footer);
    page.appendChild(measure);
  } else {
    if (header) page.appendChild(header);
    page.appendChild(doc);
    if (footer) page.appendChild(footer);
  }

  // Keep the page centred, with the comments column in the right margin (Google-Docs
  // style): an empty left spacer balances the right comments area so the page stays centred.
  const canvas = document.createElement("div");
  canvas.className = "docxedit-canvas";
  const leftSpacer = document.createElement("div");
  leftSpacer.className = "docxedit-margin";
  const rightArea = document.createElement("div");
  rightArea.className = "docxedit-margin";
  const cmtPanel = document.createElement("div");
  cmtPanel.className = "docxedit-comments";
  rightArea.appendChild(cmtPanel);
  canvas.append(leftSpacer, page, rightArea);
  scroll.appendChild(canvas);
  wrap.append(toolbar, scroll);
  container.appendChild(wrap);

  // The editable regions (body + header/footer). Toolbar actions target whichever last
  // had focus, so formatting works inside the header and footer too.
  const regions = [doc, header, footer].filter(Boolean) as HTMLElement[];
  let activeEl: HTMLElement = doc;

  // Next comment id: one past the highest already present.
  let nextCommentId = 0;
  for (const m of Array.from(wrap.querySelectorAll("[data-comment-id]"))) {
    const n = Number(m.getAttribute("data-comment-id"));
    if (Number.isFinite(n)) nextCommentId = Math.max(nextCommentId, n + 1);
  }

  // Comments side panel: one card per thread (replies grouped), anchored vertically to the
  // commented range, with reactions and a "more" toggle for long text.
  const pendingReactions: { commentId: string; emoji: string; person: string }[] = [];
  const pendingReplies: { id: string; paraId: string; parentParaId: string; author: string; date: string; text: string }[] = [];
  const pendingDone = new Map<string, boolean>(); // thread paraId -> done
  const deletedComments: string[] = []; // comment ids removed from the document
  // thread membership, kept current as replies/comments are added.
  const threadOf = new Map<string, string>(); // any comment id -> its thread id
  const threadMembers = new Map<string, string[]>();
  const registerThread = (threadId: string, memberIds: string[]) => {
    threadMembers.set(threadId, memberIds);
    for (const m of memberIds) threadOf.set(m, threadId);
  };
  let paraSeed = 0x7f000000;
  const freshParaId = () => (paraSeed++).toString(16).toUpperCase().padStart(8, "0");
  const REACT_CHOICES = ["\u{1F44D}", "❤️", "\u{1F602}", "\u{1F389}", "\u{1F440}", "\u{1F64F}"];
  const metaLine = (c: { author: string; date: string }) => (c.date ? `${c.author} – ${c.date.slice(0, 10)}` : c.author);

  const renderReactions = (row: HTMLElement, entry: CommentEntry) => {
    row.querySelectorAll(".docxedit-react").forEach((n) => n.remove());
    const addBtn = row.querySelector(".docxedit-react-add");
    for (const r of entry.reactions) {
      if (!r.people.length) continue;
      const span = document.createElement("span");
      span.className = "docxedit-react";
      span.title = r.people.join(", ");
      span.textContent = r.emoji + (r.people.length > 1 ? " " + r.people.length : "");
      row.insertBefore(span, addBtn);
    }
  };

  const buildItem = (entry: CommentEntry, isReply: boolean): HTMLElement => {
    const item = document.createElement("div");
    item.className = "docxedit-cmt-item" + (isReply ? " docxedit-cmt-reply" : "");
    const meta = document.createElement("b");
    meta.textContent = metaLine(entry);
    const text = document.createElement("div");
    text.className = "docxedit-cmt-text";
    text.textContent = entry.text;
    const more = document.createElement("button");
    more.type = "button";
    more.className = "docxedit-cmt-more is-hidden";
    more.textContent = t("more");
    more.addEventListener("click", (e) => {
      e.stopPropagation();
      text.classList.add("expanded");
      more.classList.add("is-hidden");
      positionCards();
    });
    item.append(meta, text, more);
    if (caps.commentReactions) {
      const row = document.createElement("div");
      row.className = "docxedit-cmt-react-row";
      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "docxedit-react-add";
      addBtn.title = t("addReaction");
      addBtn.textContent = "\u{1F642}+";
      addBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        openReactionPicker(addBtn, entry, row);
      });
      row.appendChild(addBtn);
      renderReactions(row, entry);
      item.appendChild(row);
    }
    return item;
  };

  const openReactionPicker = (anchor: HTMLElement, entry: CommentEntry, row: HTMLElement) => {
    document.querySelector(".docxedit-react-pop")?.remove();
    const pop = document.createElement("div");
    pop.className = "docxedit-react-pop";
    for (const emoji of REACT_CHOICES) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = emoji;
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        const person = options.author || "Author";
        const existing = entry.reactions.find((r) => r.emoji === emoji);
        if (existing) {
          if (!existing.people.includes(person)) existing.people.push(person);
        } else entry.reactions.push({ emoji, people: [person] });
        pendingReactions.push({ commentId: entry.id, emoji, person });
        renderReactions(row, entry);
        pop.remove();
        positionCards();
        mark();
      });
      pop.appendChild(b);
    }
    wrap.appendChild(pop);
    const ar = anchor.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    pop.style.left = `${Math.min(ar.left - wr.left, wrap.clientWidth - 200)}px`;
    pop.style.top = `${ar.bottom - wr.top + 4}px`;
  };

  const setActiveComment = (threadId: string | null) => {
    for (const c of Array.from(cmtPanel.children)) c.classList.toggle("active", (c as HTMLElement).dataset.commentId === threadId);
    const members = threadId ? threadMembers.get(threadId) ?? [threadId] : [];
    for (const r of Array.from(wrap.querySelectorAll(".docx-comment"))) {
      const rid = (r as HTMLElement).getAttribute("data-comment-id") ?? "";
      (r as HTMLElement).classList.toggle("active", members.includes(rid));
    }
  };

  const actionBtn = (label: string, title: string, fn: (e: Event) => void): HTMLElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "docxedit-cmt-action";
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.addEventListener("click", fn);
    return b;
  };

  const addReplyToThread = (card: HTMLElement, threadId: string, parentParaId: string, text: string) => {
    const id = String(nextCommentId++);
    const paraId = freshParaId();
    const author = options.author || "Author";
    const date = options.now || new Date().toISOString();
    pendingReplies.push({ id, paraId, parentParaId, author, date, text });
    const members = threadMembers.get(threadId) ?? [threadId];
    members.push(id);
    registerThread(threadId, members);
    const box = card.querySelector(".docxedit-cmt-replybox");
    card.insertBefore(buildItem({ id, author, date, text, reactions: [], paraId }, true), box);
    mark();
    positionCards();
  };

  const buildReplyBox = (card: HTMLElement, thread: { id: string }): HTMLElement => {
    const box = document.createElement("div");
    box.className = "docxedit-cmt-replybox";
    const showInput = () => {
      box.innerHTML = "";
      const ta = document.createElement("textarea");
      ta.className = "docxedit-cmt-replyinput";
      ta.rows = 2;
      ta.placeholder = t("reply");
      ta.addEventListener("click", (e) => e.stopPropagation());
      const send = document.createElement("button");
      send.type = "button";
      send.className = "docxedit-cmt-send";
      send.textContent = t("send");
      const commit = () => {
        const txt = ta.value.trim();
        box.replaceWith(buildReplyBox(card, thread));
        if (txt) addReplyToThread(card, thread.id, card.dataset.paraId ?? "", txt);
      };
      send.addEventListener("click", (e) => {
        e.stopPropagation();
        commit();
      });
      ta.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          commit();
        }
      });
      box.append(ta, send);
      ta.focus();
    };
    const btn2 = document.createElement("button");
    btn2.type = "button";
    btn2.className = "docxedit-cmt-replybtn";
    btn2.textContent = t("reply");
    btn2.addEventListener("click", (e) => {
      e.stopPropagation();
      showInput();
    });
    box.appendChild(btn2);
    return box;
  };

  const addThreadCard = (thread: CommentThread): HTMLElement => {
    const card = document.createElement("div");
    card.className = "docxedit-cmt-card" + (thread.resolved ? " resolved" : "");
    card.dataset.commentId = thread.id;
    card.dataset.paraId = thread.paraId || "";
    registerThread(thread.id, [thread.id, ...thread.replies.map((r) => r.id)]);

    const actions = document.createElement("div");
    actions.className = "docxedit-cmt-actions";
    actions.append(
      actionBtn("✓", t("resolve"), (e) => {
        e.stopPropagation();
        const resolved = !card.classList.contains("resolved");
        card.classList.toggle("resolved", resolved);
        pendingDone.set(card.dataset.paraId || "", resolved);
        if (resolved) setActiveComment(null);
        mark();
        positionCards();
      }),
      actionBtn("✕", t("deleteComment"), (e) => {
        e.stopPropagation();
        const members = threadMembers.get(thread.id) ?? [thread.id];
        for (const id of members) {
          deletedComments.push(id);
          // unwrap the highlight span (keep its text), then drop range/reference markers
          wrap.querySelectorAll(`.docx-comment[data-comment-id="${CSS.escape(id)}"]`).forEach((span) => {
            while (span.firstChild) span.parentNode?.insertBefore(span.firstChild, span);
            span.remove();
          });
          wrap.querySelectorAll(`.docx-comment-ref[data-comment-id="${CSS.escape(id)}"]`).forEach((n) => n.remove());
          for (const m of Array.from(wrap.querySelectorAll(".docx-cmark"))) {
            if ((m.getAttribute("data-docx-xml") ?? "").includes(`w:id="${id}"`)) m.remove();
          }
        }
        card.remove();
        mark();
        positionCards();
      }),
    );
    card.appendChild(actions);

    card.appendChild(buildItem(thread, false));
    for (const reply of thread.replies) card.appendChild(buildItem(reply, true));
    if (caps.commentReplies) card.appendChild(buildReplyBox(card, thread));

    card.addEventListener("click", () => {
      setActiveComment(thread.id);
      wrap.querySelector(`.docx-comment[data-comment-id="${CSS.escape(thread.id)}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
    });
    cmtPanel.appendChild(card);
    return card;
  };

  // Anchor each card to its range's vertical position; stack to avoid overlap.
  const positionCards = () => {
    const cards = Array.from(cmtPanel.children) as HTMLElement[];
    if (!cards.length) return;
    const panelTop = cmtPanel.getBoundingClientRect().top;
    const measured = cards.map((card) => {
      const id = card.dataset.commentId ?? "";
      const marker =
        wrap.querySelector(`.docx-comment[data-comment-id="${CSS.escape(id)}"]`) ?? wrap.querySelector(`.docx-comment-ref[data-comment-id="${CSS.escape(id)}"]`);
      const y = marker ? (marker as HTMLElement).getBoundingClientRect().top - panelTop : 0;
      return { card, y };
    });
    measured.sort((a, b) => a.y - b.y);
    let prevBottom = 0;
    for (const { card, y } of measured) {
      // reveal "more" only when the text actually overflows
      for (const item of Array.from(card.querySelectorAll(".docxedit-cmt-item"))) {
        const txt = item.querySelector(".docxedit-cmt-text") as HTMLElement | null;
        const moreBtn = item.querySelector(".docxedit-cmt-more") as HTMLElement | null;
        if (txt && moreBtn && !txt.classList.contains("expanded")) {
          moreBtn.classList.toggle("is-hidden", txt.scrollHeight <= txt.clientHeight + 2);
        }
      }
      const top = Math.max(y, prevBottom);
      card.style.top = `${top}px`;
      prevBottom = top + card.offsetHeight + 10;
    }
    cmtPanel.style.height = `${Math.max(prevBottom, page.offsetHeight)}px`;
  };

  // --- Pagination -----------------------------------------------------------
  // Measure top-level block heights, compute page breaks, then draw page cards behind the
  // body and insert inert spacer gaps so the flow lands at each page's content box. The
  // body stays one contenteditable; spacers are stripped before saving (see cleanBody).
  const PAGE_GAP = 24;
  let editingBand: HTMLElement | null = null;

  // Click a repeated header/footer to edit it in place (Word-like): the canonical band is
  // moved into the clicked position, made editable and focused; on blur it returns to the
  // hidden source and the pages re-clone its new content.
  const editBand = (src: HTMLElement, topPx: number) => {
    if (editingBand) return;
    editingBand = src;
    src.classList.add("is-editing");
    src.style.cssText = `position:absolute;top:${topPx}px;left:0;width:100%;z-index:3;background:#fff`;
    hflayer.appendChild(src);
    src.focus();
    const finish = () => {
      src.removeEventListener("blur", finish);
      src.classList.remove("is-editing");
      src.removeAttribute("style");
      measure.appendChild(src);
      editingBand = null;
      reflow();
    };
    src.addEventListener("blur", finish);
  };

  const repaginate = () => {
    if (!paginated || editingBand) return;
    for (const s of Array.from(doc.querySelectorAll(":scope > .docxedit-pagespacer"))) s.remove();
    pagelayer.replaceChildren();
    hflayer.replaceChildren();

    // measure header/footer at full page width (their own padding provides the margins)
    measure.style.width = `${geometry.widthPx}px`;
    const headerH = header ? header.offsetHeight : 0;
    const footerH = footer ? footer.offsetHeight : 0;
    const contentTop = geometry.margin.top + headerH;
    const contentBottomInset = geometry.margin.bottom + footerH;
    const contentHeight = geometry.heightPx - contentTop - contentBottomInset;
    const pageStep = geometry.heightPx + PAGE_GAP;

    // place the body inside each page's content box (overrides the CSS-var padding)
    doc.style.padding = `${contentTop}px ${geometry.margin.right}px ${contentBottomInset}px ${geometry.margin.left}px`;

    // clear the previous page-top markers so they don't skew this measurement
    for (const el of Array.from(doc.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");

    // delta of offsetTop captures each block's height plus its collapsed inter-block margin
    const kids = Array.from(doc.children).filter((c) => !c.classList.contains("docxedit-pagespacer")) as HTMLElement[];
    const tops = kids.map((k) => k.offsetTop);
    const heights = kids.map((k, i) => (i < kids.length - 1 ? tops[i + 1]! - tops[i]! : k.offsetHeight));

    // honor explicit page breaks. A manual break renders either as its own marker element
    // (break before the next block) or inside a block (break before that block).
    const forceBreakBefore = new Set<number>();
    const isManualMarker = (el: Element) =>
      el.classList.contains("docx-pagebreak") && el.getAttribute("data-docx-pagebreak") === "manual";
    kids.forEach((k, i) => {
      if (isManualMarker(k)) {
        if (i + 1 < kids.length) forceBreakBefore.add(i + 1);
      } else if (i > 0 && k.querySelector('.docx-pagebreak[data-docx-pagebreak="manual"]')) {
        forceBreakBefore.add(i);
      }
    });

    const { spacerBefore, cardCount } = paginate(heights, { pageStep, contentHeight }, forceBreakBefore);

    for (const [idx, h] of spacerBefore) {
      const sp = document.createElement("div");
      sp.className = "docxedit-pagespacer";
      sp.contentEditable = "false";
      sp.setAttribute("aria-hidden", "true");
      sp.style.height = `${h}px`;
      doc.insertBefore(sp, kids[idx]!);
      // drop the page-starting block's top margin so it aligns to the page content top
      kids[idx]!.classList.add("docxedit-pagetop");
    }

    // read-only clone of the canonical band, click to edit (its own padding gives the margins)
    const mkClone = (src: HTMLElement, topPx: number): HTMLElement => {
      const c = src.cloneNode(true) as HTMLElement;
      c.removeAttribute("contenteditable");
      c.removeAttribute("role");
      c.removeAttribute("aria-label");
      c.classList.add("docxedit-hf-clone");
      c.style.cssText = `top:${topPx}px;left:0;width:100%`;
      c.title = t("editHeaderFooter");
      c.addEventListener("mousedown", (e) => {
        e.preventDefault();
        editBand(src, topPx);
      });
      return c;
    };

    for (let p = 0; p < cardCount; p++) {
      const base = p * pageStep;
      const card = document.createElement("div");
      card.className = "docxedit-pagecard";
      card.style.top = `${base}px`;
      card.style.height = `${geometry.heightPx}px`;
      pagelayer.appendChild(card);
      const num = document.createElement("div");
      num.className = "docxedit-pagenum";
      num.textContent = `${p + 1} / ${cardCount}`;
      num.style.top = `${base + geometry.heightPx - 22}px`;
      pagelayer.appendChild(num);
      if (header) hflayer.appendChild(mkClone(header, base + geometry.margin.top));
      if (footer) hflayer.appendChild(mkClone(footer, base + geometry.heightPx - contentBottomInset));
    }

    page.style.minHeight = `${cardCount * pageStep - PAGE_GAP}px`;
  };

  // Body HTML for saving: the live doc minus pagination artifacts (inert spacers and the
  // transient page-top class the engine adds for alignment).
  const cleanBody = (): string => {
    if (!doc.querySelector(".docxedit-pagespacer, .docxedit-pagetop")) return doc.innerHTML;
    const tmp = doc.cloneNode(true) as HTMLElement;
    for (const s of Array.from(tmp.querySelectorAll(".docxedit-pagespacer"))) s.remove();
    for (const el of Array.from(tmp.querySelectorAll(".docxedit-pagetop"))) el.classList.remove("docxedit-pagetop");
    return tmp.innerHTML;
  };

  const reflow = () => {
    if (editingBand) return; // don't yank the band currently being edited
    repaginate();
    positionCards();
  };
  let reflowTimer = 0;
  const scheduleReflow = () => {
    window.clearTimeout(reflowTimer);
    reflowTimer = window.setTimeout(reflow, 150);
  };

  for (const thread of parts.comments) addThreadCard(thread);
  // Clicking commented (highlighted) text opens its thread; the inline icon is gone.
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement).closest?.(".docxedit-react-pop")) return;
    document.querySelector(".docxedit-react-pop")?.remove();
    const hit = (e.target as HTMLElement).closest?.(".docx-comment, .docx-comment-ref") as HTMLElement | null;
    if (hit) {
      const id = hit.getAttribute("data-comment-id") ?? "";
      const threadId = threadOf.get(id) ?? id;
      setActiveComment(threadId);
      cmtPanel.querySelector(`[data-comment-id="${CSS.escape(threadId)}"]`)?.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  });
  // Lay out pages + comment cards now and once the layout settles (rAF is throttled in
  // background tabs, so also use a timeout), and again whenever heights can shift: after
  // fonts and images load, on container width change, and (debounced) on every edit.
  reflow();
  requestAnimationFrame(reflow);
  setTimeout(reflow, 150);
  if (document.fonts?.ready) document.fonts.ready.then(reflow).catch(() => {});
  for (const img of Array.from(doc.querySelectorAll("img"))) img.addEventListener("load", reflow);
  let lastWidth = scroll.clientWidth;
  const repositionObserver = new ResizeObserver(() => {
    const w = scroll.clientWidth;
    if (w !== lastWidth) {
      lastWidth = w;
      reflow();
    } else {
      positionCards();
    }
  });
  repositionObserver.observe(scroll);
  for (const r of regions) r.addEventListener("input", scheduleReflow);

  // Image select: click an image to select it, drag the corner handle to resize, and use
  // the delete button (top-right) or the Delete/Backspace key to remove it.
  let selImg: HTMLImageElement | null = null;
  const imgHandle = document.createElement("div");
  imgHandle.className = "docxedit-img-handle is-hidden";
  const imgDel = document.createElement("button");
  imgDel.type = "button";
  imgDel.className = "docxedit-img-del is-hidden";
  imgDel.textContent = "✕";
  imgDel.title = t("deleteImage");
  wrap.append(imgHandle, imgDel);
  const placeHandle = () => {
    if (!selImg || !wrap.contains(selImg)) {
      imgHandle.classList.add("is-hidden");
      imgDel.classList.add("is-hidden");
      return;
    }
    const ir = selImg.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    // positions are runtime-computed; they must follow the selected image
    imgHandle.style.left = `${ir.right - wr.left - 6}px`;
    imgHandle.style.top = `${ir.bottom - wr.top - 6}px`;
    imgHandle.classList.remove("is-hidden");
    imgDel.style.left = `${ir.right - wr.left - 11}px`;
    imgDel.style.top = `${ir.top - wr.top - 11}px`;
    imgDel.classList.remove("is-hidden");
  };
  const selectImg = (img: HTMLImageElement | null) => {
    if (selImg) selImg.classList.remove("sel");
    selImg = img;
    if (selImg) selImg.classList.add("sel");
    placeHandle();
  };
  const deleteSelImg = () => {
    if (!selImg) return;
    selImg.remove();
    selImg = null;
    placeHandle();
    mark();
  };
  imgDel.addEventListener("mousedown", (e) => e.preventDefault());
  imgDel.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteSelImg();
  });
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement) === imgDel) return;
    const img = (e.target as HTMLElement).closest?.("img") as HTMLImageElement | null;
    selectImg(img && wrap.contains(img) ? img : null);
  });
  for (const r of regions)
    r.addEventListener("keydown", (e) => {
      if (selImg && (e.key === "Delete" || e.key === "Backspace")) {
        e.preventDefault();
        deleteSelImg();
      }
    });
  scroll.addEventListener("scroll", placeHandle);
  // Persist a resize: new images use the width/height attributes; existing ones carry the
  // original drawing in data-docx-xml, so update its extent (EMU) to the new size.
  const persistImgSize = (img: HTMLImageElement) => {
    const xml = img.getAttribute("data-docx-xml");
    const w = Number(img.getAttribute("width")) || 0;
    const h = Number(img.getAttribute("height")) || 0;
    if (!xml || !w) return;
    try {
      const frag = new DOMParser().parseFromString(xml, "application/xml");
      for (const tag of ["wp:extent", "a:ext"]) {
        const el = frag.getElementsByTagName(tag)[0];
        if (el) {
          el.setAttribute("cx", String(Math.round(w * 9525)));
          el.setAttribute("cy", String(Math.round(h * 9525)));
        }
      }
      img.setAttribute("data-docx-xml", new XMLSerializer().serializeToString(frag.documentElement!));
    } catch {
      /* leave as-is */
    }
  };
  imgHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    const img = selImg;
    if (!img) return;
    const startX = e.clientX;
    const rect = img.getBoundingClientRect();
    const startW = rect.width;
    const ratio = rect.height / startW || 1;
    const onMove = (ev: MouseEvent) => {
      const w = Math.max(16, Math.round(startW + (ev.clientX - startX)));
      img.setAttribute("width", String(w));
      img.setAttribute("height", String(Math.round(w * ratio)));
      placeHandle();
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      persistImgSize(img);
      mark();
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
  // Emit inline CSS (text-align, font-weight, ...) the serializer reads back, not legacy tags.
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch {
    /* not supported; legacy tags still round-trip */
  }

  const mark = () => {
    dirty = true;
    options.onChange?.();
    scheduleReflow(); // content changed: re-paginate (debounced)
  };
  for (const r of regions) {
    r.addEventListener("input", mark);
    r.addEventListener("focusin", () => {
      activeEl = r;
    });
  }
  const exec = (cmd: string, val?: string) => {
    activeEl.focus();
    document.execCommand(cmd, false, val);
    mark();
  };
  // Wrap the current selection in a span carrying one CSS property (for font size, which
  // has no execCommand equivalent in CSS mode). No-op on a collapsed selection.
  const styleSel = (prop: string, val: string) => {
    activeEl.focus();
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

  const block = document.createElement("select");
  block.title = t("paragraphStyle");
  block.setAttribute("aria-label", t("paragraphStyle"));
  for (const [v, key] of [["P", "styleParagraph"], ["H1", "styleH1"], ["H2", "styleH2"], ["H3", "styleH3"]] as const) {
    block.add(new Option(t(key), v));
  }
  block.addEventListener("mousedown", () => activeEl.focus());
  block.addEventListener("change", () => exec("formatBlock", block.value));

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
    s.addEventListener("mousedown", () => activeEl.focus());
    s.addEventListener("change", () => {
      if (s.value) fn(s.value);
      s.selectedIndex = 0;
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
  colorInput.addEventListener("mousedown", () => activeEl.focus());
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
  bgInput.addEventListener("mousedown", () => activeEl.focus());
  bgInput.addEventListener("input", () => {
    beginFormatChange();
    exec("hiliteColor", bgInput.value);
  });
  const bgClear = btn("⌫", t("none"), () => exec("hiliteColor", "transparent"), "docxedit-bg-clear");
  bgWrap.append(bgInput, bgClear);

  const FONTS = ["Arial", "Calibri", "Century", "Courier New", "Georgia", "Times New Roman", "Verdana"];
  const fontSel = pickerSelect(t("font"), FONTS.map((f) => [f, f] as [string, string]), (v) => {
    beginFormatChange();
    exec("fontName", v);
  });

  const SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48"];
  const sizeSel = pickerSelect(t("size"), SIZES.map((s) => [s, s] as [string, string]), (v) => {
    beginFormatChange();
    styleSel("fontSize", `${v}pt`);
  });

  const insertPageBreak = () => {
    activeEl.focus();
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
  const insertImage = () => {
    const sel = window.getSelection();
    const savedRange = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp";
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const buf = new Uint8Array(await file.arrayBuffer());
      const dataUrl = `data:${file.type};base64,${bytesToBase64(buf)}`;
      const probe = new Image();
      probe.onload = () => {
        const maxW = 600;
        let w = probe.naturalWidth || 200;
        let h = probe.naturalHeight || 200;
        if (w > maxW) {
          h = Math.round((h * maxW) / w);
          w = maxW;
        }
        const img = document.createElement("img");
        img.src = dataUrl;
        img.setAttribute("width", String(w));
        img.setAttribute("height", String(h));
        const range = savedRange ?? (() => {
          const r = document.createRange();
          r.selectNodeContents(activeEl);
          r.collapse(false);
          return r;
        })();
        range.collapse(false);
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
        const s2 = window.getSelection();
        if (s2) {
          s2.removeAllRanges();
          s2.addRange(range);
        }
        mark();
      };
      probe.src = dataUrl;
    });
    fileInput.click();
  };
  const imgIcon =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" aria-hidden="true">' +
    '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3" fill="currentColor" stroke="none"/>' +
    '<path d="M2 12l3.5-4 2.5 2.5L11 7l3 4"/></svg>';

  // Add a comment over the current selection: wrap it in comment-range markers and a
  // reference marker that carries the text, so the serializer can build comments.xml.
  const addComment = () => {
    activeEl.focus();
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
    const id = String(nextCommentId++);
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
      mark();
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
      mark();
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
    activeEl.focus();
  });

  // Toolbar: shared controls always shown; image/comment/page-break/track-changes are
  // gated by the adapter's capabilities so a format can hide what it cannot serialize.
  const linkBtn = btn(t("link"), t("linkAria"), () => {
    const url = prompt(t("linkPrompt"), "https://");
    if (url === null) return;
    if (url === "") exec("unlink");
    else exec("createLink", url);
  });
  const items: (Node | null)[] = [
    btn("B", t("bold"), () => { beginFormatChange(); exec("bold"); }, "docxedit-tb-bold"),
    btn("I", t("italic"), () => { beginFormatChange(); exec("italic"); }, "docxedit-tb-italic"),
    btn("U", t("underline"), () => { beginFormatChange(); exec("underline"); }, "docxedit-tb-underline"),
    caps.textColor ? colorInput : null,
    caps.textColor ? bgWrap : null,
    sep(),
    block,
    caps.fontControls ? fontSel : null,
    caps.fontControls ? sizeSel : null,
    sep(),
    btn(t("bulletedLabel"), t("bulleted"), () => exec("insertUnorderedList")),
    btn(t("numberedLabel"), t("numbered"), () => exec("insertOrderedList")),
    caps.alignment ? sep() : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 8], [2, 11]]), t("alignLeft"), () => exec("justifyLeft")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [4, 8], [3, 10]]), t("alignCenter"), () => exec("justifyCenter")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [6, 8], [3, 11]]), t("alignRight"), () => exec("justifyRight")) : null,
    caps.alignment ? iconBtn(alignIcon([[2, 12], [2, 12], [2, 12]]), t("alignJustify"), () => exec("justifyFull")) : null,
    sep(),
    caps.images ? iconBtn(imgIcon, t("insertImage"), insertImage) : null,
    caps.comments ? iconBtn(cmtIcon, t("addComment"), addComment) : null,
    caps.pageBreak ? iconBtn(pbIcon, t("insertPageBreak"), insertPageBreak) : null,
    linkBtn,
    caps.trackChanges ? sep() : null,
    caps.trackChanges ? suggestBtn : null,
    caps.trackChanges ? btn("✓", t("acceptAll"), () => resolveAll(true)) : null,
    caps.trackChanges ? btn("✕", t("rejectAll"), () => resolveAll(false)) : null,
  ];
  // Overflow menu: the toolbar is a single row; items that do not fit move into a "…"
  // popover so nothing is lost on narrow widths. The popover lives inside the toolbar so
  // the toolbar's button/sep styling (descendant selectors) still applies to pocketed items.
  const toolbarItems = items.filter((n): n is HTMLElement => n != null);
  const moreBtn = document.createElement("button");
  moreBtn.type = "button";
  moreBtn.className = "docxedit-tb-more";
  moreBtn.textContent = "⋯";
  moreBtn.title = t("moreTools");
  moreBtn.setAttribute("aria-label", t("moreTools"));
  // The popover carries the toolbar class so pocketed items keep their styling, and lives
  // in wrap (not the toolbar) so the toolbar's overflow:hidden does not clip it.
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
  };
  document.addEventListener("click", closeOverflow);
  toolbar.append(...toolbarItems, moreBtn);
  wrap.appendChild(overflow);

  const layoutToolbar = () => {
    overflow.hidden = true;
    for (const it of toolbarItems) toolbar.insertBefore(it, moreBtn); // pull everything back in
    moreBtn.style.display = "none";
    if (toolbar.scrollWidth <= toolbar.clientWidth + 1) return; // it all fits
    moreBtn.style.display = "";
    for (let i = toolbarItems.length - 1; i >= 0; i--) {
      if (toolbar.scrollWidth <= toolbar.clientWidth + 1) break;
      overflow.insertBefore(toolbarItems[i], overflow.firstChild); // pocket trailing items, in order
    }
  };
  layoutToolbar();
  requestAnimationFrame(layoutToolbar);
  setTimeout(layoutToolbar, 150);
  const toolbarObserver = new ResizeObserver(() => layoutToolbar());
  toolbarObserver.observe(toolbar);

  return {
    isDirty() {
      return dirty;
    },
    async getBytes() {
      if (!dirty) return original.slice();
      const editedParts: { path: string; html: string }[] = [];
      if (header && parts.headerPath) editedParts.push({ path: parts.headerPath, html: header.innerHTML });
      if (footer && parts.footerPath) editedParts.push({ path: parts.footerPath, html: footer.innerHTML });
      return adapter.write(cleanBody(), editedParts, {
        reactions: pendingReactions.map((r) => ({ ...r, date: options.now || new Date().toISOString() })),
        replies: pendingReplies,
        done: pendingDone,
        deletedComments,
      });
    },
    destroy() {
      for (const u of fontUrls) URL.revokeObjectURL(u);
      window.clearTimeout(reflowTimer);
      repositionObserver.disconnect();
      toolbarObserver.disconnect();
      document.removeEventListener("click", closeOverflow);
      wrap.remove();
    },
  };
}
