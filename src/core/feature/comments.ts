// Comments feature: the side panel (one card per thread), reaction pickers, reply boxes,
// resolve/delete, vertical anchoring to the commented range, and the click-to-open binding.
// Owns the comment-edit bookkeeping (new reactions/replies/resolutions/deletions) that the
// adapter writes back on save. Extracted from the engine; talks back via the deps below.
import { t } from "../i18n";
import type { CommentEdits, CommentEntry, CommentThread, Capabilities, EditorOptions } from "../types";

export interface CommentsDeps {
  wrap: HTMLElement;
  cmtPanel: HTMLElement;
  pagebox: HTMLElement;
  options: EditorOptions;
  caps: Capabilities;
  mark: () => void;
}

export function setupComments(deps: CommentsDeps) {
  const { wrap, cmtPanel, pagebox, options, caps, mark } = deps;

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
      const el = r as HTMLElement;
      el.classList.toggle("active", members.includes(el.getAttribute("data-comment-id") ?? ""));
    }
    // Select the commented text so the user sees exactly what the comment refers to. The
    // range is bracketed by the comment's markers (highlight, range marks, or reference),
    // in document order, so it works whether or not the format renders a highlight span.
    if (threadId) {
      const marks = Array.from(wrap.querySelectorAll(".docx-comment, .docx-cmark, .docx-comment-ref")).filter((el) =>
        members.includes(el.getAttribute("data-comment-id") ?? ""),
      ) as HTMLElement[];
      if (marks.length) {
        const range = document.createRange();
        range.setStartBefore(marks[0]!);
        range.setEndAfter(marks[marks.length - 1]!);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
        marks[0]!.scrollIntoView({ block: "center", behavior: "smooth" });
      }
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
    cmtPanel.style.height = `${Math.max(prevBottom, pagebox.offsetHeight)}px`;
  };

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

  // Allocate the next comment id (used by the engine when inserting a brand-new comment).
  const allocId = (): string => String(nextCommentId++);
  // The pending comment edits the adapter applies on save.
  const getEdits = (): CommentEdits => ({
    reactions: pendingReactions.map((r) => ({ ...r, date: options.now || new Date().toISOString() })),
    replies: pendingReplies,
    done: pendingDone,
    deletedComments,
  });

  return { addThreadCard, positionCards, setActiveComment, allocId, freshParaId, getEdits };
}
