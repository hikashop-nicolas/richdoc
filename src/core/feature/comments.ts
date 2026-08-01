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
  const pendingEdited: { id: string; text: string }[] = [];
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

  /**
   * Every comment item on screen, by id: its entry, its element, and its reaction row.
   *
   * A peer names a comment; without this the name reaches the file and nothing on screen,
   * which is exactly the gap that had a reply appear only after a reload.
   */
  const itemsById = new Map<string, { entry: CommentEntry; item: HTMLElement; row: HTMLElement | null; text: HTMLElement }>();

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
    // Edit in place: swap the text for a textarea; saving records the rewrite
    // for the adapter and updates the card.
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "docxedit-cmt-editbtn";
    edit.textContent = "✎";
    edit.title = t("editComment");
    edit.setAttribute("aria-label", t("editComment"));
    edit.addEventListener("click", (e) => {
      e.stopPropagation();
      if (item.querySelector(".docxedit-cmt-editinput")) return;
      const ta = document.createElement("textarea");
      ta.className = "docxedit-cmt-replyinput docxedit-cmt-editinput";
      ta.rows = 3;
      ta.value = entry.text;
      ta.addEventListener("click", (ev) => ev.stopPropagation());
      const save = document.createElement("button");
      save.type = "button";
      save.className = "docxedit-cmt-send";
      save.textContent = t("send");
      const commit = () => {
        const txt = ta.value.trim();
        ta.remove();
        save.remove();
        text.style.display = "";
        if (!txt || txt === entry.text) return;
        entry.text = txt;
        text.textContent = txt;
        // Last rewrite wins when the same comment is edited twice before a save.
        const prior = pendingEdited.find((p2) => p2.id === entry.id);
        if (prior) prior.text = txt;
        else pendingEdited.push({ id: entry.id, text: txt });
        mark();
        positionCards();
      };
      save.addEventListener("click", (ev) => {
        ev.stopPropagation();
        commit();
      });
      ta.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter" && !ev.shiftKey) {
          ev.preventDefault();
          commit();
        } else if (ev.key === "Escape") {
          ev.preventDefault();
          ta.remove();
          save.remove();
          text.style.display = "";
        }
      });
      text.style.display = "none";
      item.insertBefore(ta, more);
      item.insertBefore(save, more);
      ta.focus();
    });
    item.append(meta, edit, text, more);
    itemsById.set(entry.id, { entry, item, row: null, text });
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
      const rec = itemsById.get(entry.id);
      if (rec) rec.row = row;
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
    edited: pendingEdited,
  });

  /**
   * The pending comment edits as keyed entries, so a session can merge them.
   *
   * Keyed by what each thing IS rather than by its position in a list: two people adding
   * different replies must end up with both, and two people reacting must end up with both
   * reactions. A list would make one of the two disappear.
   */
  const editEntries = (): { kind: string; key: string; value: string }[] => {
    const out: { kind: string; key: string; value: string }[] = [];
    for (const r of pendingReactions) {
      out.push({ kind: "reaction", key: `${r.commentId}|${r.emoji}|${r.person}`, value: JSON.stringify(r) });
    }
    for (const r of pendingReplies) out.push({ kind: "reply", key: r.id, value: JSON.stringify(r) });
    for (const [paraId, value] of pendingDone) out.push({ kind: "done", key: paraId, value: String(value) });
    for (const id of deletedComments) out.push({ kind: "deleted", key: id, value: "1" });
    for (const e of pendingEdited) out.push({ kind: "edited", key: e.id, value: e.text });
    return out;
  };

  /**
   * Merge a peer's comment edits into ours, so a save from either side carries both.
   *
   * The cards on screen are NOT rebuilt: they are drawn once at mount from the document's
   * own comments, and there is no path here to add a peer's reply to the panel. So a reply
   * from someone else reaches the file and is seen after a reload, not before. That is a
   * real limit and is written down rather than left to be discovered.
   */
  /** The card for a thread, found by the paragraph id a peer names it with. */
  const cardOfPara = (paraId: string): HTMLElement | null =>
    (Array.from(cmtPanel.children) as HTMLElement[]).find((c) => c.dataset.paraId === paraId) ?? null;

  /**
   * Put a peer's comment edit on screen, mirroring what the local action does to the card.
   *
   * Each of these updates the pending state and the DOM together rather than rebuilding the
   * panel: a rebuild would throw away a reply someone is halfway through typing, and an
   * open reply box is the most likely thing to be on screen when a peer's edit arrives.
   */
  const showRemote = (kind: string, key: string, value: string): void => {
    if (kind === "reply") {
      const r = JSON.parse(value) as (typeof pendingReplies)[number];
      const card = cardOfPara(r.parentParaId);
      if (!card) return; // a thread this peer does not have; the file still gets the reply
      const threadId = card.dataset.commentId ?? "";
      const members = threadMembers.get(threadId) ?? [threadId];
      if (!members.includes(r.id)) members.push(r.id);
      registerThread(threadId, members);
      const box = card.querySelector(".docxedit-cmt-replybox");
      const item = buildItem(
        { id: r.id, author: r.author, date: r.date, text: r.text, reactions: [], paraId: r.paraId },
        true,
      );
      card.insertBefore(item, box);
      return;
    }
    if (kind === "reaction") {
      const r = JSON.parse(value) as { commentId: string; emoji: string; person: string };
      const rec = itemsById.get(r.commentId);
      if (!rec) return;
      const existing = rec.entry.reactions.find((x) => x.emoji === r.emoji);
      if (existing) {
        if (!existing.people.includes(r.person)) existing.people.push(r.person);
      } else rec.entry.reactions.push({ emoji: r.emoji, people: [r.person] });
      if (rec.row) renderReactions(rec.row, rec.entry);
      return;
    }
    if (kind === "done") {
      cardOfPara(key)?.classList.toggle("resolved", value === "true");
      return;
    }
    if (kind === "edited") {
      const rec = itemsById.get(key);
      if (!rec) return;
      rec.entry.text = value;
      rec.text.textContent = value;
      return;
    }
    if (kind === "deleted") {
      const rec = itemsById.get(key);
      rec?.item.remove();
      itemsById.delete(key);
      // The highlight in the body goes with it, as the local delete does.
      wrap.querySelectorAll(`.docx-comment[data-comment-id="${CSS.escape(key)}"]`).forEach((span) => {
        while (span.firstChild) span.parentNode?.insertBefore(span.firstChild, span);
        span.remove();
      });
      const card = (Array.from(cmtPanel.children) as HTMLElement[]).find((c) => c.dataset.commentId === key);
      card?.remove();
    }
  };

  const mergeEdits = (entries: { kind: string; key: string; value: string }[]): boolean => {
    let touched = false;
    for (const entry of entries) {
      try {
        if (entry.kind === "reaction") {
          const r = JSON.parse(entry.value) as { commentId: string; emoji: string; person: string };
          if (pendingReactions.some((x) => x.commentId === r.commentId && x.emoji === r.emoji && x.person === r.person)) continue;
          pendingReactions.push(r);
          touched = true;
        } else if (entry.kind === "reply") {
          const r = JSON.parse(entry.value) as (typeof pendingReplies)[number];
          if (pendingReplies.some((x) => x.id === r.id)) continue;
          pendingReplies.push(r);
          touched = true;
        } else if (entry.kind === "done") {
          const want = entry.value === "true";
          if (pendingDone.get(entry.key) === want) continue;
          pendingDone.set(entry.key, want);
          touched = true;
        } else if (entry.kind === "deleted") {
          if (deletedComments.includes(entry.key)) continue;
          deletedComments.push(entry.key);
          touched = true;
        } else if (entry.kind === "edited") {
          const prior = pendingEdited.find((x) => x.id === entry.key);
          if (prior?.text === entry.value) continue;
          if (prior) prior.text = entry.value;
          else pendingEdited.push({ id: entry.key, text: entry.value });
          touched = true;
        } else continue;
        // Only for an entry that was new: the screen follows the state it just changed, so
        // a duplicate delivery cannot append the same reply to the card twice.
        showRemote(entry.kind, entry.key, entry.value);
      } catch {
        /* unreadable entry; skip it rather than lose the rest */
      }
    }
    if (touched) positionCards();
    return touched;
  };

  return { addThreadCard, positionCards, setActiveComment, allocId, freshParaId, getEdits, editEntries, mergeEdits };
}
