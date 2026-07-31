import { afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { createDocxEditor } from "../adapters/docx/index";
import { BID } from "./feature/block-ids";
import type { BlockChanges, BlockPosition, RichEditor } from "./types";

// What a collaboration host needs from this editor, and nothing about how it is drawn.
//
// The three things a session cannot work without: name a block and have that name still
// mean the same block later, describe an edit as the blocks it touched, and give up undo
// so a peer's work is not taken back by someone else's Ctrl+Z. Each is asserted here
// against the real editor, because each has to hold through pagination, history and save.

beforeAll(() => {
  if (!(globalThis as unknown as { ResizeObserver?: unknown }).ResizeObserver) {
    (globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }
  const zeroRect = () =>
    ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  if (!Range.prototype.getBoundingClientRect) Range.prototype.getBoundingClientRect = zeroRect;
  if (!Range.prototype.getClientRects)
    Range.prototype.getClientRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
});

/** A docx with three paragraphs, so "which one changed" is a real question. */
const THREE = zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "_rels/.rels": strToU8("<Relationships/>"),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
      `<w:p><w:r><w:t>First.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Second.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Third.</w:t></w:r></w:p>` +
      `</w:body></w:document>`,
  ),
  "word/_rels/document.xml.rels": strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  ),
});

interface Mounted {
  editor: RichEditor;
  host: HTMLElement;
  body: HTMLElement;
  reported: BlockChanges[];
}

function mount(): Mounted {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const reported: BlockChanges[] = [];
  const editor = createDocxEditor(host, THREE, {
    paginated: false, // pagination is a view concern and has its own tests
  });
  editor.setBlockReporter((c) => reported.push(c));
  const body = host.querySelector(".docxedit-doc") as HTMLElement;
  return { editor, host, body, reported };
}

/** Type into a block the way the editor's own plumbing sees it: change it, then fire input. */
function edit(body: HTMLElement, index: number, text: string): void {
  const block = body.children[index] as HTMLElement;
  block.textContent = text;
  body.dispatchEvent(new Event("input", { bubbles: true }));
}

describe("the collaboration API", () => {
  it("gives every block an id, and keeps it across edits", () => {
    const { body, editor } = mount();
    const ids = editor.blockSnapshot().map((b) => b.id);
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size, "distinct").toBe(3);

    edit(body, 1, "Second, edited.");
    const after = editor.blockSnapshot().map((b) => b.id);
    // The point of the whole mechanism: an edit must not re-identify the document.
    expect(after).toEqual(ids);
  });

  it("reports the block that changed, not the document", () => {
    const { body, editor, reported } = mount();
    const second = editor.blockSnapshot()[1].id;

    edit(body, 1, "Second, edited.");

    expect(reported.length).toBeGreaterThan(0);
    const last = reported[reported.length - 1];
    expect(last.changed.map((b) => b.id), "one block, the one touched").toEqual([second]);
    expect(last.changed[0].html).toContain("Second, edited.");
    expect(last.removed).toEqual([]);
    expect(last.order).toHaveLength(3);
  });

  // Reporting is a subscription because it is not free: it walks every block on every
  // change. A document nobody is sharing must not pay for it.
  it("reports nothing until someone subscribes, and stops when they leave", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const reported: BlockChanges[] = [];
    const editor = createDocxEditor(host, THREE, { paginated: false });
    const body = host.querySelector(".docxedit-doc") as HTMLElement;

    edit(body, 0, "Edited before anyone asked.");
    expect(reported, "nobody listening").toEqual([]);

    editor.setBlockReporter((c) => reported.push(c));
    edit(body, 1, "Edited while listening.");
    expect(reported).toHaveLength(1);
    expect(reported[0].changed, "an edit, not the whole document").toHaveLength(1);

    editor.setBlockReporter(null);
    edit(body, 2, "Edited after they left.");
    expect(reported).toHaveLength(1);
  });

  it("reports a new block as added, leaving its neighbours alone", () => {
    const { body, editor, reported } = mount();
    const before = editor.blockSnapshot().map((b) => b.id);

    const p = document.createElement("p");
    p.textContent = "Fourth.";
    body.appendChild(p);
    body.dispatchEvent(new Event("input", { bubbles: true }));

    const last = reported[reported.length - 1];
    expect(last.changed).toHaveLength(1);
    expect(last.changed[0].html).toContain("Fourth.");
    expect(last.order).toHaveLength(4);
    expect(last.order.slice(0, 3), "the others kept their identity").toEqual(before);
  });

  it("reports a deleted block as removed", () => {
    const { body, editor, reported } = mount();
    const ids = editor.blockSnapshot().map((b) => b.id);

    body.children[1].remove();
    body.dispatchEvent(new Event("input", { bubbles: true }));

    const last = reported[reported.length - 1];
    expect(last.removed).toEqual([ids[1]]);
    expect(last.order).toEqual([ids[0], ids[2]]);
  });

  it("applies a peer's block without reporting it back as a local change", () => {
    const { body, editor, reported } = mount();
    const snapshot = editor.blockSnapshot();
    const before = reported.length;

    editor.applyRemoteBlocks({
      changed: [{ id: snapshot[2].id, html: `<p ${BID}="${snapshot[2].id}">Third, from a peer.</p>` }],
      removed: [],
      order: snapshot.map((b) => b.id),
    });

    expect(body.textContent).toContain("Third, from a peer.");
    expect(reported.length, "applying is not editing").toBe(before);
  });

  // The next local edit must describe only what this peer did. Diffing against a state
  // from before the peer's change would re-send their block as ours.
  it("does not re-report a peer's block on the next local edit", () => {
    const { body, editor, reported } = mount();
    const snapshot = editor.blockSnapshot();

    editor.applyRemoteBlocks({
      changed: [{ id: snapshot[0].id, html: `<p ${BID}="${snapshot[0].id}">First, from a peer.</p>` }],
      removed: [],
      order: snapshot.map((b) => b.id),
    });
    edit(body, 2, "Third, mine.");

    const last = reported[reported.length - 1];
    expect(last.changed.map((b) => b.id)).toEqual([snapshot[2].id]);
  });

  it("takes a peer's insertion and puts the body in their order", () => {
    const { body, editor } = mount();
    const snapshot = editor.blockSnapshot();
    const fresh = "b-new";

    editor.applyRemoteBlocks({
      changed: [{ id: fresh, html: `<p ${BID}="${fresh}">Inserted second.</p>` }],
      removed: [],
      order: [snapshot[0].id, fresh, snapshot[1].id, snapshot[2].id],
    });

    expect([...body.children].map((c) => c.textContent)).toEqual([
      "First.",
      "Inserted second.",
      "Second.",
      "Third.",
    ]);
    expect(editor.blockSnapshot().map((b) => b.id)).toEqual([
      snapshot[0].id,
      fresh,
      snapshot[1].id,
      snapshot[2].id,
    ]);
  });

  // Found in a browser, not here.
  //
  // Replacing a block's element breaks any selection inside it, and Chrome additionally
  // collapses the selection to the start of the whole body even when the caret was in a
  // different block. jsdom reproduces the first and not the second, so this covers the case
  // it can fail on: the caret sits in the very paragraph the peer edited. Without the
  // restore, every edit a peer made threw this person's cursor to the top of the document.
  it("leaves this person's caret in place when a peer edits the block it is in", () => {
    const { body, editor } = mount();
    const ids = editor.blockSnapshot().map((b) => b.id);

    const second = body.children[1] as HTMLElement;
    const range = document.createRange();
    range.setStart(second.firstChild!, 3);
    range.collapse(true);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);

    editor.applyRemoteBlocks({
      changed: [{ id: ids[1], html: `<p ${BID}="${ids[1]}">Second, rewritten by a peer.</p>` }],
      removed: [],
      order: ids,
    });

    const after = window.getSelection()!;
    expect(after.rangeCount, "there is still a caret").toBe(1);
    const block = after.getRangeAt(0).startContainer.parentElement?.closest(`[${BID}]`);
    expect(block?.getAttribute(BID), "in the block it was in").toBe(ids[1]);
    expect(after.getRangeAt(0).startOffset, "at the offset it was at").toBe(3);
  });

  it("removes a block a peer deleted", () => {
    const { body, editor } = mount();
    const snapshot = editor.blockSnapshot();

    editor.applyRemoteBlocks({
      changed: [],
      removed: [snapshot[1].id],
      order: [snapshot[0].id, snapshot[2].id],
    });

    expect([...body.children].map((c) => c.textContent)).toEqual(["First.", "Third."]);
  });

  describe("undo", () => {
    it("hands undo and redo to a host that asks for them", () => {
      const { editor } = mount();
      const calls: string[] = [];
      editor.setUndoHandler({
        undo: () => calls.push("undo"),
        redo: () => calls.push("redo"),
        canUndo: () => true,
        canRedo: () => false,
      });

      editor.undo();
      editor.redo();
      expect(calls).toEqual(["undo", "redo"]);
      expect(editor.canUndo()).toBe(true);
      expect(editor.canRedo(), "the host's answer, not the snapshot stack's").toBe(false);
    });

    it("routes the keyboard shortcut through the host too", () => {
      const { body, editor } = mount();
      const calls: string[] = [];
      editor.setUndoHandler({
        undo: () => calls.push("undo"),
        redo: () => calls.push("redo"),
        canUndo: () => true,
        canRedo: () => true,
      });

      const press = (shift: boolean): void => {
        body.dispatchEvent(
          new KeyboardEvent("keydown", { key: "z", ctrlKey: true, shiftKey: shift, bubbles: true }),
        );
      };
      press(false);
      press(true);
      expect(calls, "Ctrl+Z must not reach the snapshot history").toEqual(["undo", "redo"]);
    });

    it("takes its own history back when the handler is cleared", () => {
      const { body, editor } = mount();
      editor.setUndoHandler({
        undo: () => undefined,
        redo: () => undefined,
        canUndo: () => true,
        canRedo: () => true,
      });
      editor.setUndoHandler(null);

      edit(body, 0, "First, edited.");
      expect(body.textContent).toContain("First, edited.");
      editor.undo();
      expect(body.textContent, "its own stack again").toContain("First.");
    });
  });

  describe("presence", () => {
    // jsdom lays nothing out, so every rect is zero and a caret has nowhere to go. These
    // tests are about which carets are drawn, for whom and in what colour; where exactly
    // they land on screen is layout, and jsdom cannot answer it either way.
    let realRangeRect: typeof Range.prototype.getBoundingClientRect;
    const fakeRect = (top: number) =>
      ({ x: 10, y: top, top, left: 10, right: 12, bottom: top + 16, width: 2, height: 16, toJSON() {} }) as DOMRect;
    beforeEach(() => {
      realRangeRect = Range.prototype.getBoundingClientRect;
      Range.prototype.getBoundingClientRect = function () {
        return fakeRect(20);
      };
    });
    afterEach(() => {
      Range.prototype.getBoundingClientRect = realRangeRect;
    });

    /** Put this document's caret at `offset` characters into block `index`. */
    function putCaret(body: HTMLElement, index: number, offset: number): void {
      const block = body.children[index] as HTMLElement;
      const text = block.firstChild ?? block;
      const range = document.createRange();
      if (text.nodeType === 3) range.setStart(text, offset);
      else range.setStart(block, 0);
      range.collapse(true);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    }

    it("reports the caret as a block and an offset, not a position in the document", () => {
      const { body, editor } = mount();
      const seen: (BlockPosition | null)[] = [];
      editor.setSelectionReporter((at) => seen.push(at));

      putCaret(body, 1, 4);

      const last = seen[seen.length - 1];
      expect(last?.blockId, "the block it is in").toBe(editor.blockSnapshot()[1].id);
      expect(last?.offset, "and how far into it").toBe(4);
    });

    // The offset has to be relative to the block. A document-wide one would move every time
    // anyone typed above it, so two peers would disagree about where each other are.
    it("counts the offset from the start of its own block", () => {
      const { body, editor } = mount();
      const seen: (BlockPosition | null)[] = [];
      editor.setSelectionReporter((at) => seen.push(at));

      putCaret(body, 2, 3);
      expect(seen[seen.length - 1]?.offset).toBe(3);
      expect(seen[seen.length - 1]?.blockId).toBe(editor.blockSnapshot()[2].id);
    });

    it("stops reporting when the session lets go", () => {
      const { body, editor } = mount();
      const seen: (BlockPosition | null)[] = [];
      editor.setSelectionReporter((at) => seen.push(at));
      putCaret(body, 0, 1);
      const count = seen.length;

      editor.setSelectionReporter(null);
      putCaret(body, 1, 1);
      expect(seen.length).toBe(count);
    });

    it("draws a caret per peer, in that peer's colour", () => {
      const { host, editor } = mount();
      const ids = editor.blockSnapshot().map((b) => b.id);

      editor.setPeerCarets([
        { id: "p1", name: "Ada", colour: "rgb(255, 0, 0)", blockId: ids[0], offset: 2 },
        { id: "p2", name: "Bo", colour: "rgb(0, 0, 255)", blockId: ids[2], offset: 1 },
      ]);

      const carets = host.querySelectorAll(".docxedit-caret");
      expect(carets).toHaveLength(2);
      expect([...host.querySelectorAll(".docxedit-caret-name")].map((n) => n.textContent)).toEqual([
        "Ada",
        "Bo",
      ]);
      expect((carets[0] as HTMLElement).style.background).toBe("rgb(255, 0, 0)");
      expect((carets[1] as HTMLElement).style.background).toBe("rgb(0, 0, 255)");
    });

    it("replaces the whole set, so someone who left stops being drawn", () => {
      const { host, editor } = mount();
      const ids = editor.blockSnapshot().map((b) => b.id);
      editor.setPeerCarets([{ id: "p1", name: "Ada", colour: "#f00", blockId: ids[0], offset: 0 }]);
      expect(host.querySelectorAll(".docxedit-caret")).toHaveLength(1);

      editor.setPeerCarets([]);
      expect(host.querySelectorAll(".docxedit-caret")).toHaveLength(0);
    });

    it("draws nothing for a peer in a block this document does not have", () => {
      const { host, editor } = mount();
      editor.setPeerCarets([{ id: "p1", name: "Ada", colour: "#f00", blockId: "not-here", offset: 0 }]);
      expect(host.querySelectorAll(".docxedit-caret")).toHaveLength(0);
    });

    // The whole reason for an overlay. A peer's caret in the body would be in the undo
    // history, in the saved file, and in the next diff as a change nobody made.
    it("keeps peer carets out of the document itself", async () => {
      const { body, editor, reported } = mount();
      const ids = editor.blockSnapshot().map((b) => b.id);
      const before = reported.length;

      editor.setPeerCarets([{ id: "p1", name: "Ada", colour: "#f00", blockId: ids[0], offset: 2 }]);

      expect(body.innerHTML, "not in the body").not.toContain("docxedit-caret");
      expect(reported.length, "and not a change").toBe(before);
      const bytes = await editor.getBytes();
      expect(new TextDecoder().decode(unzipSync(bytes)["word/document.xml"])).not.toContain("Ada");
    });
  });

  // End to end, which is the guarantee that matters. It does not pin the strip itself:
  // both adapters rebuild their XML from the HTML and drop unknown attributes, so this
  // still passes with the strip removed. It pins the outcome against whichever of the two
  // is doing the work.
  it("keeps the ids out of the saved file", async () => {
    const { body, editor } = mount();
    edit(body, 0, "First, edited.");
    expect(body.innerHTML, "present in the live document").toContain(BID);

    const bytes = await editor.getBytes();
    const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]);
    expect(xml, "and absent from the file").not.toContain("rdoc-bid");
    expect(xml).toContain("First, edited.");
  });

  // The ids have to survive an undo, or the next diff calls every block new.
  it("keeps block identity across the editor's own undo", () => {
    const { body, editor } = mount();
    const ids = editor.blockSnapshot().map((b) => b.id);

    edit(body, 1, "Second, edited.");
    editor.undo();

    expect(body.textContent).toContain("Second.");
    expect(editor.blockSnapshot().map((b) => b.id), "same blocks, not new ones").toEqual(ids);
  });
});
