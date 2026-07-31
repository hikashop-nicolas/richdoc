import { beforeAll, describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { createDocxEditor } from "../adapters/docx/index";
import { BID } from "./feature/block-ids";
import type { BlockChanges, RichEditor } from "./types";

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
