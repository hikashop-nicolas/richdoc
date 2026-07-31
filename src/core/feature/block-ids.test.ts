import { describe, expect, it } from "vitest";
import { JSDOM } from "jsdom";
import {
  BID,
  assignBlockIds,
  blockIdOf,
  blockSnapshot,
  changedBlocks,
  stripBlockIds,
  topLevelBlocks,
} from "./block-ids";

// Block identity, which collaboration needs before it can name a paragraph.
//
// The rule that matters most is the last one: these ids must never reach the saved file.
// richdoc's whole claim is that an untouched part of a document comes back byte for byte,
// and a file must not gain an attribute because somebody opened it.

const bodyOf = (html: string): HTMLElement => {
  const dom = new JSDOM(`<div id="doc">${html}</div>`);
  return dom.window.document.getElementById("doc") as unknown as HTMLElement;
};

describe("block ids", () => {
  it("gives every top-level block one", () => {
    const root = bodyOf("<p>one</p><p>two</p><table><tr><td>cell</td></tr></table>");
    expect(assignBlockIds(root)).toBe(3);
    const ids = topLevelBlocks(root).map((b) => b.getAttribute(BID));
    expect(ids.every(Boolean)).toBe(true);
    expect(new Set(ids).size).toBe(3); // all different
  });

  // The important one: an id that changes is no identity at all, and "which block
  // changed" would answer "every one of them" on the next keystroke.
  it("never regenerates an id a block already has", () => {
    const root = bodyOf("<p>one</p><p>two</p>");
    assignBlockIds(root);
    const before = topLevelBlocks(root).map((b) => b.getAttribute(BID));

    root.appendChild(root.ownerDocument.createElement("p"));
    expect(assignBlockIds(root)).toBe(1); // only the new one

    const after = topLevelBlocks(root).map((b) => b.getAttribute(BID));
    expect(after.slice(0, 2)).toEqual(before);
  });

  it("reaches blocks inside pagination wrappers, which are the same paragraphs", () => {
    const root = bodyOf(
      '<div class="docxedit-colpage"><p>a</p><p>b</p></div><div class="docxedit-secpage"><p>c</p></div>',
    );
    assignBlockIds(root);
    expect(topLevelBlocks(root).map((b) => b.textContent)).toEqual(["a", "b", "c"]);
    // The wrappers themselves are not blocks: they are view furniture.
    expect(root.querySelector(".docxedit-colpage")?.hasAttribute(BID)).toBe(false);
  });

  it("skips page spacers, which are not content", () => {
    const root = bodyOf('<p>a</p><div class="docxedit-pagespacer"></div><p>b</p>');
    expect(assignBlockIds(root)).toBe(2);
    expect(root.querySelector(".docxedit-pagespacer")?.hasAttribute(BID)).toBe(false);
  });

  // Repagination moves blocks between wrappers rather than rebuilding them, so the
  // attribute rides along on the same element. This is what makes the ids survive it.
  it("survives a block being moved into a wrapper", () => {
    const root = bodyOf("<p>a</p><p>b</p>");
    assignBlockIds(root);
    const first = topLevelBlocks(root)[0];
    const id = first.getAttribute(BID);

    const page = root.ownerDocument.createElement("div");
    page.className = "docxedit-colpage";
    root.insertBefore(page, root.firstChild);
    page.appendChild(first); // repagination, in miniature

    expect(topLevelBlocks(root)[0].getAttribute(BID)).toBe(id);
    expect(assignBlockIds(root)).toBe(0); // nothing new to name
  });

  it("strips every id, which is what the saved file gets", () => {
    const root = bodyOf('<p>a</p><div class="docxedit-colpage"><p>b</p></div>');
    assignBlockIds(root);
    expect(root.querySelectorAll(`[${BID}]`).length).toBe(2);

    stripBlockIds(root);
    expect(root.querySelectorAll(`[${BID}]`).length).toBe(0);
    expect(root.innerHTML).toBe('<p>a</p><div class="docxedit-colpage"><p>b</p></div>');
  });

  it("finds the block a node sits in", () => {
    const root = bodyOf("<p>outer <b>bold</b></p><p>other</p>");
    assignBlockIds(root);
    const bold = root.querySelector("b")!;
    const first = topLevelBlocks(root)[0];

    expect(blockIdOf(bold.firstChild, root)).toBe(first.getAttribute(BID));
    expect(blockIdOf(root, root)).toBeNull(); // the body is not a block
  });

  it("reports which blocks changed, were added, and were removed", () => {
    const root = bodyOf("<p>one</p><p>two</p><p>three</p>");
    assignBlockIds(root);
    const before = blockSnapshot(root);
    const ids = topLevelBlocks(root).map((b) => b.getAttribute(BID)!);

    topLevelBlocks(root)[1].textContent = "TWO";
    topLevelBlocks(root)[2].remove();
    const added = root.ownerDocument.createElement("p");
    added.textContent = "four";
    root.appendChild(added);
    assignBlockIds(root);

    const diff = changedBlocks(before, blockSnapshot(root));
    expect(diff.changed).toEqual([ids[1]]);
    expect(diff.removed).toEqual([ids[2]]);
    expect(diff.added).toHaveLength(1);
    expect(diff.added[0]).not.toBe(ids[2]); // a new block, not the old one returning
  });

  it("reports nothing when nothing changed", () => {
    const root = bodyOf("<p>one</p><p>two</p>");
    assignBlockIds(root);
    const snap = blockSnapshot(root);
    expect(changedBlocks(snap, blockSnapshot(root))).toEqual({ changed: [], added: [], removed: [] });
  });

  // Moving a block is not a content change. A caller that reorders sees no edits, which
  // is right: the order is carried separately, exactly as it is for subtitle cues.
  it("treats a reorder as no content change", () => {
    const root = bodyOf("<p>one</p><p>two</p>");
    assignBlockIds(root);
    const before = blockSnapshot(root);

    const [first] = topLevelBlocks(root);
    root.appendChild(first); // move it to the end

    expect(changedBlocks(before, blockSnapshot(root)).changed).toEqual([]);
  });
});
