import { BID, topLevelBlocks } from "./block-ids";

// Where the other people in a shared session have their cursors.
//
// A position in this editor is a block id plus a character offset into that block, which is
// the same pair the history module already records to put a caret back after a restore.
// Nothing else would survive: a DOM node is not stable across a repagination, and a
// document-wide offset changes every time anyone types above it.
//
// The carets are drawn in an overlay rather than inserted into the body. Putting another
// person's cursor into the content would make it part of the document: it would be in the
// undo history, in the saved file, and in what the next diff reports as a change. An
// overlay costs a repaint and touches nothing.

export interface PeerCaret {
  /** Stable for the life of a peer's session. */
  id: string;
  name: string;
  colour: string;
  blockId: string;
  /** Characters from the start of the block. */
  offset: number;
}

export interface PeerCaretDeps {
  /** The scrolling element. The overlay lives in it, so carets scroll with the page. */
  scroll: HTMLElement;
  doc: HTMLElement;
}

/** A caret's position and height, in the overlay's coordinates. */
interface Placed {
  peer: PeerCaret;
  left: number;
  top: number;
  height: number;
}

export function setupPeerCarets({ scroll, doc }: PeerCaretDeps) {
  const overlay = document.createElement("div");
  overlay.className = "docxedit-carets";
  overlay.setAttribute("aria-hidden", "true"); // decoration; the peer list is the accessible view
  scroll.appendChild(overlay);

  let peers: PeerCaret[] = [];

  // Compared rather than selected: a block id arrives from another peer, and putting one
  // straight into an attribute selector makes a stray quote their problem to exploit and
  // ours to debug. Walking the blocks is also one pass for the whole set.
  const blocksById = (): Map<string, HTMLElement> => {
    const out = new Map<string, HTMLElement>();
    for (const block of topLevelBlocks(doc)) {
      const id = block.getAttribute(BID);
      if (id) out.set(id, block);
    }
    return out;
  };

  /**
   * A collapsed range at `offset` characters into `block`.
   *
   * The same walk placeCaret does, but it yields a range to measure rather than moving the
   * selection: drawing where someone else is must never move where this person is.
   */
  function rangeAt(block: HTMLElement, offset: number): Range | null {
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    let acc = 0;
    let last: Node | null = null;
    while ((node = walker.nextNode())) {
      const len = (node.textContent ?? "").length;
      if (acc + len >= offset) {
        last = node;
        break;
      }
      acc += len;
      last = node;
    }
    const range = document.createRange();
    if (last) range.setStart(last, Math.min((last.textContent ?? "").length, Math.max(0, offset - acc)));
    else range.setStart(block, 0);
    range.collapse(true);
    return range;
  }

  function place(peer: PeerCaret, blocks: Map<string, HTMLElement>): Placed | null {
    const block = blocks.get(peer.blockId);
    if (!block) return null; // a block this peer is in that we have not got: nothing to draw
    const range = rangeAt(block, peer.offset);
    if (!range) return null;

    // An empty paragraph gives a collapsed range with no height, and so does a range at the
    // very start of one. The block's own box is the honest fallback.
    let rect = range.getBoundingClientRect();
    if (!rect.height) rect = block.getBoundingClientRect();
    if (!rect.height) return null;

    const base = scroll.getBoundingClientRect();
    return {
      peer,
      left: rect.left - base.left + scroll.scrollLeft,
      top: rect.top - base.top + scroll.scrollTop,
      height: rect.height,
    };
  }

  function render(): void {
    overlay.textContent = "";
    if (!peers.length) return;

    const blocks = blocksById();
    for (const peer of peers) {
      const at = place(peer, blocks);
      if (!at) continue;

      const caret = document.createElement("div");
      caret.className = "docxedit-caret";
      caret.style.left = `${at.left}px`;
      caret.style.top = `${at.top}px`;
      caret.style.height = `${at.height}px`;
      caret.style.background = peer.colour;

      const label = document.createElement("span");
      label.className = "docxedit-caret-name";
      label.textContent = peer.name;
      label.style.background = peer.colour;
      caret.appendChild(label);
      overlay.appendChild(caret);
    }
  }

  return {
    /** Replaces the whole set, which is how presence arrives. */
    set(next: PeerCaret[]): void {
      peers = next;
      render();
    },
    /** After anything that moves text: a reflow, a zoom, a remote edit. */
    reposition(): void {
      if (peers.length) render();
    },
    teardown(): void {
      peers = [];
      overlay.remove();
    },
  };
}
