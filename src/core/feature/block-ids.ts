// Stable identity for the body's top-level blocks.
//
// Collaboration needs to name a paragraph: "this block changed", "insert after that one".
// A position cannot do it, because two people editing at once change each other's
// positions, and richdoc's body has no identity of its own; a paragraph is just wherever
// it happens to sit.
//
// So each top-level block carries a `data-rdoc-bid`. Three rules make it safe:
//
//   1. Assigned on load and after any edit that creates blocks, never regenerated for a
//      block that already has one, or "which block changed" would answer "all of them".
//   2. Carried through pagination for free: repagination moves blocks between wrappers
//      rather than rebuilding them, so the attribute rides along on the same element.
//   3. Stripped from the saved output. The file must not gain an attribute because
//      someone opened it, and richdoc's whole claim is that an untouched part comes back
//      byte for byte.

export const BID = "data-rdoc-bid";

let counter = 0;

/** Unique within this document and this session; ids are never persisted. */
function newId(): string {
  counter += 1;
  return `b${counter.toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

/**
 * Give every top-level block an id, leaving existing ids alone.
 *
 * "Top-level" means a direct child of the body, or of a pagination wrapper: a repaginated
 * document keeps its blocks inside per-page boxes, and those blocks are the same logical
 * paragraphs.
 */
export function assignBlockIds(root: HTMLElement): number {
  let added = 0;
  for (const block of topLevelBlocks(root)) {
    if (!block.hasAttribute(BID)) {
      block.setAttribute(BID, newId());
      added += 1;
    }
  }
  return added;
}

/** The blocks that carry identity: body children, and children of pagination wrappers. */
export function topLevelBlocks(root: HTMLElement): HTMLElement[] {
  const out: HTMLElement[] = [];
  // nodeType rather than `instanceof HTMLElement`: an element from another realm (a test
  // document, an iframe) is not an instance of this realm's HTMLElement, and the check
  // silently matched nothing.
  for (const child of Array.from(root.children)) {
    if (child.nodeType !== 1) continue;
    if (isWrapper(child as HTMLElement)) out.push(...topLevelBlocks(child as HTMLElement));
    else if (!isViewOnly(child as HTMLElement)) out.push(child as HTMLElement);
  }
  return out;
}

/** Pagination boxes: view-only containers whose children are the real blocks. */
const WRAPPER = ".docxedit-colpage, .docxedit-secpage, .docxedit-vband";
/** View-only furniture that is not a block and is stripped before saving anyway. */
const VIEW_ONLY = ".docxedit-pagespacer";

const isWrapper = (el: HTMLElement): boolean => el.matches(WRAPPER);
const isViewOnly = (el: HTMLElement): boolean => el.matches(VIEW_ONLY);

/** Remove every id from a tree, for the saved output. */
export function stripBlockIds(root: HTMLElement): void {
  if (root.hasAttribute(BID)) root.removeAttribute(BID);
  for (const el of Array.from(root.querySelectorAll(`[${BID}]`))) el.removeAttribute(BID);
}

/** The id of the block a node sits in, or null if it is not inside one. */
export function blockIdOf(node: Node | null, root: HTMLElement): string | null {
  let el: Node | null = node;
  while (el && el !== root) {
    if (el.nodeType === 1 && (el as HTMLElement).hasAttribute(BID)) return (el as HTMLElement).getAttribute(BID);
    el = el.parentNode;
  }
  return null;
}

/** Every block's id and current markup, for a caller that wants to diff two states. */
export function blockSnapshot(root: HTMLElement): Map<string, string> {
  const out = new Map<string, string>();
  for (const block of topLevelBlocks(root)) {
    const id = block.getAttribute(BID);
    if (id) out.set(id, block.outerHTML);
  }
  return out;
}

/**
 * Which blocks differ between two snapshots, as ids. Order is reported separately: a
 * caller that only moved blocks around sees no content change, which is correct.
 */
export function changedBlocks(before: Map<string, string>, after: Map<string, string>): {
  changed: string[];
  added: string[];
  removed: string[];
} {
  const changed: string[] = [];
  const added: string[] = [];
  for (const [id, html] of after) {
    if (!before.has(id)) added.push(id);
    else if (before.get(id) !== html) changed.push(id);
  }
  const removed = [...before.keys()].filter((id) => !after.has(id));
  return { changed, added, removed };
}
