import { createDocxAdapter } from "../adapters/docx/index";
import { createOdtAdapter } from "../adapters/odt/index";
import type { Adapter } from "./types";

// A headless round trip that matches what the EDITOR does, for the corpus checks.
//
// The bare docxToHtml/htmlToDocx pair is not enough: a document's body is only part of
// what the reader produces. Footnote and endnote bodies, the header and footer, and the
// page geometry come back on their own channels, and the writer expects them back. Calling
// the two-argument form drops them, which looks exactly like data loss but is the caller's
// omission. The adapter is the real contract, so the corpus goes through it.

export const adapterFor = (name: string, bytes: Uint8Array): Adapter =>
  name.endsWith(".docx") ? createDocxAdapter(bytes) : createOdtAdapter(bytes);

/** Read and write back, optionally editing the body first. Everything the reader returned
    that the writer accepts is handed back, so an untouched document stays untouched. */
export async function roundTrip(name: string, bytes: Uint8Array, edit?: (html: string) => string): Promise<Uint8Array> {
  const adapter = adapterFor(name, bytes);
  const doc = adapter.read();
  const body = edit ? edit(doc.body) : doc.body;
  const noEdits = { reactions: [], replies: [], done: new Map<string, boolean>(), deletedComments: [], edited: [] };
  return await adapter.write(body, [], noEdits, doc.page, [], doc.notes);
}

/** Replace the text of the first paragraph that has any, leaving the rest alone. */
export function editFirstParagraph(html: string, replacement: string): string {
  const doc = new DOMParser().parseFromString(`<div>${html}</div>`, "text/html");
  const root = doc.body.firstElementChild!;
  for (const p of Array.from(root.querySelectorAll("p"))) {
    if ((p.textContent ?? "").trim().length < 3) continue;
    p.textContent = replacement;
    return root.innerHTML;
  }
  return html;
}
