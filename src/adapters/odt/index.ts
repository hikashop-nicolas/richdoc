// odtedit: a standalone, framework-agnostic, client-side OpenDocument Text (.odt) editor.
//
// An .odt is a zip of OpenDocument XML; the body lives in content.xml. The read half
// (./read) converts it to HTML, the shared engine edits it, and the write half (./write)
// rebuilds the archive on save, preserving every other part byte-for-byte.
import { createRichEditor } from "../../core/editor";
import { unzipAsync } from "../../core/zip";
import type { Adapter, CommentEdits, CommentMarkers, EditorOptions, NewCommentMeta, NewStyle, Note, PageGeometry, RichDoc, RichEditor } from "../../core/types";
import { odtToParts } from "./read";
import { htmlToOdtAsync } from "./write";

// Public surface (also consumed by the tests).
export { odtToHtml, odtToParts } from "./read";
export { htmlToOdt } from "./write";


export type OdtEditorOptions = EditorOptions;
export type OdtEditor = RichEditor;

/** Wrap a .odt byte array as an engine adapter: parse, serialize, capabilities. */
export function createOdtAdapter(bytes: Uint8Array, preunzipped?: Record<string, Uint8Array>): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      // Let a parse failure propagate: the engine latches the editor read-only so a
      // blank surface can never overwrite the real file on save.
      const p = odtToParts(bytes, preunzipped);
      const parts = { ...p, body: p.body || "<p><br></p>" };
      return {
        body: parts.body,
        header: parts.header,
        footer: parts.footer,
        headerPath: parts.header ? "header" : undefined,
        footerPath: parts.footer ? "footer" : undefined,
        sectionBands: parts.sectionBands,
        headerEven: parts.headerEven,
        footerEven: parts.footerEven,
        headerFirst: parts.headerFirst,
        footerFirst: parts.footerFirst,
        notes: parts.notes,
        comments: parts.comments,
        page: parts.page,
        paragraphStyles: parts.paragraphStyles,
        characterStyles: parts.characterStyles,
        styleDefs: parts.styleDefs,
        styleCss: parts.styleCss,
        noteCss: parts.noteCss,
      };
    },
    write(bodyHtml: string, parts: { path: string; html: string }[], edits: CommentEdits, page?: PageGeometry, newStyles?: NewStyle[], notes?: Note[]): Promise<Uint8Array> {
      return htmlToOdtAsync(bodyHtml, original, { done: edits.done, parts, page, newStyles, notes, edited: edits.edited });
    },
    newCommentMarkers(meta: NewCommentMeta): CommentMarkers {
      const cmark = (): HTMLElement => {
        const s = document.createElement("span");
        s.className = "docx-cmark";
        s.contentEditable = "false";
        return s;
      };
      const ref = document.createElement("span");
      ref.className = "docx-comment-ref";
      ref.contentEditable = "false";
      ref.textContent = "\u{1F4AC}";
      ref.setAttribute("data-comment-id", meta.id);
      ref.setAttribute("data-comment-new", "1");
      ref.setAttribute("data-comment-paraid", meta.paraId);
      ref.setAttribute("data-comment-author", meta.author);
      if (meta.date) ref.setAttribute("data-comment-date", meta.date);
      ref.setAttribute("data-comment-text", meta.text);
      ref.setAttribute("data-comment-meta", meta.date ? `${meta.author} – ${meta.date.slice(0, 10)}` : meta.author);
      return { start: cmark(), end: cmark(), ref };
    },
    capabilities: {
      comments: true,
      commentReplies: false,
      commentReactions: false,
      trackChanges: true,
      images: true,
      tables: true,
      headerFooter: true,
      pageBreak: false,
      textColor: true,
      fontControls: true,
      alignment: true,
      verticalText: true,
      fields: true,
      equations: true, // embedded formula objects (draw:object -> MathML sub-document)
      sections: "leading",
      pageNumbering: "format", // style:num-format on the page layout; no page-layout start
      lineNumbering: "basic", // document-level text:linenumbering-configuration (no per-section / start)
      pageVAlign: false, // odt has no page-content vertical alignment
    },
  };
}

/** Mount a .odt editor in `container`: the odt adapter driving the shared engine. */
export function createOdtEditor(container: HTMLElement, bytes: Uint8Array, options: OdtEditorOptions = {}): OdtEditor {
  return createRichEditor(container, createOdtAdapter(bytes), options);
}

// Same, but inflate the .odt off the main thread first. The DOM-bound parse still runs on the
// main thread once the map is ready.
export async function createOdtEditorAsync(container: HTMLElement, bytes: Uint8Array, options: OdtEditorOptions = {}, preunzipped?: Record<string, Uint8Array>): Promise<OdtEditor> {
  let files = preunzipped;
  if (!files) {
    try {
      files = await unzipAsync(bytes);
    } catch {
      /* not a readable zip; let the sync adapter surface the parse failure (read-only latch) */
    }
  }
  return createRichEditor(container, createOdtAdapter(bytes, files), options);
}
