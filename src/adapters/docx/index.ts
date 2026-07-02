// docxedit: a standalone, framework-agnostic, client-side Office Open XML (.docx) editor.
//
// A .docx is a zip of OOXML; the document body lives in word/document.xml. The read half
// (./read) converts that body to HTML, the shared engine edits it in a contenteditable
// surface, and the write half (./write) rebuilds the archive from the edited HTML on save,
// preserving every other part (styles, headers/footers, images, numbering) byte-for-byte.
import { unzipSync } from "fflate";
import { createRichEditor } from "../../core/editor";
import type { Adapter, CommentEdits, EditorOptions, NewCommentMeta, NewStyle, Note, PageGeometry, RichDoc, RichEditor } from "../../core/types";
import "./docxedit.css";
import { W } from "./shared";
import { docxToParts, loadEmbeddedFonts, defaultFont } from "./read";
import type { DocxParts } from "./read";
import { htmlToDocx } from "./write";

// Public surface (also consumed by the tests).
export { docxToHtml, docxToParts, deobfuscateFont } from "./read";
export { htmlToDocx } from "./write";
export type { ReactionEdit, ReplyEdit } from "./write";
export type { CommentReaction, CommentEntry, CommentThread } from "../../core/types";

// ---------------------------------------------------------------------------
// docx adapter over the shared engine
// ---------------------------------------------------------------------------

export type DocxEditorOptions = EditorOptions;
export type DocxEditor = RichEditor;

/** Build the OOXML comment markers (range start/end + reference run) for a new comment. */
function docxCommentMarkers(meta: NewCommentMeta): { start: HTMLElement; end: HTMLElement; ref: HTMLElement } {
  const { id, author, date, text, paraId } = meta;
  const ns = `xmlns:w="${W}"`;
  const markerSpan = (xml: string): HTMLElement => {
    const s = document.createElement("span");
    s.className = "docx-cmark";
    s.contentEditable = "false";
    s.setAttribute("data-comment-id", id);
    s.setAttribute("data-docx-xml", xml);
    return s;
  };
  const start = markerSpan(`<w:commentRangeStart ${ns} w:id="${id}"/>`);
  const end = markerSpan(`<w:commentRangeEnd ${ns} w:id="${id}"/>`);
  const ref = document.createElement("span");
  ref.className = "docx-comment-ref";
  ref.contentEditable = "false";
  ref.textContent = "\u{1F4AC}";
  ref.setAttribute("data-comment-id", id);
  ref.setAttribute("data-comment-new", "1");
  ref.setAttribute("data-comment-paraid", paraId);
  ref.setAttribute("data-comment-author", author);
  if (date) ref.setAttribute("data-comment-date", date);
  ref.setAttribute("data-comment-text", text);
  ref.setAttribute("data-comment-meta", date ? `${author} – ${date.slice(0, 10)}` : author);
  ref.setAttribute(
    "data-docx-xml",
    `<w:r ${ns}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`,
  );
  return { start, end, ref };
}

/** Wrap a .docx byte array as an engine adapter: parse, serialize, comment markers, capabilities. */
export function createDocxAdapter(bytes: Uint8Array): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      // Let a parse failure propagate: the engine latches the editor read-only so a
      // blank surface can never overwrite the real file on save.
      const parts: DocxParts = docxToParts(bytes);
      let fontCss = "";
      let fontUrls: string[] = [];
      let defaultFontName: string | undefined;
      try {
        const fileMap = unzipSync(bytes);
        const ff = loadEmbeddedFonts(fileMap);
        fontCss = ff.css;
        fontUrls = ff.urls;
        defaultFontName = defaultFont(fileMap);
      } catch {
        /* no embedded fonts */
      }
      return { ...parts, fontCss, fontUrls, defaultFont: defaultFontName };
    },
    write(bodyHtml: string, editedParts: { path: string; html: string }[], edits: CommentEdits, page?: PageGeometry, newStyles?: NewStyle[], notes?: Note[]): Uint8Array {
      return htmlToDocx(bodyHtml, original, editedParts, { ...edits, pageGeometry: page, newStyles, notes });
    },
    newCommentMarkers: docxCommentMarkers,
    capabilities: {
      comments: true,
      commentReplies: true,
      commentReactions: true,
      trackChanges: true,
      images: true,
      tables: true,
      headerFooter: true,
      pageBreak: true,
      textColor: true,
      fontControls: true,
      alignment: true,
      verticalText: true,
      fields: true,
      equations: true,
      sections: "trailing",
      pageNumbering: "full",
      lineNumbering: "full",
      pageVAlign: true,
    },
  };
}

/** Mount a .docx editor in `container`: the docx adapter driving the shared engine. */
export function createDocxEditor(container: HTMLElement, bytes: Uint8Array, options: DocxEditorOptions = {}): DocxEditor {
  return createRichEditor(container, createDocxAdapter(bytes), options);
}
