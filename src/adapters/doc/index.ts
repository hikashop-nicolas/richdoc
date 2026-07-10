// docedit: a client-side reader + writer for the Word 97-2003 binary format (MS-DOC).
//
// Unlike docx/odt (XML-in-zip, edited in place), a .doc is an OLE compound file whose text
// position and formatting live in separate offset-indexed tables that cross-reference each
// other, so an in-place span edit is impossible. This adapter therefore reads the binary to
// HTML and, on save, regenerates a complete valid .doc from the edited HTML (from-scratch),
// mapping the subset of formatting richdoc's model represents. Round-trip is validated
// against macOS `textutil` (Apple's real Word engine).
import { createRichEditor } from "../../core/editor";
import type {
  Adapter,
  CommentEdits,
  CommentMarkers,
  EditorOptions,
  NewCommentMeta,
  RichDoc,
  RichEditor,
} from "../../core/types";
import { docToParts } from "./read";
import { htmlToDoc } from "./write";

export { docToHtml, docToParts } from "./read";
export { htmlToDoc } from "./write";
export { isCfb } from "./cfb";

export type DocEditorOptions = EditorOptions;
export type DocEditor = RichEditor;

/** Wrap a .doc byte array as an engine adapter: parse, serialize, capabilities. */
export function createDocAdapter(bytes: Uint8Array): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      const parts = docToParts(bytes);
      return {
        body: parts.body || "<p><br></p>",
        header: "",
        footer: "",
        comments: [],
        page: parts.page,
      };
    },
    // The .doc writer regenerates the whole file from the edited body HTML; parts, comment
    // edits, page geometry and styles are not part of the binary subset we round-trip yet.
    write(bodyHtml: string, _parts, _edits: CommentEdits, page): Uint8Array {
      return htmlToDoc(bodyHtml, page);
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
      return { start: cmark(), end: cmark(), ref };
    },
    capabilities: {
      comments: false,
      commentReplies: false,
      commentReactions: false,
      trackChanges: false,
      images: false,
      tables: true,
      headerFooter: false,
      pageBreak: true,
      textColor: true,
      fontControls: true,
      alignment: true,
      verticalText: true,
      fields: false,
      equations: false,
      sections: "trailing",
      pageNumbering: false,
      lineNumbering: false,
      pageVAlign: false,
    },
  };
}

/** Mount a .doc editor in `container`: the doc adapter driving the shared engine. */
export function createDocEditor(container: HTMLElement, bytes: Uint8Array, options: DocEditorOptions = {}): DocEditor {
  return createRichEditor(container, createDocAdapter(bytes), options);
}
