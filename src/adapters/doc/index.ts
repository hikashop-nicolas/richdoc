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
import { htmlToDoc, type DocComment } from "./write";

export { docToHtml, docToParts } from "./read";
export { htmlToDoc } from "./write";
export { isCfb } from "./cfb";

export type DocEditorOptions = EditorOptions;
export type DocEditor = RichEditor;

// Reconstruct the comment set to write: start from the original file's comments, add any
// authored in-session (their author/text ride on the ref markers in the body), apply text
// edits, and drop deleted ones. htmlToDoc only emits those whose ref survives in the body.
function rebuildComments(original: Uint8Array, bodyHtml: string, edits: CommentEdits): DocComment[] {
  const byId = new Map<string, DocComment>();
  for (const c of docToParts(original).comments ?? []) byId.set(c.id, { id: c.id, author: c.author, text: c.text });
  const dom = new DOMParser().parseFromString(`<body>${bodyHtml}</body>`, "text/html");
  for (const ref of Array.from(dom.querySelectorAll(".docx-comment-ref[data-comment-new]"))) {
    const id = ref.getAttribute("data-comment-id") || "";
    if (id) byId.set(id, { id, author: ref.getAttribute("data-comment-author") || "Author", text: ref.getAttribute("data-comment-text") || "" });
  }
  for (const e of edits.edited) { const c = byId.get(e.id); if (c) c.text = e.text; }
  for (const id of edits.deletedComments) byId.delete(id);
  return [...byId.values()];
}

/** Wrap a .doc byte array as an engine adapter: parse, serialize, capabilities. */
export function createDocAdapter(bytes: Uint8Array): Adapter {
  const original = bytes.slice();
  return {
    original,
    read(): RichDoc {
      const parts = docToParts(bytes);
      return {
        body: parts.body || "<p><br></p>",
        header: parts.header ?? "",
        footer: parts.footer ?? "",
        comments: parts.comments ?? [],
        page: parts.page,
        notes: parts.notes,
      };
    },
    // The .doc writer regenerates the whole file from the edited body HTML plus the footnote /
    // endnote bodies, comments and the header/footer. Comments are rebuilt from the original
    // file, overlaid with any in-session additions (data on the ref markers) and text edits.
    write(bodyHtml: string, editedParts, edits: CommentEdits, page, _styles, notes): Uint8Array {
      const comments = rebuildComments(original, bodyHtml, edits);
      const partBy = (path: string) => editedParts.find((p) => p.path === path)?.html;
      const hf = { header: partBy("header"), footer: partBy("footer") };
      return htmlToDoc(bodyHtml, page, notes, comments, hf);
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
      ref.setAttribute("data-comment-author", meta.author);
      ref.setAttribute("data-comment-text", meta.text);
      return { start: cmark(), end: cmark(), ref };
    },
    capabilities: {
      comments: true,
      commentReplies: false,
      commentReactions: false,
      trackChanges: false,
      images: true,
      tables: true,
      headerFooter: true,
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
