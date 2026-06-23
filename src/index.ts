import { strFromU8, unzipSync } from "fflate";
import { createDocxEditor, type DocxEditor } from "./adapters/docx";
import { createOdtEditor, type OdtEditor } from "./adapters/odt";

// richdoc: one rich-document editor for both .docx and .odt. Today each format is a
// self-contained adapter; the shared engine (UI, comments, track changes, passthrough)
// is being hoisted into src/core so both formats reuse it.

export { createDocxEditor, type DocxEditor } from "./adapters/docx";
export { createOdtEditor, type OdtEditor } from "./adapters/odt";
export { setLocale } from "./adapters/docx/i18n";

export type RichEditor = DocxEditor | OdtEditor;
export type EditorFormat = "docx" | "odt";

/** Detect the format from the archive (ODT declares its mimetype; otherwise assume OOXML). */
export function sniffFormat(bytes: Uint8Array): EditorFormat {
  try {
    const files = unzipSync(bytes);
    const mt = files["mimetype"];
    if (mt && strFromU8(mt).includes("opendocument.text")) return "odt";
  } catch {
    /* not a zip we can read; fall through */
  }
  return "docx";
}

export interface EditorOptions {
  onChange?: () => void;
  author?: string;
  now?: string;
}

/** Mount the right editor for the document's format. */
export function createEditor(container: HTMLElement, bytes: Uint8Array, options: EditorOptions = {}): RichEditor {
  return sniffFormat(bytes) === "odt" ? createOdtEditor(container, bytes, options) : createDocxEditor(container, bytes, options);
}
