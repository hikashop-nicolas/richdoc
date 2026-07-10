import { strFromU8, unzipSync } from "fflate";
import { createDocxEditor, type DocxEditor } from "./adapters/docx";
import { createOdtEditor, type OdtEditor } from "./adapters/odt";
import { createDocEditor, isCfb, type DocEditor } from "./adapters/doc";

// richdoc: one rich-document editor for both .docx and .odt. Today each format is a
// self-contained adapter; the shared engine (UI, comments, track changes, passthrough)
// is being hoisted into src/core so both formats reuse it.

export { createDocxEditor, type DocxEditor } from "./adapters/docx";
export { createOdtEditor, type OdtEditor } from "./adapters/odt";
export { createDocEditor, type DocEditor } from "./adapters/doc";
export { setLocale, initLocale, detectLocale, getLocale, availableLocales } from "./core/i18n";

export type RichEditor = DocxEditor | OdtEditor | DocEditor;
export type EditorFormat = "docx" | "odt" | "doc";

/** Detect the format: legacy .doc is an OLE compound file; ODT declares its mimetype in
 *  the zip; otherwise assume OOXML (.docx). */
export function sniffFormat(bytes: Uint8Array): EditorFormat {
  if (isCfb(bytes)) return "doc";
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
  const fmt = sniffFormat(bytes);
  if (fmt === "doc") return createDocEditor(container, bytes, options);
  if (fmt === "odt") return createOdtEditor(container, bytes, options);
  return createDocxEditor(container, bytes, options);
}
