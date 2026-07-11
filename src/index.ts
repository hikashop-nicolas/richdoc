import { strFromU8, unzipSync } from "fflate";
import { unzipAsync } from "./core/zip";
import { createDocxEditor, createDocxEditorAsync, type DocxEditor } from "./adapters/docx";
import { createOdtEditor, createOdtEditorAsync, type OdtEditor } from "./adapters/odt";
import { createDocEditor, createDocEditorAsync, isCfb, type DocEditor } from "./adapters/doc";

// richdoc: one rich-document editor for both .docx and .odt. Today each format is a
// self-contained adapter; the shared engine (UI, comments, track changes, passthrough)
// is being hoisted into src/core so both formats reuse it.

export { createDocxEditor, createDocxEditorAsync, type DocxEditor } from "./adapters/docx";
export { createOdtEditor, createOdtEditorAsync, type OdtEditor } from "./adapters/odt";
export { createDocEditor, createDocEditorAsync, type DocEditor } from "./adapters/doc";
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

// Same, but inflate the container off the main thread and reuse that single inflate for both
// format detection and parsing (a .docx open previously inflated three times, all on the main
// thread). A legacy .doc is not a zip, so it routes straight to the CFB editor. The DOM-bound
// parse itself still runs on the main thread once the map is ready.
export async function createEditorAsync(container: HTMLElement, bytes: Uint8Array, options: EditorOptions = {}): Promise<RichEditor> {
  if (isCfb(bytes)) return createDocEditorAsync(container, bytes, options);
  let files: Record<string, Uint8Array> | undefined;
  try {
    files = await unzipAsync(bytes);
  } catch {
    /* not a readable zip: fall through to docx, which latches read-only on the failed parse */
  }
  const mt = files?.["mimetype"];
  if (mt && strFromU8(mt).includes("opendocument.text")) return createOdtEditorAsync(container, bytes, options, files);
  return createDocxEditorAsync(container, bytes, options, files);
}
