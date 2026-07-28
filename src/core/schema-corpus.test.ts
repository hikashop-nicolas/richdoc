import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { docxToHtml, htmlToDocx } from "../adapters/docx/index";
import { odtToHtml, htmlToOdt } from "../adapters/odt/index";

// Not a test of behaviour: this writes the corpus that `npm run check:schema` validates
// against the official schemas. It lives here because only the project's own toolchain can
// compile the library, and being a test file means it goes through that toolchain rather
// than needing a build.
//
// Each fixture is read, edited (so the writer really re-emits the body rather than cloning
// it) and written beside the original for comparison.

const CORPUS = join(__dirname, "../../test-corpus");
const OUT = join(process.cwd(), ".cache", "schema-corpus");

/** Change the first paragraph with text, leaving everything else alone. */
function edit(html: string): string {
  const doc = new DOMParser().parseFromString(`<div>${html}</div>`, "text/html");
  const root = doc.body.firstElementChild!;
  for (const p of Array.from(root.querySelectorAll("p"))) {
    if ((p.textContent ?? "").trim().length < 3) continue;
    p.textContent = "schema check";
    return root.innerHTML;
  }
  return html;
}

describe.skipIf(!existsSync(CORPUS))("schema-check corpus", () => {
  it("writes every fixture back for validation", () => {
    mkdirSync(OUT, { recursive: true });
    const names = readdirSync(CORPUS).filter((n) => n.endsWith(".docx") || n.endsWith(".odt")).sort();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const bytes = new Uint8Array(readFileSync(join(CORPUS, name)));
      const out = name.endsWith(".docx")
        ? htmlToDocx(edit(docxToHtml(bytes)), bytes)
        : htmlToOdt(edit(odtToHtml(bytes)), bytes);
      writeFileSync(join(OUT, name), out);
    }
    expect(readdirSync(OUT).length).toBe(names.length);
  });
});
