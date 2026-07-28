import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { editFirstParagraph, roundTrip } from "./corpus-roundtrip";

// Not a test of behaviour: this writes the corpus that `npm run check:schema` validates
// against the official schemas. It lives here because only the project's own toolchain can
// compile the library, and being a test file means it goes through that toolchain rather
// than needing a build.
//
// Each fixture is read, edited (so the writer really re-emits the body rather than cloning
// it) and written beside the original for comparison.

const CORPUS = join(__dirname, "../../test-corpus");
const OUT = join(process.cwd(), ".cache", "schema-corpus");

describe.skipIf(!existsSync(CORPUS))("schema-check corpus", async () => {
  it("writes every fixture back for validation", async () => {
    mkdirSync(OUT, { recursive: true });
    const names = readdirSync(CORPUS).filter((n) => n.endsWith(".docx") || n.endsWith(".odt")).sort();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const bytes = new Uint8Array(readFileSync(join(CORPUS, name)));
      writeFileSync(join(OUT, name), await roundTrip(name, bytes, (h) => editFirstParagraph(h, "schema check")));
    }
    expect(readdirSync(OUT).length).toBe(names.length);
  });
});
