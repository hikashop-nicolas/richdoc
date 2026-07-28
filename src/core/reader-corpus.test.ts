import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { roundTrip } from "./corpus-roundtrip";

// Not a test of behaviour: this writes the corpus that `npm run check:reader` reads back
// with python-docx and odfpy. A NO-EDIT round trip, deliberately: input and output should
// then read identically, so the checker needs no hand-written expectations and any
// difference it reports is something richdoc changed that an outside reader can see.

const CORPUS = join(__dirname, "../../test-corpus");
const OUT = join(process.cwd(), ".cache", "reader-corpus");

describe.skipIf(!existsSync(CORPUS))("independent-reader corpus", async () => {
  it("round-trips every fixture without editing it", async () => {
    mkdirSync(OUT, { recursive: true });
    const names = readdirSync(CORPUS).filter((n) => n.endsWith(".docx") || n.endsWith(".odt")).sort();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      writeFileSync(join(OUT, name), await roundTrip(name, new Uint8Array(readFileSync(join(CORPUS, name)))));
    }
    expect(readdirSync(OUT).length).toBe(names.length);
  });
});
