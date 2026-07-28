import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { roundTrip } from "./corpus-roundtrip";

// Not a test of behaviour: this writes the corpus that `npm run check:lo` hands to
// LibreOffice. A no-edit round trip of every fixture, .doc included: the legacy writer
// rebuilds the whole file rather than preserving it, so LibreOffice is the only judge it
// can have, and this is the only corpus that covers it.

const CORPUS = join(__dirname, "../../test-corpus");
const OUT = join(process.cwd(), ".cache", "lo-corpus");

describe.skipIf(!existsSync(CORPUS))("LibreOffice-oracle corpus", () => {
  it("round-trips every fixture without editing it", async () => {
    mkdirSync(OUT, { recursive: true });
    const names = readdirSync(CORPUS).filter((n) => /\.(docx|odt|doc)$/.test(n)).sort();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      writeFileSync(join(OUT, name), await roundTrip(name, new Uint8Array(readFileSync(join(CORPUS, name)))));
    }
    expect(readdirSync(OUT).length).toBe(names.length);
    // 60s, not vitest's 5s default: these round-trip every fixture, and under the full
    // suite's parallel load that runs past five seconds and fails as a timeout.
  }, 60_000);
});
