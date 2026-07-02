import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { docxToHtml, htmlToDocx } from "./index";

// Real-file round-trip smoke corpus: every sample must read to HTML, write back,
// and re-read with its text intact. Guards the preserve-by-default merge paths
// against real Word/LibreOffice output, not just synthetic fixtures.
// demo/samples is gitignored (real files stay out of the public repo), so this
// suite runs on dev checkouts that have it and skips elsewhere (CI).
const dir = join(__dirname, "../../../demo/samples");
const samples = existsSync(dir) ? readdirSync(dir).filter((f) => f.endsWith(".docx")) : [];

describe.skipIf(!samples.length)("real-file round-trip corpus", () => {
  for (const name of samples) {
    it(`round-trips ${name}`, () => {
      const bytes = new Uint8Array(readFileSync(join(dir, name)));
      const html = docxToHtml(bytes);
      const out = htmlToDocx(html, bytes);
      const html2 = docxToHtml(out);
      // text content is preserved through a no-edit save
      const textOf = (h: string) => h.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
      expect(textOf(html2)).toBe(textOf(html));
    });
  }
});

// Same guard for the .odt samples through the odt adapter.
import { odtToHtml, htmlToOdt } from "../odt/index";
const odtSamples = existsSync(dir) ? readdirSync(dir).filter((f) => f.endsWith(".odt")) : [];
describe.skipIf(!odtSamples.length)("real-file odt round-trip corpus", () => {
  for (const name of odtSamples) {
    it(`round-trips ${name}`, () => {
      const bytes = new Uint8Array(readFileSync(join(dir, name)));
      const html = odtToHtml(bytes);
      const out = htmlToOdt(html, bytes);
      const html2 = odtToHtml(out);
      const textOf = (h: string) => h.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
      expect(textOf(html2)).toBe(textOf(html));
    });
  }
});
