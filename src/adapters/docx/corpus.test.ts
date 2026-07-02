import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { docxToHtml, htmlToDocx } from "./index";

// Real-file round-trip smoke corpus: every sample must read to HTML, write back,
// and re-read with its text intact. Guards the preserve-by-default merge paths
// against real Word/LibreOffice output, not just synthetic fixtures.
const dir = join(__dirname, "../../../demo/samples");
const samples = readdirSync(dir).filter((f) => f.endsWith(".docx"));

describe("real-file round-trip corpus", () => {
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
