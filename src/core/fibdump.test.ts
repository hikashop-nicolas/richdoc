import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { it } from "vitest";
import { parseFib } from "../adapters/doc/fib";
import { readCfb } from "../adapters/doc/cfb";
it("dump", () => {
  const out: string[] = [];
  for (const name of ["tables", "lists", "comments", "notes", "fields"])
  for (const [label, p] of [["ORIG", `test-corpus/${name}.doc`], ["OURS", `.cache/lo-corpus/${name}.doc`]] as const) {
    const cfb = readCfb(new Uint8Array(readFileSync(join(process.cwd(), p))));
    const fib = parseFib(cfb.get("WordDocument")!);
    const nz: number[] = [];
    for (let i = 0; i < 93; i++) { const e = fib.fc(i); if (e.lcb > 0) nz.push(i); }
    out.push(`${p.includes("test-corpus")?"ORIG":"OURS"} ${p.split("/").pop()}: ${nz.join(",")}`);
  }
  writeFileSync("/tmp/fib.txt", out.join("\n"));
});
