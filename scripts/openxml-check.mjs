#!/usr/bin/env node
// Validate what richdoc writes with Microsoft's own Open XML SDK.
//
// This asks a different question from the ECMA-376 check next door. That one validates
// individual XML parts against the published schemas, in isolation, and does not look at
// relationships or content types at all. The SDK opens the whole package: it checks part
// structure and the links between parts as well as the markup, and it applies Office's own
// rules rather than only what the schema says.
//
// Same baseline discipline as everywhere else: real files draw complaints too (LibreOffice
// writes charset values and numbering attributes the transitional rules reject), so the
// output is compared against the input it came from and only a NEW kind of complaint counts.
//
// Needs .NET. It uses `dotnet` when the machine has it, and otherwise runs the SDK image
// under Docker, so a checkout does not have to install a .NET SDK to run the check.

import { execFileSync, execSync } from "node:child_process";
import { existsSync, readdirSync } from "node:fs";
import { join, relative } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname.replace(/\/$/, "");
const PROJECT = "scripts/openxml-validate";
const IMAGE = "mcr.microsoft.com/dotnet/sdk:8.0";
const EXTS = /\.(docx|xlsx|pptx)$/;

function have(cmd) {
  try { execSync(`command -v ${cmd}`, { stdio: "ignore" }); return true; } catch { return false; }
}

/** Run the validator over some files, returning its output lines. */
function validate(files) {
  if (!files.length) return [];
  const rel = files.map((f) => relative(ROOT, f));
  const run = (cmd, args) =>
    execFileSync(cmd, args, { cwd: ROOT, encoding: "utf8", maxBuffer: 64 * 1024 * 1024, stdio: ["ignore", "pipe", "pipe"] });
  const out = have("dotnet")
    ? run("dotnet", ["run", "--project", PROJECT, "-c", "Release", "--", ...rel])
    : run("docker", [
        "run", "--rm",
        "-v", `${ROOT}:/w`,
        // A named volume for the package cache, so only the first run pays for the restore.
        "-v", "openxml-nuget:/root/.nuget/packages",
        "-w", "/w", IMAGE,
        "dotnet", "run", "--project", PROJECT, "-c", "Release", "--", ...rel,
      ]);
  return out.split("\n").filter((l) => l.includes("\t"));
}

/**
 * A complaint's identity, without the things that legitimately differ between two files:
 * the file name, and the positional indices in the XPath. An element inserted earlier in a
 * part shifts every index after it, which would make every later complaint look new.
 */
function kind(line) {
  const [, part, type, xpath, ...rest] = line.split("\t");
  return `${part}\t${type}\t${(xpath ?? "").replace(/\[\d+\]/g, "[]")}\t${rest.join("\t")}`;
}

if (!have("dotnet") && !have("docker")) {
  console.error("needs .NET: install the SDK (brew install --cask dotnet-sdk), or Docker to run it in a container");
  process.exit(2);
}

const corpus = join(ROOT, ".cache", "schema-corpus");
if (!existsSync(corpus)) {
  console.error("no corpus: run `npm run check:openxml`, which writes it first");
  process.exit(2);
}
const names = readdirSync(corpus).filter((n) => EXTS.test(n)).sort()
  .filter((n) => existsSync(join(ROOT, "test-corpus", n)));
if (!names.length) {
  console.error("no documents to check");
  process.exit(2);
}
console.log(`checking ${names.length} documents with the Open XML SDK ...`);

// One invocation per side rather than one per file: starting .NET is the slow part.
const before = validate(names.map((n) => join(ROOT, "test-corpus", n)));
const after = validate(names.map((n) => join(corpus, n)));

const baselineByFile = new Map();
for (const line of before) {
  const name = line.split("\t")[0].split("/").pop();
  if (!baselineByFile.has(name)) baselineByFile.set(name, new Set());
  baselineByFile.get(name).add(kind(line));
}

let introduced = 0;
for (const line of after) {
  const [path, part, type, , ...rest] = line.split("\t");
  const name = path.split("/").pop();
  if (baselineByFile.get(name)?.has(kind(line))) continue;
  introduced++;
  console.log(`${name} -> ${part} [${type}]\n  ${rest.join(" ")}`);
}

if (introduced) {
  console.error(`\n${introduced} problem(s) introduced by writing, as the Open XML SDK sees it.`);
  process.exit(1);
}
console.log("the Open XML SDK finds no problem in the output that the input did not have.");
