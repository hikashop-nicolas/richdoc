#!/usr/bin/env node
// Validate what richdoc WRITES against the official schemas.
//
// The point is not "is this file schema-perfect": real files are not. LibreOffice writes
// loext: extension attributes that the ODF schema does not allow, and Word writes
// constructs the transitional XSDs do not describe. A validator run against any real file
// reports those, and they are not defects of ours.
//
// So this compares. Every fixture is read, edited and written back (see
// src/core/schema-corpus.test.ts, which `npm run check:schema` runs first), and each XML
// part of the OUTPUT is validated against the same part of the INPUT. Only an error the
// input did not have is reported: that is one this project introduced.
//
// Schemas, downloaded once into .cache/ and not committed (they belong to ECMA and OASIS,
// and a checkout should not carry a copy of someone else's standard):
//   .docx  ECMA-376 Part 4 transitional, which is what Word writes and richdoc preserves
//   .odt   OASIS OpenDocument v1.3 RELAX NG
//
// Needs xmllint (libxml2), which macOS carries and Debian/Ubuntu package as libxml2-utils.

import { execFileSync, execSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";
import { tmpdir } from "node:os";

const ROOT = new URL("..", import.meta.url).pathname;
const CACHE = join(ROOT, ".cache");
const XSD = join(CACHE, "ooxml-xsd");
const RNG = join(CACHE, "odf", "OpenDocument-v1.3-schema.rng");
const PART4 = "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip";
const ODF_RNG = "https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-schema.rng";
const XML_XSD = "https://www.w3.org/2001/xml.xsd";

/**
 * wml.xsd and shared-math.xsd reference xml:space by ref but never import the xml
 * namespace, so xmllint cannot compile them and every validation silently reports nothing
 * at all. (sml.xsd, which sheetedit uses, does not have this problem.) Supplying the W3C
 * schema and adding the missing import makes the schema compile; without it the docx half
 * of this check is a no-op that always passes.
 */
function fixXmlNamespaceImport() {
  for (const f of ["wml.xsd", "shared-math.xsd"]) {
    const path = join(XSD, f);
    if (!existsSync(path)) continue;
    const s = readFileSync(path, "utf8");
    if (s.includes('schemaLocation="xml.xsd"')) continue;
    const m = /<(\w+):schema\b[\s\S]*?>/.exec(s);
    if (!m) continue;
    const imp = `<${m[1]}:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>`;
    writeFileSync(path, s.slice(0, m.index + m[0].length) + imp + s.slice(m.index + m[0].length));
  }
}

function have(cmd) {
  try { execSync(`command -v ${cmd}`, { stdio: "ignore" }); return true; } catch { return false; }
}

function ensureSchemas() {
  if (!existsSync(join(XSD, "wml.xsd"))) {
    console.log("fetching the ECMA-376 transitional schemas (once, into .cache/) ...");
    mkdirSync(XSD, { recursive: true });
    const work = join(tmpdir(), `ooxml-xsd-${process.pid}`);
    mkdirSync(work, { recursive: true });
    execFileSync("curl", ["-sSL", "-o", join(work, "part4.zip"), PART4]);
    execFileSync("unzip", ["-o", "-q", join(work, "part4.zip"), "OfficeOpenXML-XMLSchema-Transitional.zip", "-d", work]);
    execFileSync("unzip", ["-o", "-q", join(work, "OfficeOpenXML-XMLSchema-Transitional.zip"), "-d", XSD]);
    rmSync(work, { recursive: true, force: true });
    execFileSync("curl", ["-sSL", "-o", join(XSD, "xml.xsd"), XML_XSD]);
    fixXmlNamespaceImport();
  }
  if (!existsSync(RNG)) {
    console.log("fetching the OpenDocument 1.3 schema (once, into .cache/) ...");
    mkdirSync(join(CACHE, "odf"), { recursive: true });
    execFileSync("curl", ["-sSL", "-o", RNG, ODF_RNG]);
  }
}

/**
 * Refuse to run on a schema xmllint cannot compile. A compile failure is reported as a
 * "Schemas parser error", not a validity error, so it would be filtered out below and
 * every document would pass a check that had in fact validated nothing.
 */
function assertSchemasUsable() {
  const probe = join(tmpdir(), `schemaprobe-${process.pid}.xml`);
  writeFileSync(probe, '<?xml version="1.0"?><probe/>');
  for (const [label, mode, file] of [["wml.xsd", "--schema", join(XSD, "wml.xsd")], ["OpenDocument rng", "--relaxng", RNG]]) {
    let out = "";
    try {
      out = execFileSync("xmllint", ["--noout", mode, file, probe], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
    } catch (e) {
      out = `${e.stdout ?? ""}${e.stderr ?? ""}`;
    }
    if (/failed to compile|failed to parse/.test(out)) {
      console.error(`${label} does not compile, so this check would validate nothing:\n${out.split("\n").slice(0, 3).join("\n")}`);
      console.error("Delete .cache/ and re-run to re-fetch, or fix the schema patching in this script.");
      rmSync(probe, { force: true });
      process.exit(2);
    }
  }
  rmSync(probe, { force: true });
}

/** How a part is validated, or null for parts this check does not cover. */
function schemaFor(part, kind) {
  if (kind === "docx") {
    if (/^word\/(document|styles|numbering|settings|comments|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(part))
      return { mode: "--schema", file: join(XSD, "wml.xsd") };
    return null; // themes, fontTable, media, rels: their own schemas or not XML
  }
  if (/^(content|styles|meta|settings)\.xml$/.test(part)) return { mode: "--relaxng", file: RNG };
  return null;
}

/** Validate one part; returns error lines with file name and line numbers stripped, so the
    same complaint from two files compares equal. */
function validate(file, schema) {
  let out = "";
  try {
    out = execFileSync("xmllint", ["--noout", schema.mode, schema.file, file], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
  } catch (e) {
    out = `${e.stdout ?? ""}${e.stderr ?? ""}`;
  }
  return out
    .split("\n")
    .filter((l) => /validity error|Schemas validity error/.test(l))
    .map((l) => l.replace(/^[^:]*:\d+:\s*/, "").trim());
}

function unzipTo(zip, dir) {
  mkdirSync(dir, { recursive: true });
  execFileSync("unzip", ["-o", "-q", zip, "-d", dir]);
  const out = [];
  const walk = (d) => {
    for (const name of readdirSync(d)) {
      const p = join(d, name);
      if (statSync(p).isDirectory()) walk(p);
      else if (p.endsWith(".xml")) out.push(p);
    }
  };
  walk(dir);
  return out;
}

if (!have("xmllint")) {
  console.error("xmllint is needed (macOS has it; Debian/Ubuntu: apt-get install libxml2-utils)");
  process.exit(2);
}
ensureSchemas();
assertSchemasUsable();

const corpus = join(CACHE, "schema-corpus");
if (!existsSync(corpus)) {
  console.error("no corpus: run `npm run check:schema`, which writes it first");
  process.exit(2);
}
const names = readdirSync(corpus).filter((n) => n.endsWith(".docx") || n.endsWith(".odt")).sort();
console.log(`checking ${names.length} documents ...`);

let introduced = 0;
for (const name of names) {
  const kind = name.endsWith(".docx") ? "docx" : "odt";
  const before = join(ROOT, "test-corpus", name);
  const after = join(corpus, name);
  if (!existsSync(before)) continue;
  const work = join(tmpdir(), `schemacheck-${process.pid}-${name}`);
  const beforeParts = unzipTo(before, join(work, "in"));
  const afterParts = unzipTo(after, join(work, "out"));
  // The baseline is every complaint the INPUT draws, from any of its parts. A file that
  // already writes a loext: attribute is one where that complaint says nothing about us,
  // wherever it later turns up. What is worth reporting is a KIND of violation the file
  // did not have before.
  const baseline = new Set();
  for (const p of beforeParts) {
    const schema = schemaFor(relative(join(work, "in"), p), kind);
    if (schema) for (const e of validate(p, schema)) baseline.add(e);
  }
  for (const p of afterParts) {
    const rel = relative(join(work, "out"), p);
    const schema = schemaFor(rel, kind);
    if (!schema) continue;
    for (const e of validate(p, schema).filter((e) => !baseline.has(e))) {
      introduced++;
      console.log(`${name} -> ${rel}\n  ${e}`);
    }
  }
  rmSync(work, { recursive: true, force: true });
}

if (introduced) {
  console.error(`\n${introduced} schema violation(s) introduced by writing.`);
  process.exit(1);
}
console.log("no schema violations introduced by writing.");
