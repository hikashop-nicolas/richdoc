#!/usr/bin/env node
// Run the LibreOffice oracle. Plain python3 is enough: the script uses only the standard
// library (it shells out to soffice and parses flat ODF with xml.etree), unlike the
// independent-reader check which needs python-docx and odfpy.

import { execFileSync } from "node:child_process";
import { join } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname;
try {
  execFileSync("python3", [join(ROOT, "scripts", "lo-check.py")], { stdio: "inherit", cwd: ROOT });
} catch (e) {
  process.exit(e.status ?? 1);
}
