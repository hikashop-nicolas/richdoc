#!/usr/bin/env node
// Run the independent-reader cross-check with whichever Python has python-docx and odfpy.
//
// CI installs them into the system interpreter; a local checkout more likely wants a
// virtual environment, since a Homebrew or Debian python refuses to install into itself.
// This prefers .cache/py (which `npm run check:reader:setup` creates) and falls back to
// python3, so the command works the same either way instead of failing with an import
// error and no advice.

import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname;
const venv = join(ROOT, ".cache", "py", "bin", "python");
const python = existsSync(venv) ? venv : "python3";

try {
  execFileSync(python, [join(ROOT, "scripts", "reader-check.py")], { stdio: "inherit", cwd: ROOT });
} catch (e) {
  if (e.status === 2 && python === "python3") {
    console.error("\npython-docx and odfpy are not installed. Either `pip install python-docx odfpy`,");
    console.error("or run `npm run check:reader:setup` to put them in a virtual environment under .cache/.");
  }
  process.exit(e.status ?? 1);
}
