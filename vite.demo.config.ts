import { defineConfig } from "vite";

// GitHub Pages build of the demo. The demo/ folder is the site root and imports the
// library straight from ../src, so allow the dev server to read the parent. base "./"
// keeps asset URLs working when served from the repo subpath (…github.io/richdoc/).
// Kept separate from the default config name so Vitest still runs with its own defaults.
export default defineConfig({
  root: "demo",
  base: "./",
  server: { fs: { allow: [".."] } },
  // esbuild's dep pre-bundling mangles temml (it then errors on every command); serve its raw ESM.
  optimizeDeps: { exclude: ["temml"] },
  build: {
    outDir: "../demo-dist",
    emptyOutDir: true,
    target: "es2022",
    sourcemap: false,
  },
});
