import { defineConfig } from "vitest/config";

// DOM-based round-trip tests run under jsdom.
export default defineConfig({
  test: {
    globals: true,
    environment: "jsdom",
    include: ["src/**/*.test.ts"],
  },
});
