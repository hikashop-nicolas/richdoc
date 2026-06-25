// Self-contained i18n for richdoc so the library is a complete multilingual product on its
// own. English is bundled in core as the always-available fallback; every other language lives
// in its own file under ./locales and is loaded on demand via a dynamic import, so a bundler
// code-splits each into its own chunk and a user only ever downloads the language they use.
//
// Adding a language: create ./locales/<code>.ts (export default a Dict) and add one entry to
// LOADERS below. Nothing else changes; the chunk loads only when that locale is selected.
//
// t() is synchronous (the UI is built synchronously), so it reads whatever dict is active and
// falls back to English per key. setLocale()/initLocale() are async because they fetch a chunk;
// a host should `await` one before creating an editor if it wants a non-English first paint.
import en from "./locales/en";

export type Dict = Record<string, string>;

/** Lazy loaders for the non-English locales. Each is a separate chunk (dynamic import). */
const LOADERS: Record<string, () => Promise<{ default: Dict }>> = {
  fr: () => import("./locales/fr"),
};

const loaded: Record<string, Dict> = { en };
let active: Dict = en;
let activeCode = "en";

/** The locale codes this build can serve (English plus every lazy-loadable language). */
export function availableLocales(): string[] {
  return ["en", ...Object.keys(LOADERS)];
}

/** The base language code currently active (e.g. "fr"). */
export function getLocale(): string {
  return activeCode;
}

/** The best base-language match from the browser's preferred-languages list, "en" otherwise. */
export function detectLocale(): string {
  const prefs = (typeof navigator !== "undefined" && navigator.languages) || ["en"];
  for (const tag of prefs) {
    const base = tag.toLowerCase().split("-")[0]!;
    if (base === "en" || LOADERS[base]) return base;
  }
  return "en";
}

/** Force a locale (host escape hatch). Loads the language chunk if needed; unknown codes and
    failed loads fall back to English. Resolves once the active dict is in place. */
export async function setLocale(code: string): Promise<void> {
  const base = code.toLowerCase().split("-")[0]!;
  if (base === "en" || !LOADERS[base]) {
    active = en;
    activeCode = "en";
    return;
  }
  if (!loaded[base]) {
    try {
      loaded[base] = (await LOADERS[base]!()).default;
    } catch {
      active = en;
      activeCode = "en";
      return;
    }
  }
  active = loaded[base]!;
  activeCode = base;
}

/** Detect the browser locale and load it. Hosts call this once at startup, before mounting. */
export function initLocale(): Promise<void> {
  return setLocale(detectLocale());
}

export function t(key: string, params?: Record<string, string | number>): string {
  let s = active[key] ?? en[key] ?? key;
  if (params) s = s.replace(/\{(\w+)\}/g, (_, k: string) => (k in params ? String(params[k]) : `{${k}}`));
  return s;
}
