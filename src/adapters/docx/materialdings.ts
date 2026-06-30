// Register the bundled MaterialDings webfont (see materialdings-font.ts) so Wingdings symbol glyphs
// render without the proprietary font. The font bytes live in a separate module that is imported
// lazily, so the ~21 KB cost is only paid the first time a Wingdings glyph actually appears.
let started = false;
export function ensureDingsFont(): void {
  if (started || typeof document === "undefined" || !document.head) return;
  started = true;
  void import("./materialdings-font")
    .then(({ MATERIALDINGS_WOFF2 }) => {
      if (document.querySelector("style[data-rdoc-dings]")) return;
      const style = document.createElement("style");
      style.setAttribute("data-rdoc-dings", "");
      style.textContent = `@font-face{font-family:'MaterialDings';font-display:swap;src:url(data:font/woff2;base64,${MATERIALDINGS_WOFF2}) format('woff2')}`;
      document.head.appendChild(style);
    })
    .catch(() => { started = false; }); // let a later glyph retry if the load failed
}
