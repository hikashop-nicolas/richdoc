# Equation editor

## Status: DONE (v1, docx + odt)

Insert (LaTeX dialog + live preview), display (native 2D MathML), click-to-edit, and round-trip on
both formats all work and were browser-verified with the quadratic formula (renders as a real fraction
+ radical; re-reads as an editable equation). docx exports OMML (m:f / m:rad); odt embeds a formula
sub-document (draw:frame -> draw:object -> Object/content.xml MathML) plus its manifest entries.

**Build caveat:** esbuild's dependency pre-bundling mangles temml (it then errors on every command and
the equation shows as raw LaTeX). The demo's Vite config excludes temml from pre-bundling
(`optimizeDeps.exclude: ["temml"]`); any Vite-based consumer (e.g. omnitext dev) needs the same.

## Goal

Make equations first-class instead of opaque passthrough: insert, display, edit, and round-trip math.

## Model (DOM)

An inline, non-editable span holding MathML (the browser renders MathML natively):

```
<span class="docx-eq" data-rdoc-eq contenteditable="false" data-latex="x^2">
  <math xmlns="http://www.w3.org/1998/Math/MathML">…</math>
</span>
```

- `data-latex` is the authored source, so editing reopens the LaTeX. An imported equation has no
  `data-latex`, so opening it recovers a best-effort LaTeX from its MathML (`mathml-latex.ts`).
- MathML is the in-editor pivot every format converts to/from.

## Authoring

LaTeX input with a live MathML preview, converted with **temml** (MIT, lazy-loaded). Display is the
browser's native MathML rendering, no runtime dependency.

## Round-trip

- **docx (OMML)**: `src/adapters/docx/omml.ts` converts OMML <-> MathML for the common constructs
  (runs/mi/mn/mo, mrow, fraction m:f, scripts m:sSup/m:sSub/m:sSubSup, radical m:rad, n-ary m:nary,
  delimiters m:d, matrices m:m, accents m:acc, bars m:bar). Read: `m:oMath` -> MathML span, keeping the original OMML in `data-omml` so an
  un-edited equation rewrites verbatim (lossless). Write: a span with `data-omml` and no edit ->
  original OMML; otherwise MathML -> OMML. Constructs outside the common set stay passthrough.
- **odt**: an embedded formula object. Read: a `draw:frame` whose `draw:object` resolves to an
  `Object/content.xml` with a `math` root -> a MathML equation span, the whole frame stashed in
  `data-odt-xml` so an untouched equation re-emits verbatim (its sub-document is preserved in the
  rebuilt archive). Write: a new or edited equation writes a fresh `Formula_rdocN/content.xml`
  (the span's MathML), references it from an as-char `draw:frame` + `draw:object`, and registers
  both the directory and its content.xml in `META-INF/manifest.xml`. Capability `equations` is now
  true on both adapters. Editing an imported equation leaves its old Object sub-document orphaned in
  the archive (harmless, unreferenced); a future pass could reuse the original href instead.

## UI

`src/core/feature/equation.ts`: an "Insert equation" toolbar button (gated by caps.equations) opening a
dialog (LaTeX textarea + live preview + insert), and click-to-edit on an existing equation. The OMML
conversion lives in the docx adapter, not the engine.

## Deferred / not done

- Editing an imported odt equation orphans its original Object sub-document (unreferenced, harmless).
- OMML constructs beyond the covered set (boxes m:box, group-chars m:groupChr, equation arrays
  m:eqArr, box borders) stay verbatim passthrough.
- Delimited matrices (pmatrix / bmatrix / vmatrix / Vmatrix / Bmatrix / cases) round-trip via an
  OMML m:d wrapping the m:m, and recovery names the environment from the bracket pair. Per-column
  alignment is not preserved (matrices write centered; cases recovers but is not left-aligned).
- MathML -> LaTeX recovery (`mathml-latex.ts`) covers the supported set; constructs outside it fall
  back to text content, so editing a very exotic imported equation may lose detail.
