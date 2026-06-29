# Equation editor

## Status: DONE (v1)

Insert (LaTeX dialog + live preview), display (native 2D MathML), click-to-edit, and docx OMML
round-trip all work and were browser-verified with the quadratic formula (renders as a real fraction +
radical; exports to m:f / m:rad; re-reads as an editable equation). odt stays passthrough (gated off).

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

- `data-latex` is the authored source, so editing reopens the LaTeX. Imported equations have no
  LaTeX (editing one means retyping it).
- MathML is the in-editor pivot every format converts to/from.

## Authoring

LaTeX input with a live MathML preview, converted with **temml** (MIT, lazy-loaded). Display is the
browser's native MathML rendering, no runtime dependency.

## Round-trip

- **docx (OMML)**: `src/adapters/docx/omml.ts` converts OMML <-> MathML for the common constructs
  (runs/mi/mn/mo, mrow, fraction m:f, scripts m:sSup/m:sSub/m:sSubSup, radical m:rad, n-ary m:nary,
  delimiters m:d). Read: `m:oMath` -> MathML span, keeping the original OMML in `data-omml` so an
  un-edited equation rewrites verbatim (lossless). Write: a span with `data-omml` and no edit ->
  original OMML; otherwise MathML -> OMML. Constructs outside the common set stay passthrough.
- **odt**: DEFERRED. ODF formulas are embedded objects (a MathML sub-document + manifest + draw:frame),
  a separate heavier mechanism. Equations remain passthrough on odt and the insert button is gated off
  (capability `equations`: docx true, odt false) until the embedded-object path is built.

## UI

`src/core/feature/equation.ts`: an "Insert equation" toolbar button (gated by caps.equations) opening a
dialog (LaTeX textarea + live preview + insert), and click-to-edit on an existing equation. The OMML
conversion lives in the docx adapter, not the engine.

## Deferred / not done

- odt embedded-object round-trip.
- OMML constructs beyond the common set (matrices, accents, bars, boxes) stay verbatim passthrough.
- Recovering LaTeX from imported MathML (imported equations are edited by retyping).
- Regex-like structural editing; matrices in the authoring UI rely on LaTeX.
