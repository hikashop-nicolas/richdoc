# richdoc: legacy `.doc` (Word 97-2003 binary) read + write adapter

Date: 2026-07-10. Goal: a third richdoc adapter, `src/adapters/doc/`, that reads and
writes the Word 97-2003 binary format (MS-DOC), mapping it to richdoc's HTML model so
the whole shared engine (toolbar, editing, history) works, exactly as docx and odt do.

## Why this is different from docx/odt (and how we handle it)

docx/odt are XML-in-zip, so their writers edit in place: rewrite only the spans you
touched, keep every other byte identical. `.doc` is a binary OLE compound file whose
text position and formatting live in separate offset-indexed tables that all
cross-reference each other, so an in-place span edit is not possible. Therefore the
`.doc` writer is **from-scratch (approach 2)**: it regenerates a complete, valid `.doc`
from richdoc's HTML model on every save. This is lossy for anything the model does not
represent (comments, tracked changes, macros, exotic formatting), the same way odt's
writer is capability-gated today. That is an accepted, documented trade-off.

The read half is a normal binary parser (no such constraint) and can be fairly complete.

## Validation strategy (the thing that makes this tractable)

macOS ships `textutil`, which reads AND writes `.doc` with Apple's real document engine.
We use it as an oracle:
- **Reader tests:** `textutil -convert doc in.html -o ref.doc` gives known-good fixtures;
  our reader must extract the same text + formatting.
- **Writer tests:** our writer emits `out.doc`; `textutil -convert txt out.doc -stdout`
  (and `-convert html`) must round-trip the content. If Apple's engine reads our file,
  it is a valid `.doc`.
- We also **dissect** a minimal `textutil`-produced `.doc` and mirror its exact structure
  in the writer, rather than implementing the full MS-DOC spec blind. (Already done, see
  "Reference structure" below.)

## Reference structure (dissected from a real `textutil` min.doc)

- CFB: 512-byte sectors, v3, mini cutoff 4096. Streams (all padded to 4096 => regular
  sectors, so no mini-stream needed for the doc streams): `WordDocument`, `1Table`,
  `\x05SummaryInformation`, `\x05DocumentSummaryInformation`, plus the Root Entry's small
  mini stream. Table stream is `1Table` (FIB `fWhichTblStm` = 1).
- FIB: `wIdent` 0xA5EC, `nFib` 0x00C1 (Word 97). `ccpText` at fibRgLw index 3. Text is
  16-bit Unicode (piece `fCompressed` = 0) starting at fc 1536 (0x600) in `WordDocument`.
- Table stream layout (sequential): STSH, PlcfSed (1 section), PlcfBteChpx, PlcfBtePapx,
  Clx (piece table), Sttbfffn (font table).
- Piece table (Clx): a Pcdt (0x02) + lcb + PlcPcd = aCP[n+1] + aPcd[n]; each PCD carries
  the fc (bit 30 = fCompressed) into `WordDocument`.

This exact skeleton is what the writer reproduces, with text/formatting parameterised.

## Module layout (`src/adapters/doc/`)

- `cfb.ts` — self-contained MS-CFB reader + writer (header, FAT, directory, mini-stream).
  No new dependency (richdoc keeps only fflate + temml). Ported from a validated Python
  prototype.
- `fib.ts` — FIB read/build: the fc/lcb table (FibRgFcLcb97 field offsets), ccp fields,
  table-stream selection.
- `sprm.ts` — sprm opcode table + operand-size decoder, for the character/paragraph
  property grammar (read) and encoder (write).
- `read.ts` — bytes -> HTML: CFB -> FIB -> piece table -> text; PlcfBteChpx/PlcfBtePapx ->
  FKPs -> per-run/para formatting; STSH -> style names -> headings; lists -> HTML using
  richdoc's vocabulary (`<p>`, `<span style=...>`, `<h1-6>`/`data-rdoc-style`, `<ul>/<ol>`,
  `<a>`). Mirrors docx/odt `read.ts`.
- `write.ts` — HTML -> bytes: parse the edited body via DOMParser, walk paragraphs/runs,
  build the text stream + single-piece Clx + CHPX/PAPX FKPs + Plc tables + STSH + minimal
  Summary streams + FIB, and assemble via `cfb.ts`. Validated with `textutil`.
- `index.ts` — `createDocAdapter` / `createDocEditor`, capabilities. Mirrors odt/index.ts.
- `index.test.ts` — reader (vs textutil fixtures) + writer (textutil round-trip) + a
  reader(writer(html)) identity test.

Top-level `src/index.ts`: `sniffFormat` gains a CFB check (magic D0CF11E0A1B11AE1 =>
"doc") and `createEditor` routes `.doc` to `createDocEditor`.

## Capabilities (initial, expand later)

Target the subset that maps cleanly both directions, matching richdoc's HTML vocabulary:
- Text; **bold / italic / underline / strike**; **font family + size**; **text colour +
  highlight**; **paragraph alignment**; **left indent**; **headings** (Heading 1-6 styles);
  **bullet / numbered lists**; **hyperlinks**; **line breaks**.
- Off initially (binary `.doc` cost is high): comments, track changes, images, equations,
  tables (render existing read-only if cheap; authoring off), sections/headers/footers,
  fields, footnotes, page geometry. These become follow-ups, same staged path as odt.

## Phased build

- **A. CFB codec** (cfb.ts) + tests: round-trip a compound file byte-exact; read
  textutil's min.doc streams. (Python prototype already parses it correctly.)
- **B. Reader text** (read.ts): FIB + piece table -> text -> `<p>` split on 0x0D (para) /
  0x0C (page/section) / 0x07 (cell) / 0x0B (line break). Verify vs textutil fixtures.
- **C. Reader formatting**: CHPX (bold/italic/underline/strike/font ftc/size hps/colour
  ico+cv/highlight) and PAPX (jc alignment, dxaLeft indent, istd style -> heading) FKPs;
  list detection; hyperlinks (fldChar/HYPERLINK field or the simpler `\x13..\x14..\x15`
  field runs). Map to HTML.
- **D. Writer**: HTML -> from-scratch `.doc`. Text as 16-bit Unicode; one piece; build
  CHPX/PAPX FKPs from the run/para formatting; PlcfBteChpx/PlcfBtePapx; PlcfSed (1
  section); STSH (Normal + heading styles); Sttbfffn (fonts used); FIB; Summary streams
  (minimal, or templated from textutil). Validate: textutil reads it; reader(writer(h)) ==
  normalised h.
- **E. Wire-up**: adapter + capabilities + sniffFormat; Omnitext `.doc` format ->
  richdoc doc editor; end-to-end browser verify (open real .doc, edit, save, reopen).

## Risks / open questions

- Getting every FIB fc/lcb offset right: mitigated by mirroring the dissected reference
  and cross-checking against textutil files.
- Word (not just textutil/LibreOffice) opening our output: we optimise for textutil +
  LibreOffice (both real engines); if a Windows Word is ever available, add it to CI.
- Formatting fidelity is a subset by design; unmapped properties are dropped on write
  (documented, capability-gated), never silently corrupting the file.
- List numbering (LFO/LST tables) is complex; initial write may emit bullets/numbers as
  paragraph-level formatting good enough for textutil, refine later.
