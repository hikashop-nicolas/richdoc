# richdoc verification plan

Started 2026-07-28.

richdoc has around 5,600 lines of its own tests, and no independent judge of any kind.
Its one real-file corpus test asserts only that *text* survives a round-trip, and it
skips in CI because `demo/samples` is gitignored. So today the project checks that it
agrees with itself, and nothing checks that it agrees with the formats.

sheetedit answers this with four judges: its own round-trips, LibreOffice, the ECMA-376
schemas, and openpyxl as an independent reader. This plan brings the same discipline to
richdoc, adjusted for what richdoc actually claims.

## What richdoc claims, and therefore what to check

The central promise is the **passthrough guarantee**: anything the reader does not model
is preserved byte-for-byte, and untouched parts come back untouched. That is a stronger
and more checkable claim than "the text survives", and nothing currently tests it.

The three formats do not make the same promise, so they do not get the same judges:

| Format | Claim | Judges available |
|---|---|---|
| .docx | Preserves untouched parts; edits are surgical | Preservation diff, ECMA-376 XSD, python-docx, LibreOffice |
| .odt | Same | Preservation diff, ODF RELAX NG, odfpy, LibreOffice |
| .doc | **Rebuilds the whole file** from the edited body | LibreOffice only (preservation does not apply) |

That .doc difference is worth stating plainly: the legacy writer regenerates the file, so
a preservation check would fail by design and must not be applied to it.

## Phase 1: a committed corpus (the enabler)

Nothing else can run in CI until fixtures are in the repository. sheetedit has 38
committed fixtures; richdoc has none.

`demo/samples` cannot simply be committed: it is gitignored because real files stay out
of a public repo, and their provenance is not established. So generate a fresh corpus
that is provably synthetic:

- Author each fixture as a **flat ODF (`.fodt`) source**, which is plain XML, committed
  and reviewable as text. Obviously synthetic content, no real names or data.
- A script converts each to `.docx` and `.odt` with LibreOffice.
- **Commit the generated binaries too.** CI must not need LibreOffice merely to have
  fixtures, and a preservation test needs byte-stable inputs: regenerating them on every
  run would make the comparison meaningless.

Coverage to aim for, one fixture per area: paragraphs and runs, styles, lists, tables,
images (inline, wrapped, anchored), footnotes and endnotes, comments, tracked changes,
headers and footers, sections and columns, fields, math, RTL, vertical text.

## Phase 2: the preservation check

The one that tests the actual promise, and it needs no external tool.

For every fixture:

1. **No-edit identity.** Read and write with no edit at all. Every part in the output zip
   must be byte-identical to the input, or the difference must be a known, enumerated
   exception. This is cheap and would catch a whole class of silent damage.
2. **Single-edit blast radius.** Make one targeted edit (change the text of one
   paragraph). Assert that every part other than `word/document.xml` (or
   `content.xml`) is byte-identical, and that within that part only the intended
   paragraph differs.

Both belong in the normal `npm test` run, not behind a separate command.

## Phase 3: schema validation

Mirrors sheetedit's `scripts/schema-check.mjs`, including its baseline trick: real files
draw complaints from the official schemas, so validate the output against the schemas and
report only the violations the *input* did not already have.

- **.docx**: ECMA-376 Part 4 transitional XSDs. The exact bundle sheetedit already
  downloads into `.cache/`; `wml.xsd` covers `document.xml`, `styles.xml`, `numbering.xml`,
  `settings.xml`, `comments.xml`, `footnotes.xml` and `endnotes.xml`.
- **.odt**: the OASIS OpenDocument RELAX NG schema, via `xmllint --relaxng`. Note that
  sheetedit has no ODF validation either, so whatever is learned here should feed back to
  it for `.ods`.

Needs only xmllint, which macOS has and Ubuntu packages as libxml2-utils.

## Phase 4: independent readers

Mirrors the openpyxl cross-check: author a document exercising the writers, record what
it *meant* to write, and have a separate implementation in another language confirm it
reads back that way.

- **.docx**: python-docx (paragraphs, runs, styles, tables, sections, headers).
- **.odt**: odfpy.

Same shape as sheetedit: a vitest file authors the corpus and an `expected.json`, and a
Python script asserts agreement. Reuse `run-openpyxl.mjs`'s interpreter selection so it
works with a local virtual environment as well as CI's system Python.

## Phase 5: LibreOffice oracle

Heaviest in CI, so last. For each fixture, convert both the input and our output with
LibreOffice and compare the results, which proves an independent *implementation*
understands the file rather than merely that it parses.

Two known cautions carry over from sheetedit: LibreOffice silently drops things it does
not model (so a same-format pass, docx to docx, is what proves the file is understood),
and it is the only judge available for `.doc`.

## Order and cost

Phase 1 gates everything. Phase 2 is the highest value per hour and needs nothing
external. Phases 3 and 4 are direct ports of scripts that already exist and are proven in
sheetedit. Phase 5 is the most CI time for the least new information, so it goes last.

## Honest limits

There is no Word on this machine, and no Word in CI. Every claim these checks support is
about the formats and about other implementations, never about what Word itself does.
Say so wherever the results are reported.
