# richdoc

One client-side, framework-agnostic rich-document editor for **.docx** and
**.odt**, in the browser. A single shared engine (toolbar, contenteditable page,
comments panel, track changes, images, header/footer, passthrough) is driven by a
small per-format adapter that does the irreducible work: parse bytes into an
editable model and serialize the model back, preserving everything it does not
model.

No server, no upload: the file is read, edited and rebuilt entirely on the user's
machine.

## Usage

```js
import { createEditor } from "richdoc";

// Sniffs the format (.docx vs .odt) and mounts the right adapter on the engine.
const editor = createEditor(container, bytes, { author: "Jane" });

const out = await editor.getBytes(); // edited file, or the original bytes if untouched
editor.isDirty();
editor.destroy();
```

Format-specific entry points are also exported: `createDocxEditor`,
`createOdtEditor`, plus `sniffFormat(bytes)` and the low-level
`docxToParts` / `htmlToDocx` / `odtToHtml` / `htmlToOdt`.

## Architecture

```
src/core/        shared, format-agnostic engine
  editor.ts        createRichEditor(container, adapter, options) — all the UI
  types.ts         the Adapter / RichDoc / CommentEdits / Capabilities contract
  util.ts          shared helpers (base64, colour/size/font parsing)
  i18n.ts          shared dictionary
src/adapters/
  docx/            OOXML read/write + a thin createDocxAdapter()
  odt/             OpenDocument read/write + a thin createOdtAdapter()
src/index.ts     createEditor() — format sniff + adapter selection
```

The engine knows nothing about OOXML or ODF; an adapter knows nothing about the
toolbar. The three seams an adapter implements are `read()`, `write()` and
`newCommentMarkers()`, plus a `capabilities` flag set.

## Capabilities per format

The toolbar only offers what a format's serializer can actually persist; the rest
is hidden (existing such content is still preserved on save as passthrough).

| Feature | .docx | .odt |
|---|---|---|
| Bold / italic / underline, headings, lists, links | yes | yes |
| Passthrough preservation of unmodelled content | yes | yes |
| Comments (panel, replies, reactions, resolve) | yes | not yet |
| Track changes (suggesting, accept/reject) | yes | not yet |
| Images (insert, resize, delete) | yes | not yet |
| Header / footer editing | yes | not yet |
| Text/background colour, font family/size, alignment | yes | not yet |

The "not yet" cells for .odt are the remaining adapter work (ODF serialization of
those features via named styles); flipping a capability flag re-enables the matching
toolbar control once its `write()` path exists.

## Develop

```bash
npm install
npm test          # vitest: round-trip + mount tests
npm run typecheck
npm run build     # tsc -> dist + copy the shared CSS
node_modules/.bin/vite   # then open /demo/ to try it on a real file
```

## License

MIT
