# richdoc

One client-side, framework-agnostic rich-document editor for **.docx** and
**.odt**, in the browser. A single shared engine (paginated page view, toolbar,
contenteditable, comments panel, track changes, images, header/footer, margin
rulers, zoom, passthrough) is driven by a small per-format adapter that does the
irreducible work: parse bytes into an editable model and serialize the model
back, preserving everything it does not model.

No server, no upload: the file is read, edited and rebuilt entirely on the
user's machine.

## Usage

```js
import { createEditor } from "richdoc";

// Sniffs the format (.docx vs .odt) and mounts the right adapter on the engine.
const editor = createEditor(container, bytes, { author: "Jane" });

const out = await editor.getBytes(); // edited file, or the original bytes if untouched
editor.isDirty();
editor.destroy();
```

The package root also exports `createDocxEditor`, `createOdtEditor`,
`sniffFormat(bytes)` and `setLocale(...)`. Each adapter module additionally
exports its low-level parse/serialize functions (`docxToParts` / `htmlToDocx`,
`odtToHtml` / `htmlToOdt`).

## Architecture

The engine knows nothing about OOXML or ODF; an adapter knows nothing about the
UI. An adapter implements three seams plus a capability set: `read()`, `write()`,
`newCommentMarkers()` and `capabilities`. The engine is split by feature so the
orchestrator stays small.

```
src/core/
  editor.ts          createRichEditor(container, adapter, options): the orchestrator
                       (DOM, the pagination render loop, lifecycle, feature wiring)
  feature/
    toolbar.ts         formatting controls + the overflow "more" row
    comments.ts        comments side panel + comment-edit bookkeeping
    track-changes.ts   suggestion mode (ins/del marks, accept/reject)
    images.ts          image select / resize / delete / insert
    page-view.ts       margin rulers + zoom + the centred canvas
  page.ts            page geometry + the pure paginate()
  types.ts           the Adapter / RichDoc / CommentEdits / Capabilities contract
  util.ts, i18n.ts   shared helpers + the multilingual dictionary
src/adapters/
  docx/              OOXML, split into shared / read / write / index
  odt/               OpenDocument, split into shared / read / write / index
src/index.ts         createEditor(): format sniff + adapter selection
```

## Capabilities per format

The toolbar only offers what a format's serializer can persist; the rest is
hidden (existing such content is still preserved on save as passthrough).

| Feature | .docx | .odt |
|---|---|---|
| Bold / italic / underline, headings, lists, links | yes | yes |
| Text/background colour, font family/size, alignment | yes | yes |
| Images (insert, resize, delete) | yes | yes |
| Header / footer (edit, and create by double-clicking the margin) | yes | yes |
| Track changes (suggesting, accept/reject) | yes | yes |
| Comments | panel, replies, reactions, resolve | panel + resolve |
| Manual page breaks | yes | no |
| Passthrough preservation of unmodelled content | yes | yes |

A capability flag gates its toolbar control, so a feature appears for a format
only once that format's `write()` path can persist it.

## Develop

```bash
npm install
npm test          # vitest: round-trip + mount tests
npm run typecheck
npm run build     # tsc -> dist + copy the shared CSS
node_modules/.bin/vite   # then open /demo/ to try it on a real file
```

## License

MIT, see [LICENSE](LICENSE).
