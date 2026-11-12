# Third-party notice: WordCraft

These richdoc files are TypeScript ports of code from WordCraft, https://github.com/storytold/wordcraft
(dual-licensed MIT OR Apache-2.0), used under the MIT license below:

- src/adapters/docx/metafile.ts: the `wordcraft-metafile` crate (EMF / WMF player). WMF text, the
  SVG output and the PNG encoder were added in richdoc.
- src/adapters/docx/smartart.ts: the SmartArt drawing reader (`read/diagram.rs`) and its preset
  geometries.
- src/adapters/docx/wordart.ts: the WordArt reader and text warps (`read/wordart.rs`, `layout/wordart.rs`).
- src/adapters/docx/chart.ts and chart-spec.ts: the chart reader and chart model (`read/chart/`, `chart_spec.rs`).

MIT License

Copyright (c) 2026 ArtCraft Team and the WordCraft contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
