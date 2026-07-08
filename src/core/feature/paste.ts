// Paste pipeline: external clipboard HTML is normalized onto the editor vocabulary
// (the tags/inline styles the adapters serialize) instead of landing raw, and external
// images are inlined as data URLs so they survive a save (the writers drop non-data
// sources). Internal richdoc copies are recognized and left to the browser's default
// paste, which preserves the full vocabulary. In suggest mode the track-changes
// beforeinput handler owns pasting (plain text), so this module stands down.
import { bytesToBase64 } from "../util";

export interface PasteDeps {
  wrap: HTMLElement;
  regions: HTMLElement[];
  mark: (coalesce?: string) => void;
}

const BLOCK_TAGS = new Set([
  "address", "article", "aside", "blockquote", "caption", "dd", "div", "dl", "dt",
  "fieldset", "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5",
  "h6", "header", "hr", "li", "main", "nav", "ol", "p", "pre", "section", "table",
  "tbody", "td", "tfoot", "th", "thead", "tr", "ul",
]);
const SKIP_TAGS = new Set([
  "script", "style", "noscript", "template", "head", "meta", "link", "title",
  "iframe", "object", "embed", "video", "audio", "svg", "math", "input", "select",
  "textarea", "button", "colgroup", "col",
]);
// Inline styles the writers model (see docx/write.ts appendInline and makeParagraph).
const RUN_STYLES = ["font-weight", "font-style", "text-decoration", "text-decoration-line", "color", "background-color", "font-size", "font-family", "vertical-align"];
const BLOCK_STYLES = ["text-align", "margin-left"];
const SAFE_HREF = /^(https?:|mailto:|#)/i;
const MAX_IMG_W = 600;

type Push = (n: Node) => void;

const hasBlockDescendant = (el: Element): boolean =>
  Array.from(el.children).some((c) => BLOCK_TAGS.has(c.tagName.toLowerCase()) || hasBlockDescendant(c));

/** Clipboard HTML produced by copying inside a richdoc editor: paste it verbatim. */
export function isRichdocHtml(html: string): boolean {
  return /data-docx-(xml|rpr|ppr|numid)|docx-cell|docx-table|docx-ins|docx-del/.test(html);
}

function copyStyles(src: HTMLElement, dst: HTMLElement, props: string[]): void {
  for (const p of props) {
    const v = src.style.getPropertyValue(p).trim();
    if (!v || v === "normal" || v === "none" || v === "inherit" || v === "initial" || v === "unset") continue;
    // A transparent background would round-trip to black shading (toHex6 reads the rgb part).
    if (v === "transparent" || /^rgba\([^)]*,\s*0(\.0+)?\s*\)$/.test(v)) continue;
    if (p === "font-size" && !/^[\d.]+(px|pt)$/.test(v)) continue; // writers parse px/pt only
    dst.style.setProperty(p, v);
  }
}

function sanitizedImg(src: HTMLImageElement): HTMLImageElement | null {
  const url = src.getAttribute("src") ?? "";
  if (!url || /^javascript:/i.test(url)) return null;
  const img = document.createElement("img");
  img.setAttribute("src", url);
  const alt = src.getAttribute("alt");
  if (alt) img.setAttribute("alt", alt);
  const w = Number(src.getAttribute("width")) || Math.round(parseFloat(src.style.width)) || 0;
  const h = Number(src.getAttribute("height")) || Math.round(parseFloat(src.style.height)) || 0;
  if (w > 0) img.setAttribute("width", String(Math.min(w, MAX_IMG_W)));
  if (w > 0 && h > 0) img.setAttribute("height", String(Math.round(Math.min(w, MAX_IMG_W) * (h / w))));
  return img;
}

/** Normalize one clipboard node into editor run vocabulary, feeding `push`. */
function normalizeInlineNode(child: Node, push: Push): void {
  if (child.nodeType === 3) {
    // Collapse HTML source whitespace; it must not reach w:t as literal newlines.
    const text = (child.textContent ?? "").replace(/[\t\n\r ]+/g, " ");
    if (text) push(document.createTextNode(text));
    return;
  }
  if (child.nodeType !== 1) return;
  const el = child as HTMLElement;
  const tag = el.tagName.toLowerCase();
  if (SKIP_TAGS.has(tag)) return;
  if (tag === "br") {
    push(document.createElement("br"));
    return;
  }
  if (tag === "img") {
    const img = sanitizedImg(el as HTMLImageElement);
    if (img) push(img);
    return;
  }
  if (BLOCK_TAGS.has(tag)) {
    // A block inside inline context (malformed HTML): flatten, separated by a break.
    push(document.createElement("br"));
    normalizeInline(el, push);
    return;
  }
  if (tag === "a") {
    const href = el.getAttribute("href") ?? "";
    if (SAFE_HREF.test(href)) {
      const a = document.createElement("a");
      a.setAttribute("href", href);
      normalizeInline(el, (n) => a.appendChild(n));
      if (a.childNodes.length) push(a);
    } else {
      normalizeInline(el, push); // unsafe scheme: keep the text, drop the link
    }
    return;
  }
  if (tag === "ruby") {
    const ruby = document.createElement("ruby");
    for (const rc of Array.from(el.childNodes)) {
      const rt = rc.nodeType === 1 ? (rc as Element).tagName.toLowerCase() : "";
      if (rt === "rt") {
        const nrt = document.createElement("rt");
        normalizeInline(rc, (n) => nrt.appendChild(n));
        ruby.appendChild(nrt);
      } else if (rt !== "rp") {
        normalizeInlineNode(rc, (n) => ruby.appendChild(n));
      }
    }
    if (ruby.childNodes.length) push(ruby);
    return;
  }
  // Semantic tags map to the writer's recognized tags; the Google Docs top wrapper is
  // a <b style="font-weight:normal"> and must not bolden everything.
  const boldWrapper = (tag === "b" || tag === "strong") && /^(normal|400)/.test(el.style.fontWeight);
  let mapped: string | null = null;
  if ((tag === "b" || tag === "strong") && !boldWrapper) mapped = "b";
  else if (tag === "i" || tag === "em") mapped = "i";
  else if (tag === "u" || tag === "ins") mapped = "u";
  else if (tag === "s" || tag === "strike" || tag === "del") mapped = "s";
  else if (tag === "sub" || tag === "sup") mapped = tag;
  if (mapped) {
    const m = document.createElement(mapped);
    copyStyles(el, m, RUN_STYLES);
    normalizeInline(el, (n) => m.appendChild(n));
    if (m.childNodes.length) push(m);
    return;
  }
  // Everything else becomes a styled span, or unwraps when no modeled style survives.
  const span = document.createElement("span");
  copyStyles(el, span, RUN_STYLES);
  if (tag === "code" || tag === "kbd" || tag === "samp" || tag === "tt") span.style.fontFamily = "'Courier New', monospace";
  if (tag === "font" && el.getAttribute("color")) span.style.color = el.getAttribute("color")!;
  if (span.getAttribute("style")) {
    normalizeInline(el, (n) => span.appendChild(n));
    if (span.childNodes.length) push(span);
  } else {
    normalizeInline(el, push);
  }
}

function normalizeInline(node: Node, push: Push): void {
  for (const child of Array.from(node.childNodes)) normalizeInlineNode(child, push);
}

const blockIsEmpty = (el: Element): boolean => !el.querySelector("img,table") && !(el.textContent ?? "").trim();

function tableToEditor(src: HTMLTableElement): HTMLElement | null {
  const rows = Array.from(src.rows);
  if (!rows.length) return null;
  const table = document.createElement("table");
  table.className = "docx-table";
  table.setAttribute("contenteditable", "false");
  for (const tr of rows) {
    const nr = table.insertRow();
    for (const td of Array.from(tr.cells)) {
      const nd = nr.insertCell();
      if (td.colSpan > 1) nd.colSpan = td.colSpan;
      if (td.rowSpan > 1) nd.rowSpan = td.rowSpan;
      const cell = document.createElement("div");
      cell.className = "docx-cell";
      cell.setAttribute("contenteditable", "true");
      const blocks = normalizeBlocks(td);
      if (td.tagName.toLowerCase() === "th") {
        // Header cells keep their emphasis as plain bold runs.
        for (const b of blocks) {
          const bold = document.createElement("b");
          while (b.firstChild) bold.appendChild(b.firstChild);
          b.appendChild(bold);
        }
      }
      if (blocks.length) cell.append(...blocks);
      else cell.innerHTML = "<br>";
      nd.appendChild(cell);
    }
  }
  return table.rows.length ? table : null;
}

function listToEditor(src: HTMLElement): HTMLElement | null {
  const list = document.createElement(src.tagName.toLowerCase());
  for (const li of Array.from(src.children)) {
    if (li.tagName.toLowerCase() !== "li") continue;
    const item = document.createElement("li");
    // The writers expect nested lists inside the <li>, after its inline content.
    const nested: HTMLElement[] = [];
    for (const c of Array.from(li.childNodes)) {
      const ct = c.nodeType === 1 ? (c as Element).tagName.toLowerCase() : "";
      if (ct === "ul" || ct === "ol") {
        const nl = listToEditor(c as HTMLElement);
        if (nl) nested.push(nl);
      } else if (ct && BLOCK_TAGS.has(ct)) {
        if (item.childNodes.length) item.appendChild(document.createElement("br"));
        normalizeInline(c, (n) => item.appendChild(n));
      } else {
        normalizeInlineNode(c, (n) => item.appendChild(n));
      }
    }
    if (!blockIsEmpty(item) || nested.length) {
      item.append(...nested);
      list.appendChild(item);
    }
  }
  return list.children.length ? list : null;
}

/** Normalize a container's children into editor block elements. */
export function normalizeBlocks(container: Node): HTMLElement[] {
  const blocks: HTMLElement[] = [];
  let pending: Node[] = [];
  const push: Push = (n) => pending.push(n);
  const flush = () => {
    if (!pending.length) return;
    const p = document.createElement("p");
    p.append(...pending);
    pending = [];
    if (!blockIsEmpty(p)) blocks.push(p);
  };
  for (const child of Array.from(container.childNodes)) {
    if (child.nodeType === 3) {
      // Whitespace between blocks is layout; whitespace between inline runs is content.
      if (pending.length || (child.textContent ?? "").trim()) normalizeInlineNode(child, push);
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    if (SKIP_TAGS.has(tag) || tag === "hr") continue;
    if (!BLOCK_TAGS.has(tag) && !hasBlockDescendant(el)) {
      normalizeInlineNode(el, push);
      continue;
    }
    flush();
    if (/^h[1-6]$/.test(tag) || tag === "p") {
      const b = document.createElement(tag);
      copyStyles(el, b, BLOCK_STYLES);
      normalizeInline(el, (n) => b.appendChild(n));
      if (!blockIsEmpty(b)) blocks.push(b);
    } else if (tag === "ul" || tag === "ol") {
      const l = listToEditor(el);
      if (l) blocks.push(l);
    } else if (tag === "table") {
      const t = tableToEditor(el as HTMLTableElement);
      if (t) blocks.push(t);
    } else if (tag === "pre") {
      // Code block: one monospace paragraph per line, blank lines kept.
      for (const line of (el.textContent ?? "").replace(/\n$/, "").split("\n")) {
        const p = document.createElement("p");
        p.style.fontFamily = "'Courier New', monospace";
        if (line) p.textContent = line;
        else p.appendChild(document.createElement("br"));
        blocks.push(p);
      }
    } else if (tag === "blockquote") {
      for (const b of normalizeBlocks(el)) {
        if (!b.style.marginLeft) b.style.marginLeft = "40px";
        blocks.push(b);
      }
    } else if (hasBlockDescendant(el)) {
      // Generic container (div/section/the Google Docs <b> wrapper): splice its blocks.
      blocks.push(...normalizeBlocks(el));
    } else {
      // Inline-only container (a plain div line): one paragraph with its block styles.
      const b = document.createElement("p");
      copyStyles(el, b, BLOCK_STYLES);
      normalizeInline(el, (n) => b.appendChild(n));
      if (!blockIsEmpty(b)) blocks.push(b);
    }
  }
  flush();
  return blocks;
}

export interface NormalizedPaste {
  /** Nodes to insert: block elements, or inline nodes when `inline` is true. */
  nodes: Node[];
  inline: boolean;
  /** Images whose src is not a data URL yet (to inline asynchronously). */
  externalImages: HTMLImageElement[];
  /** True when the content is nothing but a single image (prefer the clipboard file). */
  soleImage: boolean;
}

/** Parse and normalize clipboard HTML onto the editor vocabulary. */
export function normalizeClipboardHtml(html: string): NormalizedPaste {
  const body = new DOMParser().parseFromString(html, "text/html").body;
  const blocks = normalizeBlocks(body);
  // A single unstyled paragraph pastes as inline content so it merges at the caret;
  // a styled one (alignment, blockquote indent) keeps its block identity.
  const soleP = blocks.length === 1 && blocks[0]!.tagName.toLowerCase() === "p" && !blocks[0]!.getAttribute("style");
  const nodes: Node[] = soleP ? Array.from(blocks[0]!.childNodes) : blocks;
  const allImgs: HTMLImageElement[] = [];
  for (const n of nodes) {
    if (n.nodeType !== 1) continue;
    if ((n as Element).tagName.toLowerCase() === "img") allImgs.push(n as HTMLImageElement);
    allImgs.push(...Array.from((n as Element).querySelectorAll("img")));
  }
  const text = nodes.map((n) => n.textContent ?? "").join("").trim();
  return {
    nodes,
    inline: soleP,
    externalImages: allImgs.filter((i) => !(i.getAttribute("src") ?? "").startsWith("data:")),
    soleImage: allImgs.length === 1 && !text,
  };
}

/** Cap an image's width/height attributes to the editor's insert size. */
function setImgSize(img: HTMLImageElement, natW: number, natH: number): void {
  let w = Number(img.getAttribute("width")) || natW || 200;
  let h = Number(img.getAttribute("height")) || natH || 200;
  if (w > MAX_IMG_W) {
    h = Math.round((h * MAX_IMG_W) / w);
    w = MAX_IMG_W;
  }
  img.setAttribute("width", String(w));
  img.setAttribute("height", String(h));
}

export function setupPaste(deps: PasteDeps) {
  const { wrap, regions, mark } = deps;

  const blobToDataUrl = (blob: Blob): Promise<string> =>
    new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => res(fr.result as string);
      fr.onerror = () => rej(new Error("read failed"));
      fr.readAsDataURL(blob);
    });

  const probeSize = (dataUrl: string): Promise<{ w: number; h: number }> =>
    new Promise((res) => {
      const probe = new Image();
      probe.onload = () => res({ w: probe.naturalWidth, h: probe.naturalHeight });
      probe.onerror = () => res({ w: 0, h: 0 });
      probe.src = dataUrl;
    });

  const placeCaretAfter = (node: Node): void => {
    const sel = window.getSelection();
    if (!sel) return;
    const r = document.createRange();
    r.setStartAfter(node);
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
  };

  const placeCaretAtEnd = (el: Element): void => {
    const sel = window.getSelection();
    if (!sel) return;
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
  };

  // Fetch an external image into a data URL; on failure (usually CORS) degrade to a
  // link, which is honest: the writers silently drop non-data image sources on save.
  const inlineExternal = async (img: HTMLImageElement): Promise<void> => {
    const url = img.getAttribute("src") ?? "";
    try {
      const resp = await fetch(url, { mode: "cors" });
      if (!resp.ok) throw new Error(String(resp.status));
      const blob = await resp.blob();
      if (!blob.type.startsWith("image/")) throw new Error(blob.type);
      const dataUrl = await blobToDataUrl(blob);
      if (!wrap.isConnected) return;
      const { w, h } = await probeSize(dataUrl);
      img.setAttribute("src", dataUrl);
      setImgSize(img, w, h);
    } catch {
      if (!wrap.isConnected) return;
      const a = document.createElement("a");
      a.setAttribute("href", url);
      a.textContent = img.getAttribute("alt") || url;
      img.replaceWith(a);
    }
    mark();
  };

  const insertImageFile = async (file: File, range: Range): Promise<void> => {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const dataUrl = `data:${file.type};base64,${bytesToBase64(bytes)}`;
    const { w, h } = await probeSize(dataUrl);
    if (!wrap.isConnected) return;
    const img = document.createElement("img");
    img.setAttribute("src", dataUrl);
    setImgSize(img, w, h);
    range.collapse(false);
    range.insertNode(img);
    placeCaretAfter(img);
    mark();
  };

  // The block that holds the caret: the ancestor whose parent is a block host (the
  // region, a table cell, or a pagination wrapper the layout pass may have added).
  const isHost = (el: Element, region: HTMLElement): boolean =>
    el === region || el.classList.contains("docx-cell") || el.classList.contains("docxedit-secpage") || el.classList.contains("docxedit-colpage");
  const caretBlock = (range: Range, region: HTMLElement): HTMLElement | null => {
    let el: Element | null = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : (range.startContainer as Element);
    while (el && !isHost(el, region) && !(el.parentElement && isHost(el.parentElement, region))) el = el.parentElement;
    return el && !isHost(el, region) ? (el as HTMLElement) : null;
  };

  const insertInline = (range: Range, nodes: Node[]): void => {
    const frag = document.createDocumentFragment();
    frag.append(...nodes);
    const last = frag.lastChild;
    range.insertNode(frag);
    if (last) placeCaretAfter(last);
  };

  const insertBlocks = (range: Range, blocks: HTMLElement[], region: HTMLElement): void => {
    const block = caretBlock(range, region);
    if (!block) {
      // Caret sits directly in a host (a bare table cell): insert in place.
      insertInline(range, blocks);
      return;
    }
    const li = block.closest("li");
    if (li && region.contains(li)) {
      // Inside a list: paragraphs become sibling items; tables land after the list.
      const items: HTMLElement[] = [];
      const after: HTMLElement[] = [];
      for (const b of blocks) {
        const t = b.tagName.toLowerCase();
        if (t === "ul" || t === "ol") items.push(...(Array.from(b.children) as HTMLElement[]));
        else if (t === "table") after.push(b);
        else {
          const item = document.createElement("li");
          while (b.firstChild) item.appendChild(b.firstChild);
          items.push(item);
        }
      }
      let anchor: Element = li;
      for (const it of items) {
        anchor.after(it);
        anchor = it;
      }
      const list = li.closest("ul,ol") ?? li;
      for (const t of after.reverse()) list.after(t);
      if (blockIsEmpty(li) && items.length) li.remove();
      const target = items[items.length - 1] ?? after[after.length - 1];
      if (target) placeCaretAtEnd(target);
      return;
    }
    // Split the block at the caret and place the pasted blocks in between.
    const tail = block.cloneNode(false) as HTMLElement;
    const tailRange = document.createRange();
    tailRange.selectNodeContents(block);
    tailRange.setStart(range.endContainer, range.endOffset);
    tail.appendChild(tailRange.extractContents());
    let anchor: Element = block;
    for (const b of blocks) {
      anchor.after(b);
      anchor = b;
    }
    if (!blockIsEmpty(tail)) anchor.after(tail);
    if (blockIsEmpty(block)) block.remove();
    placeCaretAtEnd(blocks[blocks.length - 1]!);
  };

  const onPaste = (e: ClipboardEvent, region: HTMLElement): void => {
    if (wrap.dataset.rdocSuggesting === "1") return; // track-changes owns paste in suggest mode
    const cd = e.clipboardData;
    if (!cd) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !region.contains(sel.getRangeAt(0).startContainer)) return;
    const html = cd.getData("text/html");
    if (html && isRichdocHtml(html)) return; // internal copy: default paste keeps the vocabulary
    const imgFile = Array.from(cd.files).find((f) => f.type.startsWith("image/")) ?? null;
    const normalized = html ? normalizeClipboardHtml(html) : null;
    // A clipboard image file wins when the HTML is nothing but that image (no CORS fetch).
    if (imgFile && (!normalized || normalized.soleImage || !normalized.nodes.length)) {
      e.preventDefault();
      const range = sel.getRangeAt(0);
      if (!range.collapsed) range.deleteContents();
      void insertImageFile(imgFile, range.cloneRange());
      return;
    }
    if (!normalized || !normalized.nodes.length) return; // plain text: default paste is fine
    e.preventDefault();
    const range = sel.getRangeAt(0);
    if (!range.collapsed) range.deleteContents();
    if (normalized.inline) insertInline(range, normalized.nodes);
    else insertBlocks(range, normalized.nodes as HTMLElement[], region);
    mark();
    for (const img of normalized.externalImages) void inlineExternal(img);
  };

  for (const r of regions) r.addEventListener("paste", (e) => onPaste(e, r));
}
