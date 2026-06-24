// odt WRITE: rebuild an .odt archive from edited HTML, preserving every untouched part.
// Pure HTML -> XML (body, styles, header/footer, comments, tracked changes, page margins);
// the read half lives in ./read.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { firstFontFamily, fontSizeToHalfPt, toHex6 } from "../../core/util";
import type { PageGeometry } from "../../core/types";
import { NS, fmtKey, FMT0, ODF_ALIGN, importPassthrough, IMG_MIME } from "./shared";
import type { Fmt } from "./shared";

const pxToCm = (px: number): string => `${Math.round((px / (96 / 2.54)) * 1000) / 1000}cm`;

/** Ensure <office:automatic-styles> exists, returning it. The new node is placed before
   `beforeTag` (office:body in content.xml, office:master-styles in styles.xml). */
function ensureAutoStyles(doc: Document, beforeTag = "office:body"): Element {
  let auto = doc.getElementsByTagName("office:automatic-styles")[0];
  if (auto) return auto;
  auto = doc.createElementNS(NS.office, "office:automatic-styles");
  const before = doc.getElementsByTagName(beforeTag)[0];
  doc.documentElement.insertBefore(auto, before ?? null);
  return auto;
}

/** Create (once) a text style for a run-formatting combo and return its name. */
function styleFor(doc: Document, auto: Element, created: Map<string, string>, f: Fmt): string | null {
  const key = fmtKey(f);
  if (!key) return null;
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_t${created.size}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "text");
  const tp = doc.createElementNS(NS.style, "style:text-properties");
  if (f.b) tp.setAttributeNS(NS.fo, "fo:font-weight", "bold");
  if (f.i) tp.setAttributeNS(NS.fo, "fo:font-style", "italic");
  if (f.u) {
    tp.setAttributeNS(NS.style, "style:text-underline-style", "solid");
    tp.setAttributeNS(NS.style, "style:text-underline-width", "auto");
    tp.setAttributeNS(NS.style, "style:text-underline-color", "font-color");
  }
  if (f.strike) {
    tp.setAttributeNS(NS.style, "style:text-line-through-style", "solid");
    tp.setAttributeNS(NS.style, "style:text-line-through-type", "single");
  }
  if (f.vertAlign) tp.setAttributeNS(NS.style, "style:text-position", f.vertAlign === "super" ? "super 58%" : "sub 58%");
  if (f.color) tp.setAttributeNS(NS.fo, "fo:color", `#${f.color}`);
  if (f.bg) tp.setAttributeNS(NS.fo, "fo:background-color", `#${f.bg}`);
  if (f.font) {
    tp.setAttributeNS(NS.fo, "fo:font-family", f.font);
    tp.setAttributeNS(NS.style, "style:font-name", f.font);
  }
  if (f.sizePt) tp.setAttributeNS(NS.fo, "fo:font-size", `${f.sizePt}pt`);
  st.appendChild(tp);
  auto.appendChild(st);
  created.set(key, name);
  return name;
}

/** Create (once) a paragraph style for an alignment and return its name. */
function paraStyleFor(doc: Document, auto: Element, created: Map<string, string>, align: string): string | null {
  const a = ODF_ALIGN[align];
  if (!a || a === "left") return null;
  const key = `p_${a}`;
  const existing = created.get(key);
  if (existing) return existing;
  const name = `OT_${key}`;
  const st = doc.createElementNS(NS.style, "style:style");
  st.setAttributeNS(NS.style, "style:name", name);
  st.setAttributeNS(NS.style, "style:family", "paragraph");
  const pp = doc.createElementNS(NS.style, "style:paragraph-properties");
  pp.setAttributeNS(NS.fo, "fo:text-align", a === "right" ? "end" : a === "center" ? "center" : a === "justify" ? "justify" : "start");
  st.appendChild(pp);
  auto.appendChild(st);
  created.set(key, name);
  return name;
}

interface RefMeta {
  author: string;
  date: string;
  text: string;
  resolved: boolean;
  paraId: string;
}
interface OdfCtx {
  doc: Document;
  auto: Element;
  created: Map<string, string>;
  files: Record<string, Uint8Array>; // the archive, so embedded images can be added
  pics: { path: string; mime: string }[]; // images added this run, for the manifest
  refMeta: Map<string, RefMeta>; // comment id -> metadata, gathered from the body refs
  rangedIds: Set<string>; // comment ids that wrap a text range (vs a point comment)
  done: Map<string, boolean>; // resolve state keyed by comment paraId
  changes: { id: string; type: "insertion" | "deletion"; author: string; date: string; deleted?: string }[]; // tracked changes to emit
}

/** Build an office:annotation element for a comment id from its gathered metadata. */
function makeAnnotation(ctx: OdfCtx, id: string): Element {
  const m = ctx.refMeta.get(id);
  const an = ctx.doc.createElementNS(NS.office, "office:annotation");
  an.setAttributeNS(NS.office, "office:name", id);
  const resolved = (m && ctx.done.get(m.paraId)) ?? m?.resolved ?? false;
  if (resolved) an.setAttributeNS(NS.loext, "loext:resolved", "true");
  if (m?.author) {
    const cr = ctx.doc.createElementNS(NS.dc, "dc:creator");
    cr.textContent = m.author;
    an.appendChild(cr);
  }
  if (m?.date) {
    const dt = ctx.doc.createElementNS(NS.dc, "dc:date");
    dt.textContent = m.date;
    an.appendChild(dt);
  }
  for (const line of (m?.text ?? "").split("\n")) {
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    if (line) p.textContent = line;
    an.appendChild(p);
  }
  return an;
}

/** Turn an <img> (data URL) into a draw:frame, embedding the bytes in the archive. */
function buildImageFrame(img: HTMLElement, ctx: OdfCtx): Element | null {
  const m = /^data:([^;]+);base64,(.+)$/.exec(img.getAttribute("src") ?? "");
  if (!m) return null; // not an embeddable image (e.g. an external URL): drop it
  const mime = m[1]!;
  const bin = atob(m[2]!);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const ext = Object.entries(IMG_MIME).find(([, v]) => v === mime)?.[0] ?? (mime.split("/")[1] || "png");
  const idx = ctx.pics.length;
  const path = `Pictures/ot_img${idx}.${ext}`;
  ctx.files[path] = bytes;
  ctx.pics.push({ path, mime });

  const wPx = parseFloat(img.getAttribute("width") ?? "") || undefined;
  const hPx = parseFloat(img.getAttribute("height") ?? "") || undefined;
  const frame = ctx.doc.createElementNS(NS.draw, "draw:frame");
  frame.setAttributeNS(NS.draw, "draw:name", `Image${idx + 1}`);
  frame.setAttributeNS(NS.text, "text:anchor-type", "as-char");
  if (wPx) frame.setAttributeNS(NS.svg, "svg:width", pxToCm(wPx));
  if (hPx) frame.setAttributeNS(NS.svg, "svg:height", pxToCm(hPx));
  const image = ctx.doc.createElementNS(NS.draw, "draw:image");
  image.setAttributeNS(NS.xlink, "xlink:href", path);
  image.setAttributeNS(NS.xlink, "xlink:type", "simple");
  image.setAttributeNS(NS.xlink, "xlink:show", "embed");
  image.setAttributeNS(NS.xlink, "xlink:actuate", "onLoad");
  frame.appendChild(image);
  return frame;
}

/** Register newly embedded images in META-INF/manifest.xml. */
function addManifestEntries(files: Record<string, Uint8Array>, pics: { path: string; mime: string }[]): void {
  if (!pics.length || !files["META-INF/manifest.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["META-INF/manifest.xml"]), "application/xml");
  const root = doc.getElementsByTagName("manifest:manifest")[0] ?? doc.documentElement;
  if (!root) return;
  const have = new Set(Array.from(doc.getElementsByTagName("manifest:file-entry")).map((e) => e.getAttribute("manifest:full-path")));
  for (const p of pics) {
    if (have.has(p.path)) continue;
    const e = doc.createElementNS(NS.manifest, "manifest:file-entry");
    e.setAttributeNS(NS.manifest, "manifest:full-path", p.path);
    e.setAttributeNS(NS.manifest, "manifest:media-type", p.mime);
    root.appendChild(e);
  }
  files["META-INF/manifest.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Append the inline content of an HTML node to an ODF block element. */
function htmlInlineToOdf(node: Node, parent: Element, f: Fmt, ctx: OdfCtx): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      const txt = child.textContent ?? "";
      if (!txt) continue;
      if (!fmtKey(f)) {
        parent.appendChild(ctx.doc.createTextNode(txt));
      } else {
        const span = ctx.doc.createElementNS(NS.text, "text:span");
        const name = styleFor(ctx.doc, ctx.auto, ctx.created, f);
        if (name) span.setAttributeNS(NS.text, "text:style-name", name);
        span.appendChild(ctx.doc.createTextNode(txt));
        parent.appendChild(span);
      }
      continue;
    }
    if (child.nodeType !== 1) continue;
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    const stash = el.getAttribute("data-odt-xml");
    if (stash) {
      const node2 = importPassthrough(ctx.doc, stash);
      if (node2) parent.appendChild(node2);
      continue;
    }
    if (tag === "br") {
      parent.appendChild(ctx.doc.createElementNS(NS.text, "text:line-break"));
      continue;
    }
    if (tag === "img") {
      const frame = buildImageFrame(el, ctx);
      if (frame) parent.appendChild(frame);
      continue;
    }
    if (el.classList.contains("docx-comment")) {
      // commented range: annotation at the start, the text, then annotation-end
      const id = el.getAttribute("data-comment-id") ?? "";
      parent.appendChild(makeAnnotation(ctx, id));
      htmlInlineToOdf(el, parent, f, ctx);
      const end = ctx.doc.createElementNS(NS.office, "office:annotation-end");
      end.setAttributeNS(NS.office, "office:name", id);
      parent.appendChild(end);
      continue;
    }
    if (el.classList.contains("docx-comment-ref")) {
      // metadata carrier: emit a point annotation only if it has no wrapped range
      const id = el.getAttribute("data-comment-id") ?? "";
      if (!ctx.rangedIds.has(id)) parent.appendChild(makeAnnotation(ctx, id));
      continue;
    }
    if (el.classList.contains("docx-ins")) {
      // tracked insertion: change-start ... text ... change-end, region built later
      const id = `ct${ctx.changes.length + 1}`;
      ctx.changes.push({ id, type: "insertion", author: el.getAttribute("data-author") ?? "", date: el.getAttribute("data-date") ?? "" });
      const start = ctx.doc.createElementNS(NS.text, "text:change-start");
      start.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(start);
      htmlInlineToOdf(el, parent, f, ctx);
      const end = ctx.doc.createElementNS(NS.text, "text:change-end");
      end.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(end);
      continue;
    }
    if (el.classList.contains("docx-del")) {
      // tracked deletion: a point marker in the body; the removed text lives in the region
      const id = `ct${ctx.changes.length + 1}`;
      ctx.changes.push({ id, type: "deletion", author: el.getAttribute("data-author") ?? "", date: el.getAttribute("data-date") ?? "", deleted: el.textContent ?? "" });
      const mark = ctx.doc.createElementNS(NS.text, "text:change");
      mark.setAttributeNS(NS.text, "text:change-id", id);
      parent.appendChild(mark);
      continue;
    }
    if (el.classList.contains("docx-cmark")) continue; // empty new-comment marker: nothing to emit
    if (tag === "a") {
      const a = ctx.doc.createElementNS(NS.text, "text:a");
      a.setAttributeNS(NS.xlink, "xlink:href", el.getAttribute("href") ?? "");
      htmlInlineToOdf(el, a, f, ctx);
      parent.appendChild(a);
      continue;
    }
    const hp = fontSizeToHalfPt(el.style.fontSize);
    const next: Fmt = {
      b: f.b || tag === "strong" || tag === "b" || /(^|;)\s*font-weight\s*:\s*(bold|[6-9]00)/.test(el.style.cssText),
      i: f.i || tag === "em" || tag === "i" || el.style.fontStyle === "italic",
      u: f.u || tag === "u" || /underline/.test(el.style.textDecoration || el.style.textDecorationLine || ""),
      strike: f.strike || tag === "s" || tag === "strike" || tag === "del" || /line-through/.test(el.style.textDecoration || el.style.textDecorationLine || ""),
      vertAlign: f.vertAlign ?? (tag === "sup" || /vertical-align:\s*super/.test(el.style.cssText) ? "super" : tag === "sub" || /vertical-align:\s*sub/.test(el.style.cssText) ? "sub" : undefined),
      color: toHex6(el.style.color) ?? f.color,
      bg: toHex6(el.style.backgroundColor) ?? f.bg,
      font: firstFontFamily(el.style.fontFamily) ?? f.font,
      sizePt: hp ? hp / 2 : f.sizePt,
    };
    htmlInlineToOdf(el, parent, next, ctx);
  }
}

function htmlListToOdf(el: HTMLElement, ctx: OdfCtx): Element {
  const list = ctx.doc.createElementNS(NS.text, "text:list");
  for (const li of Array.from(el.children)) {
    if (li.tagName.toLowerCase() !== "li") continue;
    const item = ctx.doc.createElementNS(NS.text, "text:list-item");
    const nested = li.querySelector(":scope > ul, :scope > ol");
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    htmlInlineToOdf(li, p, FMT0, ctx);
    item.appendChild(p);
    if (nested) item.appendChild(htmlListToOdf(nested as HTMLElement, ctx));
    list.appendChild(item);
  }
  return list;
}

function htmlBlockToOdf(node: Node, ctx: OdfCtx): Element | null {
  if (node.nodeType === 3) {
    if (!(node.textContent ?? "").trim()) return null;
    const p = ctx.doc.createElementNS(NS.text, "text:p");
    p.appendChild(ctx.doc.createTextNode(node.textContent ?? ""));
    return p;
  }
  if (node.nodeType !== 1) return null;
  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  const stash = el.getAttribute("data-odt-xml");
  if (stash) return importPassthrough(ctx.doc, stash);
  if (tag === "ul" || tag === "ol") return htmlListToOdf(el, ctx);
  const applyAlign = (block: Element): void => {
    const name = el.style.textAlign ? paraStyleFor(ctx.doc, ctx.auto, ctx.created, el.style.textAlign) : null;
    if (name) block.setAttributeNS(NS.text, "text:style-name", name);
  };
  const m = /^h([1-6])$/.exec(tag);
  if (m) {
    const h = ctx.doc.createElementNS(NS.text, "text:h");
    h.setAttributeNS(NS.text, "text:outline-level", String(Math.min(3, Number(m[1]))));
    applyAlign(h);
    htmlInlineToOdf(el, h, FMT0, ctx);
    return h;
  }
  // p, div, and anything else become a paragraph
  const p = ctx.doc.createElementNS(NS.text, "text:p");
  applyAlign(p);
  htmlInlineToOdf(el, p, FMT0, ctx);
  return p;
}

/** Return the document's master page, creating a minimal "Standard" one (with a page
    layout to reference) if the file has none, so a header/footer can be added from
    scratch. Most ODF files already have one; this covers the rare file that does not. */
function ensureMasterPage(doc: Document): Element | null {
  const existing = doc.getElementsByTagName("style:master-page")[0];
  if (existing) return existing;
  const root = doc.documentElement;
  if (!root) return null;
  let pl = doc.getElementsByTagName("style:page-layout")[0];
  let plName = pl?.getAttribute("style:name") ?? null;
  if (!plName) {
    pl = doc.createElementNS(NS.style, "style:page-layout");
    plName = "pm-rdoc";
    pl.setAttributeNS(NS.style, "style:name", plName);
    pl.appendChild(doc.createElementNS(NS.style, "style:page-layout-properties"));
    ensureAutoStyles(doc, "office:master-styles").appendChild(pl);
  }
  let ms = doc.getElementsByTagName("office:master-styles")[0];
  if (!ms) {
    ms = doc.createElementNS(NS.office, "office:master-styles");
    root.appendChild(ms);
  }
  const master = doc.createElementNS(NS.style, "style:master-page");
  master.setAttributeNS(NS.style, "style:name", "Standard");
  master.setAttributeNS(NS.style, "style:page-layout-name", plName);
  ms.appendChild(master);
  return master;
}

/** Write edited header/footer HTML back into the master page in styles.xml. */
function applyHeaderFooter(files: Record<string, Uint8Array>, parts: { path: string; html: string }[]): void {
  const hf = parts.filter((p) => p.path === "header" || p.path === "footer");
  if (!hf.length || !files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const master = ensureMasterPage(doc);
  if (!master) return;
  const ctx: OdfCtx = {
    doc,
    auto: ensureAutoStyles(doc, "office:master-styles"),
    created: new Map(),
    files,
    pics: [],
    refMeta: new Map(),
    rangedIds: new Set(),
    done: new Map(),
    changes: [],
  };
  for (const p of hf) {
    const tag = p.path === "header" ? "style:header" : "style:footer";
    let el = master.getElementsByTagName(tag)[0];
    if (!el) {
      el = doc.createElementNS(NS.style, tag);
      master.appendChild(el);
    }
    while (el.firstChild) el.removeChild(el.firstChild);
    const htmlDoc = new DOMParser().parseFromString(p.html || "<p><br></p>", "text/html");
    for (const node of Array.from(htmlDoc.body.childNodes)) {
      const block = htmlBlockToOdf(node, ctx);
      if (block) el.appendChild(block);
    }
    if (!el.firstChild) el.appendChild(doc.createElementNS(NS.text, "text:p"));
  }
  addManifestEntries(files, ctx.pics);
  files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Build the <text:tracked-changes> region from the changes collected while serializing. */
function buildTrackedChanges(ctx: OdfCtx): Element | null {
  if (!ctx.changes.length) return null;
  const tc = ctx.doc.createElementNS(NS.text, "text:tracked-changes");
  for (const ch of ctx.changes) {
    const region = ctx.doc.createElementNS(NS.text, "text:changed-region");
    region.setAttributeNS(NS.text, "text:id", ch.id);
    const kind = ctx.doc.createElementNS(NS.text, ch.type === "insertion" ? "text:insertion" : "text:deletion");
    const info = ctx.doc.createElementNS(NS.office, "office:change-info");
    if (ch.author) {
      const cr = ctx.doc.createElementNS(NS.dc, "dc:creator");
      cr.textContent = ch.author;
      info.appendChild(cr);
    }
    if (ch.date) {
      const dt = ctx.doc.createElementNS(NS.dc, "dc:date");
      dt.textContent = ch.date;
      info.appendChild(dt);
    }
    kind.appendChild(info);
    if (ch.type === "deletion") {
      for (const line of (ch.deleted ?? "").split("\n")) {
        const p = ctx.doc.createElementNS(NS.text, "text:p");
        if (line) p.textContent = line;
        kind.appendChild(p);
      }
    }
    region.appendChild(kind);
    tc.appendChild(region);
  }
  return tc;
}

/** Update the page-layout margins in styles.xml from edited margins (px -> cm). */
function applyPageMargins(files: Record<string, Uint8Array>, geometry: PageGeometry): void {
  if (!files["styles.xml"]) return;
  const doc = new DOMParser().parseFromString(strFromU8(files["styles.xml"]), "application/xml");
  const pl = doc.getElementsByTagName("style:page-layout")[0];
  if (!pl) return;
  let props = pl.getElementsByTagName("style:page-layout-properties")[0];
  if (!props) {
    props = doc.createElementNS(NS.style, "style:page-layout-properties");
    pl.insertBefore(props, pl.firstChild);
  }
  props.setAttributeNS(NS.fo, "fo:margin-top", pxToCm(geometry.margin.top));
  props.setAttributeNS(NS.fo, "fo:margin-right", pxToCm(geometry.margin.right));
  props.setAttributeNS(NS.fo, "fo:margin-bottom", pxToCm(geometry.margin.bottom));
  props.setAttributeNS(NS.fo, "fo:margin-left", pxToCm(geometry.margin.left));
  files["styles.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
}

/** Rebuild an .odt from edited HTML, preserving every other part of the archive. */
export function htmlToOdt(
  html: string,
  original: Uint8Array,
  opts?: { done?: Map<string, boolean>; parts?: { path: string; html: string }[]; page?: PageGeometry },
): Uint8Array {
  const files = unzipSync(original);
  const content = files["content.xml"];
  if (!content) throw new Error("not an .odt: content.xml missing");
  const doc = new DOMParser().parseFromString(strFromU8(content), "application/xml");
  const body = doc.getElementsByTagName("office:text")[0];
  if (!body) throw new Error("not an .odt: office:text missing");

  while (body.firstChild) body.removeChild(body.firstChild);
  const htmlDoc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
  // Gather comment metadata + which ids wrap a range, before serializing.
  const refMeta = new Map<string, RefMeta>();
  for (const ref of Array.from(htmlDoc.querySelectorAll(".docx-comment-ref"))) {
    const id = ref.getAttribute("data-comment-id") ?? "";
    if (!id) continue;
    refMeta.set(id, {
      author: ref.getAttribute("data-comment-author") ?? "",
      date: ref.getAttribute("data-comment-date") ?? "",
      text: ref.getAttribute("data-comment-text") ?? "",
      resolved: ref.getAttribute("data-comment-resolved") === "1",
      paraId: ref.getAttribute("data-comment-paraid") ?? id,
    });
  }
  const rangedIds = new Set(
    Array.from(htmlDoc.querySelectorAll(".docx-comment[data-comment-id]")).map((s) => s.getAttribute("data-comment-id") ?? ""),
  );
  const ctx: OdfCtx = { doc, auto: ensureAutoStyles(doc), created: new Map(), files, pics: [], refMeta, rangedIds, done: opts?.done ?? new Map(), changes: [] };
  for (const node of Array.from(htmlDoc.body.childNodes)) {
    const block = htmlBlockToOdf(node, ctx);
    if (block) body.appendChild(block);
  }
  if (!body.firstChild) body.appendChild(doc.createElementNS(NS.text, "text:p"));
  const tc = buildTrackedChanges(ctx); // tracked-changes region goes first in office:text
  if (tc) body.insertBefore(tc, body.firstChild);
  addManifestEntries(files, ctx.pics); // register any images embedded above
  if (opts?.parts) applyHeaderFooter(files, opts.parts); // header/footer -> styles.xml
  if (opts?.page) applyPageMargins(files, opts.page); // margins -> styles.xml page-layout

  const out = new XMLSerializer().serializeToString(doc);
  // Re-zip. ODF requires the "mimetype" entry first and stored (uncompressed).
  const repacked: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
  if (files["mimetype"]) repacked["mimetype"] = [files["mimetype"], { level: 0 }];
  for (const [name, data] of Object.entries(files)) {
    if (name === "mimetype") continue;
    repacked[name] = name === "content.xml" ? strToU8(out) : data;
  }
  return zipSync(repacked as Record<string, Uint8Array>);
}

