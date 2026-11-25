// The shape bar on odt drawings: their look and size are read from the shape element, edits are
// kept as data-shape-edit and the drawing is redrawn from its original element with them. New
// drawings come from the same gallery as docx, written as LibreOffice writes Word's presets.
import type { ShapeHooks, ShapeLook, ShapeWrap } from "../../core/types";
import { docxShapeHooks, galleryPreset } from "../docx/shape-hooks";
import { presetGeometry } from "../docx/drawingml-geom";
import { ooxmlGeometryXml } from "./ooxml-shape";
import { odtShapeContext } from "./read";
import { isShapeTag, shapeHtml, type Look, type ShapeCtx } from "./shapes";
import { ARROW_MARKER, applyGeometry, editOf, setEditOf, type OdtShapeEdit } from "./shape-edit";

const EDIT_STYLE = "__rdoc_edit";
const NS = 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
const WRAP_VALUE = { square: "parallel", tight: "dynamic", topbottom: "none", behind: "run-through", front: "run-through" } as const;

/** A look with an edit's fill, outline, line ends and wrap. */
function edited(base: Look, e: OdtShapeEdit, paint: boolean): Look {
  const look = { ...base };
  if (paint) {
    if ("fill" in e) { look.fill = !!e.fill; if (e.fill) look.fillColor = e.fill; }
    if ("line" in e || "lineW" in e) {
      if (!e.line || !e.lineW) look.stroke = "none";
      else { if (look.stroke === "none") look.stroke = "solid"; look.strokeColor = e.line; look.strokeWidth = e.lineW; }
    }
    if (e.ends) {
      look.markerEnd = ARROW_MARKER.name;
      look.markerEndWidth = 11.3; // 0.3 cm
      if (e.ends === "both") { look.markerStart = ARROW_MARKER.name; look.markerStartWidth = 11.3; }
    }
  }
  if (e.wrap && e.wrap !== "inline") {
    const through = e.wrap === "behind" || e.wrap === "front";
    look.wrap = WRAP_VALUE[e.wrap];
    look.runThrough = through ? (e.wrap === "behind" ? "background" : "foreground") : undefined;
    // As the writer does: running through, it keeps offsets it already had; else it is offset
    // from its paragraph. Wrapped, it is aligned across its paragraph, at its top.
    if (!through || !/^from-/.test(look.hpos ?? "")) { look.hpos = through ? "from-left" : e.align ?? "left"; look.hrel = "paragraph"; }
    if (!through || !/^from-/.test(look.vpos ?? "")) { look.vpos = through ? "from-top" : "top"; look.vrel = "paragraph"; }
  }
  return look;
}

const cm = (px: number) => `${Math.round((px * 2.54 / 96) * 1000) / 1000}cm`;

/** A gallery shape as odt XML, in line with the text: a custom shape drawn from the preset's
 *  formulas, or a plain path for the lines (which take arrowheads). */
function galleryXml(id: string, w: number, h: number): { xml: string; line: boolean } | undefined {
  const e = galleryPreset(id);
  const paths = presetGeometry(e.prst, w, h);
  if (!paths) return undefined;
  const line = paths.every((p) => !p.closed && p.fill === "none") || !!e.ends || /^(line|straightConnector1|bentConnector3|curvedConnector3)$/.test(e.prst);
  const size = `svg:width="${cm(w)}" svg:height="${cm(h)}"`;
  if (line) {
    const d = paths.map((p) => p.d).join(" ");
    return { xml: `<draw:path ${NS} text:anchor-type="as-char" svg:y="0cm" ${size} svg:viewBox="0 0 ${Math.round(w)} ${Math.round(h)}" svg:d="${d}"/>`, line };
  }
  // A Word preset as LibreOffice keeps one: its formulas, adjust values and handles.
  const geom = ooxmlGeometryXml(e.prst);
  if (!geom) return undefined;
  return { xml: `<draw:custom-shape ${NS} xmlns:drawooo="http://openoffice.org/2010/draw" text:anchor-type="as-char" svg:y="0cm" ${size}>${geom}</draw:custom-shape>`, line };
}

export function odtShapeHooks(getFiles: () => Record<string, Uint8Array>): ShapeHooks {
  let ctx: ShapeCtx | undefined;
  const sctx = () => {
    if (!ctx) {
      ctx = odtShapeContext(getFiles());
      if (!ctx.markers.has(ARROW_MARKER.name)) ctx.markers.set(ARROW_MARKER.name, { vb: ARROW_MARKER.vb.split(" ").map(Number) as [number, number, number, number], d: ARROW_MARKER.d });
    }
    return ctx;
  };
  const lastPos = new WeakMap<HTMLElement, { left: number; top: number }>();
  const boxOf = (s: HTMLElement) => s.querySelector<HTMLElement>(":scope > .odt-shape-box");
  const pos = (b: HTMLElement) => ({ left: parseFloat(b.style.left) || 0, top: parseFloat(b.style.top) || 0 });
  const rootTag = (s: HTMLElement) => /^<([\w:-]+)/.exec(s.getAttribute("data-odt-xml") ?? "")?.[1] ?? "";

  const redraw = (s: HTMLElement): void => {
    const xml = s.getAttribute("data-odt-xml");
    if (!xml) return;
    const el = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    if (!el || el.getElementsByTagName("parsererror").length) return;
    const e = editOf(s);
    applyGeometry(el, e);
    const c = sctx();
    const base = (n: string) => c.looks.get(n) ?? c.looks.get("")!;
    const temp: string[] = [];
    const restyle = (n: Element, paint: boolean) => {
      const key = `${EDIT_STYLE}${temp.length}`;
      c.looks.set(key, edited(base(n.getAttribute("draw:style-name") ?? ""), e, paint));
      n.setAttribute("draw:style-name", key);
      temp.push(key);
    };
    if (el.tagName === "draw:g") {
      // The group wraps as a whole; its members take the look.
      if (e.wrap) restyle(el, false);
      if ("fill" in e || "line" in e) for (const m of Array.from(el.getElementsByTagName("*"))) if (m.tagName !== "draw:g" && isShapeTag(m.tagName)) restyle(m, true);
    } else restyle(el, true);
    const html = shapeHtml(el, c, s.tagName === "DIV" ? "block" : "inline");
    for (const k of temp) c.looks.delete(k);
    if (!html) return;
    const tmp = document.createElement("div");
    tmp.innerHTML = html;
    const fresh = tmp.firstElementChild as HTMLElement | null;
    if (!fresh) return;
    s.replaceChildren(...Array.from(fresh.childNodes));
    // Its wrap, alignment and model attributes follow the redrawn drawing.
    for (const a of Array.from(s.attributes)) if (/^data-shape-(?!edit|dirty)/.test(a.name) && !fresh.hasAttribute(a.name)) s.removeAttribute(a.name);
    for (const a of Array.from(fresh.attributes)) if (/^data-shape-/.test(a.name)) s.setAttribute(a.name, a.value);
    const st = fresh.getAttribute("style");
    if (st) s.setAttribute("style", st); else s.removeAttribute("style");
  };

  const attr = (s: HTMLElement, n: string) => s.getAttribute(`data-shape-${n}`) ?? "";
  const color = (v: string): string | null => (!v || v === "none" ? null : v);
  const update = (s: HTMLElement, f: (e: OdtShapeEdit) => void) => {
    const e: OdtShapeEdit = { ...editOf(s) };
    f(e);
    setEditOf(s, e);
    redraw(s);
  };

  return {
    redraw,
    gallery: () => docxShapeHooks.gallery!(),
    create(id) {
      const lineish = !!galleryPreset(id).ends || /^(line|arrowLine|doubleArrowLine|elbowArrow|curvedArrow)$/.test(id);
      const w = 144, h = lineish ? 72 : 96;
      const g = galleryXml(id, w, h);
      if (!g) return undefined;
      const s = document.createElement("span");
      s.className = "odt-shape";
      s.contentEditable = "false";
      s.setAttribute("data-odt-xml", g.xml);
      s.setAttribute("data-shape-wrap", "inline");
      const ends = galleryPreset(id).ends;
      setEditOf(s, g.line ? { fill: null, line: "#4472c4", lineW: 2, ...(ends ? { ends } : {}) } : { fill: "#4472c4", line: "#2f528f", lineW: 1 });
      redraw(s);
      return s;
    },
    model: {
      selector: ".odt-shape",
      get(s): (ShapeLook & { group?: boolean; groupLook?: boolean; fixedSize?: boolean }) | undefined {
        const e = editOf(s);
        const group = s.hasAttribute("data-shape-group");
        const b = boxOf(s) ?? s;
        return {
          fill: "fill" in e ? e.fill ?? null : color(attr(s, "fill")),
          line: "line" in e ? e.line ?? null : color(attr(s, "line")),
          lineW: "lineW" in e ? e.lineW ?? 0 : Number(attr(s, "linew")) || 0,
          w: e.w ?? (Number(attr(s, "w")) || b.offsetWidth),
          h: e.h ?? (Number(attr(s, "h")) || b.offsetHeight),
          ...(group ? { group: true, groupLook: true } : {}),
        };
      },
      set(s, patch) {
        update(s, (e) => {
          if ("fill" in patch) e.fill = patch.fill ?? null;
          if ("line" in patch || "lineW" in patch) { e.line = patch.line ?? null; e.lineW = patch.lineW ?? 0; }
          if (patch.w) e.w = patch.w; // a group stretches its members to it
          if (patch.h) e.h = patch.h;
        });
      },
      posEl(s) {
        const b = /^(behind|front)$/.test(s.getAttribute("data-shape-wrap") ?? "") ? boxOf(s) : null;
        if (b) lastPos.set(s, pos(b)); // the move is measured from here when the drag ends
        return b;
      },
      moved(s) {
        const b = boxOf(s), from = lastPos.get(s);
        if (!b || !from) return;
        const now = pos(b);
        const e = editOf(s);
        e.dx = (e.dx ?? 0) + now.left - from.left;
        e.dy = (e.dy ?? 0) + now.top - from.top;
        setEditOf(s, e);
        for (const [a, d] of [["bx", now.left - from.left], ["by", now.top - from.top]] as const) {
          const v = parseFloat(s.getAttribute(`data-shape-${a}`) ?? "");
          if (Number.isFinite(v)) s.setAttribute(`data-shape-${a}`, String(Math.round((v + d) * 100) / 100));
        }
        lastPos.set(s, now);
      },
      setWrap(s, wrap, align) {
        update(s, (e) => {
          e.wrap = wrap;
          e.align = align;
          if (wrap === "inline") { delete e.dx; delete e.dy; }
        });
      },
      wraps(s): ShapeWrap[] {
        // On the page (between paragraphs) a drawing can only float; a connector stays with the
        // shapes it joins, so it does not go in line.
        if (s.tagName === "DIV") return ["behind", "front"];
        if (/^draw:(connector|measure)$/.test(rootTag(s))) return ["square", "tight", "topbottom", "behind", "front"];
        return ["inline", "square", "tight", "topbottom", "behind", "front"];
      },
    },
  };
}
