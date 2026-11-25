// The shape bar on odt drawings: their look and size are read from the shape element, edits are
// kept as data-shape-edit and the drawing is redrawn from its original element with them.
import type { ShapeHooks, ShapeLook } from "../../core/types";
import { odtShapeContext } from "./read";
import { shapeHtml, type ShapeCtx } from "./shapes";
import { applyGeometry, editOf, hasLook, setEditOf, type OdtShapeEdit } from "./shape-edit";

const EDIT_STYLE = "__rdoc_edit";

export function odtShapeHooks(getFiles: () => Record<string, Uint8Array>): ShapeHooks {
  let ctx: ShapeCtx | undefined;
  const sctx = () => (ctx ??= odtShapeContext(getFiles()));
  const lastPos = new WeakMap<HTMLElement, { left: number; top: number }>();
  const boxOf = (s: HTMLElement) => s.querySelector<HTMLElement>(":scope > .odt-shape-box");
  const pos = (b: HTMLElement) => ({ left: parseFloat(b.style.left) || 0, top: parseFloat(b.style.top) || 0 });

  const redraw = (s: HTMLElement): void => {
    const xml = s.getAttribute("data-odt-xml");
    if (!xml) return;
    const el = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    if (!el || el.getElementsByTagName("parsererror").length) return;
    const e = editOf(s);
    applyGeometry(el, e);
    const c = sctx();
    if (hasLook(e)) {
      const base = c.looks.get(el.getAttribute("draw:style-name") ?? "") ?? c.looks.get("");
      if (!base) return;
      const look = { ...base };
      if ("fill" in e) { look.fill = !!e.fill; if (e.fill) look.fillColor = e.fill; }
      if ("line" in e || "lineW" in e) {
        if (!e.line || !e.lineW) look.stroke = "none";
        else { if (look.stroke === "none") look.stroke = "solid"; look.strokeColor = e.line; look.strokeWidth = e.lineW; }
      }
      c.looks.set(EDIT_STYLE, look);
      el.setAttribute("draw:style-name", EDIT_STYLE);
    }
    const html = shapeHtml(el, c, s.tagName === "DIV" ? "block" : "inline");
    c.looks.delete(EDIT_STYLE);
    if (!html) return;
    const tmp = document.createElement("div");
    tmp.innerHTML = html;
    const fresh = tmp.firstElementChild as HTMLElement | null;
    if (!fresh) return;
    s.replaceChildren(...Array.from(fresh.childNodes));
    const st = fresh.getAttribute("style");
    if (st) s.setAttribute("style", st);
  };

  const attr = (s: HTMLElement, n: string) => s.getAttribute(`data-shape-${n}`) ?? "";
  const color = (v: string): string | null => (!v || v === "none" ? null : v);

  return {
    redraw,
    model: {
      selector: ".odt-shape",
      get(s): (ShapeLook & { group?: boolean }) | undefined {
        if (s.hasAttribute("data-shape-group")) {
          const b = boxOf(s) ?? s;
          return { fill: null, line: null, lineW: 0, w: b.offsetWidth, h: b.offsetHeight, group: true };
        }
        const e = editOf(s);
        return {
          fill: "fill" in e ? e.fill ?? null : color(attr(s, "fill")),
          line: "line" in e ? e.line ?? null : color(attr(s, "line")),
          lineW: "lineW" in e ? e.lineW ?? 0 : Number(attr(s, "linew")) || 0,
          w: e.w ?? (Number(attr(s, "w")) || 0),
          h: e.h ?? (Number(attr(s, "h")) || 0),
        };
      },
      set(s, patch) {
        const e: OdtShapeEdit = { ...editOf(s) };
        if (s.hasAttribute("data-shape-group")) return; // a group keeps its members' looks and sizes
        if ("fill" in patch) e.fill = patch.fill ?? null;
        if ("line" in patch || "lineW" in patch) { e.line = patch.line ?? null; e.lineW = patch.lineW ?? 0; }
        if (patch.w) e.w = patch.w;
        if (patch.h) e.h = patch.h;
        setEditOf(s, e);
        redraw(s);
      },
      posEl(s) {
        const b = s.getAttribute("data-shape-wrap") === "inline" ? null : boxOf(s);
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
        lastPos.set(s, now);
      },
    },
  };
}
