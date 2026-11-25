// DrawingML shape geometry (a:prstGeom, a:custGeom): the guide formulas and paths evaluated into
// SVG subpaths in the shape's own px box (y down), with where each open path starts and ends,
// for arrowheads.
import { PRESETS, type PresetPath } from "./preset-shapes";

type Pt = [number, number];
export type FillMode = "norm" | "none" | "lighten" | "lightenLess" | "darken" | "darkenLess";

export interface GeomPath {
  d: string;
  fill: FillMode;
  stroke: boolean;
  closed: boolean;
  /** The first segment (from, to) and the last (from, to), for line ends. */
  start?: [Pt, Pt];
  end?: [Pt, Pt];
}

interface Def {
  av: [string, string][];
  gd: [string, string][];
  paths: PresetPath[];
}

const DEG = 60000; // angles are in 1/60 000 degree
const rad = (a: number) => (a / DEG) * (Math.PI / 180);
const r = (n: number) => String(Math.round(n * 100) / 100);

/** A preset's outline for a `w` x `h` px shape, with `adjust` overriding its adjust values
 *  ({adj: "val 25000"}). `undefined` for an unknown preset. */
export function presetGeometry(prst: string, w: number, h: number, adjust: Record<string, string> = {}): GeomPath[] | undefined {
  const p = PRESETS[prst];
  if (!p) return undefined;
  return evaluate({ av: p[0], gd: p[1], paths: p[2] }, w, h, adjust);
}

/** An a:custGeom's outline. */
export function customGeometry(el: Element, w: number, h: number): GeomPath[] | undefined {
  const kids = (e: Element | undefined, tag: string) => (e ? Array.from(e.children).filter((c) => c.localName === tag) : []);
  const one = (e: Element | undefined, tag: string) => kids(e, tag)[0];
  const gds = (e: Element | undefined) => kids(e, "gd").map((g): [string, string] => [g.getAttribute("name") ?? "", g.getAttribute("fmla") ?? "val 0"]);
  const paths: PresetPath[] = [];
  for (const p of kids(one(el, "pathLst"), "path")) {
    const cmds: string[] = [];
    for (const c of Array.from(p.children)) {
      const pts = kids(c, "pt").flatMap((pt) => [pt.getAttribute("x") ?? "0", pt.getAttribute("y") ?? "0"]);
      const n = c.localName;
      if (n === "moveTo") cmds.push("M", ...pts);
      else if (n === "lnTo") cmds.push("L", ...pts);
      else if (n === "cubicBezTo") cmds.push("C", ...pts);
      else if (n === "quadBezTo") cmds.push("Q", ...pts);
      else if (n === "arcTo") cmds.push("A", ...["wR", "hR", "stAng", "swAng"].map((a) => c.getAttribute(a) ?? "0"));
      else if (n === "close") cmds.push("Z");
    }
    paths.push([Number(p.getAttribute("w") ?? 0), Number(p.getAttribute("h") ?? 0), p.getAttribute("fill") ?? "norm", p.getAttribute("stroke") === "false" || p.getAttribute("stroke") === "0" ? 0 : 1, cmds.join(" ")]);
  }
  return paths.length ? evaluate({ av: gds(one(el, "avLst")), gd: gds(one(el, "gdLst")), paths }, w, h, {}) : undefined;
}

function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h), ls = Math.max(w, h);
  const v = new Map<string, number>([
    ["w", w], ["h", h], ["l", 0], ["t", 0], ["r", w], ["b", h], ["hc", w / 2], ["vc", h / 2],
    ["ss", ss], ["ls", ls],
    ["cd2", 10800000], ["cd4", 5400000], ["cd8", 2700000], ["3cd4", 16200000], ["3cd8", 8100000], ["5cd8", 13500000], ["7cd8", 18900000],
  ]);
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    v.set(`wd${n}`, w / n);
    v.set(`hd${n}`, h / n);
    v.set(`ssd${n}`, ss / n);
  }
  return v;
}

function evaluate(def: Def, w: number, h: number, adjust: Record<string, string>): GeomPath[] | undefined {
  const vars = builtins(w, h);
  const arg = (t: string | undefined): number => {
    if (t === undefined) return 0;
    const n = Number(t);
    return Number.isFinite(n) ? n : vars.get(t) ?? 0;
  };
  const formula = (f: string): number => {
    const [op, ...a] = f.trim().split(/\s+/);
    const [x, y, z] = [arg(a[0]), arg(a[1]), arg(a[2])];
    switch (op) {
      case "val": return x;
      case "*/": return z ? (x * y) / z : 0;
      case "+-": return x + y - z;
      case "+/": return z ? (x + y) / z : 0;
      case "?:": return x > 0 ? y : z;
      case "abs": return Math.abs(x);
      case "at2": return (Math.atan2(y, x) * 180 / Math.PI) * DEG;
      case "cat2": return x * Math.cos(Math.atan2(z, y));
      case "sat2": return x * Math.sin(Math.atan2(z, y));
      case "cos": return x * Math.cos(rad(y));
      case "sin": return x * Math.sin(rad(y));
      case "tan": return x * Math.tan(rad(y));
      case "max": return Math.max(x, y);
      case "min": return Math.min(x, y);
      case "mod": return Math.sqrt(x * x + y * y + z * z);
      case "pin": return y < x ? x : y > z ? z : y;
      case "sqrt": return Math.sqrt(Math.max(0, x));
      default: return 0;
    }
  };
  for (const [name, f] of def.av) vars.set(name, formula(adjust[name] ?? f));
  for (const [name, f] of Object.entries(adjust)) if (!vars.has(name)) vars.set(name, formula(f));
  for (const [name, f] of def.gd) {
    const v = formula(f);
    vars.set(name, Number.isFinite(v) ? v : 0);
  }
  const out: GeomPath[] = [];
  for (const [pw, ph, fill, stroke, src] of def.paths) {
    const sx = pw ? w / pw : 1, sy = ph ? h / ph : 1;
    const toks = src.split(" ").filter(Boolean);
    let d = "";
    let cur: Pt = [0, 0], first: Pt = [0, 0];
    let start: [Pt, Pt] | undefined, end: [Pt, Pt] | undefined;
    let closed = false;
    const seg = (from: Pt, to: Pt) => {
      if (!start) start = [from, to];
      end = [from, to];
    };
    const pt = (i: number): Pt => [arg(toks[i]) * sx, arg(toks[i + 1]) * sy];
    let i = 0;
    while (i < toks.length) {
      const c = toks[i++];
      if (c === "M") { cur = first = pt(i); i += 2; d += `M${r(cur[0])},${r(cur[1])} `; }
      else if (c === "L") { const p = pt(i); i += 2; seg(cur, p); cur = p; d += `L${r(p[0])},${r(p[1])} `; }
      else if (c === "C") {
        const a = pt(i), b = pt(i + 2), p = pt(i + 4); i += 6;
        if (!start) start = [cur, a];
        end = [b, p];
        cur = p;
        d += `C${r(a[0])},${r(a[1])} ${r(b[0])},${r(b[1])} ${r(p[0])},${r(p[1])} `;
      } else if (c === "Q") {
        const a = pt(i), p = pt(i + 2); i += 4;
        if (!start) start = [cur, a];
        end = [a, p];
        cur = p;
        d += `Q${r(a[0])},${r(a[1])} ${r(p[0])},${r(p[1])} `;
      } else if (c === "A") {
        const wR = arg(toks[i]) * sx, hR = arg(toks[i + 1]) * sy, st = arg(toks[i + 2]), sw = arg(toks[i + 3]);
        i += 4;
        const arc = arcTo(cur, wR, hR, st, sw);
        if (!start) start = [cur, arc.firstCtl];
        end = [arc.lastCtl, arc.to];
        d += arc.d;
        cur = arc.to;
      } else if (c === "Z") { d += "Z "; closed = true; cur = first; }
    }
    if (d) out.push({ d: d.trim(), fill: fill as FillMode, stroke: !!stroke, closed, start, end });
  }
  return out.length ? out : undefined;
}

/** DrawingML's arcTo: from the current point along the wR x hR ellipse, from angle stAng through
 *  swAng. Angles are visual (the direction from the centre), y down. */
function arcTo(cur: Pt, wR: number, hR: number, st: number, sw: number): { d: string; to: Pt; firstCtl: Pt; lastCtl: Pt } {
  // The parametric angle of a point at visual angle `a` on the ellipse.
  const param = (a: number) => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const a0 = rad(st), a1 = rad(st + sw);
  const t0 = param(a0);
  let t1 = param(a1);
  let sweep = t1 - t0;
  if (Math.abs(sw) >= 21600000) sweep = Math.sign(sw) * 2 * Math.PI;
  else if (sw > 0 && sweep < 0) sweep += 2 * Math.PI;
  else if (sw < 0 && sweep > 0) sweep -= 2 * Math.PI;
  t1 = t0 + sweep;
  const cx = cur[0] - wR * Math.cos(t0), cy = cur[1] - hR * Math.sin(t0);
  const at = (t: number): Pt => [cx + wR * Math.cos(t), cy + hR * Math.sin(t)];
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let d = "";
  let firstCtl: Pt = cur, lastCtl: Pt = cur;
  for (let s = 0; s < n; s++) {
    const ta = t0 + s * step, tb = ta + step;
    const p0 = at(ta), p3 = at(tb);
    const c1: Pt = [p0[0] - k * wR * Math.sin(ta), p0[1] + k * hR * Math.cos(ta)];
    const c2: Pt = [p3[0] + k * wR * Math.sin(tb), p3[1] - k * hR * Math.cos(tb)];
    if (s === 0) firstCtl = c1;
    lastCtl = c2;
    d += `C${r(c1[0])},${r(c1[1])} ${r(c2[0])},${r(c2[1])} ${r(p3[0])},${r(p3[1])} `;
  }
  if (!wR || !hR || !sweep) return { d: "", to: cur, firstCtl: cur, lastCtl: cur };
  return { d, to: at(t1), firstCtl, lastCtl };
}
