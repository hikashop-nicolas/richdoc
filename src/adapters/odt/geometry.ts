// The geometry of an odt custom shape (draw:enhanced-geometry): its equations, modifiers and
// enhanced path, evaluated into plain SVG subpaths in the geometry's view box. This is what
// draws LibreOffice's ready-made shapes (block arrows, stars, callouts, flowchart symbols ...),
// whose paths are computed from the shape's size and handle positions.
import { arcTo } from "../docx/ooxml-arc";

export interface SubPath {
  d: string; // absolute M/L/C/Z commands in view-box units
  fill: boolean;
  stroke: boolean;
  /** Fill shading: darker below 0 (H -0.4, I -0.2), lighter above (J 0.4, K 0.2). */
  shade: number;
}

export interface Geometry {
  vb: [number, number, number, number];
  paths: SubPath[];
}

type Val = () => number;

/** Evaluate an enhanced geometry for a shape drawn `wPx` x `hPx`. `undefined` when it has no path. */
export function evaluateGeometry(g: Element, wPx: number, hPx: number): Geometry | undefined {
  // LibreOffice keeps the whole path (with its OOXML arcs, G) in its own attribute, and one
  // without them in the standard one.
  const path = g.getAttribute("drawooo:enhanced-path") || g.getAttribute("draw:enhanced-path") || "";
  if (!path.trim()) return undefined;
  const vbRaw = (g.getAttribute("svg:viewBox") ?? "0 0 21600 21600").trim().split(/[\s,]+/).map(Number);
  let vb: Geometry["vb"] = vbRaw.length === 4 && vbRaw.every(Number.isFinite) ? (vbRaw as Geometry["vb"]) : [0, 0, 21600, 21600];
  // LibreOffice writes Word's preset shapes (draw:type="ooxml-...") with an empty view box: their
  // paths are in the shape's own 1/100 mm (logwidth x logheight).
  if (!(vb[2] > 0) || !(vb[3] > 0)) vb = [0, 0, (wPx * 2540) / 96, (hPx * 2540) / 96];
  const modifiers = (g.getAttribute("draw:modifiers") ?? "").trim().split(/[\s,]+/).filter(Boolean).map(Number);
  const formulas = new Map<string, string>();
  for (const eq of Array.from(g.getElementsByTagName("draw:equation"))) {
    const name = eq.getAttribute("draw:name");
    if (name) formulas.set(name, eq.getAttribute("draw:formula") ?? "0");
  }
  // Constants: the view box sides and the shape's size in 1/100 mm (logwidth / logheight).
  const consts: Record<string, number> = {
    pi: Math.PI,
    left: vb[0],
    top: vb[1],
    right: vb[0] + vb[2],
    bottom: vb[1] + vb[3],
    width: vb[2],
    height: vb[3],
    xstretch: 1,
    ystretch: 1,
    hasstroke: 1,
    hasfill: 1,
    logwidth: (wPx * 2540) / 96,
    logheight: (hPx * 2540) / 96,
  };
  const memo = new Map<string, number>();
  const busy = new Set<string>();
  const equation = (name: string): number => {
    const got = memo.get(name);
    if (got !== undefined) return got;
    if (busy.has(name)) return 0; // a cycle: draw something rather than loop
    busy.add(name);
    const v = evalFormula(formulas.get(name) ?? "0");
    busy.delete(name);
    memo.set(name, v);
    return v;
  };
  const evalFormula = (src: string): number => {
    try {
      const v = parse(src, consts, equation, modifiers)();
      return Number.isFinite(v) ? v : 0;
    } catch {
      return 0;
    }
  };
  /** A path parameter: a number, ?fN, $N, or a constant name. */
  const param = (tok: string): number => {
    if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(tok)) return Number(tok);
    if (tok.startsWith("?")) return equation(tok.slice(1));
    if (tok.startsWith("$")) return modifiers[Number(tok.slice(1))] ?? 0;
    return consts[tok.toLowerCase()] ?? 0;
  };
  const paths = walkPath(path, param);
  // LibreOffice gives some subpaths their own coordinate space (drawooo:sub-view-size: a width
  // and height each, 0 0 for the shape's own), scaled into the view box.
  const sub = (g.getAttribute("drawooo:sub-view-size") ?? "").trim().split(/[\s,]+/).filter(Boolean).map(Number);
  paths.forEach((p, k) => {
    const sw = sub[2 * k] ?? 0, sh = sub[2 * k + 1] ?? 0;
    if (!(sw > 0) || !(sh > 0)) return;
    const kx = vb[2] / sw, ky = vb[3] / sh;
    p.d = p.d.replace(/(-?[\d.]+(?:e[-+]?\d+)?),(-?[\d.]+(?:e[-+]?\d+)?)/gi, (_m, xs: string, ys: string) => `${r(vb[0] + Number(xs) * kx)},${r(vb[1] + Number(ys) * ky)}`);
  });
  // LibreOffice shades some ready-made shapes by their type rather than in the path.
  const preset = PRESET_SHADES[g.getAttribute("draw:type") ?? ""];
  if (preset) paths.forEach((p, k) => { if (!p.shade && preset[k]) p.shade = preset[k]!; });
  if (g.getAttribute("draw:mirror-horizontal") === "true" || g.getAttribute("draw:mirror-vertical") === "true") {
    const mh = g.getAttribute("draw:mirror-horizontal") === "true", mv = g.getAttribute("draw:mirror-vertical") === "true";
    for (const p of paths) {
      p.d = p.d.replace(/(-?[\d.]+(?:e[-+]?\d+)?),(-?[\d.]+(?:e[-+]?\d+)?)/gi, (_m, xs: string, ys: string) => {
        const x = mh ? vb[0] * 2 + vb[2] - Number(xs) : Number(xs);
        const y = mv ? vb[1] * 2 + vb[3] - Number(ys) : Number(ys);
        return `${r(x)},${r(y)}`;
      });
    }
  }
  return paths.length ? { vb, paths } : undefined;
}

const PRESET_SHADES: Record<string, number[]> = { cube: [0, 0.2, -0.4], smiley: [0, -0.2, -0.2] };

const r = (n: number): string => String(Math.round(n * 1000) / 1000);

/** Recursive-descent parser for the odf formula language, returning a lazy evaluator. */
function parse(src: string, consts: Record<string, number>, equation: (n: string) => number, mods: number[]): Val {
  const toks = src.match(/\?[a-zA-Z_]\w*|\$\d+|[a-zA-Z_]\w*|\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+|[-+*/(),]/g) ?? [];
  let i = 0;
  const peek = () => toks[i];
  const expr = (): Val => {
    let left = term();
    while (peek() === "+" || peek() === "-") {
      const op = toks[i++];
      const a = left, b = term();
      left = op === "+" ? () => a() + b() : () => a() - b();
    }
    return left;
  };
  const term = (): Val => {
    let left = factor();
    while (peek() === "*" || peek() === "/") {
      const op = toks[i++];
      const a = left, b = factor();
      left = op === "*" ? () => a() * b() : () => { const d = b(); return d ? a() / d : 0; };
    }
    return left;
  };
  const args = (): Val[] => {
    const out: Val[] = [];
    if (peek() !== "(") return out;
    i++;
    if (peek() !== ")") {
      out.push(expr());
      while (peek() === ",") {
        i++;
        out.push(expr());
      }
    }
    if (peek() === ")") i++;
    return out;
  };
  const FN: Record<string, (a: number[]) => number> = {
    abs: (a) => Math.abs(a[0]!),
    sqrt: (a) => Math.sqrt(Math.max(0, a[0]!)),
    sin: (a) => Math.sin(a[0]!),
    cos: (a) => Math.cos(a[0]!),
    tan: (a) => Math.tan(a[0]!),
    atan: (a) => Math.atan(a[0]!),
    atan2: (a) => Math.atan2(a[0]!, a[1]!),
    min: (a) => Math.min(a[0]!, a[1]!),
    max: (a) => Math.max(a[0]!, a[1]!),
    if: (a) => (a[0]! > 0 ? a[1]! : a[2]!),
  };
  const factor = (): Val => {
    const t = toks[i++];
    if (t === undefined) return () => 0;
    if (t === "-") { const f = factor(); return () => -f(); }
    if (t === "+") return factor();
    if (t === "(") {
      const e = expr();
      if (peek() === ")") i++;
      return e;
    }
    if (/^[\d.]/.test(t)) { const n = Number(t); return () => n; }
    if (t.startsWith("?")) { const name = t.slice(1); return () => equation(name); }
    if (t.startsWith("$")) { const n = Number(t.slice(1)); return () => mods[n] ?? 0; }
    const name = t.toLowerCase();
    if (FN[name]) {
      const a = args();
      if (name === "if") return () => (a[0]!() > 0 ? a[1]!() : a[2]!()); // only the taken branch
      return () => FN[name]!(a.map((f) => f()));
    }
    const c = consts[name];
    return () => c ?? 0;
  };
  return expr();
}

type Pt = [number, number];

const KAPPA = 0.5522847498;

/** Cubic Bézier segments for an elliptical arc (center, radii, from angle a0 through sweep, radians,
 *  in odf's sense: y up, so a point is (cx + rx cos a, cy - ry sin a)). */
function arcBeziers(cx: number, cy: number, rx: number, ry: number, a0: number, sweep: number): string {
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let out = "";
  let a = a0;
  for (let s = 0; s < n; s++) {
    const b = a + step;
    const p0: Pt = [cx + rx * Math.cos(a), cy - ry * Math.sin(a)];
    const p3: Pt = [cx + rx * Math.cos(b), cy - ry * Math.sin(b)];
    const c1: Pt = [p0[0] - k * rx * Math.sin(a), p0[1] - k * ry * Math.cos(a)];
    const c2: Pt = [p3[0] + k * rx * Math.sin(b), p3[1] + k * ry * Math.cos(b)];
    out += `C${r(c1[0])},${r(c1[1])} ${r(c2[0])},${r(c2[1])} ${r(p3[0])},${r(p3[1])} `;
    a = b;
  }
  return out;
}

const angleOf = (cx: number, cy: number, rx: number, ry: number, x: number, y: number): number =>
  Math.atan2(-(y - cy) / (ry || 1), (x - cx) / (rx || 1));

/** Walk an enhanced path into subpaths of plain SVG commands. */
function walkPath(src: string, param: (t: string) => number): SubPath[] {
  const toks = src.trim().split(/[\s,]+/).filter(Boolean);
  const paths: SubPath[] = [];
  let d = "";
  let fill = true, stroke = true, shade = 0;
  let cur: Pt = [0, 0];
  let start: Pt = [0, 0];
  let i = 0;
  let cmd = "";
  const isCmd = (t: string) => /^[A-Z]$/.test(t);
  const more = () => i < toks.length && !isCmd(toks[i]!);
  const num = () => param(toks[i++] ?? "0");
  const flush = () => {
    if (d.trim()) paths.push({ d: d.trim(), fill, stroke, shade });
    d = "";
    fill = true;
    stroke = true;
    shade = 0;
  };
  const moveOrLine = (p: Pt, move: boolean) => {
    d += `${move || !d ? "M" : "L"}${r(p[0])},${r(p[1])} `;
    if (move || !d) start = p;
    cur = p;
  };
  while (i < toks.length) {
    if (isCmd(toks[i]!)) cmd = toks[i++]!;
    else if (!cmd) { i++; continue; }
    switch (cmd) {
      case "M":
        moveOrLine([num(), num()], true);
        start = cur;
        while (more()) moveOrLine([num(), num()], false);
        break;
      case "L":
        while (more()) moveOrLine([num(), num()], false);
        break;
      case "C":
        while (more()) {
          const a: Pt = [num(), num()], b: Pt = [num(), num()], p: Pt = [num(), num()];
          d += `C${r(a[0])},${r(a[1])} ${r(b[0])},${r(b[1])} ${r(p[0])},${r(p[1])} `;
          cur = p;
        }
        break;
      case "Q":
        while (more()) {
          const c: Pt = [num(), num()], p: Pt = [num(), num()];
          d += `Q${r(c[0])},${r(c[1])} ${r(p[0])},${r(p[1])} `;
          cur = p;
        }
        break;
      case "Z":
        d += "Z ";
        cur = start;
        break;
      case "N":
        flush();
        break;
      case "F":
        fill = false;
        break;
      case "S":
        stroke = false;
        break;
      case "T": // angle-ellipseto: center, radii, start and end angle (degrees)
      case "U": // angle-ellipse: the same, starting a new subpath
        while (more()) {
          const cx = num(), cy = num(), rx = num(), ry = num();
          let a0 = (num() * Math.PI) / 180, a1 = (num() * Math.PI) / 180;
          if (a1 <= a0) a1 += 2 * Math.PI;
          const p0: Pt = [cx + rx * Math.cos(a0), cy - ry * Math.sin(a0)];
          if (cmd === "U") { d += `M${r(p0[0])},${r(p0[1])} `; start = p0; } else moveOrLine(p0, false);
          d += arcBeziers(cx, cy, rx, ry, a0, a1 - a0);
          cur = [cx + rx * Math.cos(a1), cy - ry * Math.sin(a1)];
        }
        break;
      case "A": // arcto, counter-clockwise: bounding box, then points giving the start and end directions
      case "B": // arc, counter-clockwise, starting a new subpath
      case "W": // clockwise arcto
      case "V": // clockwise arc
        while (more()) {
          const x1 = num(), y1 = num(), x2 = num(), y2 = num(), x3 = num(), y3 = num(), x4 = num(), y4 = num();
          const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, rx = Math.abs(x2 - x1) / 2, ry = Math.abs(y2 - y1) / 2;
          const a0 = angleOf(cx, cy, rx, ry, x3, y3), a1 = angleOf(cx, cy, rx, ry, x4, y4);
          const cw = cmd === "W" || cmd === "V";
          let sweep = a1 - a0;
          if (cw) { if (sweep >= 0) sweep -= 2 * Math.PI; } else if (sweep <= 0) sweep += 2 * Math.PI;
          const p0: Pt = [cx + rx * Math.cos(a0), cy - ry * Math.sin(a0)];
          const newPath = cmd === "B" || cmd === "V";
          if (newPath) { d += `M${r(p0[0])},${r(p0[1])} `; start = p0; } else moveOrLine(p0, false);
          d += arcBeziers(cx, cy, rx, ry, a0, sweep);
          cur = [cx + rx * Math.cos(a0 + sweep), cy - ry * Math.sin(a0 + sweep)];
        }
        break;
      case "X": // elliptical quadrant, starting horizontally
      case "Y": // elliptical quadrant, starting vertically
        {
          let horiz = cmd === "X";
          while (more()) {
            const p: Pt = [num(), num()];
            const c1: Pt = horiz ? [cur[0] + KAPPA * (p[0] - cur[0]), cur[1]] : [cur[0], cur[1] + KAPPA * (p[1] - cur[1])];
            const c2: Pt = horiz ? [p[0], p[1] - KAPPA * (p[1] - cur[1])] : [p[0] - KAPPA * (p[0] - cur[0]), p[1]];
            d += `C${r(c1[0])},${r(c1[1])} ${r(c2[0])},${r(c2[1])} ${r(p[0])},${r(p[1])} `;
            cur = p;
            horiz = !horiz; // consecutive quadrants alternate
          }
        }
        break;
      case "G": // arcangleto (OOXML's arcTo): radii, start angle and swing, from the current point
        while (more()) {
          const wr = num(), hr = num();
          let st = num(), sw = num();
          if (Math.abs(st) > 720 || Math.abs(sw) > 720) { st /= 60000; sw /= 60000; } // 60000ths of a degree
          const arc = arcTo(cur, wr, hr, st * 60000, sw * 60000);
          d += arc.d;
          cur = arc.to;
        }
        break;
      case "H": shade = -0.4; break;
      case "I": shade = -0.2; break;
      case "J": shade = 0.4; break;
      case "K": shade = 0.2; break;
      default:
        // Anything unknown: skip its parameters.
        while (more()) i++;
    }
  }
  flush();
  return paths;
}
