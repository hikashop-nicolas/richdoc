// A Word preset shape as LibreOffice writes one in odt (draw:type="ooxml-<preset>"): its guides
// become draw:equation formulas over the shape's size (logwidth / logheight, 1/100 mm), its
// adjust values draw:modifiers ($0, $1 ...), its paths an enhanced path, and its handles
// draw:handle. The shape then keeps its proportions when resized, and its handles in LibreOffice.
import { PRESETS } from "../docx/preset-shapes";

const BUILTIN: Record<string, string> = {
  w: "logwidth", h: "logheight", l: "0", t: "0", r: "logwidth", b: "logheight",
  hc: "logwidth/2", vc: "logheight/2", ss: "min(logwidth,logheight)", ls: "max(logwidth,logheight)",
  cd2: "10800000", cd4: "5400000", cd8: "2700000", "3cd4": "16200000", "3cd8": "8100000", "5cd8": "13500000", "7cd8": "18900000",
};
for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
  BUILTIN[`wd${n}`] = `logwidth/${n}`;
  BUILTIN[`hd${n}`] = `logheight/${n}`;
  BUILTIN[`ssd${n}`] = `min(logwidth,logheight)/${n}`;
}
const SHADE: Record<string, string> = { darken: "H", darkenLess: "I", lighten: "J", lightenLess: "K" };
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/** The draw:enhanced-geometry element (as XML) for a Word preset, or undefined if unknown. */
export function ooxmlGeometryXml(prst: string): string | undefined {
  const p = PRESETS[prst];
  if (!p) return undefined;
  const [av, gd, paths, handles] = p;
  const eqs: string[] = [];
  const names = new Map<string, string>(); // guide / adjust name -> ?fN / $N
  av.forEach(([n], i) => names.set(n, `$${i}`));
  const cache = new Map<string, string>();
  /** A formula as an equation: its ?fN. */
  const eq = (formula: string): string => {
    const got = cache.get(formula);
    if (got) return got;
    const ref = `?f${eqs.length} `;
    eqs.push(formula);
    cache.set(formula, ref);
    return ref;
  };
  /** An argument as a term of a formula. */
  const term = (a: string | undefined): string => {
    if (a === undefined) return "0";
    if (/^-?\d+(\.\d+)?$/.test(a)) return a;
    return names.get(a) ?? (BUILTIN[a] ? `(${BUILTIN[a]})` : "0");
  };
  /** An argument as a path or handle parameter: a number, $N or ?fN (anything else an equation). */
  const param = (a: string): string => {
    if (/^-?\d+(\.\d+)?$/.test(a)) return a;
    const n = names.get(a);
    if (n) return n;
    return BUILTIN[a] === "0" ? "0" : eq(BUILTIN[a] ?? "0");
  };
  const formula = (f: string): string => {
    const [op, ...args] = f.trim().split(/\s+/);
    const [x, y, z] = [term(args[0]), term(args[1]), term(args[2])];
    const deg = (v: string) => `pi*${v}/10800000`;
    switch (op) {
      case "val": return x;
      case "*/": return `${x}*${y}/${z}`;
      case "+-": return `${x}+${y}-${z}`;
      case "+/": return `(${x}+${y})/${z}`;
      case "?:": return `if(${x},${y},${z})`;
      case "abs": return `abs(${x})`;
      case "at2": return `10800000*atan2(${y},${x})/pi`;
      case "cat2": return `${x}*cos(atan2(${z},${y}))`;
      case "sat2": return `${x}*sin(atan2(${z},${y}))`;
      case "cos": return `${x}*cos(${deg(y)})`;
      case "sin": return `${x}*sin(${deg(y)})`;
      case "tan": return `${x}*tan(${deg(y)})`;
      case "max": return `max(${x},${y})`;
      case "min": return `min(${x},${y})`;
      case "mod": return `sqrt(${x}*${x}+${y}*${y}+${z}*${z})`;
      case "pin": return `if(${x}-${y},${x},if(${y}-${z},${z},${y}))`;
      case "sqrt": return `sqrt(${x})`;
      default: return "0";
    }
  };
  for (const [n, f] of gd) names.set(n, eq(formula(f)));

  // Paths: in the shape's own size, or scaled from their own w x h.
  let full = "", plain = "";
  for (const [pw, ph, fill, stroke, src] of paths) {
    const toks = src.split(" ").filter(Boolean);
    const x = (t: string) => (pw ? eq(`${term(t)}*logwidth/${pw}`) : param(t));
    const y = (t: string) => (ph ? eq(`${term(t)}*logheight/${ph}`) : param(t));
    let out = SHADE[fill] ? `${SHADE[fill]} ` : "", fallback = out;
    for (let i = 0; i < toks.length;) {
      const c = toks[i++]!;
      const n = { M: 2, L: 2, C: 6, Q: 4, A: 4, Z: 0 }[c] ?? 0;
      const a = toks.slice(i, i + n);
      i += n;
      if (c === "A") {
        // OOXML's arcTo: radii, then start and swing in degrees (LibreOffice's G).
        out += `G ${x(a[0]!)} ${y(a[1]!)} ${eq(`${term(a[2])}/60000`)} ${eq(`${term(a[3])}/60000`)} `;
        continue; // the standard path has no such arc
      }
      const pts = a.map((t, k) => (k % 2 ? y(t) : x(t))).join(" ");
      const seg = `${c}${pts ? ` ${pts}` : ""} `;
      out += seg;
      fallback += seg;
    }
    const tail = `${fill === "none" ? "F " : ""}${stroke ? "" : "S "}N `;
    full += out + tail;
    plain += fallback + tail;
  }
  const mods = av.map(([, f]) => term(f.split(/\s+/)[1])).join(" ");
  // Handles, as LibreOffice writes them: their place and the range of the value they move.
  const hs = handles.map(([kind, ref, min, max, ref2, min2, max2, hx, hy]) => {
    const px = param(hx), py = param(hy);
    let attrs = `draw:handle-position="${px.trim()} ${py.trim()}"`;
    const range = (axis: string, lo: string, hi: string) => ` draw:handle-${axis}-minimum="${param(lo || "0").trim()}" draw:handle-${axis}-maximum="${param(hi || "0").trim()}"`;
    if (kind === "xy") {
      if (ref) attrs += range("range-x", min, max);
      if (ref2) attrs += range("range-y", min2, max2);
    } else if (ref) attrs += range("radius-range", min, max);
    return `<draw:handle ${attrs}/>`;
  });
  const equations = eqs.map((f, i) => `<draw:equation draw:name="f${i}" draw:formula="${esc(f)}"/>`).join("");
  return `<draw:enhanced-geometry svg:viewBox="0 0 0 0" draw:type="ooxml-${prst}"${mods ? ` draw:modifiers="${mods}"` : ""} draw:enhanced-path="${plain.trim()}" drawooo:enhanced-path="${full.trim()}">${equations}${hs.join("")}</draw:enhanced-geometry>`;
}
