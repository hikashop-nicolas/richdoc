// Recover an editable LaTeX approximation from presentation MathML. An imported equation (from a
// docx OMML or an odt formula object) has no authored LaTeX, so clicking it to edit would otherwise
// open an empty box. This walks the MathML and rebuilds LaTeX for the common constructs (the inverse
// of what temml emits for our subset); unknown elements fall back to their text content. It is
// best-effort: re-rendering the recovered LaTeX yields an equivalent equation, not the original bytes.

// Operators / symbols that have a LaTeX command, keyed by the Unicode character MathML uses.
const SYM: Record<string, string> = {
  "−": "-", "±": "\\pm", "∓": "\\mp", "×": "\\times", "÷": "\\div",
  "⋅": "\\cdot", "∗": "*", "≤": "\\le", "≥": "\\ge", "≠": "\\ne",
  "≈": "\\approx", "≡": "\\equiv", "∝": "\\propto", "→": "\\to",
  "⇒": "\\Rightarrow", "⇔": "\\Leftrightarrow", "∞": "\\infty", "∑": "\\sum",
  "∏": "\\prod", "∫": "\\int", "∂": "\\partial", "√": "\\sqrt", "∇": "\\nabla",
  "∈": "\\in", "∉": "\\notin", "⊂": "\\subset", "⊆": "\\subseteq",
  "∪": "\\cup", "∩": "\\cap", "∀": "\\forall", "∃": "\\exists", "∅": "\\emptyset",
  "⋯": "\\cdots", "…": "\\ldots", "′": "'", "·": "\\cdot",
  // Greek (lower-case)
  "α": "\\alpha", "β": "\\beta", "γ": "\\gamma", "δ": "\\delta", "ε": "\\epsilon",
  "ζ": "\\zeta", "η": "\\eta", "θ": "\\theta", "κ": "\\kappa", "λ": "\\lambda",
  "μ": "\\mu", "ν": "\\nu", "ξ": "\\xi", "π": "\\pi", "ρ": "\\rho",
  "σ": "\\sigma", "τ": "\\tau", "φ": "\\phi", "χ": "\\chi", "ψ": "\\psi", "ω": "\\omega",
  // Greek (upper-case)
  "Γ": "\\Gamma", "Δ": "\\Delta", "Θ": "\\Theta", "Λ": "\\Lambda", "Ξ": "\\Xi",
  "Π": "\\Pi", "Σ": "\\Sigma", "Φ": "\\Phi", "Ψ": "\\Psi", "Ω": "\\Omega",
};

// Standard operators that LaTeX uprights via a command (mathrm), so recovery keeps them upright.
const FUNCS = new Set([
  "sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "coth",
  "arcsin", "arccos", "arctan", "log", "ln", "lg", "exp", "lim", "limsup",
  "liminf", "max", "min", "sup", "inf", "det", "dim", "gcd", "deg", "arg", "ker", "hom", "Pr",
]);
// Invisible operators MathML inserts for layout (function application, invisible times, ...): dropped.
const INVISIBLE = /[⁡⁢⁣⁤​]/g;

// A token (mi/mn/mo): drop invisible operators, map known symbols to commands (with a trailing space
// so "\pm b" stays two tokens), upright known function names, pass plain characters through.
const tokenLatex = (raw: string): string => {
  const s = raw.replace(INVISIBLE, "").trim();
  if (!s) return "";
  if (FUNCS.has(s)) return `\\${s} `;
  if (SYM[s]) return /^\\[a-zA-Z]+$/.test(SYM[s]) ? `${SYM[s]} ` : SYM[s];
  return s;
};

const kidsOf = (el: Element): Element[] => Array.from(el.children);

// Bracket pair -> the LaTeX matrix environment temml builds it from (cases has an empty close).
const matrixEnv = (open: string, close: string): string | null => {
  if (open === "(" && close === ")") return "pmatrix";
  if (open === "[" && close === "]") return "bmatrix";
  if (open === "|" && close === "|") return "vmatrix";
  if (open === "‖" && close === "‖") return "Vmatrix";
  if (open === "{") return close === "}" ? "Bmatrix" : "cases";
  return null;
};
// The body of an mtable as LaTeX rows (cells joined by &, rows by \\).
const tableBody = (table: Element): string =>
  kidsOf(table)
    .filter((r) => r.localName === "mtr")
    .map((r) => kidsOf(r).filter((c) => c.localName === "mtd").map(nodeLatex).join(" & "))
    .join(" \\\\ ");
// An mrow of exactly [open mo, mtable, close mo] -> the matrix environment, or null for a plain mrow.
const fencedMatrixLatex = (el: Element): string | null => {
  const k = kidsOf(el);
  if (k.length !== 3 || k[0]!.localName !== "mo" || k[1]!.localName !== "mtable" || k[2]!.localName !== "mo") return null;
  const open = (k[0]!.textContent ?? "").trim();
  const close = (k[2]!.textContent ?? "").trim();
  const env = matrixEnv(open, close);
  return env ? `\\begin{${env}}${tableBody(k[1]!)}\\end{${env}}` : `${open}\\begin{matrix}${tableBody(k[1]!)}\\end{matrix}${close}`;
};

// A horizontal brace as the over/under glyph -> its LaTeX command (temml nests these:
// mover[mover[body, ⏞], label] for \overbrace{body}^{label}).
const BRACE: Record<string, string> = { "⏞": "\\overbrace", "⏟": "\\underbrace" };
// If `el` is the inner brace mover/munder, return [command, body]; else null.
const braceInner = (el: Element): [string, Element] | null => {
  const k = kidsOf(el);
  const g = k.length === 2 && k[1]!.localName === "mo" ? (k[1]!.textContent ?? "").trim() : "";
  return BRACE[g] ? [BRACE[g]!, k[0]!] : null;
};

// An accent glyph (the mo over a base) -> its LaTeX command.
const ACCENT: Record<string, string> = {
  "^": "\\hat", "ˆ": "\\hat", "~": "\\tilde", "˜": "\\tilde", "‾": "\\bar",
  "ˉ": "\\bar", "¯": "\\bar", "→": "\\vec", "⃗": "\\vec", "˙": "\\dot",
  ".": "\\dot", "¨": "\\ddot", "ˇ": "\\check",
};

// A LaTeX argument: braces only when it is more than a single character / command (so "x^2" stays
// "x^{2}" but "(a+b)^n" becomes "{...}^{n}" with the base braced where needed).
const arg = (el: Element | undefined): string => {
  const s = el ? nodeLatex(el) : "";
  return s.length <= 1 || /^\\[a-zA-Z]+ ?$/.test(s) ? s.trim() : s;
};
// A script base: like arg but kept inline (sup/sub add their own braces around the script, not the base).
const base = (el: Element | undefined): string => {
  const s = el ? nodeLatex(el) : "";
  return s.length <= 1 || /^\\[a-zA-Z]+ ?$/.test(s.trim()) ? s.trim() : `{${s}}`;
};

function nodeLatex(el: Element): string {
  const k = kidsOf(el);
  switch (el.localName) {
    case "mrow":
      return fencedMatrixLatex(el) ?? k.map(nodeLatex).join("");
    case "math":
    case "mstyle":
    case "mpadded":
      return k.map(nodeLatex).join("");
    case "semantics": // presentation child first; the annotation (StarMath / TeX) is dropped
      return k.length ? nodeLatex(k[0]!) : "";
    case "mi":
    case "mn":
    case "mo":
    case "ms":
      return tokenLatex(el.textContent ?? "");
    case "mtext": {
      const s = (el.textContent ?? "").replace(INVISIBLE, "");
      return s.trim() ? `\\text{${s}}` : s;
    }
    case "mspace":
      return " ";
    case "mfrac":
      return `\\frac{${arg(k[0])}}{${arg(k[1])}}`;
    case "msup":
      return `${base(k[0])}^{${arg(k[1])}}`;
    case "msub":
      return `${base(k[0])}_{${arg(k[1])}}`;
    case "msubsup":
      return `${base(k[0])}_{${arg(k[1])}}^{${arg(k[2])}}`;
    case "munder": {
      const inner = braceInner(el); // bare \underbrace{body}
      if (inner) return `${inner[0]}{${arg(inner[1])}}`;
      const labeled = k[0] ? braceInner(k[0]) : null; // \underbrace{body}_{label}
      if (labeled) return `${labeled[0]}{${arg(labeled[1])}}_{${arg(k[1])}}`;
      return `${base(k[0])}_{${arg(k[1])}}`;
    }
    case "mover": {
      const inner = braceInner(el); // bare \overbrace{body}
      if (inner) return `${inner[0]}{${arg(inner[1])}}`;
      const labeled = k[0] ? braceInner(k[0]) : null; // \overbrace{body}^{label}
      if (labeled) return `${labeled[0]}{${arg(labeled[1])}}^{${arg(k[1])}}`;
      const over = k[1];
      const cmd = over && over.localName === "mo" ? ACCENT[(over.textContent ?? "").trim()] : undefined;
      return cmd ? `${cmd}{${arg(k[0])}}` : `${base(k[0])}^{${arg(k[1])}}`;
    }
    case "menclose": {
      const notation = el.getAttribute("notation") ?? "";
      const cmd = notation.includes("bottom") || notation === "underline" ? "\\underline" : "\\overline";
      return `${cmd}{${k.map(nodeLatex).join("")}}`;
    }
    case "mtable":
      return `\\begin{matrix}${tableBody(el)}\\end{matrix}`;
    case "munderover":
      return `${base(k[0])}_{${arg(k[1])}}^{${arg(k[2])}}`;
    case "msqrt":
      return `\\sqrt{${k.map(nodeLatex).join("")}}`;
    case "mroot":
      return `\\sqrt[${arg(k[1])}]{${arg(k[0])}}`;
    case "mfenced": { // legacy delimiter element: re-wrap its children in parentheses
      const open = el.getAttribute("open") ?? "(";
      const close = el.getAttribute("close") ?? ")";
      return `${open}${k.map(nodeLatex).join(el.getAttribute("separators") ?? ",")}${close}`;
    }
    case "annotation":
    case "annotation-xml":
      return "";
    default:
      return k.length ? k.map(nodeLatex).join("") : tokenLatex(el.textContent ?? "");
  }
}

/** Best-effort LaTeX for a MathML <math> element, for re-editing an imported equation. */
export function mathmlToLatex(math: Element): string {
  try {
    return nodeLatex(math).replace(/\s+/g, " ").trim();
  } catch {
    return "";
  }
}
