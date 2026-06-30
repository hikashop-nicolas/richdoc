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

// A token (mi/mn/mo): map known symbols to commands (with a trailing space so "\pm b" stays two
// tokens), pass plain characters through.
const tokenLatex = (raw: string): string => {
  const s = raw.trim();
  if (!s) return "";
  if (SYM[s]) return /^\\[a-zA-Z]+$/.test(SYM[s]) ? `${SYM[s]} ` : SYM[s];
  return s;
};

const kidsOf = (el: Element): Element[] => Array.from(el.children);

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
    case "math":
    case "mrow":
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
      const s = el.textContent ?? "";
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
    case "munder":
      return `${base(k[0])}_{${arg(k[1])}}`;
    case "mover": {
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
      return `\\begin{matrix}${k
        .filter((r) => r.localName === "mtr")
        .map((r) => kidsOf(r).filter((c) => c.localName === "mtd").map(nodeLatex).join(" & "))
        .join(" \\\\ ")}\\end{matrix}`;
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
