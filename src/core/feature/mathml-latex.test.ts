import { describe, expect, it } from "vitest";
import { mathmlToLatex } from "./mathml-latex";

const NS = ' xmlns="http://www.w3.org/1998/Math/MathML"';
// Parse a MathML string the way the editor holds it (HTML-parsed foreign content).
const recover = (inner: string): string => {
  const doc = new DOMParser().parseFromString(`<math${NS}>${inner}</math>`, "text/html");
  return mathmlToLatex(doc.querySelector("math")!);
};

describe("mathmlToLatex", () => {
  it("recovers a fraction", () => {
    expect(recover("<mfrac><mi>a</mi><mi>b</mi></mfrac>")).toBe("\\frac{a}{b}");
  });

  it("recovers superscripts, subscripts, and both", () => {
    expect(recover("<msup><mi>x</mi><mn>2</mn></msup>")).toBe("x^{2}");
    expect(recover("<msub><mi>x</mi><mn>1</mn></msub>")).toBe("x_{1}");
    expect(recover("<msubsup><mi>x</mi><mn>1</mn><mn>2</mn></msubsup>")).toBe("x_{1}^{2}");
  });

  it("recovers roots", () => {
    expect(recover("<msqrt><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></msqrt>")).toBe("\\sqrt{x+1}");
    expect(recover("<mroot><mi>x</mi><mn>3</mn></mroot>")).toBe("\\sqrt[3]{x}");
  });

  it("maps operators and Greek letters to commands", () => {
    expect(recover("<mi>&#945;</mi><mo>&#177;</mo><mi>&#946;</mi>")).toBe("\\alpha \\pm \\beta");
    expect(recover("<mi>x</mi><mo>&#8804;</mo><mi>y</mi>")).toBe("x\\le y");
    expect(recover("<mo>&#8722;</mo><mi>b</mi>")).toBe("-b"); // U+2212 minus -> hyphen
  });

  it("unwraps semantics and drops the StarMath annotation (odt formula shape)", () => {
    const odt =
      "<semantics><mrow><mi>x</mi><mo>=</mo><mi>y</mi></mrow>" +
      '<annotation encoding="StarMath 5.0">x = y</annotation></semantics>';
    expect(recover(odt)).toBe("x=y");
  });

  it("recovers the quadratic formula's structure", () => {
    const inner =
      "<mi>x</mi><mo>=</mo><mfrac>" +
      "<mrow><mo>&#8722;</mo><mi>b</mi><mo>&#177;</mo><msqrt><mrow>" +
      "<msup><mi>b</mi><mn>2</mn></msup><mo>&#8722;</mo><mn>4</mn><mi>a</mi><mi>c</mi>" +
      "</mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac>";
    const latex = recover(inner);
    expect(latex).toContain("\\frac{");
    expect(latex).toContain("\\sqrt{");
    expect(latex).toContain("\\pm");
    expect(latex).toContain("b^{2}");
    expect(latex).toContain("{2a}"); // denominator
  });

  it("recovers a matrix", () => {
    const inner =
      "<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>" +
      "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>";
    expect(recover(inner)).toBe("\\begin{matrix}a & b \\\\ c & d\\end{matrix}");
  });

  it("recovers accents and overline", () => {
    expect(recover('<mover accent="true"><mi>x</mi><mo>^</mo></mover>')).toBe("\\hat{x}");
    expect(recover('<mover accent="true"><mi>v</mi><mo>&#8594;</mo></mover>')).toBe("\\vec{v}"); // U+2192
    expect(recover('<mover accent="true"><mi>x</mi><mo>&#8254;</mo></mover>')).toBe("\\bar{x}"); // U+203E
    expect(recover('<menclose notation="top"><mrow><mi>A</mi><mi>B</mi></mrow></menclose>')).toBe("\\overline{AB}");
  });

  it("recovers delimited matrices as their named environment", () => {
    const tbl = "<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>";
    expect(recover(`<mrow><mo>(</mo>${tbl}<mo>)</mo></mrow>`)).toBe("\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}");
    expect(recover(`<mrow><mo>[</mo>${tbl}<mo>]</mo></mrow>`)).toBe("\\begin{bmatrix}a & b \\\\ c & d\\end{bmatrix}");
    expect(recover(`<mrow><mo>|</mo>${tbl}<mo>|</mo></mrow>`)).toBe("\\begin{vmatrix}a & b \\\\ c & d\\end{vmatrix}");
  });

  it("recovers a cases block (open brace, empty close)", () => {
    const tbl =
      "<mtable><mtr><mtd><mn>1</mn></mtd><mtd><mrow><mi>x</mi><mo>&gt;</mo><mn>0</mn></mrow></mtd></mtr>" +
      "<mtr><mtd><mn>0</mn></mtd><mtd><mrow><mi>x</mi><mo>&#8804;</mo><mn>0</mn></mrow></mtd></mtr></mtable>";
    expect(recover(`<mrow><mo>{</mo>${tbl}<mo></mo></mrow>`)).toBe("\\begin{cases}1 & x>0 \\\\ 0 & x\\le 0\\end{cases}");
  });

  it("falls back to text for unknown constructs and empty math", () => {
    expect(recover("")).toBe("");
    expect(recover("<munknown><mi>q</mi></munknown>")).toBe("q");
  });
});
