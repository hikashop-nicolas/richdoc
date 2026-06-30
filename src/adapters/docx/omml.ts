// OMML (Office Math, docx) <-> MathML conversion for the common constructs: runs, grouping,
// fractions, scripts, radicals, n-ary operators (sum / integral / product), delimiters, matrices,
// accents, bars and over/under braces. MathML is the editor's in-document representation; this is the only place that
// knows OMML. Anything outside the common set is read best-effort (the original OMML is kept verbatim
// for a lossless rewrite) and, when authored, falls back to a plain run. Element matching is by
// localName so the prefix ("m") is irrelevant.
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MML = "http://www.w3.org/1998/Math/MathML";
const NARY_OVER = "∑∏⋃⋂⋁⋀∐"; // operators whose limits sit above/below (vs. at the corner)
const OPERATOR = /[+\-*/=<>±×÷·∗∘≤≥≠≈≡→←↔∈∉⊂⊆∪∩∧∨¬∀∃∇∂∞%!|,;:(){}[\]]/;
// An accent over a base: the spacing glyph MathML (temml) uses <-> the combining char OMML stores.
const ACCENT_TO_COMBINING: Record<string, string> = {
  "^": "̂", "ˆ": "̂", "~": "̃", "˜": "̃", "∼": "̃",
  "ˉ": "̄", "‾": "̄", "¯": "̄", "→": "⃗", "⃗": "⃗",
  "˙": "̇", ".": "̇", "¨": "̈", "ˇ": "̌",
};
const COMBINING_TO_GLYPH: Record<string, string> = {
  "̂": "^", "̃": "~", "̄": "‾", "⃗": "→", "̇": "˙", "̈": "¨", "̌": "ˇ",
};
const BRACE_OVER = "⏞"; // over-brace glyph (U+23DE), an OMML group-char positioned on top
const BRACE_UNDER = "⏟"; // under-brace glyph (U+23DF), positioned on the bottom

const INVISIBLE = /[⁡-⁤]/g; // function application / invisible times / separator / plus
const escXml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escAttr = (s: string): string => escXml(s).replace(/"/g, "&quot;");
const ln = (el: Element): string => el.localName;
const elemChildren = (el: Element): Element[] => Array.from(el.children);

// --- MathML -> OMML (authored equations) -------------------------------------------------------

export function mathmlToOmml(math: Element): string {
  return `<m:oMath xmlns:m="${M}">${mmlSeq(Array.from(math.childNodes))}</m:oMath>`;
}

const mRun = (text: string): string => { const t = text.replace(INVISIBLE, ""); return t ? `<m:r><m:t xml:space="preserve">${escXml(t)}</m:t></m:r>` : ""; };
// mathvariant (or the MathML default that a multi-letter mi is upright) -> OMML run style. A math run
// is italic by default, so only a non-italic style needs an m:sty; this keeps functions (sin, log)
// and \mathrm upright in Word instead of italicised.
const STY: Record<string, string> = { normal: "p", bold: "b", italic: "i", "bold-italic": "bi" };
const miRun = (el: Element): string => {
  const t = (el.textContent ?? "").replace(INVISIBLE, "");
  if (!t) return "";
  const mv = el.getAttribute("mathvariant") ?? (t.length > 1 ? "normal" : "");
  const sty = STY[mv];
  const pr = sty && sty !== "i" ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : "";
  return `<m:r>${pr}<m:t xml:space="preserve">${escXml(t)}</m:t></m:r>`;
};
// A slot (numerator, base, sub, ...) holding one MathML node -> its OMML, flattening a wrapping mrow.
const mmlSlot = (node: Element | null): string => {
  if (!node) return "";
  return ln(node) === "mrow" ? mmlSeq(Array.from(node.childNodes)) : mmlSeq([node]);
};
const naryOmml = (chr: string, sub: Element | null, sup: Element | null, rest: ChildNode[]): string => {
  const pr = `<m:naryPr><m:chr m:val="${escAttr(chr)}"/><m:limLoc m:val="${NARY_OVER.includes(chr) ? "undOvr" : "subSup"}"/>` +
    `${sub ? "" : '<m:subHide m:val="1"/>'}${sup ? "" : '<m:supHide m:val="1"/>'}</m:naryPr>`;
  return `<m:nary>${pr}<m:sub>${mmlSlot(sub)}</m:sub><m:sup>${mmlSlot(sup)}</m:sup><m:e>${mmlSeq(rest)}</m:e></m:nary>`;
};
// An mtable -> an OMML matrix (m:m). The column count comes from the widest row.
const matrixOmml = (table: Element): string => {
  const rows = elemChildren(table).filter((c) => ln(c) === "mtr");
  const cols = Math.max(0, ...rows.map((r) => elemChildren(r).filter((c) => ln(c) === "mtd").length));
  const mr = rows.map((r) => `<m:mr>${elemChildren(r).filter((c) => ln(c) === "mtd").map((c) => `<m:e>${mmlSeq(Array.from(c.childNodes))}</m:e>`).join("")}</m:mr>`).join("");
  return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${mr}</m:m>`;
};
// An over/underbrace -> an OMML group character over (top) or under (bottom) the base.
const groupChrOmml = (chr: string, body: string): string => {
  const pos = chr === BRACE_OVER ? "top" : "bot";
  return `<m:groupChr><m:groupChrPr><m:chr m:val="${escAttr(chr)}"/><m:pos m:val="${pos}"/><m:vertJc m:val="${pos === "top" ? "bot" : "top"}"/></m:groupChrPr><m:e>${body}</m:e></m:groupChr>`;
};
// A delimited matrix: an mrow of exactly [open mo, mtable, close mo] (what temml emits for
// pmatrix / bmatrix / vmatrix / cases). Returns the brackets + table, or null if it is a plain mrow.
const fencedMatrix = (el: Element): { open: string; close: string; table: Element } | null => {
  const k = elemChildren(el);
  if (k.length === 3 && ln(k[0]!) === "mo" && ln(k[1]!) === "mtable" && ln(k[2]!) === "mo")
    return { open: (k[0]!.textContent ?? "").trim(), close: (k[2]!.textContent ?? "").trim(), table: k[1]! };
  return null;
};

function mmlSeq(nodes: ChildNode[]): string {
  let out = "";
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i]!;
    if (n.nodeType === 3) { const t = n.textContent ?? ""; if (t.trim()) out += mRun(t); continue; }
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    const kids = elemChildren(el);
    switch (ln(el)) {
      case "mi": out += miRun(el); break;
      case "mn": case "mo": case "mtext": case "ms": out += mRun(el.textContent ?? ""); break;
      case "mrow": case "mstyle": case "mpadded": {
        const fm = fencedMatrix(el); // pmatrix / bmatrix / cases: brackets around a matrix
        if (fm) out += `<m:d><m:dPr><m:begChr m:val="${escAttr(fm.open)}"/><m:endChr m:val="${escAttr(fm.close)}"/></m:dPr><m:e>${matrixOmml(fm.table)}</m:e></m:d>`;
        else out += mmlSeq(Array.from(el.childNodes));
        break;
      }
      case "semantics": out += mmlSeq(Array.from((kids[0] ?? el).childNodes)); break; // presentation child
      case "mfrac": out += `<m:f><m:num>${mmlSlot(kids[0] ?? null)}</m:num><m:den>${mmlSlot(kids[1] ?? null)}</m:den></m:f>`; break;
      case "msqrt": out += `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${mmlSeq(Array.from(el.childNodes))}</m:e></m:rad>`; break;
      case "mroot": out += `<m:rad><m:deg>${mmlSlot(kids[1] ?? null)}</m:deg><m:e>${mmlSlot(kids[0] ?? null)}</m:e></m:rad>`; break;
      case "msub": case "msup": case "msubsup": case "munder": case "mover": case "munderover": {
        const base = kids[0] ?? null;
        const tag = ln(el);
        // An accent (hat / bar / vec / ...): mover whose script is a single accent glyph -> m:acc.
        if (tag === "mover" && kids[1] && ln(kids[1]!) === "mo") {
          const comb = ACCENT_TO_COMBINING[(kids[1]!.textContent ?? "").trim()];
          if (comb) { out += `<m:acc><m:accPr><m:chr m:val="${escAttr(comb)}"/></m:accPr><m:e>${mmlSlot(base)}</m:e></m:acc>`; break; }
        }
        // An over/underbrace: mover/munder with a brace glyph (bare), or a label wrapping one.
        if (tag === "mover" || tag === "munder") {
          const g = kids[1] && ln(kids[1]!) === "mo" ? (kids[1]!.textContent ?? "").trim() : "";
          if (g === BRACE_OVER || g === BRACE_UNDER) { out += groupChrOmml(g, mmlSlot(base)); break; }
          const bk = base ? elemChildren(base) : []; // labeled: base is itself a brace mover/munder
          const bg = bk[1] && ln(bk[1]!) === "mo" ? (bk[1]!.textContent ?? "").trim() : "";
          if (bg === BRACE_OVER || bg === BRACE_UNDER) {
            const gc = groupChrOmml(bg, mmlSlot(bk[0] ?? null)), lim = mmlSlot(kids[1] ?? null);
            out += bg === BRACE_OVER ? `<m:limUpp><m:e>${gc}</m:e><m:lim>${lim}</m:lim></m:limUpp>` : `<m:limLow><m:e>${gc}</m:e><m:lim>${lim}</m:lim></m:limLow>`;
            break;
          }
        }
        const sub = tag === "msup" || tag === "mover" ? null : kids[1] ?? null;
        const sup = tag === "msub" || tag === "munder" ? null : tag === "msup" || tag === "mover" ? kids[1] ?? null : kids[2] ?? null;
        const chr = base && ln(base) === "mo" ? (base.textContent ?? "").trim() : "";
        // An n-ary operator (sum / integral / ...) absorbs the rest of the row as its body.
        if (chr && (NARY_OVER.includes(chr) || "∫∬∭∮".includes(chr))) { out += naryOmml(chr, sub, sup, nodes.slice(i + 1)); return out; }
        const e = `<m:e>${mmlSlot(base)}</m:e>`;
        if (sub && sup) out += `<m:sSubSup>${e}<m:sub>${mmlSlot(sub)}</m:sub><m:sup>${mmlSlot(sup)}</m:sup></m:sSubSup>`;
        else if (sub) out += `<m:sSub>${e}<m:sub>${mmlSlot(sub)}</m:sub></m:sSub>`;
        else if (sup) out += `<m:sSup>${e}<m:sup>${mmlSlot(sup)}</m:sup></m:sSup>`;
        else out += mmlSlot(base);
        break;
      }
      case "mfenced": {
        const open = el.getAttribute("open") ?? "(", close = el.getAttribute("close") ?? ")";
        out += `<m:d><m:dPr><m:begChr m:val="${escAttr(open)}"/><m:endChr m:val="${escAttr(close)}"/></m:dPr><m:e>${mmlSeq(Array.from(el.childNodes))}</m:e></m:d>`;
        break;
      }
      case "mtable": out += matrixOmml(el); break;
      case "menclose": { // overline / underline enclosure -> a bar
        const notation = el.getAttribute("notation") ?? "";
        const pos = notation.includes("bottom") || notation === "underline" ? "bot" : "top";
        out += `<m:bar><m:barPr><m:pos m:val="${pos}"/></m:barPr><m:e>${mmlSeq(Array.from(el.childNodes))}</m:e></m:bar>`;
        break;
      }
      default: out += mRun(el.textContent ?? ""); // unknown: keep its text
    }
  }
  return out;
}

// --- OMML -> MathML (display / re-edit of imported equations) -----------------------------------

export function ommlToMathml(oMath: Element): string {
  return `<math xmlns="${MML}" display="inline">${ommlSeq(elemChildren(oMath))}</math>`;
}

// OMML run style -> the MathML mathvariant put on identifiers (italic is the default, so omitted).
const VARIANT: Record<string, string> = { p: "normal", b: "bold", bi: "bold-italic" };
const mmlTok = (raw: string, variant?: string): string => {
  // Split a run into identifier / number / operator / other tokens for sensible MathML spacing.
  const text = raw.replace(INVISIBLE, "");
  const mv = variant ? ` mathvariant="${variant}"` : "";
  let out = "";
  for (const m of text.matchAll(/(\d+\.?\d*)|([A-Za-z]+)|(\s+)|(.)/g)) {
    if (m[1]) out += `<mn>${escXml(m[1])}</mn>`;
    else if (m[2]) out += `<mi${mv}>${escXml(m[2])}</mi>`;
    else if (m[3]) out += "<mspace width=\"0.2em\"/>";
    else if (m[4]) out += OPERATOR.test(m[4]) ? `<mo>${escXml(m[4])}</mo>` : `<mtext>${escXml(m[4])}</mtext>`;
  }
  return out;
};
const child = (el: Element, name: string): Element | undefined => elemChildren(el).find((c) => ln(c) === name);
const ommlSlot = (slot: Element | undefined): string => {
  if (!slot) return "<mrow></mrow>";
  const inner = ommlSeq(elemChildren(slot));
  return elemChildren(slot).length === 1 ? inner || "<mrow></mrow>" : `<mrow>${inner}</mrow>`;
};

function ommlSeq(els: Element[]): string {
  let out = "";
  for (const el of els) {
    switch (ln(el)) {
      case "r": { // a run; m:rPr/m:sty carries an upright/bold style for its identifiers
        const sty = child(child(el, "rPr") ?? el, "sty")?.getAttribute("m:val");
        out += mmlTok(child(el, "t")?.textContent ?? el.textContent ?? "", sty ? VARIANT[sty] : undefined);
        break;
      }
      case "f": out += `<mfrac>${ommlSlot(child(el, "num"))}${ommlSlot(child(el, "den"))}</mfrac>`; break;
      case "sSup": out += `<msup>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "sup"))}</msup>`; break;
      case "sSub": out += `<msub>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "sub"))}</msub>`; break;
      case "sSubSup": out += `<msubsup>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "sub"))}${ommlSlot(child(el, "sup"))}</msubsup>`; break;
      case "rad": {
        const degHidden = !!child(child(el, "radPr") ?? el, "degHide") || !child(el, "deg") || !elemChildren(child(el, "deg")!).length;
        out += degHidden ? `<msqrt>${ommlSlot(child(el, "e"))}</msqrt>` : `<mroot>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "deg"))}</mroot>`;
        break;
      }
      case "nary": {
        const pr = child(el, "naryPr");
        const chr = (pr && child(pr, "chr")?.getAttribute("m:val")) || "∫";
        const over = NARY_OVER.includes(chr);
        const sub = child(el, "sub"), sup = child(el, "sup");
        const op = `<mo>${escXml(chr)}</mo>`;
        let big: string;
        if (sub && sup) big = `<${over ? "munderover" : "msubsup"}>${op}${ommlSlot(sub)}${ommlSlot(sup)}</${over ? "munderover" : "msubsup"}>`;
        else if (sub) big = `<${over ? "munder" : "msub"}>${op}${ommlSlot(sub)}</${over ? "munder" : "msub"}>`;
        else if (sup) big = `<${over ? "mover" : "msup"}>${op}${ommlSlot(sup)}</${over ? "mover" : "msup"}>`;
        else big = op;
        out += `<mrow>${big}${ommlSlot(child(el, "e"))}</mrow>`;
        break;
      }
      case "d": {
        const pr = child(el, "dPr");
        const beg = (pr && child(pr, "begChr")?.getAttribute("m:val")) ?? "(";
        const end = (pr && child(pr, "endChr")?.getAttribute("m:val")) ?? ")";
        const delim = (c: string): string => `<mo fence="true" stretchy="true">${escXml(c)}</mo>`;
        const body = elemChildren(el).filter((c) => ln(c) === "e").map((e) => ommlSlot(e)).join(`<mo>,</mo>`);
        out += `<mrow>${delim(beg)}${body}${delim(end)}</mrow>`;
        break;
      }
      case "func": out += `<mrow>${ommlSlot(child(el, "fName"))}${ommlSlot(child(el, "e"))}</mrow>`; break;
      case "m": { // a matrix: m:mr rows of m:e cells
        const rows = elemChildren(el).filter((c) => ln(c) === "mr");
        const mtr = rows.map((r) => `<mtr>${elemChildren(r).filter((c) => ln(c) === "e").map((c) => `<mtd>${ommlSeq(elemChildren(c))}</mtd>`).join("")}</mtr>`).join("");
        out += `<mtable>${mtr}</mtable>`;
        break;
      }
      case "acc": { // an accent over a base
        const comb = (child(child(el, "accPr") ?? el, "chr")?.getAttribute("m:val")) || "̂";
        out += `<mover accent="true">${ommlSlot(child(el, "e"))}<mo>${escXml(COMBINING_TO_GLYPH[comb] ?? comb)}</mo></mover>`;
        break;
      }
      case "bar": { // an overline / underline
        const pos = (child(child(el, "barPr") ?? el, "pos")?.getAttribute("m:val")) || "top";
        out += `<menclose notation="${pos === "bot" ? "bottom" : "top"}">${ommlSlot(child(el, "e"))}</menclose>`;
        break;
      }
      case "groupChr": { // an over/under brace (no label)
        const pr = child(el, "groupChrPr");
        const chr = (pr && child(pr, "chr")?.getAttribute("m:val")) || BRACE_UNDER;
        const tag = ((pr && child(pr, "pos")?.getAttribute("m:val")) || "bot") === "top" ? "mover" : "munder";
        out += `<${tag}>${ommlSlot(child(el, "e"))}<mo stretchy="true">${escXml(chr)}</mo></${tag}>`;
        break;
      }
      case "limUpp": out += `<mover>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "lim"))}</mover>`; break; // a brace with a label above
      case "limLow": out += `<munder>${ommlSlot(child(el, "e"))}${ommlSlot(child(el, "lim"))}</munder>`; break; // ... or below
      case "e": case "num": case "den": case "sub": case "sup": case "deg": case "fName": case "lim": out += ommlSeq(elemChildren(el)); break;
      case "rPr": case "naryPr": case "fPr": case "dPr": case "radPr": case "ctrlPr": case "mPr": case "accPr": case "barPr": case "groupChrPr": break; // properties, no glyphs
      default: { const t = el.textContent ?? ""; if (t.trim()) out += mmlTok(t); }
    }
  }
  return out;
}
