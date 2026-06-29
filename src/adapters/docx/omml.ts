// OMML (Office Math, docx) <-> MathML conversion for the common constructs: runs, grouping,
// fractions, scripts, radicals, n-ary operators (sum / integral / product) and delimiters. MathML is
// the editor's in-document representation; this is the only place that knows OMML. Anything outside the
// common set is read best-effort (the original OMML is kept verbatim for a lossless rewrite) and, when
// authored, falls back to a plain run. Element matching is by localName so the prefix ("m") is irrelevant.
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MML = "http://www.w3.org/1998/Math/MathML";
const NARY_OVER = "∑∏⋃⋂⋁⋀∐"; // operators whose limits sit above/below (vs. at the corner)
const OPERATOR = /[+\-*/=<>±×÷·∗∘≤≥≠≈≡→←↔∈∉⊂⊆∪∩∧∨¬∀∃∇∂∞%!|,;:(){}[\]]/;

const escXml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escAttr = (s: string): string => escXml(s).replace(/"/g, "&quot;");
const ln = (el: Element): string => el.localName;
const elemChildren = (el: Element): Element[] => Array.from(el.children);

// --- MathML -> OMML (authored equations) -------------------------------------------------------

export function mathmlToOmml(math: Element): string {
  return `<m:oMath xmlns:m="${M}">${mmlSeq(Array.from(math.childNodes))}</m:oMath>`;
}

const mRun = (text: string): string => (text ? `<m:r><m:t xml:space="preserve">${escXml(text)}</m:t></m:r>` : "");
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

function mmlSeq(nodes: ChildNode[]): string {
  let out = "";
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i]!;
    if (n.nodeType === 3) { const t = n.textContent ?? ""; if (t.trim()) out += mRun(t); continue; }
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    const kids = elemChildren(el);
    switch (ln(el)) {
      case "mi": case "mn": case "mo": case "mtext": case "ms": out += mRun(el.textContent ?? ""); break;
      case "mrow": case "mstyle": case "mpadded": out += mmlSeq(Array.from(el.childNodes)); break;
      case "semantics": out += mmlSeq(Array.from((kids[0] ?? el).childNodes)); break; // presentation child
      case "mfrac": out += `<m:f><m:num>${mmlSlot(kids[0] ?? null)}</m:num><m:den>${mmlSlot(kids[1] ?? null)}</m:den></m:f>`; break;
      case "msqrt": out += `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${mmlSeq(Array.from(el.childNodes))}</m:e></m:rad>`; break;
      case "mroot": out += `<m:rad><m:deg>${mmlSlot(kids[1] ?? null)}</m:deg><m:e>${mmlSlot(kids[0] ?? null)}</m:e></m:rad>`; break;
      case "msub": case "msup": case "msubsup": case "munder": case "mover": case "munderover": {
        const base = kids[0] ?? null;
        const tag = ln(el);
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
      default: out += mRun(el.textContent ?? ""); // unknown: keep its text
    }
  }
  return out;
}

// --- OMML -> MathML (display / re-edit of imported equations) -----------------------------------

export function ommlToMathml(oMath: Element): string {
  return `<math xmlns="${MML}" display="inline">${ommlSeq(elemChildren(oMath))}</math>`;
}

const mmlTok = (text: string): string => {
  // Split a run into identifier / number / operator / other tokens for sensible MathML spacing.
  let out = "";
  for (const m of text.matchAll(/(\d+\.?\d*)|([A-Za-z]+)|(\s+)|(.)/g)) {
    if (m[1]) out += `<mn>${escXml(m[1])}</mn>`;
    else if (m[2]) out += `<mi>${escXml(m[2])}</mi>`;
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
      case "r": out += mmlTok(child(el, "t")?.textContent ?? el.textContent ?? ""); break;
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
        const body = elemChildren(el).filter((c) => ln(c) === "e").map((e) => ommlSlot(e)).join(`<mo>,</mo>`);
        out += `<mrow><mo>${escXml(beg)}</mo>${body}<mo>${escXml(end)}</mo></mrow>`;
        break;
      }
      case "func": out += `<mrow>${ommlSlot(child(el, "fName"))}${ommlSlot(child(el, "e"))}</mrow>`; break;
      case "e": case "num": case "den": case "sub": case "sup": case "deg": case "fName": out += ommlSeq(elemChildren(el)); break;
      case "rPr": case "naryPr": case "fPr": case "dPr": case "radPr": case "ctrlPr": break; // properties, no glyphs
      default: { const t = el.textContent ?? ""; if (t.trim()) out += mmlTok(t); }
    }
  }
  return out;
}
