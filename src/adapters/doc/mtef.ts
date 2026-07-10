// Decode an MS Equation 3.0 object (the "Equation Native" OLE stream of a binary .doc) into
// presentation MathML, for the common constructs LibreOffice/Word emit. The format is MTEF v3
// (MathType Equation Format): a 28-byte OLE header, a 5-byte MTEF header, then a record stream.
// Records were reverse-engineered against LibreOffice-authored oracles (x^2, x_i, a/b, sqrt x,
// a+b). To honour "never emit wrong math", anything outside the verified subset (unknown record
// or template, a non-ASCII character whose MathType code may not be Unicode) aborts the whole
// parse and returns null, so the reader shows a plain placeholder instead of a fabricated formula.

// Record tags (low value = type; the high nibble carries options we do not need here).
const REC_END = 0x00;
const REC_LINE = 0x01; // opens a slot; its contents run until a matching END
const REC_CHAR = 0x02; // tag + typeface byte + character (uint16)
const REC_TMPL = 0x03; // tag + selector byte + variation (uint16), then its slots, then END
const REC_NULLLINE = 0x11; // an empty slot (a LINE with the null option); no contents, no END
// Size records carry no operand and only change the current font size, never structure/content.
const SIZE_RECS = new Set([0x0a, 0x0b, 0x0c, 0x0e, 0x0f]); // FULL, SUB, SUB2, SYM, SUBSYM

// Template selectors we can map to MathML. (These byte values only appear right after a TMPL
// tag, so they never collide with the size records that share some numeric values.)
const TMPL_SCRIPTS = 0x0f; // slots in stream order: [subscript, superscript]
const TMPL_FRACTION = 0x0e; // slots: [numerator, denominator]
const TMPL_RADICAL = 0x0d; // slots: [radicand, index]

type Node =
  | { t: "char"; c: string }
  | { t: "group"; kids: Node[] }
  | { t: "tmpl"; sel: number; base: Node | null; slots: (Node[] | null)[] };

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Decode the "Equation Native" stream to a MathML string, or null if it is not a parseable
    MTEF v3 equation within our verified subset. */
export function mtefToMathml(eq: Uint8Array): string | null {
  if (eq.length < 33) return null;
  const dv = new DataView(eq.buffer, eq.byteOffset, eq.byteLength);
  const cbHdr = dv.getUint16(0, true); // EQNOLEFILEHDR.cbHdr (always 28)
  const start = cbHdr;
  if (start + 5 > eq.length || eq[start] !== 3) return null; // only MTEF v3 (Equation Editor 3.0)
  let pos = start + 5; // skip the 5-byte MTEF header (version, platform, product, ver, subver)

  const fail = Symbol("fail");
  // Parse the contents of one slot/line: records until an END (which it consumes) or buffer end.
  const parseList = (): Node[] | typeof fail => {
    const nodes: Node[] = [];
    while (pos < eq.length) {
      const b = eq[pos];
      if (b === REC_END) { pos++; return nodes; }
      if (SIZE_RECS.has(b)) { pos++; continue; }
      if (b === REC_NULLLINE) { pos++; nodes.push({ t: "group", kids: [] }); continue; }
      if (b === REC_LINE) { pos++; const sub = parseList(); if (sub === fail) return fail; nodes.push({ t: "group", kids: sub }); continue; }
      if (b === REC_CHAR) {
        if (pos + 4 > eq.length) return fail;
        const code = eq[pos + 2] | (eq[pos + 3] << 8);
        pos += 4;
        if (code < 0x20 || code > 0x7e) return fail; // only characters whose MTEF code is ASCII
        nodes.push({ t: "char", c: String.fromCharCode(code) });
        continue;
      }
      if (b === REC_TMPL) {
        if (pos + 4 > eq.length) return fail;
        const sel = eq[pos + 1];
        pos += 4; // tag + selector + variation(uint16)
        if (sel !== TMPL_SCRIPTS && sel !== TMPL_FRACTION && sel !== TMPL_RADICAL) return fail;
        const slots = parseTemplateSlots();
        if (slots === fail) return fail;
        // A scripts template raises/lowers the element that precedes it in the line.
        const base = sel === TMPL_SCRIPTS ? nodes.pop() ?? null : null;
        nodes.push({ t: "tmpl", sel, base, slots });
        continue;
      }
      return fail; // an unknown record: abort rather than guess
    }
    return nodes;
  };
  // A template's body is a list of slots (LINE or NULLLINE), size records skipped, closed by END.
  const parseTemplateSlots = (): (Node[] | null)[] | typeof fail => {
    const slots: (Node[] | null)[] = [];
    while (pos < eq.length) {
      const b = eq[pos];
      if (b === REC_END) { pos++; return slots; }
      if (SIZE_RECS.has(b)) { pos++; continue; }
      if (b === REC_NULLLINE) { pos++; slots.push(null); continue; }
      if (b === REC_LINE) { pos++; const sub = parseList(); if (sub === fail) return fail; slots.push(sub); continue; }
      return fail;
    }
    return fail; // ran off the end without closing the template
  };

  const top = parseList();
  if (top === fail) return null;
  const inner = `<mrow>${top.map(render).join("")}</mrow>`;
  return `<math xmlns="http://www.w3.org/1998/Math/MathML">${inner}</math>`;
}

function slotMathml(slot: Node[] | null): string {
  if (!slot || !slot.length) return "<mrow></mrow>";
  return slot.length === 1 ? render(slot[0]) : `<mrow>${slot.map(render).join("")}</mrow>`;
}

function render(n: Node): string {
  if (n.t === "char") {
    const c = esc(n.c);
    if (/[A-Za-z]/.test(n.c)) return `<mi>${c}</mi>`;
    if (/[0-9]/.test(n.c)) return `<mn>${c}</mn>`;
    return `<mo>${c}</mo>`;
  }
  if (n.t === "group") {
    if (!n.kids.length) return "";
    return n.kids.length === 1 ? render(n.kids[0]) : `<mrow>${n.kids.map(render).join("")}</mrow>`;
  }
  // template
  const base = n.base ? render(n.base) : "<mrow></mrow>";
  if (n.sel === TMPL_SCRIPTS) {
    const sub = n.slots[0];
    const sup = n.slots[1];
    const hasSub = sub && sub.length;
    const hasSup = sup && sup.length;
    if (hasSub && hasSup) return `<msubsup>${base}${slotMathml(sub)}${slotMathml(sup)}</msubsup>`;
    if (hasSup) return `<msup>${base}${slotMathml(sup)}</msup>`;
    return `<msub>${base}${slotMathml(sub)}</msub>`;
  }
  if (n.sel === TMPL_FRACTION) return `<mfrac>${slotMathml(n.slots[0])}${slotMathml(n.slots[1])}</mfrac>`;
  // radical: an empty index slot is a square root, otherwise an n-th root.
  const index = n.slots[1];
  if (index && index.length) return `<mroot>${slotMathml(n.slots[0])}${slotMathml(index)}</mroot>`;
  return `<msqrt>${slotMathml(n.slots[0])}</msqrt>`;
}
