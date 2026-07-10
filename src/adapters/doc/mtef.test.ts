import { describe, expect, it } from "vitest";
import { mtefToMathml } from "./mtef";

// Wrap a raw MTEF v3 record stream in an "Equation Native" buffer: a 28-byte OLE header (only
// cbHdr matters to us) followed by the MTEF header + records. The record bytes below are the
// exact sequences LibreOffice emits, verified against real .doc oracles.
function eqNative(records: number[]): Uint8Array {
  const header = [0x1c, 0, 0, 0, 0, 0, 0, 0]; // cbHdr = 28, rest ignored
  while (header.length < 28) header.push(0);
  const mtef = [0x03, 0x01, 0x01, 0x03, 0x00, ...records]; // MTEF v3 header + records
  return Uint8Array.from([...header, ...mtef]);
}
const CH = (code: number) => [0x02, 0x83, code, 0x00]; // a CHAR record for an ASCII code

describe("mtef -> mathml", () => {
  it("decodes a plain sequence a+b", () => {
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, ...CH(0x61), ...CH(0x2b), ...CH(0x62), 0x00, 0x00, 0x00]));
    expect(m).toContain("<mi>a</mi><mo>+</mo><mi>b</mi>");
  });
  it("decodes a superscript x^2 as msup", () => {
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, ...CH(0x78), 0x03, 0x0f, 0x00, 0x00, 0x0b, 0x11, 0x01, ...CH(0x32), 0x00, 0x00, 0x0a, 0x00, 0x00, 0x00]));
    expect(m).toContain("<msup><mi>x</mi><mn>2</mn></msup>");
  });
  it("decodes a subscript x_i as msub", () => {
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, ...CH(0x78), 0x03, 0x0f, 0x01, 0x00, 0x0b, 0x01, ...CH(0x69), 0x00, 0x11, 0x00, 0x0a, 0x00, 0x00, 0x00]));
    expect(m).toContain("<msub><mi>x</mi><mi>i</mi></msub>");
  });
  it("decodes a fraction a/b as mfrac", () => {
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, 0x03, 0x0e, 0x00, 0x00, 0x0a, 0x01, ...CH(0x61), 0x00, 0x0a, 0x01, ...CH(0x62), 0x00, 0x00, 0x00, 0x00, 0x00]));
    expect(m).toContain("<mfrac><mi>a</mi><mi>b</mi></mfrac>");
  });
  it("decodes a square root as msqrt", () => {
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, 0x03, 0x0d, 0x00, 0x00, 0x01, ...CH(0x78), 0x00, 0x11, 0x00, 0x00, 0x00, 0x00]));
    expect(m).toContain("<msqrt><mi>x</mi></msqrt>");
  });
  it("bails (returns null) on a non-ASCII character rather than emitting wrong math", () => {
    // A CHAR whose MathType code is 0x03b1 (Greek alpha): outside the safe ASCII subset.
    const m = mtefToMathml(eqNative([0x0a, 0x01, 0x0a, 0x01, 0x02, 0x86, 0xb1, 0x03, 0x00, 0x00]));
    expect(m).toBeNull();
  });
  it("bails on an unknown record", () => {
    expect(mtefToMathml(eqNative([0x0a, 0x01, 0x07, 0x99, 0x00]))).toBeNull(); // 0x07 = RULER, not handled
  });
  it("returns null for a non-MTEF-v3 stream", () => {
    const bad = eqNative([]);
    bad[28] = 0x05; // MTEF version 5, not supported
    expect(mtefToMathml(bad)).toBeNull();
  });
});
