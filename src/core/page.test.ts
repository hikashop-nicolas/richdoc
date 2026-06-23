import { describe, expect, it } from "vitest";
import { defaultPageGeometry, PAGE_SIZES, paginate } from "./page";

describe("page geometry defaults", () => {
  it("A4 and Letter sizes in px at 96 dpi", () => {
    expect(PAGE_SIZES.a4).toEqual({ widthPx: 794, heightPx: 1123 });
    expect(PAGE_SIZES.letter).toEqual({ widthPx: 816, heightPx: 1056 });
    expect(defaultPageGeometry("a4").margin).toEqual({ top: 96, right: 96, bottom: 96, left: 96 });
    expect(defaultPageGeometry().widthPx).toBe(794); // default is A4
  });
});

describe("paginate", () => {
  const M = { pageStep: 1100, contentHeight: 1000 };

  it("keeps everything on one page when it fits", () => {
    const r = paginate([100, 100, 100], M);
    expect(r.spacerBefore.size).toBe(0);
    expect(r.cardCount).toBe(1);
  });

  it("breaks before the block that would overflow", () => {
    const r = paginate([400, 400, 400, 400], M);
    // page 0 holds blocks 0,1 (800); block 2 would make 1200 > 1000 -> next page
    expect([...r.spacerBefore.keys()]).toEqual([2]);
    expect(r.spacerBefore.get(2)).toBe(300); // 1*1100 - 800
    expect(r.cardCount).toBe(2);
  });

  it("an oversized block overflows its card; the next block resumes after it", () => {
    const r = paginate([1500, 200], M);
    // block 0 (1500) overflows page 0 into card 1; block 1 must land on card 2
    expect(r.spacerBefore.get(1)).toBe(700); // 2*1100 - 1500
    expect(r.cardCount).toBe(3);
  });

  it("a single oversized block still draws enough cards", () => {
    const r = paginate([1500], M);
    expect(r.spacerBefore.size).toBe(0);
    expect(r.cardCount).toBe(2); // ceil(1500/1100)
  });

  it("empty document is one card", () => {
    const r = paginate([], M);
    expect(r.cardCount).toBe(1);
    expect(r.spacerBefore.size).toBe(0);
  });

  it("degenerate metrics fall back to a single card", () => {
    expect(paginate([100, 100], { pageStep: 0, contentHeight: 0 }).cardCount).toBe(1);
  });
});
