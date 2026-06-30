import { describe, expect, it } from "vitest";
import { formatPageNumber } from "./util";

describe("formatPageNumber", () => {
  it("formats decimal (and treats no format / unknown as decimal)", () => {
    expect(formatPageNumber(7, "decimal")).toBe("7");
    expect(formatPageNumber(7, undefined)).toBe("7");
    expect(formatPageNumber(7, "ordinal")).toBe("7");
  });

  it("formats roman numerals in both cases", () => {
    expect(formatPageNumber(4, "lowerRoman")).toBe("iv");
    expect(formatPageNumber(2026, "upperRoman")).toBe("MMXXVI");
    expect(formatPageNumber(49, "lowerRoman")).toBe("xlix");
  });

  it("formats letters bijectively in both cases", () => {
    expect(formatPageNumber(1, "lowerLetter")).toBe("a");
    expect(formatPageNumber(26, "upperLetter")).toBe("Z");
    expect(formatPageNumber(27, "lowerLetter")).toBe("aa");
    expect(formatPageNumber(28, "upperLetter")).toBe("AB");
  });

  it("returns the number unchanged below 1", () => {
    expect(formatPageNumber(0, "upperRoman")).toBe("0");
  });
});
