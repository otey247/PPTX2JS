import { normalizeColor, resolveThemeColor, applyLumAdjust } from "../normalize/colors";

describe("colors", () => {
  describe("normalizeColor", () => {
    test("returns uppercase 6-digit hex", () => {
      expect(normalizeColor("ff0000")).toBe("FF0000");
      expect(normalizeColor("#ff0000")).toBe("FF0000");
      expect(normalizeColor("FF0000")).toBe("FF0000");
    });

    test("strips alpha from 8-digit ARGB", () => {
      expect(normalizeColor("FFFF0000")).toBe("FF0000");
    });

    test("returns undefined for undefined input", () => {
      expect(normalizeColor(undefined)).toBeUndefined();
    });

    test("returns undefined for invalid hex", () => {
      expect(normalizeColor("ZZZ")).toBeUndefined();
    });
  });

  describe("resolveThemeColor", () => {
    test("returns mapped color", () => {
      const map = { accent1: "4472C4", dk1: "000000" };
      expect(resolveThemeColor("accent1", map)).toBe("4472C4");
    });

    test("returns undefined for unknown scheme color", () => {
      expect(resolveThemeColor("unknown", {})).toBeUndefined();
    });
  });

  describe("applyLumAdjust", () => {
    test("returns base color when no adjustments", () => {
      expect(applyLumAdjust("FF0000")).toBe("FF0000");
    });

    test("darkens color with lumMod < 100000", () => {
      // 50% darker: R=Math.round(255 * 0.5)=128=0x80, G=0, B=0 → "800000"
      const result = applyLumAdjust("FF0000", 50000);
      expect(result).toBe("800000");
    });

    test("lightens color with lumOff > 0", () => {
      const result = applyLumAdjust("000000", undefined, 100000);
      expect(result).toBe("FFFFFF");
    });
  });
});
