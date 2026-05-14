import { emuToInches, emuToPt, halfPtToPt, inchesToEmu, EMU_PER_INCH } from "../normalize/units";

describe("units", () => {
  test("emuToInches converts correctly", () => {
    expect(emuToInches(EMU_PER_INCH)).toBe(1);
    expect(emuToInches(EMU_PER_INCH * 2)).toBe(2);
    expect(emuToInches(0)).toBe(0);
    expect(emuToInches("914400")).toBe(1);
  });

  test("emuToPt converts correctly", () => {
    // 12700 EMU = 1 pt
    expect(emuToPt(12700)).toBe(1);
    expect(emuToPt(0)).toBe(0);
  });

  test("halfPtToPt converts correctly", () => {
    // 2400 hundredths-of-a-point = 24pt
    expect(halfPtToPt(2400)).toBe(24);
    expect(halfPtToPt(1200)).toBe(12);
  });

  test("inchesToEmu converts correctly", () => {
    expect(inchesToEmu(1)).toBe(EMU_PER_INCH);
    expect(inchesToEmu(0)).toBe(0);
  });

  test("round-trips correctly", () => {
    const inches = 3.5;
    expect(emuToInches(inchesToEmu(inches))).toBeCloseTo(inches);
  });
});
