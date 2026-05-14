/**
 * Unit conversion utilities for OOXML → PptxGenJS.
 * PowerPoint OOXML uses EMUs (English Metric Units).
 * PptxGenJS uses inches by default.
 */

export const EMU_PER_INCH = 914400;
export const EMU_PER_PT = 12700;

/**
 * Convert EMU value to inches.
 */
export function emuToInches(value: string | number): number {
  return Number(value) / EMU_PER_INCH;
}

/**
 * Convert EMU value to points.
 */
export function emuToPt(value: string | number): number {
  return Number(value) / EMU_PER_PT;
}

/**
 * Convert OOXML font size (hundredths of a point) to points.
 * The OOXML `sz` attribute stores font sizes as hundredths of a point,
 * so 2400 means 24pt.
 */
export function ooxmlSzToPt(value: string | number): number {
  return Number(value) / 100;
}

/**
 * Convert inches to EMU.
 */
export function inchesToEmu(value: number): number {
  return Math.round(value * EMU_PER_INCH);
}
