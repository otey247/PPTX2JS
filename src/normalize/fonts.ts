/**
 * Font normalization utilities for OOXML → PptxGenJS.
 */

/**
 * Resolve a font typeface, handling "+mj-lt" and "+mn-lt" theme font references.
 */
export function resolveFontFace(
  typeface: string | undefined,
  majorFont?: string,
  minorFont?: string
): string | undefined {
  if (!typeface) return undefined;
  if (typeface === "+mj-lt") return majorFont ?? "Calibri Light";
  if (typeface === "+mn-lt") return minorFont ?? "Calibri";
  if (typeface.startsWith("+")) return undefined; // Unknown theme font ref
  return typeface;
}
