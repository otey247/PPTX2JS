/**
 * Color normalization utilities for OOXML → PptxGenJS.
 */

/**
 * Normalize a color value to a 6-digit hex string (without the leading #).
 * Handles sRGB hex colors from OOXML (already 6-digit hex).
 */
export function normalizeColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  // Strip leading # if present
  const hex = value.replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) {
    return hex;
  }
  // Handle 8-digit ARGB (OOXML sometimes uses AARRGGBB)
  if (/^[0-9A-F]{8}$/.test(hex)) {
    return hex.slice(2); // Drop alpha channel
  }
  return undefined;
}

/**
 * Resolve a theme color reference to a hex value.
 * themeColorMap is built from the theme XML during extraction.
 */
export function resolveThemeColor(
  schemeClr: string,
  themeColorMap: Record<string, string>
): string | undefined {
  return themeColorMap[schemeClr];
}

/**
 * Parse a lumMod/lumOff tint/shade adjustment on a theme color.
 * Returns the adjusted hex color string.
 */
export function applyLumAdjust(
  baseHex: string,
  lumMod?: number,
  lumOff?: number
): string {
  if (lumMod === undefined && lumOff === undefined) return baseHex;

  const r = parseInt(baseHex.slice(0, 2), 16);
  const g = parseInt(baseHex.slice(2, 4), 16);
  const b = parseInt(baseHex.slice(4, 6), 16);

  const mod = lumMod !== undefined ? lumMod / 100000 : 1;
  const off = lumOff !== undefined ? lumOff / 100000 : 0;

  const clamp = (v: number) => Math.min(255, Math.max(0, Math.round(v)));
  const nr = clamp(r * mod + off * 255);
  const ng = clamp(g * mod + off * 255);
  const nb = clamp(b * mod + off * 255);

  return [nr, ng, nb]
    .map((v) => v.toString(16).padStart(2, "0").toUpperCase())
    .join("");
}
