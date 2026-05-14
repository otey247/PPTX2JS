/**
 * Extract theme data (colors, fonts) from theme XML.
 */

import JSZip from "jszip";
import { ThemeModel } from "../types";
import { XML_PARSER } from "../ingest/read-pptx";
import { normalizeColor } from "../normalize/colors";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

/**
 * Extract theme from a theme XML part path (e.g. "ppt/theme/theme1.xml").
 */
export async function extractTheme(
  zip: JSZip,
  themePath: string
): Promise<ThemeModel | null> {
  try {
    const xml = await zip.file(themePath)?.async("text");
    if (!xml) return null;

    const parsed = XML_PARSER.parse(xml) as Record<string, unknown>;
    const themeRoot = parsed["a:theme"] as Record<string, unknown> | undefined;
    if (!themeRoot) return null;

    const themeName = getAttr(themeRoot, "name");
    const themeElements = themeRoot["a:themeElements"] as
      | Record<string, unknown>
      | undefined;
    if (!themeElements) return null;

    const colors: Record<string, string> = {};

    // Extract color scheme
    const clrScheme = themeElements["a:clrScheme"] as
      | Record<string, unknown>
      | undefined;
    if (clrScheme) {
      const colorNames = [
        "a:dk1",
        "a:lt1",
        "a:dk2",
        "a:lt2",
        "a:accent1",
        "a:accent2",
        "a:accent3",
        "a:accent4",
        "a:accent5",
        "a:accent6",
        "a:hlink",
        "a:folHlink",
      ];

      // Map OOXML color keys to PptxGenJS / common scheme names
      const keyMap: Record<string, string> = {
        "a:dk1": "dk1",
        "a:lt1": "lt1",
        "a:dk2": "dk2",
        "a:lt2": "lt2",
        "a:accent1": "accent1",
        "a:accent2": "accent2",
        "a:accent3": "accent3",
        "a:accent4": "accent4",
        "a:accent5": "accent5",
        "a:accent6": "accent6",
        "a:hlink": "hlink",
        "a:folHlink": "folHlink",
      };

      for (const cn of colorNames) {
        const node = clrScheme[cn] as Record<string, unknown> | undefined;
        if (!node) continue;

        // The color may be a:srgbClr or a:sysClr
        const srgb = node["a:srgbClr"] as Record<string, unknown> | undefined;
        if (srgb) {
          const hex = normalizeColor(getAttr(srgb, "val"));
          if (hex) colors[keyMap[cn]] = hex;
        }

        const sysClr = node["a:sysClr"] as Record<string, unknown> | undefined;
        if (sysClr) {
          const lastClr = getAttr(sysClr, "lastClr");
          if (lastClr) {
            const hex = normalizeColor(lastClr);
            if (hex) colors[keyMap[cn]] = hex;
          }
        }
      }
    }

    // Extract font scheme
    const fontSchemeNode = themeElements["a:fontScheme"] as
      | Record<string, unknown>
      | undefined;
    let majorFont: string | undefined;
    let minorFont: string | undefined;

    if (fontSchemeNode) {
      const majorFontNode = fontSchemeNode["a:majorFont"] as
        | Record<string, unknown>
        | undefined;
      const minorFontNode = fontSchemeNode["a:minorFont"] as
        | Record<string, unknown>
        | undefined;

      if (majorFontNode) {
        const latin = majorFontNode["a:latin"] as Record<string, unknown> | undefined;
        majorFont = latin ? getAttr(latin, "typeface") : undefined;
      }
      if (minorFontNode) {
        const latin = minorFontNode["a:latin"] as Record<string, unknown> | undefined;
        minorFont = latin ? getAttr(latin, "typeface") : undefined;
      }
    }

    return {
      name: themeName,
      colors: Object.keys(colors).length > 0 ? colors : undefined,
      fontScheme:
        majorFont || minorFont ? { majorFont, minorFont } : undefined,
    };
  } catch {
    return null;
  }
}

/**
 * Find the theme path for a presentation via presentation.xml.rels.
 */
export function getThemePath(
  presentationRels: Record<string, unknown>
): string | null {
  const relsObj = presentationRels["Relationships"] as
    | Record<string, unknown>
    | undefined;
  const relItems = relsObj?.["Relationship"];
  if (!relItems) return null;

  const rels = Array.isArray(relItems) ? relItems : [relItems];
  const themeRel = rels.find((r) => {
    const type = (r as Record<string, unknown>)["@_Type"] as string;
    return type?.endsWith("/theme");
  }) as Record<string, unknown> | undefined;

  if (!themeRel) return null;
  const target = (themeRel["@_Target"] as string).replace(/^\.\.\//, "ppt/");
  return target.startsWith("ppt/") ? target : `ppt/${target}`;
}
