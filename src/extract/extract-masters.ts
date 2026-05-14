/**
 * Extract slide master data from slide master XML.
 */

import JSZip from "jszip";
import { MasterModel, FillModel } from "../types";
import { XML_PARSER } from "../ingest/read-pptx";
import { normalizeTarget } from "../ingest/resolve-slides";
import { normalizeColor, resolveThemeColor } from "../normalize/colors";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function extractBackground(
  cSld: Record<string, unknown>,
  themeColorMap: Record<string, string>
): FillModel | undefined {
  const bg = cSld["p:bg"] as Record<string, unknown> | undefined;
  if (!bg) return undefined;

  const bgPr = bg["p:bgPr"] as Record<string, unknown> | undefined;
  if (!bgPr) return undefined;

  const solidFill = bgPr["a:solidFill"] as Record<string, unknown> | undefined;
  if (solidFill) {
    const srgb = solidFill["a:srgbClr"] as Record<string, unknown> | undefined;
    if (srgb) {
      const hex = normalizeColor(getAttr(srgb, "val"));
      if (hex) return { type: "solid", color: hex };
    }

    const schemeClr = solidFill["a:schemeClr"] as Record<string, unknown> | undefined;
    if (schemeClr) {
      const color = resolveThemeColor(getAttr(schemeClr, "val") ?? "", themeColorMap);
      if (color) return { type: "solid", color };
    }
  }

  return undefined;
}

/**
 * Extract a slide master from its XML path.
 */
export async function extractMaster(
  zip: JSZip,
  masterPath: string,
  themeColorMap: Record<string, string>
): Promise<MasterModel | null> {
  try {
    const xml = await zip.file(masterPath)?.async("text");
    if (!xml) return null;

    const parsed = XML_PARSER.parse(xml) as Record<string, unknown>;
    const sldMaster = parsed["p:sldMaster"] as Record<string, unknown> | undefined;
    if (!sldMaster) return null;

    const cSld = sldMaster["p:cSld"] as Record<string, unknown> | undefined;
    const background = cSld
      ? extractBackground(cSld, themeColorMap)
      : undefined;

    return {
      id: masterPath,
      background,
      elements: [], // Master element extraction is complex; provide empty for Phase 1
    };
  } catch {
    return null;
  }
}

/**
 * Find all slide master paths from presentation.xml.rels.
 */
export function getMasterPaths(
  presentationRels: Record<string, unknown>
): string[] {
  const relsObj = presentationRels["Relationships"] as
    | Record<string, unknown>
    | undefined;
  const relItems = relsObj?.["Relationship"];
  if (!relItems) return [];

  const rels = Array.isArray(relItems) ? relItems : [relItems];
  return rels
    .filter((r) => {
      const type = (r as Record<string, unknown>)["@_Type"] as string;
      return type?.endsWith("/slideMaster");
    })
    .map((r) => normalizeTarget((r as Record<string, unknown>)["@_Target"] as string));
}
