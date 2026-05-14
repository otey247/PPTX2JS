/**
 * Content-types utilities for PPTX packages.
 * Parses [Content_Types].xml to determine part types.
 */

import JSZip from "jszip";
import { XML_PARSER } from "./read-pptx";

export type ContentTypeMap = {
  defaults: Map<string, string>; // extension → contentType
  overrides: Map<string, string>; // partName → contentType
};

/**
 * Load and parse [Content_Types].xml from the zip.
 */
export async function loadContentTypes(zip: JSZip): Promise<ContentTypeMap> {
  const text = await zip.file("[Content_Types].xml")?.async("text");
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  if (!text) return { defaults, overrides };

  const parsed = XML_PARSER.parse(text) as Record<string, unknown>;
  const types = (parsed as Record<string, Record<string, unknown>>)["Types"];
  if (!types) return { defaults, overrides };

  const defaultItems = types["Default"];
  if (defaultItems) {
    const items = Array.isArray(defaultItems) ? defaultItems : [defaultItems];
    for (const item of items as Array<Record<string, string>>) {
      defaults.set(item["@_Extension"], item["@_ContentType"]);
    }
  }

  const overrideItems = types["Override"];
  if (overrideItems) {
    const items = Array.isArray(overrideItems) ? overrideItems : [overrideItems];
    for (const item of items as Array<Record<string, string>>) {
      overrides.set(item["@_PartName"], item["@_ContentType"]);
    }
  }

  return { defaults, overrides };
}
