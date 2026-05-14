/**
 * Relationship file utilities.
 * Handles loading and querying .rels files from the PPTX zip.
 */

import JSZip from "jszip";
import { XML_PARSER } from "./read-pptx";

/**
 * Derive the .rels path for a given XML part path.
 * e.g. "ppt/slides/slide1.xml" → "ppt/slides/_rels/slide1.xml.rels"
 */
export function getRelsPath(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = partPath.slice(0, lastSlash + 1);
  const file = partPath.slice(lastSlash + 1);
  return `${dir}_rels/${file}.rels`;
}

/**
 * Load and parse a .rels file from the zip. Returns null if absent.
 */
export async function loadRels(
  zip: JSZip,
  partPath: string
): Promise<Record<string, unknown> | null> {
  const relsPath = getRelsPath(partPath);
  const text = await zip.file(relsPath)?.async("text");
  if (!text) return null;
  return XML_PARSER.parse(text) as Record<string, unknown>;
}
