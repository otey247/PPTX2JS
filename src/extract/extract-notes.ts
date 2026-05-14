/**
 * Extract speaker notes from notes slide XML.
 */

import JSZip from "jszip";
import { XML_PARSER } from "../ingest/read-pptx";

/**
 * Extract plain text notes from a notes slide XML part.
 */
export async function extractNotes(
  zip: JSZip,
  notesPath: string | undefined
): Promise<string | undefined> {
  if (!notesPath) return undefined;

  try {
    const xml = await zip.file(notesPath)?.async("text");
    if (!xml) return undefined;

    const parsed = XML_PARSER.parse(xml) as Record<string, unknown>;
    const notesSl = (parsed as Record<string, unknown>)["p:notes"] as
      | Record<string, unknown>
      | undefined;
    if (!notesSl) return undefined;

    const cSld = notesSl["p:cSld"] as Record<string, unknown> | undefined;
    const spTree = cSld?.["p:spTree"] as Record<string, unknown> | undefined;
    if (!spTree) return undefined;

    const spNodes = spTree["p:sp"];
    if (!spNodes) return undefined;

    const sps = Array.isArray(spNodes) ? spNodes : [spNodes];
    const textParts: string[] = [];

    for (const sp of sps as Array<Record<string, unknown>>) {
      // Skip the slide image placeholder (idx 0, type "sldImg")
      const nvSpPr = sp["p:nvSpPr"] as Record<string, unknown> | undefined;
      const nvPr = nvSpPr?.["p:nvPr"] as Record<string, unknown> | undefined;
      const ph = nvPr?.["p:ph"] as Record<string, unknown> | undefined;
      if (ph && (ph["@_type"] === "sldImg" || ph["@_idx"] === "0")) continue;

      const txBody = sp["p:txBody"] as Record<string, unknown> | undefined;
      if (!txBody) continue;

      const paras = txBody["a:p"];
      if (!paras) continue;

      const paraList = Array.isArray(paras) ? paras : [paras];
      for (const para of paraList as Array<Record<string, unknown>>) {
        const runNodes = para["a:r"];
        if (!runNodes) continue;
        const runs = Array.isArray(runNodes) ? runNodes : [runNodes];
        for (const r of runs as Array<Record<string, unknown>>) {
          const t = r["a:t"];
          if (typeof t === "string" || typeof t === "number") {
            textParts.push(String(t));
          }
        }
        textParts.push("\n");
      }
    }

    const notes = textParts.join("").trim();
    return notes.length > 0 ? notes : undefined;
  } catch {
    return undefined;
  }
}
