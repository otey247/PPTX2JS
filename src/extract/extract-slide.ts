/**
 * Top-level slide extractor.
 * Coordinates extraction of all elements from a single slide XML.
 */

import JSZip from "jszip";
import {
  SlideModel,
  SlideElement,
  FillModel,
  type UnsupportedElement,
} from "../types";
import { normalizeColor, resolveThemeColor } from "../normalize/colors";
import { extractTextElement } from "./extract-text";
import { extractShapeElement, extractConnectorElement } from "./extract-shapes";
import { extractImageElement } from "./extract-images";
import { extractTableElement } from "./extract-tables";
import { extractChartElement } from "./extract-charts";
import { extractNotes } from "./extract-notes";
import { loadRels } from "../ingest/relationships";
import { resolveSlideRelationships } from "../ingest/resolve-slides";
import { XML_PARSER } from "../ingest/read-pptx";

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
      const color = resolveThemeColor(
        getAttr(schemeClr, "val") ?? "",
        themeColorMap
      );
      if (color) return { type: "solid", color };
    }
  }

  return undefined;
}

async function extractSpTree(
  spTree: Record<string, unknown>,
  zip: JSZip,
  slideRels: Record<string, unknown> | null,
  imagePaths: Map<string, string>,
  hyperlinkUrls: Map<string, string>,
  themeColorMap: Record<string, string>,
  majorFont?: string,
  minorFont?: string
): Promise<SlideElement[]> {
  const elements: SlideElement[] = [];

  // p:sp — text boxes and shapes
  const spNodes = spTree["p:sp"];
  if (spNodes) {
    const sps = Array.isArray(spNodes) ? spNodes : [spNodes];
    for (const sp of sps as Array<Record<string, unknown>>) {
      // Prefer text extraction when the shape is a plain rect (the default text box
      // preset) AND has a text body — those elements have richer text properties
      // (valign, margin, align, placeholder) that would be lost in the shape branch.
      const spPr = sp["p:spPr"] as Record<string, unknown> | undefined;
      const prstGeom = spPr?.["a:prstGeom"] as Record<string, unknown> | undefined;
      const prst = prstGeom ? (prstGeom["@_prst"] as string | undefined) : undefined;
      const hasTxBody = !!sp["p:txBody"];
      const isRectOrNoGeom = !prstGeom || prst === "rect";

      if (hasTxBody && isRectOrNoGeom) {
        const text = extractTextElement(sp, themeColorMap, hyperlinkUrls, majorFont, minorFont);
        if (text) {
          elements.push(text);
          continue;
        }
      }

      // Try shape extraction (non-rect geometry or no text body)
      const shape = extractShapeElement(sp, themeColorMap, majorFont, minorFont);
      if (shape) {
        elements.push(shape);
        continue;
      }

      // Fallback: try text extraction for any remaining p:sp with a text body
      if (hasTxBody) {
        const text = extractTextElement(sp, themeColorMap, hyperlinkUrls, majorFont, minorFont);
        if (text) elements.push(text);
      }
    }
  }

  // p:pic — images
  const picNodes = spTree["p:pic"];
  if (picNodes) {
    const pics = Array.isArray(picNodes) ? picNodes : [picNodes];
    for (const pic of pics as Array<Record<string, unknown>>) {
      const img = await extractImageElement(pic, zip, imagePaths);
      if (img) elements.push(img);
    }
  }

  // p:graphicFrame — tables or charts
  const gfNodes = spTree["p:graphicFrame"];
  if (gfNodes) {
    const gfs = Array.isArray(gfNodes) ? gfNodes : [gfNodes];
    for (const gf of gfs as Array<Record<string, unknown>>) {
      // Check if it's a chart
      const graphic = gf["a:graphic"] as Record<string, unknown> | undefined;
      const graphicData = graphic?.["a:graphicData"] as
        | Record<string, unknown>
        | undefined;
      const uri = graphicData
        ? getAttr(graphicData, "uri") ?? ""
        : "";

      if (uri.includes("chart")) {
        const chart = await extractChartElement(gf, zip, slideRels);
        if (chart) {
          elements.push(chart);
        } else {
          elements.push({
            type: "unsupported",
            reason: "Chart extraction failed or unsupported chart type",
            rawXmlTag: "p:graphicFrame[chart]",
          } as UnsupportedElement);
        }
      } else if (uri.includes("table")) {
        const table = extractTableElement(gf, themeColorMap);
        if (table) {
          elements.push(table);
        } else {
          elements.push({
            type: "unsupported",
            reason: "Table extraction failed",
            rawXmlTag: "p:graphicFrame[table]",
          } as UnsupportedElement);
        }
      } else {
        // SmartArt, diagrams, etc.
        elements.push({
          type: "unsupported",
          reason: `Unsupported graphic type: ${uri}`,
          rawXmlTag: "p:graphicFrame",
        } as UnsupportedElement);
      }
    }
  }

  // p:cxnSp — connectors
  const cxnNodes = spTree["p:cxnSp"];
  if (cxnNodes) {
    const cxns = Array.isArray(cxnNodes) ? cxnNodes : [cxnNodes];
    for (const cxn of cxns as Array<Record<string, unknown>>) {
      const connector = extractConnectorElement(cxn, themeColorMap);
      if (connector) elements.push(connector);
    }
  }

  // p:grpSp — grouped shapes (flatten one level)
  const grpNodes = spTree["p:grpSp"];
  if (grpNodes) {
    const grps = Array.isArray(grpNodes) ? grpNodes : [grpNodes];
    for (const grp of grps as Array<Record<string, unknown>>) {
      // Recursively extract children of the group
      const grpSpTree = grp as Record<string, unknown>;
      const children = await extractSpTree(
        grpSpTree,
        zip,
        slideRels,
        imagePaths,
        hyperlinkUrls,
        themeColorMap,
        majorFont,
        minorFont
      );
      elements.push(...children);
    }
  }

  return elements;
}

/**
 * Extract a full slide model from a slide XML path.
 */
export async function extractSlide(
  zip: JSZip,
  slidePath: string,
  slideWidth: number,
  slideHeight: number,
  themeColorMap: Record<string, string>,
  majorFont?: string,
  minorFont?: string
): Promise<SlideModel> {
  const slideId = slidePath;

  const xml = await zip.file(slidePath)?.async("text");
  if (!xml) {
    return {
      id: slideId,
      width: slideWidth,
      height: slideHeight,
      elements: [],
    };
  }

  const parsed = XML_PARSER.parse(xml) as Record<string, unknown>;
  const sld = parsed["p:sld"] as Record<string, unknown> | undefined;
  if (!sld) {
    return {
      id: slideId,
      width: slideWidth,
      height: slideHeight,
      elements: [],
    };
  }

  // Load slide relationships
  const slideRels = await loadRels(zip, slidePath);
  const { notesPath, imagePaths, hyperlinkUrls } = resolveSlideRelationships(slideRels);

  const cSld = sld["p:cSld"] as Record<string, unknown> | undefined;
  const spTree = cSld?.["p:spTree"] as Record<string, unknown> | undefined;

  const background = cSld ? extractBackground(cSld, themeColorMap) : undefined;

  const elements = spTree
    ? await extractSpTree(
        spTree,
        zip,
        slideRels,
        imagePaths,
        hyperlinkUrls,
        themeColorMap,
        majorFont,
        minorFont
      )
    : [];

  const notes = await extractNotes(zip, notesPath);

  return {
    id: slideId,
    width: slideWidth,
    height: slideHeight,
    background,
    elements,
    notes,
  };
}
