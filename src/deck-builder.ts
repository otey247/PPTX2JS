/**
 * Build the normalized DeckModel from a parsed PPTX package.
 */

import JSZip from "jszip";
import { DeckModel, ThemeModel } from "./types";
import { PptxPackage } from "./ingest/read-pptx";
import { resolveSlidePaths } from "./ingest/resolve-slides";
import { emuToInches } from "./normalize/units";
import { extractSlide } from "./extract/extract-slide";
import { extractTheme, getThemePath } from "./extract/extract-theme";
import { extractMaster, getMasterPaths } from "./extract/extract-masters";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function getSlideSize(
  presentation: Record<string, unknown>
): { width: number; height: number } {
  const pres = (presentation as Record<string, Record<string, unknown>>)[
    "p:presentation"
  ];
  if (!pres) return { width: 10, height: 7.5 };

  const sldSz = pres["p:sldSz"] as Record<string, unknown> | undefined;
  if (!sldSz) return { width: 10, height: 7.5 };

  const cx = getAttr(sldSz, "cx");
  const cy = getAttr(sldSz, "cy");

  return {
    width: cx ? emuToInches(cx) : 10,
    height: cy ? emuToInches(cy) : 7.5,
  };
}

/**
 * Build the full DeckModel from the PPTX package.
 */
export async function extractDeckModel(pkg: PptxPackage): Promise<DeckModel> {
  const { zip, presentation, presentationRels } = pkg;

  // Slide size
  const { width, height } = getSlideSize(presentation);

  // Theme
  let theme: ThemeModel | undefined;
  const themePath = getThemePath(presentationRels);
  if (themePath) {
    const extractedTheme = await extractTheme(zip, themePath);
    theme = extractedTheme ?? undefined;
  }

  const themeColorMap: Record<string, string> = theme?.colors ?? {};
  const majorFont = theme?.fontScheme?.majorFont;
  const minorFont = theme?.fontScheme?.minorFont;

  // Slide masters
  const masterPaths = getMasterPaths(presentationRels);
  const masters = await Promise.all(
    masterPaths.map((p) => extractMaster(zip, p, themeColorMap))
  );

  // Slides
  const slidePaths = resolveSlidePaths(presentation, presentationRels);
  const slides = await Promise.all(
    slidePaths.map((p) =>
      extractSlide(zip, p, width, height, themeColorMap, majorFont, minorFont)
    )
  );

  return {
    layout: { width, height, name: "CUSTOM" },
    theme,
    masters: masters.filter((m) => m !== null) as DeckModel["masters"],
    layouts: [],
    slides,
  };
}
