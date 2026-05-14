/**
 * Extract text elements (p:sp with text body) from slide XML.
 */

import { TextElement, TextRunModel, FillModel, LineModel } from "../types";
import { emuToInches, ooxmlSzToPt } from "../normalize/units";
import { normalizeColor, resolveThemeColor, applyLumAdjust } from "../normalize/colors";
import { resolveFontFace } from "../normalize/fonts";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function extractFill(
  spPr: Record<string, unknown>,
  themeColorMap: Record<string, string>
): FillModel | undefined {
  // Solid fill
  const solidFill = spPr["a:solidFill"] as Record<string, unknown> | undefined;
  if (solidFill) {
    const color = extractColor(solidFill, themeColorMap);
    if (color) return { type: "solid", color };
  }
  // No fill
  if (spPr["a:noFill"] !== undefined) return { type: "none" };
  return undefined;
}

function extractColor(
  fillNode: Record<string, unknown>,
  themeColorMap: Record<string, string>
): string | undefined {
  const srgb = fillNode["a:srgbClr"] as Record<string, unknown> | undefined;
  if (srgb) {
    return normalizeColor(getAttr(srgb, "val"));
  }

  const schemeClr = fillNode["a:schemeClr"] as Record<string, unknown> | undefined;
  if (schemeClr) {
    const schemeName = getAttr(schemeClr, "val") ?? "";
    let baseColor = resolveThemeColor(schemeName, themeColorMap);
    if (baseColor) {
      const lumMod = (schemeClr["a:lumMod"] as Record<string, unknown> | undefined);
      const lumOff = (schemeClr["a:lumOff"] as Record<string, unknown> | undefined);
      const lumModVal = lumMod ? getAttr(lumMod, "val") : undefined;
      const lumOffVal = lumOff ? getAttr(lumOff, "val") : undefined;
      baseColor = applyLumAdjust(
        baseColor,
        lumModVal !== undefined ? Number(lumModVal) : undefined,
        lumOffVal !== undefined ? Number(lumOffVal) : undefined
      );
    }
    return baseColor;
  }

  const prstClr = fillNode["a:prstClr"] as Record<string, unknown> | undefined;
  if (prstClr) {
    return normalizeColor(getAttr(prstClr, "val"));
  }

  return undefined;
}

function extractLine(
  spPr: Record<string, unknown>,
  themeColorMap: Record<string, string>
): LineModel | undefined {
  const ln = spPr["a:ln"] as Record<string, unknown> | undefined;
  if (!ln) return undefined;

  const line: LineModel = {};

  const w = getAttr(ln, "w");
  if (w) line.width = Math.round(Number(w) / 12700); // EMU → pt

  const solidFill = ln["a:solidFill"] as Record<string, unknown> | undefined;
  if (solidFill) {
    line.color = extractColor(solidFill, themeColorMap);
  }

  const prstDash = ln["a:prstDash"] as Record<string, unknown> | undefined;
  if (prstDash) {
    const val = getAttr(prstDash, "val") ?? "solid";
    if (val === "dash" || val === "dot" || val === "dashDot" || val === "solid") {
      line.dashType = val;
    }
  }

  return Object.keys(line).length > 0 ? line : undefined;
}

function extractTextRuns(
  txBody: Record<string, unknown>,
  themeColorMap: Record<string, string>,
  hyperlinkUrls: Map<string, string>,
  majorFont?: string,
  minorFont?: string
): TextRunModel[] {
  const runs: TextRunModel[] = [];
  const paragraphs = txBody["a:p"];
  if (!paragraphs) return runs;

  const paras = Array.isArray(paragraphs) ? paragraphs : [paragraphs];

  for (let pIdx = 0; pIdx < paras.length; pIdx++) {
    const para = paras[pIdx] as Record<string, unknown>;

    const pPr = para["a:pPr"] as Record<string, unknown> | undefined;

    const runNodes = para["a:r"];
    const breakNodes = para["a:br"];

    const rNodes = runNodes
      ? Array.isArray(runNodes)
        ? runNodes
        : [runNodes]
      : [];
    const brNodes = breakNodes
      ? Array.isArray(breakNodes)
        ? breakNodes
        : [breakNodes]
      : [];

    for (const r of rNodes as Array<Record<string, unknown>>) {
      const rPr = r["a:rPr"] as Record<string, unknown> | undefined;
      const t = r["a:t"];
      const text = typeof t === "string" ? t : typeof t === "number" ? String(t) : "";

      const run: TextRunModel = { text };

      if (rPr) {
        if (getAttr(rPr, "b") === "1" || getAttr(rPr, "b") === "true")
          run.bold = true;
        if (getAttr(rPr, "i") === "1" || getAttr(rPr, "i") === "true")
          run.italic = true;
        if (getAttr(rPr, "u") && getAttr(rPr, "u") !== "none")
          run.underline = true;
        if (getAttr(rPr, "strike") && getAttr(rPr, "strike") !== "noStrike")
          run.strike = true;

        const sz = getAttr(rPr, "sz");
        if (sz) run.fontSize = ooxmlSzToPt(sz);

        const solidFill = rPr["a:solidFill"] as Record<string, unknown> | undefined;
        if (solidFill) run.color = extractColor(solidFill, themeColorMap);

        const latin = rPr["a:latin"] as Record<string, unknown> | undefined;
        if (latin) {
          run.fontFace = resolveFontFace(
            getAttr(latin, "typeface"),
            majorFont,
            minorFont
          );
        }

        const hlinkClick = rPr["a:hlinkClick"] as Record<string, unknown> | undefined;
        if (hlinkClick) {
          // Resolve the r:id relationship to the actual URL
          const rId = getAttr(hlinkClick, "r:id");
          if (rId) {
            run.hyperlink = hyperlinkUrls.get(rId) ?? rId;
          }
        }
      }

      runs.push(run);
    }

    // Set breakLine on the last run of each paragraph where a:br elements are present
    for (const _br of brNodes) {
      if (runs.length > 0) {
        runs[runs.length - 1].breakLine = true;
      } else {
        runs.push({ text: "", breakLine: true });
      }
    }

    // Add paragraph break between paragraphs (not after the last one)
    if (pIdx < paras.length - 1 && runs.length > 0) {
      runs[runs.length - 1].breakLine = true;
    }
  }

  return runs;
}

function extractTransform(
  spPr: Record<string, unknown>
): { x: number; y: number; w: number; h: number; rotation?: number; flipH?: boolean; flipV?: boolean } {
  const xfrm = spPr["a:xfrm"] as Record<string, unknown> | undefined;
  if (!xfrm) return { x: 0, y: 0, w: 1, h: 1 };

  const off = xfrm["a:off"] as Record<string, unknown> | undefined;
  const ext = xfrm["a:ext"] as Record<string, unknown> | undefined;

  const x = off ? emuToInches(getAttr(off, "x") ?? "0") : 0;
  const y = off ? emuToInches(getAttr(off, "y") ?? "0") : 0;
  const w = ext ? emuToInches(getAttr(ext, "cx") ?? "0") : 1;
  const h = ext ? emuToInches(getAttr(ext, "cy") ?? "0") : 1;

  const rot = getAttr(xfrm, "rot");
  const rotation = rot ? Math.round(Number(rot) / 60000) : undefined; // 60000ths of a degree

  const flipH = getAttr(xfrm, "flipH") === "1" || getAttr(xfrm, "flipH") === "true";
  const flipV = getAttr(xfrm, "flipV") === "1" || getAttr(xfrm, "flipV") === "true";

  return { x, y, w, h, rotation, flipH: flipH || undefined, flipV: flipV || undefined };
}

/**
 * Extract text element from a p:sp node.
 */
export function extractTextElement(
  sp: Record<string, unknown>,
  themeColorMap: Record<string, string>,
  hyperlinkUrls: Map<string, string>,
  majorFont?: string,
  minorFont?: string
): TextElement | null {
  const spPr = sp["p:spPr"] as Record<string, unknown> | undefined;
  const txBody = sp["p:txBody"] as Record<string, unknown> | undefined;

  if (!txBody) return null;
  if (!spPr) return null;

  const transform = extractTransform(spPr);
  const textRuns = extractTextRuns(txBody, themeColorMap, hyperlinkUrls, majorFont, minorFont);
  const fill = extractFill(spPr, themeColorMap);
  const line = extractLine(spPr, themeColorMap);

  // Body properties
  const bodyPr = txBody["a:bodyPr"] as Record<string, unknown> | undefined;
  let margin: number | undefined;
  let valign: "top" | "mid" | "bottom" | undefined;

  if (bodyPr) {
    const inset = getAttr(bodyPr, "insTl");
    if (inset) margin = emuToInches(inset);
    const anchor = getAttr(bodyPr, "anchor");
    if (anchor === "t") valign = "top";
    else if (anchor === "ctr") valign = "mid";
    else if (anchor === "b") valign = "bottom";
  }

  // Placeholder detection
  const nvSpPr = sp["p:nvSpPr"] as Record<string, unknown> | undefined;
  const nvPr = nvSpPr?.["p:nvPr"] as Record<string, unknown> | undefined;
  const ph = nvPr?.["p:ph"] as Record<string, unknown> | undefined;
  const isPlaceholder = !!ph;
  const placeholderType = ph ? getAttr(ph, "type") : undefined;

  // First paragraph alignment
  const firstPara = Array.isArray(txBody["a:p"])
    ? (txBody["a:p"] as Array<Record<string, unknown>>)[0]
    : (txBody["a:p"] as Record<string, unknown> | undefined);

  let align: "left" | "center" | "right" | "justify" | undefined;
  if (firstPara) {
    const pPr = firstPara["a:pPr"] as Record<string, unknown> | undefined;
    const algn = pPr ? getAttr(pPr, "algn") : undefined;
    if (algn === "ctr") align = "center";
    else if (algn === "r") align = "right";
    else if (algn === "just") align = "justify";
    else if (algn === "l") align = "left";
  }

  return {
    type: "text",
    ...transform,
    textRuns,
    fill,
    line,
    margin,
    valign,
    align,
    isPlaceholder,
    placeholderType,
  };
}
