/**
 * Extract shape elements (p:sp without text, p:cxnSp) from slide XML.
 */

import { ShapeElement, TextRunModel, FillModel, LineModel } from "../types";
import { emuToInches, halfPtToPt } from "../normalize/units";
import { normalizeColor, resolveThemeColor, applyLumAdjust } from "../normalize/colors";
import { resolveFontFace } from "../normalize/fonts";
import { normalizeShapeType } from "../normalize/shape-map";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function extractColor(
  fillNode: Record<string, unknown>,
  themeColorMap: Record<string, string>
): string | undefined {
  const srgb = fillNode["a:srgbClr"] as Record<string, unknown> | undefined;
  if (srgb) return normalizeColor(getAttr(srgb, "val"));

  const schemeClr = fillNode["a:schemeClr"] as Record<string, unknown> | undefined;
  if (schemeClr) {
    const schemeName = getAttr(schemeClr, "val") ?? "";
    let baseColor = resolveThemeColor(schemeName, themeColorMap);
    if (baseColor) {
      const lumMod = schemeClr["a:lumMod"] as Record<string, unknown> | undefined;
      const lumOff = schemeClr["a:lumOff"] as Record<string, unknown> | undefined;
      baseColor = applyLumAdjust(
        baseColor,
        lumMod ? Number(getAttr(lumMod, "val")) : undefined,
        lumOff ? Number(getAttr(lumOff, "val")) : undefined
      );
    }
    return baseColor;
  }

  return undefined;
}

function extractFill(
  spPr: Record<string, unknown>,
  themeColorMap: Record<string, string>
): FillModel | undefined {
  const solidFill = spPr["a:solidFill"] as Record<string, unknown> | undefined;
  if (solidFill) {
    const color = extractColor(solidFill, themeColorMap);
    if (color) return { type: "solid", color };
  }
  if (spPr["a:noFill"] !== undefined) return { type: "none" };
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
  if (w) line.width = Math.round(Number(w) / 12700);

  const solidFill = ln["a:solidFill"] as Record<string, unknown> | undefined;
  if (solidFill) line.color = extractColor(solidFill, themeColorMap);

  const prstDash = ln["a:prstDash"] as Record<string, unknown> | undefined;
  if (prstDash) {
    const val = getAttr(prstDash, "val") ?? "solid";
    if (val === "dash" || val === "dot" || val === "dashDot" || val === "solid") {
      line.dashType = val;
    }
  }

  return Object.keys(line).length > 0 ? line : undefined;
}

function extractTransform(spPr: Record<string, unknown>): {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
} {
  const xfrm = spPr["a:xfrm"] as Record<string, unknown> | undefined;
  if (!xfrm) return { x: 0, y: 0, w: 1, h: 1 };

  const off = xfrm["a:off"] as Record<string, unknown> | undefined;
  const ext = xfrm["a:ext"] as Record<string, unknown> | undefined;

  const x = emuToInches(getAttr(off ?? {}, "x") ?? "0");
  const y = emuToInches(getAttr(off ?? {}, "y") ?? "0");
  const w = emuToInches(getAttr(ext ?? {}, "cx") ?? "0");
  const h = emuToInches(getAttr(ext ?? {}, "cy") ?? "0");

  const rot = getAttr(xfrm, "rot");
  const rotation = rot ? Math.round(Number(rot) / 60000) : undefined;
  const flipH = getAttr(xfrm, "flipH") === "1" || getAttr(xfrm, "flipH") === "true";
  const flipV = getAttr(xfrm, "flipV") === "1" || getAttr(xfrm, "flipV") === "true";

  return { x, y, w, h, rotation, flipH: flipH || undefined, flipV: flipV || undefined };
}

function extractShapeTextRuns(
  txBody: Record<string, unknown>,
  themeColorMap: Record<string, string>,
  majorFont?: string,
  minorFont?: string
): TextRunModel[] {
  const runs: TextRunModel[] = [];
  const paras = txBody["a:p"];
  if (!paras) return runs;
  const paraList = Array.isArray(paras) ? paras : [paras];

  for (let pIdx = 0; pIdx < paraList.length; pIdx++) {
    const para = paraList[pIdx] as Record<string, unknown>;
    const runNodes = para["a:r"];
    if (!runNodes) continue;

    const rs = Array.isArray(runNodes) ? runNodes : [runNodes];
    for (const r of rs as Array<Record<string, unknown>>) {
      const rPr = r["a:rPr"] as Record<string, unknown> | undefined;
      const t = r["a:t"];
      const text = typeof t === "string" ? t : typeof t === "number" ? String(t) : "";
      const run: TextRunModel = { text };

      if (rPr) {
        if (getAttr(rPr, "b") === "1" || getAttr(rPr, "b") === "true") run.bold = true;
        if (getAttr(rPr, "i") === "1" || getAttr(rPr, "i") === "true") run.italic = true;

        const sz = getAttr(rPr, "sz");
        if (sz) run.fontSize = halfPtToPt(sz);

        const solidFill = rPr["a:solidFill"] as Record<string, unknown> | undefined;
        if (solidFill) run.color = extractColor(solidFill, themeColorMap);

        const latin = rPr["a:latin"] as Record<string, unknown> | undefined;
        if (latin)
          run.fontFace = resolveFontFace(
            getAttr(latin, "typeface"),
            majorFont,
            minorFont
          );
      }

      runs.push(run);
    }

    if (pIdx < paraList.length - 1 && runs.length > 0) {
      runs[runs.length - 1].breakLine = true;
    }
  }

  return runs;
}

/**
 * Extract a shape element from a p:sp node (when it has a prstGeom, not just text).
 */
export function extractShapeElement(
  sp: Record<string, unknown>,
  themeColorMap: Record<string, string>,
  majorFont?: string,
  minorFont?: string
): ShapeElement | null {
  const spPr = sp["p:spPr"] as Record<string, unknown> | undefined;
  if (!spPr) return null;

  const prstGeom = spPr["a:prstGeom"] as Record<string, unknown> | undefined;
  const custGeom = spPr["a:custGeom"] as Record<string, unknown> | undefined;
  if (!prstGeom && !custGeom) return null;

  const shapeType = normalizeShapeType(
    prstGeom ? getAttr(prstGeom, "prst") : "custGeom"
  );

  const transform = extractTransform(spPr);
  const fill = extractFill(spPr, themeColorMap);
  const line = extractLine(spPr, themeColorMap);

  const txBody = sp["p:txBody"] as Record<string, unknown> | undefined;
  const textRuns = txBody
    ? extractShapeTextRuns(txBody, themeColorMap, majorFont, minorFont)
    : undefined;

  return {
    type: "shape",
    ...transform,
    shapeType,
    fill,
    line,
    textRuns: textRuns && textRuns.length > 0 ? textRuns : undefined,
  };
}

/**
 * Extract a connector shape (p:cxnSp) as a shape element.
 */
export function extractConnectorElement(
  cxnSp: Record<string, unknown>,
  themeColorMap: Record<string, string>
): ShapeElement | null {
  const spPr = cxnSp["p:spPr"] as Record<string, unknown> | undefined;
  if (!spPr) return null;

  const transform = extractTransform(spPr);
  const line = extractLine(spPr, themeColorMap);

  return {
    type: "shape",
    ...transform,
    shapeType: "line",
    line,
  };
}
