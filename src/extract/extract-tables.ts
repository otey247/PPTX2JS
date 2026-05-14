/**
 * Extract table elements (p:graphicFrame containing a:tbl) from slide XML.
 */

import { TableElement, TableCellModel } from "../types";
import { emuToInches, halfPtToPt } from "../normalize/units";
import { normalizeColor, resolveThemeColor } from "../normalize/colors";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function extractCellColor(
  fillNode: Record<string, unknown>,
  themeColorMap: Record<string, string>
): string | undefined {
  const solidFill = fillNode["a:solidFill"] as Record<string, unknown> | undefined;
  if (!solidFill) return undefined;

  const srgb = solidFill["a:srgbClr"] as Record<string, unknown> | undefined;
  if (srgb) return normalizeColor(getAttr(srgb, "val"));

  const schemeClr = solidFill["a:schemeClr"] as Record<string, unknown> | undefined;
  if (schemeClr) return resolveThemeColor(getAttr(schemeClr, "val") ?? "", themeColorMap);

  return undefined;
}

function extractCell(
  tc: Record<string, unknown>,
  themeColorMap: Record<string, string>
): TableCellModel {
  const txBody = tc["a:txBody"] as Record<string, unknown> | undefined;
  let text = "";
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let fontSize: number | undefined;
  let color: string | undefined;
  let align: "left" | "center" | "right" | undefined;

  if (txBody) {
    const paras = txBody["a:p"];
    const paraList = paras ? (Array.isArray(paras) ? paras : [paras]) : [];
    const texts: string[] = [];

    for (const para of paraList as Array<Record<string, unknown>>) {
      const pPr = para["a:pPr"] as Record<string, unknown> | undefined;
      if (pPr) {
        const algn = getAttr(pPr, "algn");
        if (algn === "ctr") align = "center";
        else if (algn === "r") align = "right";
        else if (algn === "l") align = "left";
      }

      const runs = para["a:r"];
      if (!runs) continue;
      const runList = Array.isArray(runs) ? runs : [runs];

      for (const r of runList as Array<Record<string, unknown>>) {
        const rPr = r["a:rPr"] as Record<string, unknown> | undefined;
        const t = r["a:t"];
        texts.push(typeof t === "string" ? t : typeof t === "number" ? String(t) : "");

        if (rPr && bold === undefined) {
          if (getAttr(rPr, "b") === "1" || getAttr(rPr, "b") === "true") bold = true;
          if (getAttr(rPr, "i") === "1" || getAttr(rPr, "i") === "true") italic = true;
          const sz = getAttr(rPr, "sz");
          if (sz && fontSize === undefined) fontSize = halfPtToPt(sz);

          const solidFill = rPr["a:solidFill"] as Record<string, unknown> | undefined;
          if (solidFill && color === undefined) {
            const srgb = solidFill["a:srgbClr"] as Record<string, unknown> | undefined;
            if (srgb) color = normalizeColor(getAttr(srgb, "val"));
          }
        }
      }
      texts.push("\n");
    }

    text = texts.join("").replace(/\n$/, "");
  }

  const tcPr = tc["a:tcPr"] as Record<string, unknown> | undefined;
  const fill = tcPr ? extractCellColor(tcPr, themeColorMap) : undefined;
  const colspan = tcPr ? Number(getAttr(tcPr, "gridSpan") ?? "1") : 1;
  const rowspan = tcPr ? Number(getAttr(tcPr, "rowSpan") ?? "1") : 1;

  return {
    text,
    bold,
    italic,
    fontSize,
    color,
    fill,
    align,
    colspan: colspan > 1 ? colspan : undefined,
    rowspan: rowspan > 1 ? rowspan : undefined,
  };
}

/**
 * Extract a table element from a p:graphicFrame that wraps an a:tbl.
 */
export function extractTableElement(
  graphicFrame: Record<string, unknown>,
  themeColorMap: Record<string, string>
): TableElement | null {
  const xfrm = (graphicFrame["p:xfrm"] ?? graphicFrame["a:xfrm"]) as
    | Record<string, unknown>
    | undefined;

  const off = xfrm?.["a:off"] as Record<string, unknown> | undefined;
  const ext = xfrm?.["a:ext"] as Record<string, unknown> | undefined;

  const x = emuToInches(getAttr(off ?? {}, "x") ?? "0");
  const y = emuToInches(getAttr(off ?? {}, "y") ?? "0");
  const w = emuToInches(getAttr(ext ?? {}, "cx") ?? "0");
  const h = emuToInches(getAttr(ext ?? {}, "cy") ?? "0");

  // Navigate to a:tbl inside the graphic
  const graphic = graphicFrame["a:graphic"] as Record<string, unknown> | undefined;
  const graphicData = graphic?.["a:graphicData"] as Record<string, unknown> | undefined;
  const tbl = graphicData?.["a:tbl"] as Record<string, unknown> | undefined;
  if (!tbl) return null;

  const tblGrid = tbl["a:tblGrid"] as Record<string, unknown> | undefined;
  const gridCols = tblGrid?.["a:gridCol"];
  const colWidths: number[] = [];
  if (gridCols) {
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols];
    for (const col of cols as Array<Record<string, unknown>>) {
      colWidths.push(emuToInches(getAttr(col, "w") ?? "0"));
    }
  }

  const trNodes = tbl["a:tr"];
  if (!trNodes) return null;
  const trList = Array.isArray(trNodes) ? trNodes : [trNodes];

  const rows: TableCellModel[][] = [];
  const rowHeights: number[] = [];

  for (const tr of trList as Array<Record<string, unknown>>) {
    const rowH = getAttr(tr, "h");
    if (rowH) rowHeights.push(emuToInches(rowH));

    const tcNodes = tr["a:tc"];
    if (!tcNodes) {
      rows.push([]);
      continue;
    }
    const tcList = Array.isArray(tcNodes) ? tcNodes : [tcNodes];
    const row = tcList.map((tc) => extractCell(tc as Record<string, unknown>, themeColorMap));
    rows.push(row);
  }

  return {
    type: "table",
    x,
    y,
    w,
    h,
    rows,
    colWidths: colWidths.length > 0 ? colWidths : undefined,
    rowHeights: rowHeights.length > 0 ? rowHeights : undefined,
  };
}
