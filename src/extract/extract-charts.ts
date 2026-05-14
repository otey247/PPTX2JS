/**
 * Extract chart elements from p:graphicFrame nodes that reference chart parts.
 */

import JSZip from "jszip";
import { ChartElement, ChartDataSeries } from "../types";
import { normalizeTarget } from "../ingest/resolve-slides";
import { emuToInches } from "../normalize/units";
import { XML_PARSER } from "../ingest/read-pptx";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

/**
 * Map OOXML chart type element names to PptxGenJS chart type strings.
 */
const CHART_TYPE_MAP: Record<string, string> = {
  "c:barChart": "bar",
  "c:lineChart": "line",
  "c:pieChart": "pie",
  "c:areaChart": "area",
  "c:doughnutChart": "doughnut",
  "c:scatterChart": "scatter",
  "c:bubbleChart": "bubble",
  "c:radarChart": "radar",
};

function extractNumericValues(valNode: Record<string, unknown>): number[] {
  const numRef = valNode["c:numRef"] as Record<string, unknown> | undefined;
  const numLit = valNode["c:numLit"] as Record<string, unknown> | undefined;
  const src = numRef?.["c:numCache"] ?? numLit;
  if (!src) return [];

  const ptNodes = (src as Record<string, unknown>)["c:pt"];
  if (!ptNodes) return [];
  const pts = Array.isArray(ptNodes) ? ptNodes : [ptNodes];
  return pts.map((pt) => Number((pt as Record<string, unknown>)["c:v"] ?? 0));
}

function extractStringValues(catNode: Record<string, unknown>): string[] {
  const strRef = catNode["c:strRef"] as Record<string, unknown> | undefined;
  const strLit = catNode["c:strLit"] as Record<string, unknown> | undefined;
  const src = strRef?.["c:strCache"] ?? strLit;
  if (!src) return [];

  const ptNodes = (src as Record<string, unknown>)["c:pt"];
  if (!ptNodes) return [];
  const pts = Array.isArray(ptNodes) ? ptNodes : [ptNodes];
  return pts.map((pt) => String((pt as Record<string, unknown>)["c:v"] ?? ""));
}

function extractSeriesData(chartData: Record<string, unknown>): ChartDataSeries[] {
  for (const [key, value] of Object.entries(CHART_TYPE_MAP)) {
    const chartTypeNode = chartData[key] as
      | Record<string, unknown>
      | undefined;
    if (!chartTypeNode) continue;

    const serNodes = chartTypeNode["c:ser"];
    if (!serNodes) return [];
    const serList = Array.isArray(serNodes) ? serNodes : [serNodes];

    return serList.map((ser) => {
      const s = ser as Record<string, unknown>;
      const txNode = s["c:tx"] as Record<string, unknown> | undefined;
      const strRef = txNode?.["c:strRef"] as Record<string, unknown> | undefined;
      const nameCache = strRef?.["c:strCache"] as Record<string, unknown> | undefined;
      const namePts = nameCache?.["c:pt"];
      const name = namePts
        ? String(
            (
              (Array.isArray(namePts) ? namePts[0] : namePts) as Record<
                string,
                unknown
              >
            )["c:v"] ?? "Series"
          )
        : "Series";

      const catNode = s["c:cat"] as Record<string, unknown> | undefined;
      const labels = catNode ? extractStringValues(catNode) : undefined;

      const valNode = s["c:val"] as Record<string, unknown> | undefined;
      const values = valNode ? extractNumericValues(valNode) : [];

      return { name, labels, values };
    });
  }

  return [];
}

/**
 * Resolve chart XML path from a graphicFrame's relationship.
 */
export function getChartRelPath(
  graphicFrame: Record<string, unknown>,
  slideRels: Record<string, unknown> | null
): string | null {
  if (!slideRels) return null;

  // Find chart relationship ID in the graphicFrame
  const graphic = graphicFrame["a:graphic"] as Record<string, unknown> | undefined;
  const graphicData = graphic?.["a:graphicData"] as Record<string, unknown> | undefined;
  const chart = graphicData?.["c:chart"] as Record<string, unknown> | undefined;
  if (!chart) return null;

  const rId = getAttr(chart, "r:id");
  if (!rId) return null;

  const relsObj = slideRels["Relationships"] as Record<string, unknown> | undefined;
  const relItems = relsObj?.["Relationship"];
  if (!relItems) return null;

  const rels = Array.isArray(relItems) ? relItems : [relItems];
  const rel = rels.find(
    (r) => (r as Record<string, unknown>)["@_Id"] === rId
  ) as Record<string, unknown> | undefined;
  if (!rel) return null;

  return normalizeTarget(rel["@_Target"] as string);
}

/**
 * Extract a chart element from a p:graphicFrame that references a chart XML part.
 */
export async function extractChartElement(
  graphicFrame: Record<string, unknown>,
  zip: JSZip,
  slideRels: Record<string, unknown> | null
): Promise<ChartElement | null> {
  const xfrm = (graphicFrame["p:xfrm"] ?? graphicFrame["a:xfrm"]) as
    | Record<string, unknown>
    | undefined;
  const off = xfrm?.["a:off"] as Record<string, unknown> | undefined;
  const ext = xfrm?.["a:ext"] as Record<string, unknown> | undefined;

  const x = emuToInches(getAttr(off ?? {}, "x") ?? "0");
  const y = emuToInches(getAttr(off ?? {}, "y") ?? "0");
  const w = emuToInches(getAttr(ext ?? {}, "cx") ?? "0");
  const h = emuToInches(getAttr(ext ?? {}, "cy") ?? "0");

  const chartPath = getChartRelPath(graphicFrame, slideRels);
  if (!chartPath) return null;

  try {
    const chartXml = await zip.file(chartPath)?.async("text");
    if (!chartXml) return null;

    const chartObj = XML_PARSER.parse(chartXml) as Record<string, unknown>;
    const chartSpace = chartObj["c:chartSpace"] as Record<string, unknown> | undefined;
    if (!chartSpace) return null;

    const chartNode = chartSpace["c:chart"] as Record<string, unknown> | undefined;
    if (!chartNode) return null;

    const plotArea = chartNode["c:plotArea"] as Record<string, unknown> | undefined;
    if (!plotArea) return null;

    // Determine chart type
    let chartType = "bar";
    for (const key of Object.keys(CHART_TYPE_MAP)) {
      if (plotArea[key] !== undefined) {
        chartType = CHART_TYPE_MAP[key];
        break;
      }
    }

    const data = extractSeriesData(plotArea);

    return { type: "chart", x, y, w, h, chartType, data };
  } catch {
    return null;
  }
}
