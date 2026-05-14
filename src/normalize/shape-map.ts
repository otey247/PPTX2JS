/**
 * Maps OOXML preset shape names to PptxGenJS shape type strings.
 * PptxGenJS uses the same OOXML preset names but accessed via pptx.ShapeType.
 */

/**
 * Normalize an OOXML preset shape name.
 * PptxGenJS shape names match OOXML preset names from the `a:prstGeom` prst attribute.
 */
export function normalizeShapeType(ooxmlPreset: string | undefined): string {
  if (!ooxmlPreset) return "rect";
  return ooxmlPreset;
}

/**
 * Common OOXML preset shape names that map directly to PptxGenJS shapes.
 */
export const SUPPORTED_SHAPES = new Set([
  "rect",
  "roundRect",
  "ellipse",
  "triangle",
  "rtTriangle",
  "parallelogram",
  "trapezoid",
  "diamond",
  "pentagon",
  "hexagon",
  "heptagon",
  "octagon",
  "decagon",
  "dodecagon",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "ribbon",
  "ribbon2",
  "chevron",
  "pentagon",
  "notchedRightArrow",
  "homePlate",
  "leftArrow",
  "rightArrow",
  "upArrow",
  "downArrow",
  "leftRightArrow",
  "upDownArrow",
  "quadArrow",
  "callout1",
  "callout2",
  "callout3",
  "line",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
]);
