/**
 * Normalized deck model types.
 * This is the intermediate representation between raw OOXML and PptxGenJS output.
 */

export type FillModel =
  | { type: "solid"; color: string }
  | { type: "gradient"; stops: { color: string; position: number }[] }
  | { type: "none" };

export type LineModel = {
  color?: string;
  width?: number;
  dashType?: "solid" | "dash" | "dot" | "dashDot";
};

export type ThemeModel = {
  name?: string;
  colors?: Record<string, string>;
  fontScheme?: {
    majorFont?: string;
    minorFont?: string;
  };
};

export type MasterModel = {
  id: string;
  background?: FillModel;
  elements: SlideElement[];
};

export type LayoutModel = {
  id: string;
  name?: string;
  width: number;
  height: number;
  elements: SlideElement[];
};

export type TextRunModel = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFace?: string;
  fontSize?: number;
  color?: string;
  highlight?: string;
  hyperlink?: string;
  breakLine?: boolean;
};

export type TextElement = {
  type: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  textRuns: TextRunModel[];
  margin?: number;
  rotation?: number;
  align?: "left" | "center" | "right" | "justify";
  valign?: "top" | "mid" | "bottom";
  fill?: FillModel;
  line?: LineModel;
  isPlaceholder?: boolean;
  placeholderType?: string;
};

export type ShapeElement = {
  type: "shape";
  x: number;
  y: number;
  w: number;
  h: number;
  shapeType: string;
  fill?: FillModel;
  line?: LineModel;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  textRuns?: TextRunModel[];
  align?: "left" | "center" | "right" | "justify";
  valign?: "top" | "mid" | "bottom";
};

export type ImageElement = {
  type: "image";
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  dataUri?: string;
  path?: string;
  mimeType?: string;
  altText?: string;
  hyperlink?: string;
};

export type TableCellModel = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  fontFace?: string;
  fontSize?: number;
  color?: string;
  fill?: string;
  align?: "left" | "center" | "right";
  valign?: "top" | "mid" | "bottom";
  colspan?: number;
  rowspan?: number;
};

export type TableElement = {
  type: "table";
  x: number;
  y: number;
  w: number;
  h: number;
  rows: TableCellModel[][];
  colWidths?: number[];
  rowHeights?: number[];
  border?: LineModel;
  fill?: string;
  color?: string;
  fontSize?: number;
};

export type ChartDataSeries = {
  name: string;
  labels?: string[];
  values: number[];
};

export type ChartElement = {
  type: "chart";
  x: number;
  y: number;
  w: number;
  h: number;
  chartType: string;
  data: ChartDataSeries[];
  options?: Record<string, unknown>;
};

export type GroupElement = {
  type: "group";
  x: number;
  y: number;
  w: number;
  h: number;
  children: SlideElement[];
};

export type UnsupportedElement = {
  type: "unsupported";
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  reason: string;
  rawXmlTag?: string;
  fallbackImage?: string;
};

export type SlideElement =
  | TextElement
  | ShapeElement
  | ImageElement
  | TableElement
  | ChartElement
  | GroupElement
  | UnsupportedElement;

export type SlideModel = {
  id: string;
  width: number;
  height: number;
  background?: FillModel;
  elements: SlideElement[];
  notes?: string;
  layoutId?: string;
  masterId?: string;
};

export type DeckModel = {
  layout: {
    width: number;
    height: number;
    name?: string;
  };
  theme?: ThemeModel;
  masters: MasterModel[];
  layouts: LayoutModel[];
  slides: SlideModel[];
};
