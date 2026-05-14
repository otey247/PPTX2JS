/**
 * Generate a recreated .pptx file from a DeckModel using PptxGenJS.
 */

// eslint-disable-next-line @typescript-eslint/no-require-imports
const pptxgen = require("pptxgenjs");

import {
  DeckModel,
  SlideElement,
  TextElement,
  ShapeElement,
  ImageElement,
  TableElement,
  ChartElement,
} from "../types";

function mapAlign(
  align?: "left" | "center" | "right" | "justify"
): "left" | "center" | "right" | "justify" | undefined {
  return align;
}

function mapValign(
  valign?: "top" | "mid" | "bottom"
): "top" | "middle" | "bottom" | undefined {
  if (valign === "mid") return "middle";
  return valign as "top" | "bottom" | undefined;
}

function addTextElement(slide: ReturnType<typeof pptxgen.prototype.addSlide>, el: TextElement): void {
  if (el.textRuns.length === 0) return;

  const textItems = el.textRuns.map((run) => ({
    text: run.text,
    options: {
      bold: run.bold,
      italic: run.italic,
      underline: run.underline ? { style: "sng" } : undefined,
      strike: run.strike,
      fontFace: run.fontFace,
      fontSize: run.fontSize,
      color: run.color,
      breakLine: run.breakLine,
      hyperlink: run.hyperlink ? { url: run.hyperlink } : undefined,
    },
  }));

  const opts: Record<string, unknown> = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    align: mapAlign(el.align),
    valign: mapValign(el.valign),
    margin: el.margin ?? 0,
  };

  if (el.rotation !== undefined) opts.rotate = el.rotation;

  if (el.fill?.type === "solid") {
    opts.fill = { color: el.fill.color };
  } else if (el.fill?.type === "none") {
    opts.fill = { type: "none" };
  }

  if (el.line?.color) {
    opts.line = {
      color: el.line.color,
      width: el.line.width ?? 1,
      dashType: el.line.dashType ?? "solid",
    };
  }

  slide.addText(textItems, opts);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function addShapeElement(slide: ReturnType<typeof pptxgen.prototype.addSlide>, el: ShapeElement, pptx: any): void {
  const opts: Record<string, unknown> = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
  };

  if (el.rotation !== undefined) opts.rotate = el.rotation;
  if (el.flipH) opts.flipH = true;
  if (el.flipV) opts.flipV = true;

  if (el.fill?.type === "solid") {
    opts.fill = { color: el.fill.color };
  } else if (el.fill?.type === "none") {
    opts.fill = { type: "none" };
  }

  if (el.line) {
    opts.line = {
      color: el.line.color,
      width: el.line.width ?? 1,
      dashType: el.line.dashType ?? "solid",
    };
  }

  // Resolve shape type via PptxGenJS ShapeType enum; fall back to "rect"
  const pptxShapeType: string = pptx.ShapeType[el.shapeType] ?? pptx.ShapeType.rect;

  if (el.textRuns && el.textRuns.length > 0) {
    const textItems = el.textRuns.map((run) => ({
      text: run.text,
      options: {
        bold: run.bold,
        italic: run.italic,
        fontFace: run.fontFace,
        fontSize: run.fontSize,
        color: run.color,
        breakLine: run.breakLine,
      },
    }));
    slide.addText(textItems, { ...opts, shape: pptxShapeType });
  } else {
    slide.addShape(pptxShapeType, opts);
  }
}

function addImageElement(slide: ReturnType<typeof pptxgen.prototype.addSlide>, el: ImageElement): void {
  if (!el.dataUri && !el.path) return;

  const opts: Record<string, unknown> = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
  };

  if (el.rotation !== undefined) opts.rotate = el.rotation;
  if (el.flipH) opts.flipH = true;
  if (el.flipV) opts.flipV = true;
  if (el.altText) opts.altText = el.altText;
  if (el.hyperlink) opts.hyperlink = { url: el.hyperlink };

  if (el.dataUri) {
    opts.data = el.dataUri;
  } else if (el.path) {
    opts.path = el.path;
  }

  slide.addImage(opts);
}

function addTableElement(slide: ReturnType<typeof pptxgen.prototype.addSlide>, el: TableElement): void {
  if (el.rows.length === 0) return;

  const rows = el.rows.map((row) =>
    row.map((cell) => ({
      text: cell.text,
      options: {
        bold: cell.bold,
        italic: cell.italic,
        fontFace: cell.fontFace,
        fontSize: cell.fontSize,
        color: cell.color,
        fill: cell.fill ? { color: cell.fill } : undefined,
        align: cell.align,
        valign: cell.valign,
        colspan: cell.colspan,
        rowspan: cell.rowspan,
      },
    }))
  );

  const opts: Record<string, unknown> = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
  };

  if (el.colWidths) opts.colW = el.colWidths;
  if (el.rowHeights) opts.rowH = el.rowHeights;
  if (el.border) {
    opts.border = {
      color: el.border.color,
      pt: el.border.width ?? 1,
    };
  }
  if (el.fill) opts.fill = { color: el.fill };
  if (el.color) opts.color = el.color;
  if (el.fontSize) opts.fontSize = el.fontSize;

  slide.addTable(rows, opts);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function addChartElement(slide: ReturnType<typeof pptxgen.prototype.addSlide>, el: ChartElement, pptx: any): void {
  if (el.data.length === 0) return;

  // PptxGenJS chart data format
  const data = el.data.map((series) => ({
    name: series.name,
    labels: series.labels ?? series.values.map((_, i) => String(i + 1)),
    values: series.values,
  }));

  const opts: Record<string, unknown> = {
    x: el.x,
    y: el.y,
    w: el.w,
    h: el.h,
    ...(el.options ?? {}),
  };

  // Resolve chart type via PptxGenJS ChartType enum; fall back to bar
  const pptxChartType: string = pptx.ChartType[el.chartType] ?? pptx.ChartType.bar;
  slide.addChart(pptxChartType, data, opts);
}

function processElement(
  slide: ReturnType<typeof pptxgen.prototype.addSlide>,
  el: SlideElement,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pptx: any
): void {
  switch (el.type) {
    case "text":
      addTextElement(slide, el as TextElement);
      break;
    case "shape":
      addShapeElement(slide, el as ShapeElement, pptx);
      break;
    case "image":
      addImageElement(slide, el as ImageElement);
      break;
    case "table":
      addTableElement(slide, el as TableElement);
      break;
    case "chart":
      addChartElement(slide, el as ChartElement, pptx);
      break;
    case "group":
      // Flatten group: process children directly
      for (const child of el.children) {
        processElement(slide, child, pptx);
      }
      break;
    case "unsupported":
      // Log unsupported elements but do not fail
      break;
  }
}

/**
 * Build and configure a PptxGenJS instance from a DeckModel.
 * Shared by generatePptx (buffer) and generatePptxFile (disk).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildPptxInstance(deck: DeckModel): any {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pptx = new (pptxgen as any)();

  // Define custom layout
  pptx.defineLayout({
    name: "CUSTOM",
    width: deck.layout.width,
    height: deck.layout.height,
  });
  pptx.layout = "CUSTOM";

  for (const sourceSlide of deck.slides) {
    const slide = pptx.addSlide();

    // Background
    if (sourceSlide.background?.type === "solid") {
      slide.background = { color: sourceSlide.background.color };
    }

    // Elements
    for (const el of sourceSlide.elements) {
      try {
        processElement(slide, el, pptx);
      } catch {
        // Skip elements that fail to render without crashing the whole deck
      }
    }

    // Speaker notes
    if (sourceSlide.notes) {
      slide.addNotes(sourceSlide.notes);
    }
  }

  return pptx;
}

/**
 * Generate a .pptx file from a DeckModel using PptxGenJS.
 * Returns a Buffer containing the generated file.
 */
export async function generatePptx(deck: DeckModel): Promise<Buffer> {
  const pptx = buildPptxInstance(deck);
  const result = await pptx.write({ outputType: "nodebuffer" });
  return result as Buffer;
}

/**
 * Write a .pptx file to disk from a DeckModel.
 */
export async function generatePptxFile(
  deck: DeckModel,
  outPath: string
): Promise<void> {
  const pptx = buildPptxInstance(deck);
  await pptx.writeFile({ fileName: outPath });
}
