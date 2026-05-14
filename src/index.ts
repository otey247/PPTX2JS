/**
 * Public API for pptx2js.
 */

export { readPptx } from "./ingest/read-pptx";
export { resolveSlidePaths } from "./ingest/resolve-slides";
export { extractDeckModel } from "./deck-builder";
export { generatePptx, generatePptxFile } from "./generate/generate-pptxgen";
export { generateCode } from "./generate/generate-code";
export { validateDeck } from "./validate/unsupported-report";
export type {
  DeckModel,
  SlideModel,
  SlideElement,
  TextElement,
  ShapeElement,
  ImageElement,
  TableElement,
  ChartElement,
  GroupElement,
  UnsupportedElement,
  ThemeModel,
  MasterModel,
  FillModel,
  LineModel,
  TextRunModel,
} from "./types";

import * as fs from "fs";
import * as path from "path";
import { readPptx } from "./ingest/read-pptx";
import { extractDeckModel } from "./deck-builder";
import { generatePptxFile } from "./generate/generate-pptxgen";
import { generateCode } from "./generate/generate-code";
import { validateDeck } from "./validate/unsupported-report";

export type RecreateOptions = {
  /** Output directory for generated files (default: "./generated") */
  outDir?: string;
  /** Emit deck.json (default: true) */
  emitDeckJson?: boolean;
  /** Emit recreate.ts (default: true) */
  emitCode?: boolean;
  /** Emit recreated.pptx (default: true) */
  emitPptx?: boolean;
};

export type RecreateResult = {
  deckJsonPath?: string;
  codePath?: string;
  pptxPath?: string;
  report: ReturnType<typeof validateDeck>;
};

/**
 * Full pipeline: read a .pptx file, extract a deck model, and generate outputs.
 */
export async function recreatePptx(
  inputPath: string,
  options: RecreateOptions = {}
): Promise<RecreateResult> {
  const {
    outDir = "./generated",
    emitDeckJson = true,
    emitCode = true,
    emitPptx = true,
  } = options;

  const buffer = await fs.promises.readFile(inputPath);
  const pkg = await readPptx(buffer);
  const deck = await extractDeckModel(pkg);
  const report = validateDeck(deck);

  await fs.promises.mkdir(outDir, { recursive: true });

  const result: RecreateResult = { report };

  if (emitDeckJson) {
    result.deckJsonPath = path.join(outDir, "deck.json");
    await fs.promises.writeFile(
      result.deckJsonPath,
      JSON.stringify(deck, null, 2),
      "utf-8"
    );
  }

  if (emitCode) {
    result.codePath = path.join(outDir, "recreate.ts");
    await fs.promises.writeFile(result.codePath, generateCode(deck), "utf-8");
  }

  if (emitPptx) {
    result.pptxPath = path.join(outDir, "recreated.pptx");
    await generatePptxFile(deck, result.pptxPath);
  }

  return result;
}
