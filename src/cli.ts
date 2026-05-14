#!/usr/bin/env node
/**
 * CLI entry point for pptx2js.
 *
 * Usage:
 *   pptx2js <input.pptx> [options]
 *
 * Options:
 *   --out <dir>        Output directory (default: ./generated)
 *   --no-deck-json     Skip emitting deck.json
 *   --no-code          Skip emitting recreate.ts
 *   --no-pptx          Skip emitting recreated.pptx
 *   --verbose          Show full validation report
 */

import { Command } from "commander";
import * as path from "path";
import { recreatePptx } from "./index";

const program = new Command();

program
  .name("pptx2js")
  .description(
    "Convert a .pptx file into a normalized deck.json, a PptxGenJS TypeScript script, and a recreated .pptx"
  )
  .version("1.0.0")
  .argument("<input>", "Path to the input .pptx file")
  .option("--out <dir>", "Output directory", "./generated")
  .option("--no-deck-json", "Skip emitting deck.json")
  .option("--no-code", "Skip emitting recreate.ts")
  .option("--no-pptx", "Skip emitting recreated.pptx")
  .option("--verbose", "Print the full validation report")
  .action(async (input: string, opts: { out: string; deckJson: boolean; code: boolean; pptx: boolean; verbose: boolean }) => {
    const inputPath = path.resolve(input);

    console.log(`\npptx2js — converting: ${inputPath}`);
    console.log(`Output directory:     ${path.resolve(opts.out)}\n`);

    try {
      const result = await recreatePptx(inputPath, {
        outDir: opts.out,
        emitDeckJson: opts.deckJson !== false,
        emitCode: opts.code !== false,
        emitPptx: opts.pptx !== false,
      });

      const { report } = result;
      console.log(`✓ Slides:    ${report.totalSlides}`);
      console.log(`✓ Elements:  ${report.totalElements}`);

      if (report.unsupportedElements.length > 0) {
        console.warn(`⚠ Unsupported: ${report.unsupportedElements.length} element(s) skipped`);
        if (opts.verbose) {
          for (const u of report.unsupportedElements) {
            console.warn(
              `    Slide ${u.slideIndex + 1}: ${u.reason}${u.rawXmlTag ? ` [${u.rawXmlTag}]` : ""}`
            );
          }
        }
      }

      if (result.deckJsonPath) console.log(`\n→ deck.json:     ${result.deckJsonPath}`);
      if (result.codePath) console.log(`→ recreate.ts:   ${result.codePath}`);
      if (result.pptxPath) console.log(`→ recreated.pptx: ${result.pptxPath}`);

      console.log("\nDone.");
    } catch (err) {
      console.error("Error:", err instanceof Error ? err.message : String(err));
      process.exit(1);
    }
  });

program.parse(process.argv);
