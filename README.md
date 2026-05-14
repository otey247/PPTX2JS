# PPTX2JS

A **PowerPoint-to-PptxGenJS reconstruction tool** that reads a `.pptx` file, extracts its slide structure into a normalized JSON representation, and generates a [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) script that recreates the deck.

> **Note:** This is a reconstruction tool, not a perfect importer. Supported elements are recreated as editable PptxGenJS objects. Unsupported or highly complex elements are skipped with warnings.

## Architecture

```
.pptx file
   │
   ▼
PPTX Package Reader (JSZip + fast-xml-parser)
   │
   ▼
OOXML Extractors
   - slides, layouts, masters, themes
   - text, shapes, images, tables, charts, notes
   │
   ▼
Normalized Deck Model (deck.json)
   │
   ▼
PptxGenJS Code Generator
   - recreate.ts   ← editable TypeScript script
   - recreated.pptx ← recreated PowerPoint file
```

## Installation

```bash
npm install
npm run build
```

## CLI Usage

```bash
# Convert a .pptx file (generates deck.json, recreate.ts, recreated.pptx in ./generated)
node dist/cli.js input.pptx

# Specify output directory
node dist/cli.js input.pptx --out ./my-output

# Verbose mode (shows unsupported elements per slide)
node dist/cli.js input.pptx --verbose

# Skip individual outputs
node dist/cli.js input.pptx --no-deck-json
node dist/cli.js input.pptx --no-code
node dist/cli.js input.pptx --no-pptx
```

## Programmatic API

```typescript
import { recreatePptx } from "pptx2js";

const result = await recreatePptx("input.pptx", {
  outDir: "./generated",
  emitDeckJson: true,
  emitCode: true,
  emitPptx: true,
});

console.log(`Slides: ${result.report.totalSlides}`);
console.log(`Elements: ${result.report.totalElements}`);
console.log(`Unsupported: ${result.report.unsupportedElements.length}`);
```

### Lower-level API

```typescript
import { readPptx, extractDeckModel, generatePptx, generateCode } from "pptx2js";
import * as fs from "fs";

const buffer = await fs.promises.readFile("input.pptx");
const pkg = await readPptx(buffer);
const deck = await extractDeckModel(pkg);

// deck.json — normalized intermediate model
await fs.promises.writeFile("deck.json", JSON.stringify(deck, null, 2));

// recreate.ts — editable TypeScript script
const code = generateCode(deck);
await fs.promises.writeFile("recreate.ts", code);

// recreated.pptx — regenerated PowerPoint
const pptxBuffer = await generatePptx(deck);
await fs.promises.writeFile("recreated.pptx", pptxBuffer);
```

## Three Outputs

Every conversion produces three files:

| Output | Description |
|--------|-------------|
| `deck.json` | Machine-readable normalized deck model. Inspect, diff, or patch before regenerating. |
| `recreate.ts` | Editable TypeScript source that calls PptxGenJS APIs to recreate the deck. |
| `recreated.pptx` | The regenerated PowerPoint file. |

## Supported Elements

| OOXML Node | Extracted As |
|-----------|--------------|
| `p:sp` (text box) | `TextElement` → `slide.addText()` |
| `p:sp` (shape with `prstGeom`) | `ShapeElement` → `slide.addShape()` |
| `p:pic` | `ImageElement` → `slide.addImage()` |
| `p:graphicFrame` + `a:tbl` | `TableElement` → `slide.addTable()` |
| `p:graphicFrame` + chart | `ChartElement` → `slide.addChart()` |
| `p:cxnSp` | `ShapeElement` (line) → `slide.addShape()` |
| `p:grpSp` | Flattened group → children extracted individually |
| Notes slide | `slide.addNotes()` |
| Background fill | `slide.background` |
| Theme colors + fonts | Used for color/font resolution |

Unsupported elements (SmartArt, animations, custom geometry, etc.) are recorded as `UnsupportedElement` entries in `deck.json` with a reason string.

## Folder Structure

```
src/
  cli.ts                     ← CLI entry point
  index.ts                   ← Public API
  types.ts                   ← DeckModel type definitions
  deck-builder.ts            ← Builds DeckModel from PptxPackage
  ingest/
    read-pptx.ts             ← Opens zip, parses presentation.xml
    resolve-slides.ts        ← Resolves slide order via .rels
    relationships.ts         ← .rels file loader
    content-types.ts         ← [Content_Types].xml parser
  extract/
    extract-slide.ts         ← Top-level slide extractor
    extract-text.ts          ← Text elements (p:sp with txBody)
    extract-shapes.ts        ← Shape elements (p:sp, p:cxnSp)
    extract-images.ts        ← Image elements (p:pic)
    extract-tables.ts        ← Table elements (p:graphicFrame/a:tbl)
    extract-charts.ts        ← Chart elements (p:graphicFrame/c:chart)
    extract-notes.ts         ← Speaker notes
    extract-theme.ts         ← Theme colors and fonts
    extract-masters.ts       ← Slide masters
  normalize/
    units.ts                 ← EMU → inches conversion
    colors.ts                ← Color normalization
    fonts.ts                 ← Font face resolution
    shape-map.ts             ← OOXML shape name mapping
  generate/
    generate-pptxgen.ts      ← PptxGenJS file generator
    generate-code.ts         ← TypeScript source code generator
  validate/
    unsupported-report.ts    ← Validation and unsupported element report
```

## Development

```bash
npm run build     # Compile TypeScript
npm test          # Run tests
npm run lint      # Run ESLint
```

## Fidelity Notes

This tool achieves good fidelity for:
- Text boxes with rich formatting (bold, italic, color, font size, underline)
- Basic shapes with solid fills and borders
- Embedded images (encoded as base64 data URIs)
- Tables
- Basic charts (bar, line, pie, area, doughnut)
- Speaker notes
- Slide backgrounds
- Theme color resolution

Complex features not fully supported in this version:
- Animations and transitions
- SmartArt / diagrams
- Complex gradients
- Grouped object transforms
- Slide layout/master inheritance
- Embedded videos
