/**
 * Generate a TypeScript source file that recreates the deck using PptxGenJS.
 * This provides an editable, human-readable script as one of the three outputs.
 */

import { DeckModel, SlideElement, SlideModel } from "../types";

function indent(code: string, spaces = 2): string {
  return code
    .split("\n")
    .map((line) => " ".repeat(spaces) + line)
    .join("\n");
}

function escapeStr(s: string): string {
  return s
    .replace(/\\/g, "\\\\")
    .replace(/`/g, "\\`")
    .replace(/\${/g, "\\${");
}

function elementToCode(el: SlideElement, slideVar: string): string {
  switch (el.type) {
    case "text": {
      const runs = el.textRuns
        .map((r) => {
          const opts: string[] = [];
          if (r.bold) opts.push("bold: true");
          if (r.italic) opts.push("italic: true");
          if (r.underline) opts.push('underline: { style: "sng" }');
          if (r.strike) opts.push("strike: true");
          if (r.fontFace) opts.push(`fontFace: "${escapeStr(r.fontFace)}"`);
          if (r.fontSize) opts.push(`fontSize: ${r.fontSize}`);
          if (r.color) opts.push(`color: "${r.color}"`);
          if (r.breakLine) opts.push("breakLine: true");
          return `{ text: \`${escapeStr(r.text)}\`${opts.length > 0 ? `, options: { ${opts.join(", ")} }` : ""} }`;
        })
        .join(",\n    ");

      const opts: string[] = [
        `x: ${el.x}`,
        `y: ${el.y}`,
        `w: ${el.w}`,
        `h: ${el.h}`,
      ];
      if (el.align) opts.push(`align: "${el.align}"`);
      if (el.valign) opts.push(`valign: "${el.valign === "mid" ? "middle" : el.valign}"`);
      if (el.margin !== undefined) opts.push(`margin: ${el.margin}`);
      if (el.rotation !== undefined) opts.push(`rotate: ${el.rotation}`);
      if (el.fill?.type === "solid") opts.push(`fill: { color: "${el.fill.color}" }`);
      if (el.line?.color) opts.push(`line: { color: "${el.line.color}", width: ${el.line.width ?? 1} }`);

      return `${slideVar}.addText(\n  [\n    ${runs}\n  ],\n  { ${opts.join(", ")} }\n);`;
    }

    case "shape": {
      const opts: string[] = [
        `x: ${el.x}`,
        `y: ${el.y}`,
        `w: ${el.w}`,
        `h: ${el.h}`,
      ];
      if (el.rotation !== undefined) opts.push(`rotate: ${el.rotation}`);
      if (el.flipH) opts.push("flipH: true");
      if (el.flipV) opts.push("flipV: true");
      if (el.fill?.type === "solid") opts.push(`fill: { color: "${el.fill.color}" }`);
      if (el.fill?.type === "none") opts.push(`fill: { type: "none" }`);
      if (el.line?.color) opts.push(`line: { color: "${el.line.color}", width: ${el.line.width ?? 1} }`);

      if (el.textRuns && el.textRuns.length > 0) {
        const runs = el.textRuns.map((r) => `{ text: \`${escapeStr(r.text)}\` }`).join(", ");
        return `${slideVar}.addText([${runs}], { shape: "${el.shapeType}", ${opts.join(", ")} });`;
      }

      return `${slideVar}.addShape("${el.shapeType}", { ${opts.join(", ")} });`;
    }

    case "image": {
      if (!el.dataUri && !el.path) return `// Skipped image: no data`;
      const opts: string[] = [
        `x: ${el.x}`,
        `y: ${el.y}`,
        `w: ${el.w}`,
        `h: ${el.h}`,
      ];
      if (el.rotation !== undefined) opts.push(`rotate: ${el.rotation}`);
      if (el.altText) opts.push(`altText: "${escapeStr(el.altText)}"`);
      const dataProp = el.dataUri
        ? `data: \`${el.dataUri}\``
        : `path: "${escapeStr(el.path ?? "")}"`;
      return `${slideVar}.addImage({ ${dataProp}, ${opts.join(", ")} });`;
    }

    case "table": {
      if (el.rows.length === 0) return `// Skipped empty table`;
      const rows = el.rows
        .map(
          (row) =>
            `[${row.map((c) => `{ text: \`${escapeStr(c.text)}\` }`).join(", ")}]`
        )
        .join(",\n    ");
      const opts: string[] = [
        `x: ${el.x}`,
        `y: ${el.y}`,
        `w: ${el.w}`,
        `h: ${el.h}`,
      ];
      return `${slideVar}.addTable(\n  [\n    ${rows}\n  ],\n  { ${opts.join(", ")} }\n);`;
    }

    case "chart": {
      if (el.data.length === 0) return `// Skipped empty chart`;
      const data = el.data.map((s) => {
        const labels = s.labels ? `labels: [${s.labels.map((l) => `"${escapeStr(l)}"`).join(", ")}], ` : "";
        return `{ name: "${escapeStr(s.name)}", ${labels}values: [${s.values.join(", ")}] }`;
      });
      const opts: string[] = [
        `x: ${el.x}`,
        `y: ${el.y}`,
        `w: ${el.w}`,
        `h: ${el.h}`,
      ];
      return `${slideVar}.addChart("${el.chartType}", [\n  ${data.join(",\n  ")}\n], { ${opts.join(", ")} });`;
    }

    case "group":
      return el.children.map((c) => elementToCode(c, slideVar)).join("\n");

    case "unsupported":
      return `// Unsupported element: ${el.reason}`;

    default:
      return `// Unknown element type`;
  }
}

function slideToCode(slide: SlideModel, index: number): string {
  const slideVar = `slide${index + 1}`;
  const lines: string[] = [`const ${slideVar} = pptx.addSlide();`];

  if (slide.background?.type === "solid") {
    lines.push(`${slideVar}.background = { color: "${slide.background.color}" };`);
  }

  for (const el of slide.elements) {
    lines.push(elementToCode(el, slideVar));
  }

  if (slide.notes) {
    lines.push(`${slideVar}.addNotes(\`${escapeStr(slide.notes)}\`);`);
  }

  return lines.join("\n");
}

/**
 * Generate a TypeScript source file that reproduces the deck using PptxGenJS.
 */
export function generateCode(deck: DeckModel): string {
  const parts: string[] = [
    `/**`,
    ` * Auto-generated by pptx2js`,
    ` * Recreates the original deck using PptxGenJS.`,
    ` */`,
    ``,
    `import pptxgen from "pptxgenjs";`,
    ``,
    `async function main() {`,
    `  const pptx = new pptxgen();`,
    ``,
    `  pptx.defineLayout({`,
    `    name: "CUSTOM",`,
    `    width: ${deck.layout.width},`,
    `    height: ${deck.layout.height},`,
    `  });`,
    `  pptx.layout = "CUSTOM";`,
    ``,
  ];

  for (let i = 0; i < deck.slides.length; i++) {
    parts.push(indent(slideToCode(deck.slides[i], i), 2));
    parts.push("");
  }

  parts.push(`  await pptx.writeFile({ fileName: "recreated.pptx" });`);
  parts.push(`  console.log("Wrote recreated.pptx");`);
  parts.push(`}`);
  parts.push(``);
  parts.push(`main().catch(console.error);`);
  parts.push(``);

  return parts.join("\n");
}
