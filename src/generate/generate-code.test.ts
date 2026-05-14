import { generateCode } from "../generate/generate-code";
import { DeckModel } from "../types";

const simpleDeck: DeckModel = {
  layout: { width: 10, height: 7.5, name: "CUSTOM" },
  masters: [],
  layouts: [],
  slides: [
    {
      id: "slide1",
      width: 10,
      height: 7.5,
      elements: [
        {
          type: "text",
          x: 1,
          y: 1,
          w: 8,
          h: 1.5,
          textRuns: [{ text: "Hello World", bold: true, fontSize: 24, color: "FF0000" }],
          align: "center",
        },
        {
          type: "shape",
          x: 2,
          y: 3,
          w: 3,
          h: 2,
          shapeType: "rect",
          fill: { type: "solid", color: "4472C4" },
        },
      ],
      notes: "Speaker notes here",
    },
  ],
};

describe("generateCode", () => {
  test("generates valid TypeScript source", () => {
    const code = generateCode(simpleDeck);
    expect(code).toContain("import pptxgen");
    expect(code).toContain("pptx.defineLayout");
    expect(code).toContain('width: 10');
    expect(code).toContain('height: 7.5');
    expect(code).toContain("addSlide");
    expect(code).toContain("Hello World");
    expect(code).toContain("addText");
    expect(code).toContain("addShape");
    expect(code).toContain("Speaker notes here");
    expect(code).toContain("addNotes");
  });

  test("handles empty deck", () => {
    const emptyDeck: DeckModel = {
      layout: { width: 13.333, height: 7.5 },
      masters: [],
      layouts: [],
      slides: [],
    };
    const code = generateCode(emptyDeck);
    expect(code).toContain("import pptxgen");
    expect(code).toContain("writeFile");
  });

  test("handles unsupported elements gracefully", () => {
    const deckWithUnsupported: DeckModel = {
      ...simpleDeck,
      slides: [
        {
          id: "slide1",
          width: 10,
          height: 7.5,
          elements: [
            {
              type: "unsupported",
              reason: "SmartArt not supported",
              rawXmlTag: "p:graphicFrame",
            },
          ],
        },
      ],
    };
    const code = generateCode(deckWithUnsupported);
    expect(code).toContain("Unsupported element");
    expect(code).toContain("SmartArt not supported");
  });

  test("escapes backticks in text", () => {
    const deckWithBacktick: DeckModel = {
      ...simpleDeck,
      slides: [
        {
          id: "slide1",
          width: 10,
          height: 7.5,
          elements: [
            {
              type: "text",
              x: 0,
              y: 0,
              w: 5,
              h: 1,
              textRuns: [{ text: "Hello `World`" }],
            },
          ],
        },
      ],
    };
    const code = generateCode(deckWithBacktick);
    expect(code).toContain("Hello \\`World\\`");
  });
});
