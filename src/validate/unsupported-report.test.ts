import { validateDeck } from "../validate/unsupported-report";
import { DeckModel } from "../types";

const deck: DeckModel = {
  layout: { width: 10, height: 7.5 },
  masters: [],
  layouts: [],
  slides: [
    {
      id: "slide1",
      width: 10,
      height: 7.5,
      elements: [
        { type: "text", x: 0, y: 0, w: 5, h: 1, textRuns: [{ text: "Hello" }] },
        {
          type: "unsupported",
          reason: "SmartArt not supported",
          rawXmlTag: "p:graphicFrame",
        },
      ],
    },
    {
      id: "slide2",
      width: 10,
      height: 7.5,
      elements: [
        { type: "shape", x: 0, y: 0, w: 3, h: 2, shapeType: "rect" },
      ],
    },
  ],
};

describe("validateDeck", () => {
  test("counts slides and elements correctly", () => {
    const report = validateDeck(deck);
    expect(report.totalSlides).toBe(2);
    expect(report.totalElements).toBe(3);
  });

  test("lists unsupported elements", () => {
    const report = validateDeck(deck);
    expect(report.unsupportedElements).toHaveLength(1);
    expect(report.unsupportedElements[0].slideIndex).toBe(0);
    expect(report.unsupportedElements[0].reason).toBe("SmartArt not supported");
    expect(report.unsupportedElements[0].rawXmlTag).toBe("p:graphicFrame");
  });

  test("adds warning for unsupported elements", () => {
    const report = validateDeck(deck);
    expect(report.warnings.some((w) => w.includes("unsupported"))).toBe(true);
  });

  test("warns on empty deck", () => {
    const emptyDeck: DeckModel = {
      layout: { width: 10, height: 7.5 },
      masters: [],
      layouts: [],
      slides: [],
    };
    const report = validateDeck(emptyDeck);
    expect(report.warnings.some((w) => w.includes("No slides"))).toBe(true);
  });
});
