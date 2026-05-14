/**
 * Validate output by reporting unsupported elements and warnings.
 */

import { DeckModel, UnsupportedElement } from "../types";

export type ValidationReport = {
  totalSlides: number;
  totalElements: number;
  unsupportedElements: Array<{
    slideIndex: number;
    reason: string;
    rawXmlTag?: string;
  }>;
  warnings: string[];
};

/**
 * Scan the deck model for unsupported or potentially problematic elements.
 */
export function validateDeck(deck: DeckModel): ValidationReport {
  const report: ValidationReport = {
    totalSlides: deck.slides.length,
    totalElements: 0,
    unsupportedElements: [],
    warnings: [],
  };

  for (let i = 0; i < deck.slides.length; i++) {
    const slide = deck.slides[i];
    report.totalElements += slide.elements.length;

    for (const el of slide.elements) {
      if (el.type === "unsupported") {
        const u = el as UnsupportedElement;
        report.unsupportedElements.push({
          slideIndex: i,
          reason: u.reason,
          rawXmlTag: u.rawXmlTag,
        });
      }
    }
  }

  if (deck.slides.length === 0) {
    report.warnings.push("No slides were extracted from the PPTX.");
  }

  if (report.unsupportedElements.length > 0) {
    report.warnings.push(
      `${report.unsupportedElements.length} unsupported element(s) were skipped.`
    );
  }

  return report;
}
