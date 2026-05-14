/**
 * Resolve the ordered list of slide XML paths from presentation.xml.
 * PowerPoint stores slide order as a list of r:id references in sldIdLst,
 * which must be resolved through presentation.xml.rels.
 */

type Relationship = {
  "@_Id": string;
  "@_Type": string;
  "@_Target": string;
};

/**
 * Normalize an OOXML relationship Target to a zip-relative path under "ppt/".
 * Handles absolute paths (e.g. /ppt/slides/slide1.xml) and relative paths
 * (e.g. ../slides/slide1.xml or slides/slide1.xml).
 */
export function normalizeTarget(target: string): string {
  // Strip leading slash for absolute OOXML paths (e.g. /ppt/slides/...)
  const stripped = target.startsWith("/") ? target.slice(1) : target;
  // Replace leading "../" which is relative to ppt/ directory
  const resolved = stripped.replace(/^\.\.\//, "ppt/");
  return resolved.startsWith("ppt/") ? resolved : `ppt/${resolved}`;
}

function getRelationships(presentationRels: Record<string, unknown>): Relationship[] {
  const rels = (presentationRels as { Relationships?: { Relationship?: unknown } })
    .Relationships?.Relationship;
  if (!rels) return [];
  return Array.isArray(rels) ? (rels as Relationship[]) : [rels as Relationship];
}

function getSlideIds(presentation: Record<string, unknown>): Array<{ "@_r:id": string }> {
  const pres = (presentation as Record<string, Record<string, unknown>>)[
    "p:presentation"
  ];
  if (!pres) throw new Error("Invalid PPTX: missing p:presentation element");

  const sldIdLst = (pres as Record<string, unknown>)["p:sldIdLst"] as
    | Record<string, unknown>
    | undefined;
  if (!sldIdLst) return [];

  const slideIds = sldIdLst["p:sldId"];
  if (!slideIds) return [];

  return Array.isArray(slideIds)
    ? (slideIds as Array<{ "@_r:id": string }>)
    : [slideIds as { "@_r:id": string }];
}

/**
 * Returns ordered list of slide paths within the zip (e.g. "ppt/slides/slide1.xml").
 */
export function resolveSlidePaths(
  presentation: Record<string, unknown>,
  presentationRels: Record<string, unknown>
): string[] {
  const slideIds = getSlideIds(presentation);
  const relationships = getRelationships(presentationRels);

  return slideIds.map((slide) => {
    const relId = slide["@_r:id"];
    if (!relId) {
      throw new Error("Slide entry missing r:id attribute");
    }

    const rel = relationships.find((r) => r["@_Id"] === relId);
    if (!rel) {
      throw new Error(`Missing relationship for slide r:id="${relId}"`);
    }

    // Target may be relative like "../slides/slide1.xml" or absolute like "/ppt/slides/slide1.xml"
    return normalizeTarget(rel["@_Target"]);
  });
}

/**
 * Resolve the layout, notes, image, and hyperlink paths/URLs for a slide via its .rels file.
 */
export function resolveSlideRelationships(
  slideRels: Record<string, unknown> | null
): {
  layoutPath?: string;
  notesPath?: string;
  imagePaths: Map<string, string>;
  hyperlinkUrls: Map<string, string>;
} {
  const imagePaths = new Map<string, string>();
  const hyperlinkUrls = new Map<string, string>();
  if (!slideRels) return { imagePaths, hyperlinkUrls };

  const rels = getRelationships(slideRels as Record<string, unknown>);
  let layoutPath: string | undefined;
  let notesPath: string | undefined;

  for (const rel of rels) {
    const type = rel["@_Type"] ?? "";
    const target = rel["@_Target"] ?? "";
    const id = rel["@_Id"] ?? "";

    if (type.endsWith("/slideLayout")) {
      layoutPath = normalizeTarget(target);
    } else if (type.endsWith("/notesSlide")) {
      notesPath = normalizeTarget(target);
    } else if (type.endsWith("/image")) {
      imagePaths.set(id, normalizeTarget(target));
    } else if (type.endsWith("/hyperlink")) {
      // Hyperlink targets are actual URLs; store them as-is (no path normalization)
      hyperlinkUrls.set(id, target);
    }
  }

  return { layoutPath, notesPath, imagePaths, hyperlinkUrls };
}
