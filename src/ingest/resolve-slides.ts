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

    // Target may be relative like "../slides/slide1.xml" or "slides/slide1.xml"
    const target = rel["@_Target"].replace(/^\.\.\//, "ppt/");
    if (target.startsWith("ppt/")) return target;
    return `ppt/${target}`;
  });
}

/**
 * Resolve the layout and master paths for a given slide via its .rels file.
 */
export function resolveSlideRelationships(
  slideRels: Record<string, unknown> | null
): { layoutPath?: string; notesPath?: string; imagePaths: Map<string, string> } {
  const imagePaths = new Map<string, string>();
  if (!slideRels) return { imagePaths };

  const rels = getRelationships(slideRels as Record<string, unknown>);
  let layoutPath: string | undefined;
  let notesPath: string | undefined;

  for (const rel of rels) {
    const type = rel["@_Type"] ?? "";
    const target = rel["@_Target"] ?? "";
    const id = rel["@_Id"] ?? "";

    if (type.endsWith("/slideLayout")) {
      layoutPath = target.replace(/^\.\.\//, "ppt/");
      if (!layoutPath.startsWith("ppt/")) layoutPath = `ppt/${layoutPath}`;
    } else if (type.endsWith("/notesSlide")) {
      notesPath = target.replace(/^\.\.\//, "ppt/");
      if (!notesPath.startsWith("ppt/")) notesPath = `ppt/${notesPath}`;
    } else if (type.endsWith("/image")) {
      const imgPath = target.replace(/^\.\.\//, "ppt/");
      imagePaths.set(id, imgPath.startsWith("ppt/") ? imgPath : `ppt/${imgPath}`);
    }
  }

  return { layoutPath, notesPath, imagePaths };
}
