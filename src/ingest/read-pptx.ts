import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const XML_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  parseTagValue: true,
  trimValues: true,
  isArray: (tagName) =>
    [
      "p:sldId",
      "Relationship",
      "p:sp",
      "p:pic",
      "p:graphicFrame",
      "p:grpSp",
      "p:cxnSp",
      "a:r",
      "a:br",
      "a:p",
      "a:tr",
      "a:tc",
      "c:ser",
      "c:val",
      "c:cat",
      "c:pt",
    ].includes(tagName),
});

export type PptxPackage = {
  zip: JSZip;
  presentation: Record<string, unknown>;
  presentationRels: Record<string, unknown>;
  parser: XMLParser;
};

/**
 * Read a .pptx file buffer and return parsed presentation XML and relationships.
 */
export async function readPptx(buffer: Buffer): Promise<PptxPackage> {
  const zip = await JSZip.loadAsync(buffer);

  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const presentationRelsXml = await zip
    .file("ppt/_rels/presentation.xml.rels")
    ?.async("text");

  if (!presentationXml) {
    throw new Error("Invalid PPTX: missing ppt/presentation.xml");
  }
  if (!presentationRelsXml) {
    throw new Error("Invalid PPTX: missing ppt/_rels/presentation.xml.rels");
  }

  const presentation = XML_PARSER.parse(presentationXml) as Record<
    string,
    unknown
  >;
  const presentationRels = XML_PARSER.parse(presentationRelsXml) as Record<
    string,
    unknown
  >;

  return { zip, presentation, presentationRels, parser: XML_PARSER };
}

/**
 * Parse arbitrary XML from the zip and return an object tree.
 */
export async function parseZipXml(
  zip: JSZip,
  path: string
): Promise<Record<string, unknown> | null> {
  const text = await zip.file(path)?.async("text");
  if (!text) return null;
  return XML_PARSER.parse(text) as Record<string, unknown>;
}

export { XML_PARSER };
