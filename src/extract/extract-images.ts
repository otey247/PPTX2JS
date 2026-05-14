/**
 * Extract image elements (p:pic) from slide XML.
 */

import JSZip from "jszip";
import { ImageElement } from "../types";
import { emuToInches } from "../normalize/units";

function getAttr(obj: Record<string, unknown>, key: string): string | undefined {
  return obj[`@_${key}`] as string | undefined;
}

function getMimeType(ext: string): string {
  const map: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    tiff: "image/tiff",
    tif: "image/tiff",
    svg: "image/svg+xml",
    webp: "image/webp",
    wmf: "image/x-wmf",
    emf: "image/x-emf",
  };
  return map[ext.toLowerCase()] ?? "image/png";
}

/**
 * Extract an image element from a p:pic node.
 * Loads the image binary from the zip and encodes it as a data URI.
 */
export async function extractImageElement(
  pic: Record<string, unknown>,
  zip: JSZip,
  imagePaths: Map<string, string>
): Promise<ImageElement | null> {
  const spPr = pic["p:spPr"] as Record<string, unknown> | undefined;
  if (!spPr) return null;

  const xfrm = spPr["a:xfrm"] as Record<string, unknown> | undefined;
  if (!xfrm) return null;

  const off = xfrm["a:off"] as Record<string, unknown> | undefined;
  const ext = xfrm["a:ext"] as Record<string, unknown> | undefined;

  const x = emuToInches(getAttr(off ?? {}, "x") ?? "0");
  const y = emuToInches(getAttr(off ?? {}, "y") ?? "0");
  const w = emuToInches(getAttr(ext ?? {}, "cx") ?? "0");
  const h = emuToInches(getAttr(ext ?? {}, "cy") ?? "0");

  const rot = getAttr(xfrm, "rot");
  const rotation = rot ? Math.round(Number(rot) / 60000) : undefined;
  const flipH = getAttr(xfrm, "flipH") === "1" || getAttr(xfrm, "flipH") === "true";
  const flipV = getAttr(xfrm, "flipV") === "1" || getAttr(xfrm, "flipV") === "true";

  // Resolve image relationship ID
  const nvPicPr = pic["p:nvPicPr"] as Record<string, unknown> | undefined;
  const cNvPr = nvPicPr?.["p:cNvPr"] as Record<string, unknown> | undefined;
  const altText = cNvPr ? getAttr(cNvPr, "descr") : undefined;

  const blipFill = pic["p:blipFill"] as Record<string, unknown> | undefined;
  const blip = blipFill?.["a:blip"] as Record<string, unknown> | undefined;
  const rEmbed = blip
    ? (getAttr(blip, "r:embed") ?? getAttr(blip, "r_embed"))
    : undefined;

  if (!rEmbed) return { type: "image", x, y, w, h, rotation, altText };

  const imgPath = imagePaths.get(rEmbed);
  if (!imgPath) return { type: "image", x, y, w, h, rotation, altText };

  try {
    const ext2 = imgPath.split(".").pop() ?? "png";
    const mimeType = getMimeType(ext2);
    const data = await zip.file(imgPath)?.async("base64");
    if (!data) return { type: "image", x, y, w, h, rotation, altText };

    const dataUri = `data:${mimeType};base64,${data}`;
    return {
      type: "image",
      x,
      y,
      w,
      h,
      rotation,
      flipH: flipH || undefined,
      flipV: flipV || undefined,
      dataUri,
      mimeType,
      altText,
    };
  } catch {
    return { type: "image", x, y, w, h, rotation, altText };
  }
}
