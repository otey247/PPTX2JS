import { extractTextElement } from "../extract/extract-text";

const themeColorMap = { accent1: "4472C4" };

function makeSpNode(opts: {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  text?: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: string;
  hasPrstGeom?: boolean;
}) {
  const emu = (inches: number) => String(Math.round(inches * 914400));

  return {
    "p:spPr": {
      ...(opts.hasPrstGeom
        ? { "a:prstGeom": { "@_prst": "rect" } }
        : {}),
      "a:xfrm": {
        "a:off": {
          "@_x": emu(opts.x ?? 1),
          "@_y": emu(opts.y ?? 1),
        },
        "a:ext": {
          "@_cx": emu(opts.w ?? 4),
          "@_cy": emu(opts.h ?? 1),
        },
      },
    },
    "p:txBody": {
      "a:p": [
        {
          "a:r": [
            {
              "a:t": opts.text ?? "Hello",
              "a:rPr": {
                ...(opts.bold ? { "@_b": "1" } : {}),
                ...(opts.italic ? { "@_i": "1" } : {}),
                ...(opts.fontSize
                  ? { "@_sz": String(opts.fontSize * 100) }
                  : {}),
                ...(opts.color
                  ? { "a:solidFill": { "a:srgbClr": { "@_val": opts.color } } }
                  : {}),
              },
            },
          ],
        },
      ],
    },
  };
}

describe("extractTextElement", () => {
  test("extracts basic text element", () => {
    const sp = makeSpNode({ text: "Hello World", x: 1, y: 2, w: 5, h: 1.5 });
    const el = extractTextElement(sp, themeColorMap);
    expect(el).not.toBeNull();
    expect(el!.type).toBe("text");
    expect(el!.x).toBeCloseTo(1);
    expect(el!.y).toBeCloseTo(2);
    expect(el!.w).toBeCloseTo(5);
    expect(el!.h).toBeCloseTo(1.5);
    expect(el!.textRuns[0].text).toBe("Hello World");
  });

  test("extracts bold text", () => {
    const sp = makeSpNode({ text: "Bold", bold: true });
    const el = extractTextElement(sp, themeColorMap);
    expect(el!.textRuns[0].bold).toBe(true);
  });

  test("extracts italic text", () => {
    const sp = makeSpNode({ text: "Italic", italic: true });
    const el = extractTextElement(sp, themeColorMap);
    expect(el!.textRuns[0].italic).toBe(true);
  });

  test("extracts font size in points", () => {
    // sz=2400 → 24pt
    const sp = makeSpNode({ fontSize: 24 });
    const el = extractTextElement(sp, themeColorMap);
    expect(el!.textRuns[0].fontSize).toBe(24);
  });

  test("extracts color", () => {
    const sp = makeSpNode({ color: "FF0000" });
    const el = extractTextElement(sp, themeColorMap);
    expect(el!.textRuns[0].color).toBe("FF0000");
  });

  test("returns null when no txBody", () => {
    const sp = { "p:spPr": { "a:xfrm": {} } };
    const el = extractTextElement(sp, themeColorMap);
    expect(el).toBeNull();
  });
});
