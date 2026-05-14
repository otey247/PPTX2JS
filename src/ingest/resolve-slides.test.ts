import { resolveSlidePaths, resolveSlideRelationships, normalizeTarget } from "../ingest/resolve-slides";

const samplePresentation = {
  "p:presentation": {
    "p:sldIdLst": {
      "p:sldId": [
        { "@_id": "256", "@_r:id": "rId2" },
        { "@_id": "257", "@_r:id": "rId3" },
      ],
    },
  },
};

const sampleRels = {
  Relationships: {
    Relationship: [
      {
        "@_Id": "rId2",
        "@_Type":
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        "@_Target": "slides/slide1.xml",
      },
      {
        "@_Id": "rId3",
        "@_Type":
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        "@_Target": "slides/slide2.xml",
      },
      {
        "@_Id": "rId1",
        "@_Type":
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        "@_Target": "theme/theme1.xml",
      },
    ],
  },
};

describe("normalizeTarget", () => {
  test("handles relative ../ prefix", () => {
    expect(normalizeTarget("../slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });

  test("handles absolute path starting with /ppt/", () => {
    expect(normalizeTarget("/ppt/slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });

  test("handles path already starting with ppt/", () => {
    expect(normalizeTarget("ppt/slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });

  test("prepends ppt/ for bare relative paths", () => {
    expect(normalizeTarget("slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });
});

describe("resolveSlidePaths", () => {
  test("returns ordered slide paths", () => {
    const paths = resolveSlidePaths(samplePresentation, sampleRels);
    expect(paths).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
  });

  test("handles absolute Target paths (e.g. /ppt/slides/slide1.xml)", () => {
    const absRels = {
      Relationships: {
        Relationship: [
          {
            "@_Id": "rId2",
            "@_Type":
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            "@_Target": "/ppt/slides/slide1.xml",
          },
          {
            "@_Id": "rId3",
            "@_Type":
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            "@_Target": "/ppt/slides/slide2.xml",
          },
        ],
      },
    };
    const paths = resolveSlidePaths(samplePresentation, absRels);
    expect(paths).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
  });

  test("throws when relationship is missing", () => {
    const badRels = { Relationships: { Relationship: [] } };
    expect(() => resolveSlidePaths(samplePresentation, badRels)).toThrow();
  });

  test("returns empty array when no slides", () => {
    const emptyPres = { "p:presentation": { "p:sldIdLst": {} } };
    const paths = resolveSlidePaths(emptyPres, sampleRels);
    expect(paths).toEqual([]);
  });
});

describe("resolveSlideRelationships", () => {
  const slideRels = {
    Relationships: {
      Relationship: [
        {
          "@_Id": "rId1",
          "@_Type":
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
          "@_Target": "../slideLayouts/slideLayout1.xml",
        },
        {
          "@_Id": "rId2",
          "@_Type":
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "@_Target": "../media/image1.png",
        },
        {
          "@_Id": "rId3",
          "@_Type":
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
          "@_Target": "../notesSlides/notesSlide1.xml",
        },
        {
          "@_Id": "rId4",
          "@_Type":
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
          "@_Target": "https://example.com/page",
        },
      ],
    },
  };

  test("extracts layout path", () => {
    const { layoutPath } = resolveSlideRelationships(slideRels);
    expect(layoutPath).toBe("ppt/slideLayouts/slideLayout1.xml");
  });

  test("extracts notes path", () => {
    const { notesPath } = resolveSlideRelationships(slideRels);
    expect(notesPath).toBe("ppt/notesSlides/notesSlide1.xml");
  });

  test("extracts image paths map", () => {
    const { imagePaths } = resolveSlideRelationships(slideRels);
    expect(imagePaths.get("rId2")).toBe("ppt/media/image1.png");
  });

  test("extracts hyperlink URL map (resolves actual URL, not r:id)", () => {
    const { hyperlinkUrls } = resolveSlideRelationships(slideRels);
    expect(hyperlinkUrls.get("rId4")).toBe("https://example.com/page");
  });

  test("handles null input", () => {
    const { layoutPath, imagePaths, hyperlinkUrls } = resolveSlideRelationships(null);
    expect(layoutPath).toBeUndefined();
    expect(imagePaths.size).toBe(0);
    expect(hyperlinkUrls.size).toBe(0);
  });
});
