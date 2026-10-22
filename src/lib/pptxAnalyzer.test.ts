import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { analyzePptx, replaceFont } from "./pptxAnalyzer";

async function fixture(font = "Rare Canva Font") {
  const zip = new JSZip();
  zip.file("ppt/slides/slide1.xml", `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:rPr><a:latin typeface="${font}"/><a:ea typeface="${font}"/></a:rPr><a:t>Hello Canva</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: "uint8array" });
}

async function embeddedFixture(font = "Tlab 월광소나타") {
  const zip = new JSZip();
  zip.file("ppt/slides/slide1.xml", `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:rPr><a:latin typeface="${font}"/></a:rPr><a:t>내장 글꼴</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  zip.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:r="r"><p:embeddedFontLst><p:embeddedFont><p:font typeface="${font}"/><p:regular r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>`);
  zip.file("ppt/fonts/font1.fntdata", new Uint8Array([1, 2, 3]));
  return zip.generateAsync({ type: "uint8array" });
}

describe("pptx analyzer", () => {
  it("finds a missing font per slide", async () => {
    const result = await analyzePptx(await fixture(), "sample.pptx", ["Arial"]);
    expect(result.slideCount).toBe(1);
    expect(result.fonts[0].family).toBe("Rare Canva Font");
    expect(result.slides[0].risks[0].category).toBe("missing-font");
  });

  it("replaces fonts without rebuilding the presentation", async () => {
    const updated = await replaceFont(await fixture(), "Rare Canva Font", "Arial");
    const result = await analyzePptx(updated, "sample.pptx", ["Arial"]);
    expect(result.fonts[0].family).toBe("Arial");
    expect(result.slides[0].risks).toHaveLength(0);
  });

  it("does not report a font embedded in the presentation as missing", async () => {
    const result = await analyzePptx(await embeddedFixture(), "embedded.pptx", []);
    expect(result.fonts[0]).toMatchObject({ family: "Tlab 월광소나타", installed: true, embedded: true });
    expect(result.slides[0].risks).not.toEqual(expect.arrayContaining([expect.objectContaining({ category: "missing-font" })]));
  });
});
