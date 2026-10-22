import { describe, expect, it } from "vitest";
import { recommendFonts } from "./fontMatcher";

describe("font recommendations", () => {
  it("prefers handwriting-like Korean fonts for decorative source fonts", () => {
    const result = recommendFonts("Tlab 월광소나타", ["Arial", "맑은 고딕", "나눔손글씨 펜", "Noto Serif KR"]);
    expect(result[0]).toBe("나눔손글씨 펜");
  });

  it("only returns fonts installed on the computer", () => {
    expect(recommendFonts("Unknown Gothic", ["Arial", "Calibri"])).toEqual(["Arial", "Calibri"]);
  });
});
