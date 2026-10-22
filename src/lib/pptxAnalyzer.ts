import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type { FontUsage, PresentationAnalysis, Risk, Severity, SlideAnalysis } from "../types";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
});

const GENERIC_FONT_NAMES = new Set(["+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea", "+mj-cs", "+mn-cs"]);
const PORTABLE_FONTS = new Set([
  "arial", "calibri", "times new roman", "verdana", "tahoma", "trebuchet ms",
  "맑은 고딕", "malgun gothic", "noto sans kr", "나눔고딕", "nanumgothic",
]);

function arrayify<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function walk(value: unknown, visit: (key: string, node: unknown) => void): void {
  if (!value || typeof value !== "object") return;
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    visit(key, child);
    if (Array.isArray(child)) child.forEach((item) => walk(item, visit));
    else walk(child, visit);
  }
}

function getTexts(doc: unknown): string[] {
  const texts: string[] = [];
  walk(doc, (key, value) => {
    if (key === "a:t" && typeof value === "string") texts.push(value);
  });
  return texts;
}

function getFonts(doc: unknown): string[] {
  const fonts: string[] = [];
  walk(doc, (key, value) => {
    if (!["a:latin", "a:ea", "a:cs", "a:sym", "a:rPr", "a:defRPr", "a:endParaRPr"].includes(key) || !value || typeof value !== "object") return;
    const attrs = value as Record<string, unknown>;
    const font = attrs["@_typeface"];
    if (typeof font === "string" && font.trim() && !GENERIC_FONT_NAMES.has(font)) fonts.push(font.trim());
  });
  return fonts;
}

function normalizeFontName(font: string) {
  return font.normalize("NFKC").replace(/^@/, "").replace(/\s+/g, " ").trim().toLocaleLowerCase();
}

function getEmbeddedFonts(doc: unknown): string[] {
  const fonts: string[] = [];
  walk(doc, (key, value) => {
    if (key !== "p:embeddedFont" || !value || typeof value !== "object") return;
    for (const font of arrayify((value as Record<string, unknown>)["p:font"])) {
      if (!font || typeof font !== "object") continue;
      const typeface = (font as Record<string, unknown>)["@_typeface"];
      if (typeof typeface === "string" && typeface.trim()) fonts.push(typeface.trim());
    }
  });
  return fonts;
}

function getTextBoxDensityRisks(doc: unknown, slideNumber: number): Risk[] {
  const risks: Risk[] = [];
  let index = 0;
  walk(doc, (key, value) => {
    if (key !== "p:sp" || !value || typeof value !== "object") return;
    const shape = value as Record<string, unknown>;
    const texts = getTexts(shape);
    const text = texts.join("");
    if (text.length < 70) return;
    index += 1;
    const lineBreaks = texts.length;
    if (text.length > 180 || lineBreaks > 8) {
      risks.push({
        id: `density-${slideNumber}-${index}`,
        slideNumber,
        severity: text.length > 300 ? "danger" : "warning",
        category: "text-density",
        title: "텍스트 상자가 빽빽합니다",
        detail: `${text.length}자의 텍스트가 한 개체에 들어 있습니다. 글꼴 대체 시 줄바꿈이나 넘침이 발생할 수 있습니다.`,
      });
    }
  });
  return risks;
}

function getEffectRisks(xml: string, slideNumber: number): Risk[] {
  const patterns: Array<[RegExp, string]> = [
    [/<a:glow\b/, "광선 효과"],
    [/<a:reflection\b/, "반사 효과"],
    [/<a:prstTxWarp\b/, "변형 텍스트"],
    [/<p:video\b|<a:videoFile\b/, "비디오"],
  ];
  return patterns.flatMap(([pattern, label], index) => pattern.test(xml) ? [{
    id: `effect-${slideNumber}-${index}`,
    slideNumber,
    severity: "warning" as const,
    category: "unsupported-effect" as const,
    title: `${label} 호환성을 확인하세요`,
    detail: `${label}는 Canva와 PowerPoint에서 다르게 표시되거나 정적으로 변환될 수 있습니다.`,
  }] : []);
}

function getLayoutRisks(xml: string, slideNumber: number, slideWidth: number, slideHeight: number): Risk[] {
  const risks: Risk[] = [];
  if (/vertOverflow="clip"/i.test(xml)) {
    risks.push({
      id: `text-box-clip-${slideNumber}`,
      slideNumber,
      severity: "warning",
      category: "text-box",
      title: "잘림 설정이 있는 텍스트 상자",
      detail: "PowerPoint에서 글꼴이나 줄바꿈이 달라지면 상자 밖의 텍스트가 숨겨질 수 있습니다.",
    });
  }
  const frames = xml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g) ?? [];
  frames.forEach((frame, index) => {
    if (!/<a:tbl\b/.test(frame)) return;
    const transform = frame.match(/<a:off x="(\d+)" y="(\d+)"\s*\/>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/);
    if (!transform || !slideWidth || !slideHeight) return;
    const [, x, y, width, height] = transform.map(Number);
    const rightGap = slideWidth - (x + width);
    const bottomGap = slideHeight - (y + height);
    const outside = rightGap < 0 || bottomGap < 0 || x < 0 || y < 0;
    const nearEdge = rightGap < slideWidth * 0.015 || bottomGap < slideHeight * 0.015 || x < slideWidth * 0.015 || y < slideHeight * 0.015;
    if (!outside && !nearEdge) return;
    risks.push({
      id: `table-edge-${slideNumber}-${index}`,
      slideNumber,
      severity: outside ? "danger" : "warning",
      category: "table",
      title: outside ? "표가 슬라이드 영역을 벗어납니다" : "표가 슬라이드 가장자리에 가깝습니다",
      detail: outside ? "PowerPoint 변환 과정에서 표의 행이나 열이 잘릴 가능성이 높습니다." : "렌더링 차이로 표 테두리나 셀 내용이 잘리지 않는지 비교 이미지를 확인하세요.",
    });
  });
  return risks;
}

function severityFor(score: number): Severity {
  if (score >= 75) return "safe";
  if (score >= 45) return "warning";
  return "danger";
}

export async function analyzePptx(bytes: Uint8Array, fileName: string, installedFonts: string[]): Promise<PresentationAnalysis> {
  const zip = await JSZip.loadAsync(bytes);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
  const slideSize = presentationXml?.match(/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  const slideWidth = Number(slideSize?.[1] || 0);
  const slideHeight = Number(slideSize?.[2] || 0);
  const embedded = new Set(presentationXml ? getEmbeddedFonts(parser.parse(presentationXml)).map(normalizeFontName) : []);
  const installed = new Set(installedFonts.map(normalizeFontName));
  const usage = new Map<string, { count: number; slides: Set<number> }>();
  const slides: SlideAnalysis[] = [];

  for (let i = 0; i < slideFiles.length; i += 1) {
    const slideNumber = i + 1;
    const xml = await zip.file(slideFiles[i])!.async("string");
    const doc = parser.parse(xml);
    const texts = getTexts(doc);
    const fonts = getFonts(doc);
    const risks: Risk[] = [...getTextBoxDensityRisks(doc, slideNumber), ...getEffectRisks(xml, slideNumber), ...getLayoutRisks(xml, slideNumber, slideWidth, slideHeight)];

    for (const font of fonts) {
      const current = usage.get(font) ?? { count: 0, slides: new Set<number>() };
      current.count += 1;
      current.slides.add(slideNumber);
      usage.set(font, current);
    }

    for (const font of new Set(fonts)) {
      const normalized = normalizeFontName(font);
      if (!installed.has(normalized) && !embedded.has(normalized)) {
        risks.push({
          id: `font-${slideNumber}-${font}`,
          slideNumber,
          severity: PORTABLE_FONTS.has(normalized) ? "warning" : "danger",
          category: "missing-font",
          title: `‘${font}’ 글꼴을 찾을 수 없습니다`,
          detail: "현재 PC에서 대체 글꼴이 사용되어 글자 폭과 줄바꿈이 달라질 수 있습니다.",
          fontFamily: font,
        });
      }
    }

    const penalty = risks.reduce((sum, risk) => sum + (risk.severity === "danger" ? 28 : 12), 0);
    const score = Math.max(0, 100 - penalty);
    slides.push({
      number: slideNumber,
      title: texts[0]?.trim() || `슬라이드 ${slideNumber}`,
      textPreview: texts.join(" ").slice(0, 150),
      score,
      severity: severityFor(score),
      risks,
    });
  }

  const fonts: FontUsage[] = [...usage.entries()].map(([family, data]) => ({
    family,
    installed: installed.has(normalizeFontName(family)) || embedded.has(normalizeFontName(family)),
    embedded: embedded.has(normalizeFontName(family)),
    occurrences: data.count,
    slides: [...data.slides].sort((a, b) => a - b),
  })).sort((a, b) => Number(a.installed) - Number(b.installed) || b.occurrences - a.occurrences);
  const score = slides.length ? Math.round(slides.reduce((sum, slide) => sum + slide.score, 0) / slides.length) : 100;

  return { fileName, slideCount: slides.length, score, fonts, slides, analyzedAt: new Date().toISOString() };
}

export async function replaceFont(bytes: Uint8Array, from: string, to: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const xmlFiles = Object.keys(zip.files).filter((name) => name.endsWith(".xml"));
  const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`([\"'])${escaped}\\1`, "g");
  await Promise.all(xmlFiles.map(async (name) => {
    const entry = zip.file(name);
    if (!entry) return;
    const xml = await entry.async("string");
    const updated = xml.replace(pattern, (_match, quote: string) => `${quote}${to}${quote}`);
    if (updated !== xml) zip.file(name, updated);
  }));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export async function replaceFontOnSlide(bytes: Uint8Array, from: string, to: string, slideNumber: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const name = `ppt/slides/slide${slideNumber}.xml`;
  const entry = zip.file(name);
  if (!entry) throw new Error(`${slideNumber}번 슬라이드를 찾지 못했습니다.`);
  const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(["'])${escaped}\\1`, "g");
  const xml = await entry.async("string");
  zip.file(name, xml.replace(pattern, (_match, quote: string) => `${quote}${to}${quote}`));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
