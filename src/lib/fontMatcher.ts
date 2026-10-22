const groups = [
  { keywords: ["명조", "바탕", "serif", "times", "myeongjo"], preferred: ["Noto Serif KR", "나눔명조", "바탕", "함초롬바탕", "Times New Roman"] },
  { keywords: ["손글씨", "캘리", "붓", "펜", "소나타", "script", "hand", "calligraphy"], preferred: ["나눔손글씨 펜", "휴먼편지체", "함초롬돋움", "궁서", "Noto Serif KR"] },
  { keywords: ["고딕", "돋움", "sans", "gothic", "grotesk"], preferred: ["Noto Sans KR", "맑은 고딕", "나눔고딕", "Arial", "Calibri"] },
  { keywords: ["mono", "code", "고정폭"], preferred: ["D2Coding", "Consolas", "Courier New"] },
];

function normalized(value: string) {
  return value.normalize("NFKC").replace(/\s+/g, " ").trim().toLocaleLowerCase();
}

export function recommendFonts(source: string, installedFonts: string[], limit = 5) {
  const installed = new Map(installedFonts.map((font) => [normalized(font), font]));
  const sourceName = normalized(source);
  const group = groups.find((item) => item.keywords.some((keyword) => sourceName.includes(keyword)));
  const preferred = [...(group?.preferred ?? []), "Noto Sans KR", "맑은 고딕", "나눔고딕", "Arial", "Calibri"];
  const results: string[] = [];
  for (const candidate of preferred) {
    const actual = installed.get(normalized(candidate));
    if (actual && normalized(actual) !== sourceName && !results.includes(actual)) results.push(actual);
    if (results.length >= limit) return results;
  }
  for (const font of installedFonts) {
    if (normalized(font) !== sourceName && !results.includes(font)) results.push(font);
    if (results.length >= limit) break;
  }
  return results;
}
