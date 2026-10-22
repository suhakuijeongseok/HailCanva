import { BrowserWindow, session } from "electron";
import { resolveCanvaLink } from "./canva";

export type LinkRisk = {
  id: string;
  slideNumber: number;
  severity: "warning" | "danger";
  category: "text-density" | "small-text" | "long-token" | "edge-proximity" | "media";
  title: string;
  detail: string;
};

export type LinkSlideAnalysis = {
  number: number;
  title: string;
  textPreview: string;
  score: number;
  severity: "safe" | "warning" | "danger";
  risks: LinkRisk[];
};

export type PublicDesignAnalysis = {
  fileName: string;
  slideCount: number;
  score: number;
  fonts: Array<{ family: string; installed: boolean; occurrences: number; slides: number[] }>;
  slides: LinkSlideAnalysis[];
  analyzedAt: string;
  sourceType: "public-link";
  confidence: "preflight";
  scannedSlideCount: number;
  scanNotice?: string;
};

type PageSnapshot = {
  page: number;
  activePage: number | null;
  text: string[];
  elements: Array<{
    text: string;
    fontSize: number;
    fontFamily: string;
    nearEdge: boolean;
  }>;
  hasMedia: boolean;
};

const wait = (milliseconds: number) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function clickText(window: BrowserWindow, pattern: string) {
  return window.webContents.executeJavaScript(`(() => {
    const pattern = new RegExp(${JSON.stringify(pattern)}, 'i');
    const node = [...document.querySelectorAll('button,[role="button"]')].find((item) => pattern.test((item.textContent || '').trim()));
    if (!node) return false; node.click(); return true;
  })()`);
}

async function readDocumentInfo(window: BrowserWindow) {
  return window.webContents.executeJavaScript(`(() => {
    const text = document.body?.innerText || '';
    const pageNumbers = [...document.querySelectorAll('[aria-label]')]
      .map((node) => (node.getAttribute('aria-label') || '').match(/^(\\d+)\\s*페이지$/)?.[1])
      .filter(Boolean).map(Number);
    const counter = text.match(/페이지\\s*[  ]*(\\d+)\\s*\\/\\s*(\\d+)/);
    return {
      title: document.title.replace(/\\s*-\\s*프레젠테이션.*$/, '').trim() || 'Canva 프레젠테이션',
      pageCount: Math.max(counter ? Number(counter[2]) : 0, ...pageNumbers, 1),
      denied: /request access|access denied|you need permission|권한 요청|액세스 권한|page not found|페이지를 찾을 수 없습니다/i.test(text),
      login: /\\/login(?:[/?#]|$)/i.test(location.pathname)
    };
  })()` ) as Promise<{ title: string; pageCount: number; denied: boolean; login: boolean }>;
}

async function selectPage(window: BrowserWindow, page: number) {
  const point = await window.webContents.executeJavaScript(`(() => {
    const labels = [${JSON.stringify(`${page} 페이지`)}, ${JSON.stringify(`${page}페이지`)}];
    const candidates = [...document.querySelectorAll('[aria-label]')].filter((node) => labels.includes(node.getAttribute('aria-label')));
    const thumbnail = candidates.find((node) => {
      const rect = node.getBoundingClientRect(); return rect.width > 20 && rect.height > 20;
    }) || candidates[0];
    if (!thumbnail) return null;
    thumbnail.scrollIntoView({ block: 'center', inline: 'nearest' });
    thumbnail.focus();
    const rect = thumbnail.getBoundingClientRect();
    return { x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
  })()`);
  if (!point) return false;
  if (!window.webContents.debugger.isAttached()) window.webContents.debugger.attach("1.3");
  await window.webContents.debugger.sendCommand("Input.dispatchMouseEvent", { type: "mousePressed", x: point.x, y: point.y, button: "left", clickCount: 1 });
  await window.webContents.debugger.sendCommand("Input.dispatchMouseEvent", { type: "mouseReleased", x: point.x, y: point.y, button: "left", clickCount: 1 });
  await window.webContents.debugger.sendCommand("Input.dispatchKeyEvent", { type: "keyDown", key: "Enter", code: "Enter", windowsVirtualKeyCode: 13 });
  await window.webContents.debugger.sendCommand("Input.dispatchKeyEvent", { type: "keyUp", key: "Enter", code: "Enter", windowsVirtualKeyCode: 13 });
  return true;
}

async function readCurrentPage(window: BrowserWindow, page: number): Promise<PageSnapshot> {
  return window.webContents.executeJavaScript(`(() => {
    const page = ${page};
    const exact = new RegExp('^' + page + '\\s*페이지$');
    const pageGroups = [...document.querySelectorAll('[role="group"][aria-label]')]
      .filter((node) => /^(\d+)\s*페이지$/.test(node.getAttribute('aria-label') || ''));
    const pageGroup = pageGroups.find((node) => exact.test(node.getAttribute('aria-label') || ''));
    const appNode = document.querySelector('[role="application"][aria-label*="Canvas" i], [role="application"][aria-label*="캔버스" i]');
    let canvas = appNode;
    while (canvas && canvas.parentElement) {
      const rect = canvas.getBoundingClientRect();
      if (rect.width > 500 && rect.height > 250) break;
      canvas = canvas.parentElement;
    }
    const canvasRect = canvas?.getBoundingClientRect();
    const seen = new Set();
    const elements = canvasRect ? [...canvas.querySelectorAll('span,p')].flatMap((node) => {
      if (node.children.length > 0) return [];
      const value = (node.textContent || '').trim().replace(/\\s+/g, ' ');
      const rect = node.getBoundingClientRect();
      if (!value || rect.width < 1 || rect.height < 1) return [];
      const key = value + ':' + Math.round(rect.x) + ':' + Math.round(rect.y);
      if (seen.has(key)) return []; seen.add(key);
      const style = getComputedStyle(node);
      const margin = Math.max(5, Math.min(canvasRect.width, canvasRect.height) * 0.015);
      return [{
        text: value,
        fontSize: parseFloat(style.fontSize) || 0,
        fontFamily: style.fontFamily || '',
        nearEdge: rect.left - canvasRect.left < margin || canvasRect.right - rect.right < margin || rect.top - canvasRect.top < margin || canvasRect.bottom - rect.bottom < margin
      }];
    }) : [];
    const fallbackText = pageGroup ? (pageGroup.textContent || '').trim() : '';
    const visibleText = [...new Set(elements.map((item) => item.text))];
    return {
      page,
      activePage: pageGroups.map((node) => {
        const rect = node.getBoundingClientRect();
        const match = (node.getAttribute('aria-label') || '').match(/^(\d+)\s*페이지$/);
        return { page: match ? Number(match[1]) : null, area: rect.width * rect.height };
      }).sort((a, b) => b.area - a.area)[0]?.page || null,
      text: visibleText.length ? visibleText : (fallbackText ? [fallbackText] : []),
      elements,
      hasMedia: Boolean(canvas?.querySelector('video, audio'))
    };
  })()`);
}

function analyzeSnapshot(snapshot: PageSnapshot, zoom = 0.5): LinkSlideAnalysis {
  const combined = snapshot.text.join(" ").trim();
  const risks: LinkRisk[] = [];
  const add = (risk: Omit<LinkRisk, "id" | "slideNumber">) => risks.push({ ...risk, id: `${risk.category}-${snapshot.page}-${risks.length}`, slideNumber: snapshot.page });

  if (combined.length > 500) add({ severity: "danger", category: "text-density", title: "텍스트가 매우 많습니다", detail: `${combined.length}자의 텍스트가 있습니다. PowerPoint 변환 시 줄바꿈과 넘침 위험이 큽니다.` });
  else if (combined.length > 260) add({ severity: "warning", category: "text-density", title: "텍스트 밀도가 높습니다", detail: `${combined.length}자의 텍스트가 있습니다. 발표 화면의 가독성과 줄바꿈을 확인하세요.` });

  const smallest = snapshot.elements.length ? Math.min(...snapshot.elements.map((item) => item.fontSize / Math.max(zoom, 0.1))) : 999;
  if (smallest < 13) add({ severity: "danger", category: "small-text", title: "매우 작은 글자가 감지됐습니다", detail: `화면 배율을 보정한 추정 글자 크기가 약 ${smallest.toFixed(1)}px입니다.` });
  else if (smallest < 17) add({ severity: "warning", category: "small-text", title: "작은 글자가 포함되어 있습니다", detail: `화면 배율을 보정한 추정 글자 크기가 약 ${smallest.toFixed(1)}px입니다.` });

  const longToken = combined.split(/\s+/).find((token) => token.length >= 35);
  if (longToken) add({ severity: "warning", category: "long-token", title: "줄바꿈이 어려운 긴 문자열", detail: `“${longToken.slice(0, 45)}${longToken.length > 45 ? "…" : ""}”는 PowerPoint에서 텍스트 상자를 넘을 수 있습니다.` });
  if (snapshot.elements.some((item) => item.nearEdge)) add({ severity: "warning", category: "edge-proximity", title: "슬라이드 가장자리에 가까운 텍스트", detail: "글꼴이나 렌더링 차이로 텍스트가 잘릴 가능성이 있습니다." });
  if (snapshot.hasMedia) add({ severity: "danger", category: "media", title: "동영상 또는 오디오가 포함되어 있습니다", detail: "Canva 미디어는 PowerPoint 내보내기에서 재생되지 않거나 정적 요소로 바뀔 수 있습니다." });

  const penalty = risks.reduce((sum, risk) => sum + (risk.severity === "danger" ? 28 : 12), 0);
  const score = Math.max(0, 100 - penalty);
  return {
    number: snapshot.page,
    title: snapshot.text[0]?.slice(0, 70) || `슬라이드 ${snapshot.page}`,
    textPreview: combined.slice(0, 180),
    score,
    severity: score >= 75 ? "safe" : score >= 45 ? "warning" : "danger",
    risks,
  };
}

export async function analyzePublicCanvaDesign(link: string): Promise<PublicDesignAnalysis> {
  const resolvedLink = await resolveCanvaLink(link);
  const window = new BrowserWindow({ show: false, width: 1440, height: 960, webPreferences: { session: session.fromPartition("persist:hailcanva-public"), sandbox: true, contextIsolation: true } });
  try {
    await window.loadURL(resolvedLink);
    await wait(5_000);
    await clickText(window, "^(모든 쿠키 허용|accept all)");
    await wait(700);
    const info = await readDocumentInfo(window);
    if (info.denied || info.login) throw new Error("공개적으로 접근할 수 없는 Canva 링크입니다.");
    const zoomText = await window.webContents.executeJavaScript(`document.body?.innerText || ''`) as string;
    const zoomMatch = zoomText.match(/(?:^|\n)(\d{1,3})%(?:\n|$)/);
    const zoom = zoomMatch ? Number(zoomMatch[1]) / 100 : 0.5;
    const snapshots: PageSnapshot[] = [];
    let scanNotice: string | undefined;
    for (let page = 1; page <= Math.min(info.pageCount, 100); page += 1) {
      await selectPage(window, page);
      await wait(page === 1 ? 500 : 420);
      const snapshot = await readCurrentPage(window, page);
      if (page > 1 && snapshot.activePage !== page) {
        scanNotice = `Canva가 자동 페이지 전환을 허용하지 않아 전체 ${info.pageCount}장 중 ${snapshots.length}장만 검사했습니다.`;
        break;
      }
      snapshots.push(snapshot);
    }
    const slides = snapshots.map((snapshot) => analyzeSnapshot(snapshot, zoom));
    const fontMap = new Map<string, { count: number; slides: Set<number> }>();
    for (const snapshot of snapshots) for (const element of snapshot.elements) {
      const family = element.fontFamily.split(",")[0]?.replace(/["']/g, "").trim();
      if (!family || /^YAF[A-Za-z0-9_-]+$/i.test(family)) continue;
      const item = fontMap.get(family) ?? { count: 0, slides: new Set<number>() };
      item.count += 1; item.slides.add(snapshot.page); fontMap.set(family, item);
    }
    const fonts = [...fontMap.entries()].map(([family, item]) => ({ family, installed: true, occurrences: item.count, slides: [...item.slides] }));
    const score = slides.length ? Math.round(slides.reduce((sum, slide) => sum + slide.score, 0) / slides.length) : 100;
    return { fileName: info.title, slideCount: info.pageCount, scannedSlideCount: slides.length, scanNotice, score, fonts, slides, analyzedAt: new Date().toISOString(), sourceType: "public-link", confidence: "preflight" };
  } finally {
    if (window.webContents.debugger.isAttached()) window.webContents.debugger.detach();
    if (!window.isDestroyed()) window.destroy();
  }
}
