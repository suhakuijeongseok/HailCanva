import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { app, BrowserWindow, session } from "electron";
import { resolveCanvaLink } from "./canva";

export type PublicCanvaInspection = {
  access: "public" | "private" | "unreachable";
  download: "pptx" | "disabled" | "failed" | "consent_required";
  message: string;
  file?: { name: string; path: string; bytes: Uint8Array };
};

const wait = (milliseconds: number) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function clickMatching(window: BrowserWindow, patterns: string[]) {
  return window.webContents.executeJavaScript(`(() => {
    const patterns = ${JSON.stringify(patterns)}.map((value) => new RegExp(value, "i"));
    const nodes = [...document.querySelectorAll('button, a, select, [role="button"], [role="combobox"], [role="menuitem"], [role="option"]')];
    const visible = (node) => {
      const style = getComputedStyle(node); const rect = node.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    };
    const target = nodes.find((node) => {
      const label = [node.textContent, node.getAttribute('aria-label'), node.getAttribute('title'), node.getAttribute('data-testid')].filter(Boolean).join(' ').trim();
      return visible(node) && patterns.some((pattern) => pattern.test(label));
    });
    if (!target) return false;
    target.click(); return true;
  })()`);
}

async function pageSignals(window: BrowserWindow) {
  return window.webContents.executeJavaScript(`(() => ({
    url: location.href,
    title: document.title,
    text: (document.body?.innerText || '').slice(0, 25000),
    visualCount: document.querySelectorAll('canvas, svg, img').length
  }))()` ) as Promise<{ url: string; title: string; text: string; visualCount: number }>;
}

function watchForPptxDownload(downloadSession: Electron.Session, destination: string, timeoutMs = 45_000) {
  let cancel: () => void = () => {};
  const promise = new Promise<{ path: string; name: string } | null>((resolve) => {
    let completed = false;
    const finish = (value: { path: string; name: string } | null) => {
      if (completed) return;
      completed = true;
      clearTimeout(timeout);
      downloadSession.removeListener("will-download", listener);
      resolve(value);
    };
    const listener = (_event: Electron.Event, item: Electron.DownloadItem) => {
      const originalName = item.getFilename();
      const mime = item.getMimeType().toLowerCase();
      const looksLikePptx = originalName.toLowerCase().endsWith(".pptx") || mime.includes("powerpoint") || mime.includes("presentationml");
      if (!looksLikePptx) return;
      const name = originalName.toLowerCase().endsWith(".pptx") ? originalName : `${originalName}.pptx`;
      const filePath = path.join(destination, `${crypto.randomUUID()}-${name.replace(/[<>:"/\\|?*]/g, "_")}`);
      item.setSavePath(filePath);
      item.once("done", (_doneEvent, state) => finish(state === "completed" ? { path: filePath, name } : null));
    };
    downloadSession.on("will-download", listener);
    const timeout = setTimeout(() => finish(null), timeoutMs);
    cancel = () => finish(null);
  });
  return { promise, cancel };
}

async function attemptDownload(window: BrowserWindow) {
  // Canva changes its menus frequently, so use accessible labels rather than CSS class names.
  let clicked = await clickMatching(window, ["^download$", "^다운로드$", "download.*design", "디자인.*다운로드"]);
  if (!clicked) {
    clicked = await clickMatching(window, ["^share$", "^공유$", "share.*design"]);
    if (clicked) {
      await wait(1_200);
      clicked = await clickMatching(window, ["^download$", "^다운로드$", "download.*design", "디자인.*다운로드"]);
    }
  }
  if (!clicked) {
    clicked = await clickMatching(window, ["^file$", "^파일$"]);
    if (clicked) {
      await wait(1_200);
      clicked = await clickMatching(window, ["^download$", "^다운로드$", "download.*design", "디자인.*다운로드"]);
    }
  }
  if (!clicked) return false;
  await wait(1_500);
  const formatMenuOpened = await clickMatching(window, [
    "^file type$", "^파일 형식$", "file type.*(?:pdf|powerpoint)", "파일 형식.*(?:pdf|파워포인트)",
    "^pdf standard$", "^pdf 표준$", "^microsoft powerpoint$",
  ]);
  if (formatMenuOpened) await wait(700);
  const powerpointSelected = await clickMatching(window, ["microsoft powerpoint", "powerpoint", "pptx", "파워포인트"]);
  if (!powerpointSelected) return false;
  await wait(900);
  return clickMatching(window, ["^download$", "^다운로드$", "download.*pages", "페이지.*다운로드"]);
}

function isAccessDenied(signals: { url: string; text: string }) {
  const normalized = signals.text.toLocaleLowerCase();
  const deniedPhrases = [
    "request access to this design",
    "you need permission to access",
    "you don't have access to this design",
    "access denied",
    "이 디자인에 대한 액세스 권한 요청",
    "이 디자인에 액세스할 수 없습니다",
    "이 디자인을 보려면 권한이 필요합니다",
    "페이지를 찾을 수 없습니다",
    "page not found",
  ];
  return /\/login(?:[/?#]|$)/i.test(new URL(signals.url).pathname)
    || deniedPhrases.some((phrase) => normalized.includes(phrase));
}

async function waitForCanvaPage(window: BrowserWindow, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  let latest = await pageSignals(window);
  while (Date.now() < deadline) {
    latest = await pageSignals(window);
    if (isAccessDenied(latest) || hasConsentGate(latest.text)) return latest;
    const meaningfulText = latest.text.trim().length >= 40;
    if (meaningfulText && latest.visualCount > 0) return latest;
    await wait(500);
  }
  return latest;
}

async function dismissCookieBanner(window: BrowserWindow) {
  const clicked = await clickMatching(window, ["^accept all", "^allow all", "모든 쿠키 허용", "모두 허용", "모두 동의"]);
  if (clicked) await wait(900);
}

function hasConsentGate(text: string) {
  const normalized = text.toLocaleLowerCase();
  const hasTermsTitle = normalized.includes("canva 이용 약관") || normalized.includes("canva terms of use");
  const hasContinueButton = normalized.includes("동의 및 계속하기") || normalized.includes("agree and continue");
  return hasTermsTitle && hasContinueButton;
}

async function waitForConsentGate(window: BrowserWindow, timeoutMs = 6_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const signals = await pageSignals(window);
    if (hasConsentGate(signals.text)) return true;
    await wait(500);
  }
  return false;
}

function publicSession() {
  // App-owned persistent storage keeps only the consent the user grants in the visible handoff window.
  return session.fromPartition("persist:hailcanva-public", { cache: false });
}

export async function openCanvaConsentWindow(link: string): Promise<PublicCanvaInspection["file"] | null> {
  const resolvedLink = await resolveCanvaLink(link);
  const webSession = publicSession();
  const window = new BrowserWindow({
    show: true,
    width: 1180,
    height: 820,
    title: "HailCanva - Canva 동의 확인",
    autoHideMenuBar: true,
    webPreferences: {
      session: webSession,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  const tempDirectory = await fs.mkdtemp(path.join(app.getPath("temp"), "hailcanva-interactive-"));
  const downloadWatcher = watchForPptxDownload(webSession, tempDirectory, 180_000);
  let closed = false;
  const closedPromise = new Promise<null>((resolve) => window.once("closed", () => { closed = true; resolve(null); }));
  try {
    await window.loadURL(resolvedLink);
    await wait(2_000);
    await dismissCookieBanner(window);
    window.show();
    window.focus();

    // Never accept terms or sign in for the user. Once the user has completed
    // those steps (or the editor was already available), drive the export in
    // this same visible, authenticated web session.
    for (let attempt = 0; attempt < 6 && !closed; attempt += 1) {
      const signals = await pageSignals(window);
      const loginPage = /\/login(?:[/?#]|$)/i.test(new URL(signals.url).pathname);
      if (!loginPage && !hasConsentGate(signals.text) && !isAccessDenied(signals)) {
        const initiated = await attemptDownload(window);
        if (initiated) break;
      }
      await Promise.race([wait(2_500), closedPromise]);
    }

    const downloaded = await Promise.race([downloadWatcher.promise, closedPromise]);
    if (!downloaded) return null;
    const bytes = await fs.readFile(downloaded.path);
    if (!window.isDestroyed()) window.destroy();
    return { name: downloaded.name, path: link, bytes };
  } finally {
    downloadWatcher.cancel();
    if (!window.isDestroyed()) window.destroy();
    await fs.rm(tempDirectory, { recursive: true, force: true }).catch(() => undefined);
  }
}

export async function inspectPublicCanvaLink(link: string): Promise<PublicCanvaInspection> {
  const resolvedLink = await resolveCanvaLink(link);
  const isolatedSession = publicSession();
  const window = new BrowserWindow({
    show: false,
    width: 1440,
    height: 960,
    webPreferences: {
      session: isolatedSession,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      images: true,
    },
  });

  const tempDirectory = await fs.mkdtemp(path.join(app.getPath("temp"), "hailcanva-"));
  const downloadWatcher = watchForPptxDownload(isolatedSession, tempDirectory);
  try {
    await window.loadURL(resolvedLink, { userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36" });
    await wait(1_500);
    await dismissCookieBanner(window);
    const signals = await waitForCanvaPage(window);
    if (isAccessDenied(signals)) {
      return {
        access: "private",
        download: "disabled",
        message: "로그인이나 접근 권한이 필요한 Canva 링크입니다. 소유자가 공개 보기·다운로드를 허용해야 합니다.",
      };
    }
    if (signals.text.trim().length < 40 && signals.visualCount === 0) {
      return {
        access: "unreachable",
        download: "failed",
        message: "Canva 디자인 화면이 제한 시간 안에 로드되지 않았습니다. 인터넷 연결을 확인한 뒤 다시 시도해 주세요.",
      };
    }
    if (hasConsentGate(signals.text)) {
      return {
        access: "public",
        download: "consent_required",
        message: "공개 열람은 가능하지만 Canva 이용 약관 동의가 필요합니다. 동의 창에서 직접 확인한 뒤 검사를 다시 실행해 주세요.",
      };
    }

    const initiated = await attemptDownload(window);
    if (!initiated) {
      if (await waitForConsentGate(window)) {
        return {
          access: "public",
          download: "consent_required",
          message: "Canva가 다운로드 전에 이용 약관 동의를 요구합니다. 동의 창에서 직접 확인한 뒤 검사를 다시 실행해 주세요.",
        };
      }
      return {
        access: "public",
        download: "disabled",
        message: "공개 열람은 가능하지만 PPTX 다운로드 버튼을 찾지 못했습니다. 소유자가 다운로드를 허용했는지 확인해 주세요.",
      };
    }
    const downloaded = await downloadWatcher.promise;
    if (!downloaded) {
      if (await waitForConsentGate(window, 2_000)) {
        return {
          access: "public",
          download: "consent_required",
          message: "Canva가 PowerPoint 내보내기 전에 이용 약관 동의를 요구합니다. 동의 창에서 직접 확인한 뒤 검사를 다시 실행해 주세요.",
        };
      }
      return {
        access: "public",
        download: "failed",
        message: "다운로드 메뉴는 열렸지만 PPTX 파일을 받지 못했습니다. PowerPoint 형식 다운로드가 허용되어 있는지 확인해 주세요.",
      };
    }
    const bytes = await fs.readFile(downloaded.path);
    return {
      access: "public",
      download: "pptx",
      message: "공개 Canva 링크에서 PPTX를 내려받았습니다.",
      file: { name: downloaded.name, path: link, bytes },
    };
  } catch (error) {
    return {
      access: "unreachable",
      download: "failed",
      message: error instanceof Error ? `Canva 링크를 열지 못했습니다: ${error.message}` : "Canva 링크를 열지 못했습니다.",
    };
  } finally {
    downloadWatcher.cancel();
    if (!window.isDestroyed()) window.destroy();
    isolatedSession.clearCache().catch(() => undefined);
  }
}
