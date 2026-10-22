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
    const nodes = [...document.querySelectorAll('button, a, [role="button"], [role="menuitem"], [role="option"]')];
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

function watchForPptxDownload(downloadSession: Electron.Session, destination: string) {
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
    const timeout = setTimeout(() => finish(null), 45_000);
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
  if (!clicked) return false;
  await wait(1_500);
  await clickMatching(window, ["microsoft powerpoint", "powerpoint", "pptx", "파워포인트"]);
  await wait(900);
  await clickMatching(window, ["^download$", "^다운로드$", "download.*pages", "페이지.*다운로드"]);
  return true;
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

export async function openCanvaConsentWindow(link: string): Promise<void> {
  const resolvedLink = await resolveCanvaLink(link);
  const window = new BrowserWindow({
    show: true,
    width: 1180,
    height: 820,
    title: "HailCanva - Canva 동의 확인",
    autoHideMenuBar: true,
    webPreferences: {
      session: publicSession(),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  await window.loadURL(resolvedLink);
  await wait(2_000);
  await dismissCookieBanner(window);
  window.show();
  window.focus();
  await new Promise<void>((resolve) => window.once("closed", () => resolve()));
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
    await wait(5_000);
    await dismissCookieBanner(window);
    const signals = await pageSignals(window);
    const normalized = `${signals.url}\n${signals.title}\n${signals.text}`.toLocaleLowerCase();
    const denied = ["request access", "access denied", "you need permission", "권한 요청", "액세스 권한", "페이지를 찾을 수 없습니다", "page not found"]
      .some((phrase) => normalized.includes(phrase));
    const loginRedirect = /\/login(?:[/?#]|$)/i.test(new URL(signals.url).pathname);
    if (denied || loginRedirect || signals.visualCount === 0) {
      return {
        access: "private",
        download: "disabled",
        message: "로그인이나 접근 권한이 필요한 Canva 링크입니다. 소유자가 공개 보기·다운로드를 허용해야 합니다.",
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
