import crypto from "node:crypto";
import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import { app, safeStorage, shell } from "electron";
import JSZip from "jszip";

const API_BASE = "https://api.canva.com/rest/v1";
const AUTHORIZE_URL = "https://www.canva.com/api/oauth/authorize";

type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
};

type ExportJob = {
  id: string;
  status: "in_progress" | "success" | "failed";
  urls?: string[];
  error?: { code?: string; message?: string };
};

export type CanvaConnection = {
  connected: boolean;
  expiresAt?: number;
  configured: boolean;
};

let accessToken: string | null = null;
let refreshToken: string | null = null;
let tokenExpiresAt = 0;
let refreshPromise: Promise<void> | null = null;
const resolvedLinkCache = new Map<string, string>();

function tokenPath() {
  return path.join(app.getPath("userData"), "canva-session.bin");
}

async function persistRefreshToken() {
  if (!refreshToken) {
    await fs.rm(tokenPath(), { force: true });
    return;
  }
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error("Windows 보안 저장소를 사용할 수 없어 Canva 연결을 저장하지 못했습니다.");
  }
  const encrypted = safeStorage.encryptString(JSON.stringify({ refreshToken }));
  const target = tokenPath();
  const temporary = `${target}.tmp`;
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(temporary, encrypted);
  await fs.rename(temporary, target);
}

async function clearStoredToken() {
  accessToken = null;
  refreshToken = null;
  tokenExpiresAt = 0;
  await fs.rm(tokenPath(), { force: true });
}

export async function initializeCanvaAuth() {
  try {
    if (!safeStorage.isEncryptionAvailable()) return;
    const encrypted = await fs.readFile(tokenPath());
    const stored = JSON.parse(safeStorage.decryptString(encrypted)) as { refreshToken?: string };
    refreshToken = stored.refreshToken || null;
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code !== "ENOENT") await clearStoredToken();
    return;
  }
  if (refreshToken) await refreshAccessToken().catch(() => undefined);
}

function config() {
  const clientId = process.env.CANVA_CLIENT_ID?.trim();
  const clientSecret = process.env.CANVA_CLIENT_SECRET?.trim();
  const port = Number(process.env.CANVA_OAUTH_PORT || 47821);
  const redirectUri = process.env.CANVA_REDIRECT_URI?.trim() || `http://127.0.0.1:${port}/oauth/callback`;
  return { clientId, clientSecret, port, redirectUri };
}

export function connectionStatus(): CanvaConnection {
  const { clientId, clientSecret } = config();
  return {
    connected: Boolean(refreshToken || (accessToken && tokenExpiresAt > Date.now() + 30_000)),
    expiresAt: tokenExpiresAt || undefined,
    configured: Boolean(clientId && clientSecret),
  };
}

async function requestToken(body: URLSearchParams): Promise<TokenResponse & { message?: string }> {
  const { clientId, clientSecret } = config();
  if (!clientId || !clientSecret) throw new Error("CANVA_CLIENT_ID와 CANVA_CLIENT_SECRET 환경변수를 먼저 설정해 주세요.");
  const response = await fetch(`${API_BASE}/oauth/token`, {
    method: "POST",
    headers: {
      Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString("base64")}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });
  const token = await response.json() as TokenResponse & { message?: string };
  if (!response.ok || !token.access_token) {
    const error = new Error(token.message || "Canva 토큰을 발급받지 못했습니다.") as Error & { status?: number };
    error.status = response.status;
    throw error;
  }
  return token;
}

async function applyToken(token: TokenResponse) {
  accessToken = token.access_token;
  tokenExpiresAt = Date.now() + token.expires_in * 1000;
  if (token.refresh_token) refreshToken = token.refresh_token;
  await persistRefreshToken();
}

async function refreshAccessToken() {
  if (refreshPromise) return refreshPromise;
  if (!refreshToken) throw new Error("Canva 계정을 다시 연결해 주세요.");
  refreshPromise = (async () => {
    try {
      const token = await requestToken(new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken! }));
      await applyToken(token);
    } catch (error) {
      const status = (error as Error & { status?: number }).status;
      accessToken = null;
      tokenExpiresAt = 0;
      if (status === 400 || status === 401) await clearStoredToken();
      const guidance = status === 400 || status === 401 ? "다시 연결해 주세요." : "네트워크 연결을 확인한 뒤 다시 시도해 주세요.";
      throw new Error(`저장된 Canva 연결을 갱신하지 못했습니다. ${guidance} ${error instanceof Error ? error.message : ""}`.trim());
    } finally {
      refreshPromise = null;
    }
  })();
  return refreshPromise;
}

async function ensureAccessToken(force = false) {
  if (!force && accessToken && tokenExpiresAt > Date.now() + 60_000) return;
  await refreshAccessToken();
}

function base64Url(bytes: Buffer) {
  return bytes.toString("base64").replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
}

function htmlPage(title: string, message: string, success: boolean) {
  const color = success ? "#2f9e68" : "#d94e3d";
  return `<!doctype html><html lang="ko"><meta charset="utf-8"><title>${title}</title>
  <body style="margin:0;background:#f3f3ed;font-family:system-ui;color:#1d1d19;display:grid;place-items:center;height:100vh">
  <main style="background:white;padding:42px;border-radius:20px;text-align:center;box-shadow:0 20px 60px #0001;max-width:440px">
  <div style="font-size:36px;color:${color}">${success ? "✓" : "!"}</div><h1>${title}</h1><p>${message}</p>
  <p style="color:#777;font-size:13px">이 창을 닫고 HailCanva로 돌아가세요.</p></main></body></html>`;
}

export async function connectCanva(): Promise<CanvaConnection> {
  const { clientId, clientSecret, port, redirectUri } = config();
  if (!clientId || !clientSecret) {
    throw new Error("CANVA_CLIENT_ID와 CANVA_CLIENT_SECRET 환경변수를 먼저 설정해 주세요.");
  }

  const verifier = base64Url(crypto.randomBytes(64));
  const challenge = base64Url(crypto.createHash("sha256").update(verifier).digest());
  const state = base64Url(crypto.randomBytes(32));
  const scopes = "design:meta:read design:content:read";
  const authUrl = new URL(AUTHORIZE_URL);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("scope", scopes);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("client_id", clientId);
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("redirect_uri", redirectUri);

  return new Promise<CanvaConnection>((resolve, reject) => {
    let settled = false;
    const finish = (error?: Error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      server.close();
      if (error) reject(error);
      else resolve(connectionStatus());
    };

    const server = http.createServer(async (request, response) => {
      try {
        const callbackUrl = new URL(request.url || "/", `http://127.0.0.1:${port}`);
        if (callbackUrl.pathname !== "/oauth/callback") {
          response.writeHead(404).end();
          return;
        }
        const returnedState = callbackUrl.searchParams.get("state");
        const code = callbackUrl.searchParams.get("code");
        const oauthError = callbackUrl.searchParams.get("error");
        if (oauthError) throw new Error(`Canva 인증이 취소되었거나 거부되었습니다: ${oauthError}`);
        if (!code || returnedState !== state) throw new Error("Canva 인증 응답을 검증하지 못했습니다.");

        const body = new URLSearchParams({
          grant_type: "authorization_code",
          code,
          code_verifier: verifier,
          redirect_uri: redirectUri,
        });
        const tokenBody = await requestToken(body);
        await applyToken(tokenBody);
        response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        response.end(htmlPage("Canva 연결 완료", "디자인 링크를 바로 검사할 수 있습니다.", true));
        finish();
      } catch (error) {
        const message = error instanceof Error ? error.message : "알 수 없는 인증 오류";
        response.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
        response.end(htmlPage("Canva 연결 실패", message, false));
        finish(new Error(message));
      }
    });

    server.once("error", (error) => finish(new Error(`OAuth 콜백 서버를 열지 못했습니다: ${error.message}`)));
    server.listen(port, "127.0.0.1", () => void shell.openExternal(authUrl.toString()));
    const timeout = setTimeout(() => finish(new Error("Canva 인증 시간이 초과되었습니다.")), 180_000);
  });
}

export function parseCanvaDesignId(input: string): string {
  const trimmed = input.trim();
  let url: URL;
  try { url = new URL(trimmed); } catch { throw new Error("올바른 Canva 링크를 입력해 주세요."); }
  if (url.hostname !== "www.canva.com" && url.hostname !== "canva.com") throw new Error("canva.com 링크만 사용할 수 있습니다.");
  const match = url.pathname.match(/\/(?:design|d)\/([^/?#]+)/i);
  if (!match?.[1]) throw new Error("링크에서 Canva 디자인 ID를 찾지 못했습니다.");
  return decodeURIComponent(match[1]);
}

export async function resolveCanvaLink(input: string): Promise<string> {
  const cached = resolvedLinkCache.get(input.trim());
  if (cached) return cached;
  let url: URL;
  try { url = new URL(input.trim()); } catch { throw new Error("올바른 Canva 링크를 입력해 주세요."); }
  const allowedInput = ["canva.com", "www.canva.com", "canva.link"].includes(url.hostname.toLocaleLowerCase());
  if (!allowedInput || url.protocol !== "https:") throw new Error("canva.com 또는 canva.link의 HTTPS 링크만 사용할 수 있습니다.");
  if (url.hostname.toLocaleLowerCase() !== "canva.link") {
    parseCanvaDesignId(url.toString());
    return url.toString();
  }
  const response = await fetch(url, { method: "HEAD", redirect: "follow", headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36" } });
  const resolved = new URL(response.url);
  if (!["canva.com", "www.canva.com"].includes(resolved.hostname.toLocaleLowerCase())) {
    throw new Error("Canva 단축 링크의 최종 디자인 주소를 확인하지 못했습니다.");
  }
  parseCanvaDesignId(resolved.toString());
  const result = resolved.toString();
  resolvedLinkCache.set(input.trim(), result);
  return result;
}

async function canvaFetch<T>(path: string, init?: RequestInit): Promise<T> {
  await ensureAccessToken();
  let response = await fetch(`${API_BASE}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", ...init?.headers },
  });
  if (response.status === 401) {
    await ensureAccessToken(true);
    response = await fetch(`${API_BASE}${path}`, {
      ...init,
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", ...init?.headers },
    });
  }
  const body = await response.json() as T & { message?: string };
  if (!response.ok) throw new Error(body.message || `Canva API 요청 실패 (${response.status})`);
  return body;
}

async function waitForExport(jobId: string): Promise<ExportJob> {
  for (let attempt = 0; attempt < 90; attempt += 1) {
    const result = await canvaFetch<{ job: ExportJob }>(`/exports/${encodeURIComponent(jobId)}`);
    if (result.job.status === "success") return result.job;
    if (result.job.status === "failed") throw new Error(result.job.error?.message || "Canva PPTX 내보내기에 실패했습니다.");
    await new Promise((resolve) => setTimeout(resolve, 1_000));
  }
  throw new Error("Canva 내보내기 시간이 초과되었습니다.");
}

async function createExport(designId: string, format: Record<string, unknown>) {
  const created = await canvaFetch<{ job: ExportJob }>("/exports", {
    method: "POST",
    body: JSON.stringify({ design_id: designId, format }),
  });
  return created.job.status === "success" ? created.job : waitForExport(created.job.id);
}

async function fetchExportImages(urls: string[]) {
  const images: Uint8Array[] = [];
  for (const url of urls) {
    const response = await fetch(url);
    if (!response.ok) throw new Error("Canva 원본 이미지를 다운로드하지 못했습니다.");
    const bytes = new Uint8Array(await response.arrayBuffer());
    const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    if (!isZip) {
      images.push(bytes);
      continue;
    }
    const zip = await JSZip.loadAsync(bytes);
    const entries = Object.values(zip.files)
      .filter((entry) => !entry.dir && /\.(png|jpe?g)$/i.test(entry.name))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
    for (const entry of entries) images.push(await entry.async("uint8array"));
  }
  return images;
}

export async function downloadCanvaPageImages(link: string) {
  const resolvedLink = await resolveCanvaLink(link);
  const designId = parseCanvaDesignId(resolvedLink);
  const job = await createExport(designId, { type: "png", as_single_image: false });
  if (!job.urls?.length) throw new Error("Canva가 페이지 이미지 주소를 반환하지 않았습니다.");
  return fetchExportImages(job.urls);
}

export async function downloadCanvaPptx(link: string): Promise<{ name: string; path: string; bytes: Uint8Array }> {
  const resolvedLink = await resolveCanvaLink(link);
  const designId = parseCanvaDesignId(resolvedLink);
  const metadata = await canvaFetch<{ design: { title?: string; design_types?: string[] } }>(`/designs/${encodeURIComponent(designId)}`);
  if (metadata.design.design_types && !metadata.design.design_types.includes("presentation")) {
    throw new Error("이 Canva 디자인은 프레젠테이션 형식이 아닙니다.");
  }
  const job = await createExport(designId, { type: "pptx" });
  const downloadUrl = job.urls?.[0];
  if (!downloadUrl) throw new Error("Canva가 PPTX 다운로드 주소를 반환하지 않았습니다.");
  const response = await fetch(downloadUrl);
  if (!response.ok) throw new Error("Canva PPTX 파일을 다운로드하지 못했습니다.");
  const safeTitle = (metadata.design.title || "Canva-presentation").replace(/[<>:"/\\|?*]/g, "_");
  return { name: `${safeTitle}.pptx`, path: link, bytes: new Uint8Array(await response.arrayBuffer()) };
}
