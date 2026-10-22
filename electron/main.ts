import "dotenv/config";
import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { connectCanva, connectionStatus, downloadCanvaPptx } from "./canva";
import { inspectPublicCanvaLink, openCanvaConsentWindow } from "./publicCanva";

const execFileAsync = promisify(execFile);

function createWindow() {
  const window = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 1050,
    minHeight: 680,
    backgroundColor: "#f5f5f0",
    titleBarStyle: "hiddenInset",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  const devUrl = process.env.VITE_DEV_SERVER_URL;
  if (devUrl) void window.loadURL(devUrl);
  else void window.loadFile(path.join(__dirname, "../dist/index.html"));
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

ipcMain.handle("file:pickPptx", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "PowerPoint", extensions: ["pptx"] }],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const filePath = result.filePaths[0];
  const bytes = await fs.readFile(filePath);
  return { name: path.basename(filePath), path: filePath, bytes };
});

ipcMain.handle("font:list", async () => {
  if (process.platform !== "win32") return [];
  const script = [
    "$fonts = New-Object System.Drawing.Text.InstalledFontCollection",
    "$fonts.Families | ForEach-Object { $_.Name } | Sort-Object -Unique | ConvertTo-Json -Compress",
  ].join("; ");
  try {
    const { stdout } = await execFileAsync("powershell.exe", ["-NoProfile", "-Command", script], {
      windowsHide: true,
      maxBuffer: 4 * 1024 * 1024,
    });
    const parsed = JSON.parse(stdout.trim() || "[]");
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
});

ipcMain.handle("file:savePptx", async (_event, name: string, bytes: Uint8Array) => {
  const result = await dialog.showSaveDialog({
    defaultPath: name.replace(/\.pptx$/i, "_fixed.pptx"),
    filters: [{ name: "PowerPoint", extensions: ["pptx"] }],
  });
  if (result.canceled || !result.filePath) return null;
  await fs.writeFile(result.filePath, Buffer.from(bytes));
  return result.filePath;
});

ipcMain.handle("canva:status", () => connectionStatus());
ipcMain.handle("canva:connect", () => connectCanva());
ipcMain.handle("canva:downloadPptx", (_event, link: string) => downloadCanvaPptx(link));
ipcMain.handle("canva:inspectPublic", (_event, link: string) => inspectPublicCanvaLink(link));
ipcMain.handle("canva:openConsent", (_event, link: string) => openCanvaConsentWindow(link));
