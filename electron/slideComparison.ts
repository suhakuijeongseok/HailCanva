import { app } from "electron";
import { execFile } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { downloadCanvaPageImages } from "./canva";

const execFileAsync = promisify(execFile);

function pngDataUrl(bytes: Uint8Array) {
  return `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
}

export async function renderPowerPoint(bytes: Uint8Array) {
  if (process.platform !== "win32") throw new Error("PowerPoint 이미지 비교는 Windows에서만 지원합니다.");
  const tempRoot = path.resolve(app.getPath("temp"));
  const directory = await fs.mkdtemp(path.join(tempRoot, "hailcanva-compare-"));
  const resolvedDirectory = path.resolve(directory);
  if (path.dirname(resolvedDirectory) !== tempRoot || !path.basename(resolvedDirectory).startsWith("hailcanva-compare-")) {
    throw new Error("안전한 임시 폴더를 만들지 못했습니다.");
  }
  const input = path.join(directory, `${crypto.randomUUID()}.pptx`);
  const output = path.join(directory, "rendered");
  await fs.mkdir(output);
  await fs.writeFile(input, bytes);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$powerpoint = New-Object -ComObject PowerPoint.Application",
    "$presentation = $null",
    "try {",
    "  $presentation = $powerpoint.Presentations.Open($env:HAILCANVA_PPTX_INPUT, $true, $false, $false)",
    "  $width = 1600",
    "  $height = [Math]::Round($width * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth)",
    "  for ($index = 1; $index -le $presentation.Slides.Count; $index++) {",
    "    $target = Join-Path $env:HAILCANVA_PPTX_OUTPUT ($index.ToString('D4') + '.png')",
    "    $presentation.Slides.Item($index).Export($target, 'PNG', $width, $height)",
    "  }",
    "} finally {",
    "  if ($presentation) { $presentation.Close() }",
    "  $powerpoint.Quit()",
    "}",
  ].join("\n");
  try {
    await execFileAsync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script], {
      windowsHide: true,
      timeout: 120_000,
      maxBuffer: 4 * 1024 * 1024,
      env: { ...process.env, HAILCANVA_PPTX_INPUT: input, HAILCANVA_PPTX_OUTPUT: output },
    });
    const files = (await fs.readdir(output)).filter((name) => name.endsWith(".png")).sort();
    return Promise.all(files.map(async (name) => pngDataUrl(await fs.readFile(path.join(output, name)))));
  } catch (error) {
    throw new Error(`PowerPoint로 슬라이드를 렌더링하지 못했습니다. ${error instanceof Error ? error.message : ""}`.trim());
  } finally {
    await fs.rm(resolvedDirectory, { recursive: true, force: true });
  }
}

export async function createSlideComparisons(link: string, pptxBytes: Uint8Array) {
  const [canvaImages, powerpointImages] = await Promise.all([
    downloadCanvaPageImages(link),
    renderPowerPoint(pptxBytes),
  ]);
  const count = Math.min(canvaImages.length, powerpointImages.length);
  return Array.from({ length: count }, (_, index) => ({
    slideNumber: index + 1,
    canvaImage: pngDataUrl(canvaImages[index]),
    powerpointImage: powerpointImages[index],
  }));
}
