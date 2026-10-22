export type Severity = "safe" | "warning" | "danger";

export interface FontUsage {
  family: string;
  installed: boolean;
  embedded?: boolean;
  occurrences: number;
  slides: number[];
}

export interface Risk {
  id: string;
  slideNumber: number;
  severity: Exclude<Severity, "safe">;
  category: "missing-font" | "text-density" | "text-box" | "table" | "unsupported-effect" | "small-text" | "long-token" | "edge-proximity" | "media" | "visual-difference";
  title: string;
  detail: string;
  fontFamily?: string;
}

export interface SlideAnalysis {
  number: number;
  title: string;
  textPreview: string;
  score: number;
  severity: Severity;
  risks: Risk[];
}

export interface PresentationAnalysis {
  fileName: string;
  slideCount: number;
  score: number;
  fonts: FontUsage[];
  slides: SlideAnalysis[];
  analyzedAt: string;
  sourceType?: "pptx" | "public-link";
  confidence?: "preflight" | "exact";
  scannedSlideCount?: number;
  scanNotice?: string;
}

export interface PickedFile {
  name: string;
  path: string;
  bytes: Uint8Array;
}

export interface SlideComparison {
  slideNumber: number;
  canvaImage: string;
  powerpointImage: string;
  differencePercent: number;
  edgeDifferencePercent: number;
}

declare global {
  interface Window {
    hailCanva?: {
      pickPptx(): Promise<PickedFile | null>;
      listFonts(): Promise<string[]>;
      savePptx(name: string, bytes: Uint8Array): Promise<string | null>;
      canvaStatus(): Promise<{ connected: boolean; configured: boolean; expiresAt?: number }>;
      connectCanva(): Promise<{ connected: boolean; configured: boolean; expiresAt?: number }>;
      downloadCanvaPptx(link: string): Promise<PickedFile>;
      inspectPublicCanva(link: string): Promise<{
        access: "public" | "private" | "unreachable";
        download: "pptx" | "disabled" | "failed" | "consent_required";
        message: string;
        file?: PickedFile;
      }>;
      openCanvaConsent(link: string): Promise<void>;
      analyzePublicCanva(link: string): Promise<PresentationAnalysis>;
      compareCanvaSlides(link: string, bytes: Uint8Array): Promise<SlideComparison[]>;
    };
  }
}
