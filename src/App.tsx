import { useEffect, useMemo, useRef, useState } from "react";
import { AlertTriangle, Check, ChevronRight, ExternalLink, FileUp, FileWarning, Link2, LoaderCircle, LogIn, Save, ShieldCheck, Type, X } from "lucide-react";
import { analyzePptx, replaceFont, replaceFontOnSlide } from "./lib/pptxAnalyzer";
import { recommendFonts } from "./lib/fontMatcher";
import type { PickedFile, PresentationAnalysis, Severity, SlideComparison } from "./types";

const replacementOptions = ["맑은 고딕", "Noto Sans KR", "Arial", "Calibri"];

const problemGroups = [
  { key: "font", label: "글꼴 문제", categories: ["missing-font"] },
  { key: "text", label: "텍스트 상자", categories: ["text-density", "text-box", "small-text", "long-token", "edge-proximity"] },
  { key: "table", label: "표·화면 차이", categories: ["table", "visual-difference"] },
  { key: "effect", label: "효과·미디어", categories: ["unsupported-effect", "media"] },
] as const;

function addVisualRisks(analysis: PresentationAnalysis, comparisons: SlideComparison[]): PresentationAnalysis {
  const slides = analysis.slides.map((slide) => {
    const comparison = comparisons.find((item) => item.slideNumber === slide.number);
    if (!comparison || slide.risks.some((risk) => risk.category === "visual-difference")) return slide;
    const edgeProblem = comparison.edgeDifferencePercent >= 5;
    const generalProblem = comparison.differencePercent >= 3.5;
    if (!edgeProblem && !generalProblem) return slide;
    const risk = {
      id: `visual-difference-${slide.number}`,
      slideNumber: slide.number,
      severity: (comparison.differencePercent >= 10 || comparison.edgeDifferencePercent >= 12 ? "danger" : "warning") as "danger" | "warning",
      category: "visual-difference" as const,
      title: edgeProblem ? "변환 후 가장자리 또는 표 내용이 잘릴 수 있습니다" : "Canva 원본과 PowerPoint 화면 차이가 큽니다",
      detail: `전체 화면의 ${comparison.differencePercent.toFixed(1)}%, 가장자리 영역의 ${comparison.edgeDifferencePercent.toFixed(1)}%가 다릅니다. 두 이미지를 비교해 표, 텍스트 상자, 줄바꿈을 확인하세요.`,
    };
    const risks = [...slide.risks, risk];
    const penalty = risks.reduce((sum, item) => sum + (item.severity === "danger" ? 28 : 12), 0);
    const score = Math.max(0, 100 - penalty);
    return { ...slide, risks, score, severity: score >= 75 ? "safe" as const : score >= 45 ? "warning" as const : "danger" as const };
  });
  const score = slides.length ? Math.round(slides.reduce((sum, slide) => sum + slide.score, 0) / slides.length) : analysis.score;
  return { ...analysis, slides, score };
}

function badgeLabel(severity: Severity) {
  return severity === "safe" ? "안전" : severity === "warning" ? "주의" : "위험";
}

function riskSolution(category: PresentationAnalysis["slides"][number]["risks"][number]["category"]) {
  const solutions = {
    "missing-font": "아래에서 비슷한 설치 글꼴을 선택하세요. 이 카드의 교체는 현재 슬라이드에만 적용됩니다.",
    "text-density": "내용을 두 개 이상의 텍스트 상자로 나누거나 문장을 줄이고, 상자의 높이를 넉넉하게 확보하세요.",
    "text-box": "PowerPoint에서 텍스트 상자를 넓힌 뒤 자동 맞춤을 확인하고, 잘림 옵션 대신 크기 조정을 사용하세요.",
    table: "표의 행 높이와 열 너비를 늘리고 글자 크기를 확인하세요. 모양 유지가 중요하면 표를 이미지로 고정하는 방법도 좋습니다.",
    "unsupported-effect": "효과를 단순화하거나 Canva에서 해당 요소를 PNG로 내보내 PowerPoint의 이미지로 교체하세요.",
    "small-text": "본문 글자 크기를 키우고 내용이 많다면 슬라이드를 나누세요.",
    "long-token": "긴 URL이나 영문 문자열에 공백 또는 수동 줄바꿈을 넣어 상자 안에서 줄이 바뀌게 하세요.",
    "edge-proximity": "요소를 슬라이드 안쪽으로 옮겨 가장자리에서 최소 3% 정도의 여백을 확보하세요.",
    media: "동영상·오디오는 PowerPoint에서 다시 삽입하고, 발표 PC에서 재생 여부와 코덱을 확인하세요.",
    "visual-difference": "두 이미지를 위아래로 비교해 달라진 표·텍스트 상자를 찾고, PowerPoint에서 크기와 줄바꿈을 직접 조정하세요.",
  };
  return solutions[category];
}

export default function App() {
  const [source, setSource] = useState<PickedFile | null>(null);
  const [analysis, setAnalysis] = useState<PresentationAnalysis | null>(null);
  const [selectedSlide, setSelectedSlide] = useState(1);
  const [workingBytes, setWorkingBytes] = useState<Uint8Array | null>(null);
  const [loading, setLoading] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [savedMessage, setSavedMessage] = useState("");
  const [canvaLink, setCanvaLink] = useState("");
  const [canvaConnected, setCanvaConnected] = useState(false);
  const [canvaConfigured, setCanvaConfigured] = useState(true);
  const [connectionMessage, setConnectionMessage] = useState("");
  const [needsConsent, setNeedsConsent] = useState(false);
  const [slideComparisons, setSlideComparisons] = useState<SlideComparison[]>([]);
  const [comparisonError, setComparisonError] = useState("");
  const [installedFonts, setInstalledFonts] = useState<string[]>([]);
  const [batchFontChoices, setBatchFontChoices] = useState<Record<string, string>>({});
  const [problemFilter, setProblemFilter] = useState<string | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const comparisonJob = useRef(0);

  const currentSlide = analysis?.slides.find((slide) => slide.number === selectedSlide);
  const currentComparison = slideComparisons.find((slide) => slide.slideNumber === selectedSlide);
  const riskCount = analysis?.slides.reduce((sum, slide) => sum + slide.risks.length, 0) ?? 0;
  const missingFonts = analysis?.fonts.filter((font) => !font.installed) ?? [];
  const level = analysis ? (analysis.score >= 75 ? "safe" : analysis.score >= 45 ? "warning" : "danger") : "safe";
  const affectedSlides = analysis?.slides.filter((slide) => slide.risks.length > 0).length ?? 0;
  const categoryCounts = useMemo(() => problemGroups.map((group) => ({
    ...group,
    count: analysis?.slides.reduce((sum, slide) => sum + slide.risks.filter((risk) => (group.categories as readonly string[]).includes(risk.category)).length, 0) ?? 0,
  })), [analysis]);
  const visibleSlides = useMemo(() => {
    if (!analysis || !problemFilter) return analysis?.slides ?? [];
    const group = problemGroups.find((item) => item.key === problemFilter);
    return group ? analysis.slides.filter((slide) => slide.risks.some((risk) => (group.categories as readonly string[]).includes(risk.category))) : analysis.slides;
  }, [analysis, problemFilter]);

  useEffect(() => {
    void refreshCanvaStatus();
    void window.hailCanva?.listFonts().then((fonts) => setInstalledFonts(fonts));
  }, []);

  const analyze = async (file: PickedFile) => {
    const job = ++comparisonJob.current;
    setLoading(true);
    setSavedMessage("");
    setBatchFontChoices({});
    setProblemFilter(null);
    setSlideComparisons([]);
    setComparisonError("");
    try {
      const availableFonts = installedFonts.length ? installedFonts : await window.hailCanva?.listFonts() ?? [];
      setInstalledFonts(availableFonts);
      const result = await analyzePptx(file.bytes, file.name, availableFonts);
      setSource(file);
      setWorkingBytes(file.bytes);
      setAnalysis(result);
      setSelectedSlide(result.slides[0]?.number ?? 1);
      if (/^https?:\/\/(?:(?:www\.)?canva\.com|canva\.link)\//i.test(file.path) && window.hailCanva) {
        setConnectionMessage("Canva 원본과 PowerPoint 변환 이미지를 만들고 있습니다.");
        void window.hailCanva.compareCanvaSlides(file.path, file.bytes).then((comparisons) => {
          if (comparisonJob.current !== job) return;
          setSlideComparisons(comparisons);
          setAnalysis((current) => current ? addVisualRisks(current, comparisons) : current);
          setConnectionMessage(comparisons.length ? "" : "비교할 슬라이드 이미지를 만들지 못했습니다.");
        }).catch((error) => {
          if (comparisonJob.current !== job) return;
          const message = error instanceof Error ? error.message : "슬라이드 비교 이미지를 만들지 못했습니다.";
          setComparisonError(message);
          setConnectionMessage(message);
        });
      }
    } catch (error) {
      alert(`파일을 분석하지 못했습니다.\n${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    } finally {
      setLoading(false);
    }
  };

  const acceptWebFile = async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      alert(".pptx 파일을 선택해 주세요.");
      return;
    }
    await analyze({ name: file.name, path: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
  };

  const pickFile = async () => {
    if (window.hailCanva) {
      const file = await window.hailCanva.pickPptx();
      if (file) await analyze(file);
    } else fileInput.current?.click();
  };

  const refreshCanvaStatus = async () => {
    const status = await window.hailCanva?.canvaStatus();
    if (status) {
      setCanvaConnected(status.connected);
      setCanvaConfigured(status.configured);
    }
    return status;
  };

  const connect = async () => {
    if (!window.hailCanva) return;
    setLoading(true);
    setConnectionMessage("브라우저에서 Canva 접근을 허용해 주세요.");
    try {
      const status = await window.hailCanva.connectCanva();
      setCanvaConnected(status.connected);
      setCanvaConfigured(status.configured);
      setConnectionMessage(status.connected ? "Canva 계정이 연결되었습니다." : "Canva 연결을 확인해 주세요.");
    } catch (error) {
      setConnectionMessage(error instanceof Error ? error.message : "Canva 연결에 실패했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const analyzeCanvaLink = async () => {
    if (!window.hailCanva || !canvaLink.trim()) return;
    setLoading(true);
    setNeedsConsent(false);
    setConnectionMessage("Canva 연결 상태와 PPTX 다운로드 가능 여부를 확인하고 있습니다.");
    try {
      const status = await refreshCanvaStatus();
      if (status?.connected) {
        setConnectionMessage("저장된 Canva 계정으로 PPTX를 가져오고 있습니다.");
        try {
          const file = await window.hailCanva.downloadCanvaPptx(canvaLink);
          await analyze(file);
          return;
        } catch {
          setConnectionMessage(
            "연결된 계정에는 이 디자인의 직접 API 권한이 없습니다. 공유 링크 권한으로 다시 확인하고 있습니다.",
          );
        }
      }

      let publicResult = await window.hailCanva.inspectPublicCanva(canvaLink);
      if (publicResult.file) {
        setConnectionMessage(publicResult.message);
        await analyze(publicResult.file);
        return;
      }
      if (publicResult.download === "consent_required") {
        setNeedsConsent(true);
        setConnectionMessage("Canva 이용 약관 확인 창을 엽니다. 내용을 직접 확인한 뒤 창을 닫아 주세요.");
        const interactiveFile = await window.hailCanva.openCanvaConsent(canvaLink);
        if (interactiveFile) {
          setNeedsConsent(false);
          setConnectionMessage("Canva 웹 창에서 PPTX를 가져왔습니다.");
          await analyze(interactiveFile);
          return;
        }
        setConnectionMessage("동의 상태와 다운로드 가능 여부를 다시 확인하고 있습니다.");
        publicResult = await window.hailCanva.inspectPublicCanva(canvaLink);
        if (publicResult.file) {
          setNeedsConsent(false);
          setConnectionMessage(publicResult.message);
          await analyze(publicResult.file);
          return;
        }
        if (publicResult.download === "consent_required") {
          setConnectionMessage("Canva 약관 동의가 완료되지 않았습니다. 아래 버튼을 눌러 다시 확인해 주세요.");
          return;
        }
        setNeedsConsent(publicResult.download !== "pptx");
      }

      if (publicResult.download === "disabled" || publicResult.download === "failed") {
        setNeedsConsent(true);
        setConnectionMessage(
          `${publicResult.message} 아래 버튼으로 Canva 웹 창을 열어 같은 계정으로 로그인한 뒤 창을 닫고 다시 시도해 주세요.`,
        );
      } else {
        setConnectionMessage(publicResult.message);
      }
    } catch (error) {
      setConnectionMessage(error instanceof Error ? error.message : "Canva 링크를 가져오지 못했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const analyzePublicLink = async () => {
    if (!window.hailCanva || !canvaLink.trim()) return;
    setLoading(true);
    setNeedsConsent(false);
    setConnectionMessage("Canva 연결 상태를 확인하고 있습니다.");
    try {
      const status = await refreshCanvaStatus();
      if (status?.connected) {
        setConnectionMessage("연결된 Canva 계정으로 전체 PPTX를 가져와 검사하고 있습니다.");
        try {
          const file = await window.hailCanva.downloadCanvaPptx(canvaLink);
          await analyze(file);
          return;
        } catch {
          setConnectionMessage(
            "연결된 계정에는 이 디자인의 직접 API 권한이 없습니다. 공유 링크 화면을 분석하고 있습니다.",
          );
        }
      }

      setConnectionMessage("공개 Canva 슬라이드의 텍스트와 배치를 수집하고 있습니다.");
      const result = await window.hailCanva.analyzePublicCanva(canvaLink);
      setSource(null);
      setWorkingBytes(null);
      setAnalysis(result);
      setSelectedSlide(result.slides[0]?.number ?? 1);
      setConnectionMessage("");
    } catch (error) {
      const message = error instanceof Error ? error.message : "공개 Canva 디자인을 분석하지 못했습니다.";
      if (/약관|로그인|권한 창/.test(message)) setNeedsConsent(true);
      setConnectionMessage(message);
    } finally {
      setLoading(false);
    }
  };

  const completeCanvaConsent = async () => {
    if (!window.hailCanva || !canvaLink.trim()) return;
    setLoading(true);
    setConnectionMessage("열린 Canva 창에서 약관 동의 또는 로그인을 완료한 뒤 창을 닫아 주세요.");
    try {
      const interactiveFile = await window.hailCanva.openCanvaConsent(canvaLink);
      if (interactiveFile) {
        setNeedsConsent(false);
        setConnectionMessage("Canva 웹 창에서 PPTX를 가져왔습니다.");
        await analyze(interactiveFile);
        return;
      }
      setNeedsConsent(false);
      setLoading(false);
      await analyzeCanvaLink();
    } catch (error) {
      setConnectionMessage(error instanceof Error ? error.message : "Canva 동의 창을 열지 못했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const fixFont = async (from: string, to: string, slideNumber?: number) => {
    if (!workingBytes || !source) return;
    setLoading(true);
    try {
      const updated = slideNumber === undefined ? await replaceFont(workingBytes, from, to) : await replaceFontOnSlide(workingBytes, from, to, slideNumber);
      setWorkingBytes(updated);
      const availableFonts = installedFonts.length ? installedFonts : await window.hailCanva?.listFonts() ?? [to];
      const result = await analyzePptx(updated, source.name, [...availableFonts, to]);
      setAnalysis(result);
    } finally {
      setLoading(false);
    }
  };

  const save = async () => {
    if (!workingBytes || !source) return;
    const saved = await window.hailCanva?.savePptx(source.name, workingBytes);
    if (saved) setSavedMessage(`저장 완료: ${saved}`);
  };

  const fontSummary = useMemo(() => analysis?.fonts.slice(0, 5) ?? [], [analysis]);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand"><span className="brand-mark">H</span><span>HailCanva</span><small>PowerPoint Preflight</small></div>
        {analysis && <button className="secondary-button" onClick={() => { setAnalysis(null); setSource(null); setSlideComparisons([]); setComparisonError(""); }}><X size={16}/> 다른 파일</button>}
      </header>

      {!analysis ? (
        <main className="landing">
          <section className="intro">
            <span className="eyebrow">CANVA → POWERPOINT</span>
            <h1><span>발표 전에,</span><span><em>깨질 곳부터</em></span><span>찾으세요.</span></h1>
            <p>공개 Canva 링크만으로 먼저 위험 요소를 찾고, 필요할 때만 PPTX 정밀 검사를 실행합니다.</p>
            <div className="feature-row">
              <span><Check size={16}/> 로컬 분석</span><span><Check size={16}/> 파일 외부 전송 없음</span><span><Check size={16}/> 수정본 저장</span>
            </div>
          </section>
          <section className="import-panel">
            <div className="link-import">
              <div className="link-heading"><span className="link-icon"><Link2/></span><div><h2>Canva 링크로 검사</h2><p>다운로드 없이 슬라이드의 텍스트와 배치를 먼저 확인합니다</p></div></div>
              <label htmlFor="canva-link">Canva 디자인 링크</label>
              <div className="link-input-row">
                <input id="canva-link" value={canvaLink} onChange={(event) => setCanvaLink(event.target.value)} placeholder="canva.com/design/... 또는 canva.link/..." onFocus={() => void refreshCanvaStatus()}/>
              </div>
              <div className="link-actions">
                <button className="primary-button" onClick={() => void analyzePublicLink()} disabled={loading || !canvaLink.trim()}>링크 사전 검사 <ChevronRight size={17}/></button>
                <button className="secondary-button" onClick={() => void analyzeCanvaLink()} disabled={loading || !canvaLink.trim()}>PPTX 정밀 검사</button>
              </div>
              <div className="connection-row">
                <span className={`connection-dot ${canvaConnected ? "online" : ""}`}/>
                <span>{canvaConnected ? "Canva 연결됨 · 소유하거나 직접 공유받은 디자인 정밀 검사" : "공개 링크는 계정 연결 없이 검사"}</span>
                {!canvaConnected && canvaConfigured && <button onClick={() => void connect()}><LogIn size={14}/> 비공개 링크 연결</button>}
              </div>
              {connectionMessage && <p className="connection-message">{connectionMessage}</p>}
              {needsConsent && <button className="consent-button" onClick={() => void completeCanvaConsent()} disabled={loading}><ExternalLink size={14}/> Canva 웹 권한 창 열기</button>}
            </div>
            <div className="divider"><span>또는</span></div>
            <div
            className={`drop-zone compact ${dragging ? "dragging" : ""}`}
            onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => { event.preventDefault(); setDragging(false); const file = event.dataTransfer.files[0]; if (file) void acceptWebFile(file); }}
          >
            {loading ? <LoaderCircle className="spin upload-icon"/> : <FileUp className="upload-icon"/>}
            <h2>{loading ? "파일을 검사하고 있습니다" : "PPTX 직접 검사"}</h2>
            <p>파일을 놓거나 컴퓨터에서 선택하세요</p>
            <button className="secondary-button" onClick={pickFile} disabled={loading}>파일 선택 <ChevronRight size={18}/></button>
            <input ref={fileInput} hidden type="file" accept=".pptx" onChange={(event) => { const file = event.target.files?.[0]; if (file) void acceptWebFile(file); }}/>
            </div>
          </section>
        </main>
      ) : (
        <main className="dashboard">
          <section className="summary-panel">
            <div>
              <span className="eyebrow">{analysis.sourceType === "public-link" ? "LINK PREFLIGHT COMPLETE" : "ANALYSIS COMPLETE"}</span>
              <h1>{analysis.fileName}</h1>
              {analysis.scanNotice && <p className="scan-notice">{analysis.scanNotice}</p>}
            </div>
            <div className={`score-card ${level}`}>
              <div className="score-ring"><strong>{analysis.score}</strong><span>/ 100</span></div>
              <div><b>{analysis.sourceType === "public-link" ? "사전 위험 점수" : "호환성 점수"}</b><span>{level === "safe" ? "발표 준비가 양호합니다" : "위험 요소를 확인하세요"}</span></div>
            </div>
          </section>

          <section className="stat-grid">
            <article><ShieldCheck/><div><strong>{affectedSlides}/{analysis.slideCount}</strong><span>확인이 필요한 슬라이드</span></div></article>
            <article><FileWarning/><div><strong>{riskCount}</strong><span>발견된 위험 요소</span></div></article>
            <article><Type/><div><strong>{analysis.sourceType === "public-link" ? analysis.fonts.length : missingFonts.length}</strong><span>{analysis.sourceType === "public-link" ? "감지된 글꼴" : "누락된 글꼴"}</span></div></article>
          </section>

          <section className="issue-panel">
            <div className="issue-heading"><div><b>문제 분류</b><span>종류별 문제를 확인하고 반복 작업을 한 번에 처리합니다</span></div></div>
            <div className="issue-categories">
              {categoryCounts.map((group) => <button className={problemFilter === group.key ? "active" : ""} disabled={!group.count} key={group.key} onClick={() => {
                const next = problemFilter === group.key ? null : group.key;
                setProblemFilter(next);
                if (next) {
                  const sourceGroup = problemGroups.find((item) => item.key === next);
                  const first = analysis.slides.find((slide) => slide.risks.some((risk) => sourceGroup && (sourceGroup.categories as readonly string[]).includes(risk.category)));
                  if (first) setSelectedSlide(first.number);
                }
              }}><span>{group.label}</span><strong>{group.count}</strong></button>)}
            </div>
            {!!missingFonts.length && <div className="batch-fonts">
              <b>누락 글꼴 일괄 교체</b>
              {missingFonts.map((font) => {
                const suggestions = recommendFonts(font.family, installedFonts);
                return <div className="batch-font-row" key={font.family}>
                  <span><b>{font.family}</b><small>{font.slides.length}개 슬라이드</small></span>
                  <select value={batchFontChoices[font.family] ?? ""} onChange={(event) => setBatchFontChoices((current) => ({ ...current, [font.family]: event.target.value }))}>
                    <option value="" disabled>비슷한 글꼴 선택</option>
                    {suggestions.map((candidate) => <option key={candidate}>{candidate}</option>)}
                  </select>
                  <button className="secondary-button" disabled={!batchFontChoices[font.family]} onClick={() => void fixFont(font.family, batchFontChoices[font.family])}>전체 교체</button>
                </div>;
              })}
            </div>}
          </section>

          <section className="workspace">
            <aside className="slide-list">
              <div className="section-title"><span>슬라이드</span><small>{analysis.scannedSlideCount === undefined ? analysis.slideCount : `${analysis.scannedSlideCount}/${analysis.slideCount} 검사`}</small></div>
              {visibleSlides.map((slide) => (
                <button key={slide.number} className={selectedSlide === slide.number ? "selected" : ""} onClick={() => setSelectedSlide(slide.number)}>
                  <span className="slide-number">{String(slide.number).padStart(2, "0")}</span>
                  <span className="slide-copy"><b>{slide.title}</b><small>{slide.risks.length ? `${slide.risks.length}개 확인 필요` : "문제 없음"}</small></span>
                  <span className={`status-dot ${slide.severity}`}/>
                </button>
              ))}
            </aside>

            <div className="detail-panel">
              <div className="detail-header">
                <div><span className="eyebrow">SLIDE {currentSlide?.number}</span><h2>{currentSlide?.title}</h2></div>
                <span className={`severity-badge ${currentSlide?.severity}`}>{badgeLabel(currentSlide?.severity ?? "safe")}</span>
              </div>
              {currentComparison && <div className="visual-comparison">
                <figure><figcaption>Canva 원본</figcaption><img src={currentComparison.canvaImage} alt={`Canva 원본 ${currentComparison.slideNumber}페이지`}/></figure>
                <figure><figcaption>PowerPoint 변환</figcaption><img src={currentComparison.powerpointImage} alt={`PowerPoint 변환 ${currentComparison.slideNumber}페이지`}/></figure>
                <div className="visual-difference-note">화면 차이 {currentComparison.differencePercent.toFixed(1)}% · 가장자리 차이 {currentComparison.edgeDifferencePercent.toFixed(1)}%</div>
              </div>}
              {!currentComparison && comparisonError && <div className="comparison-error"><AlertTriangle size={16}/><span>{comparisonError}</span></div>}
              <div className="text-preview"><span>PPTX 추출 텍스트</span><small>이미지에서 보이지 않는 잘림이나 누락을 원문과 대조하는 용도입니다.</small><p>{currentSlide?.textPreview || "텍스트가 없는 슬라이드입니다."}</p></div>
              <div className="risks">
                <h3>위험 요소</h3>
                {!currentSlide?.risks.length && <div className="empty-state"><Check/><div><b>발견된 문제가 없습니다</b><span>기본 호환성 검사 결과 안전합니다.</span></div></div>}
                {currentSlide?.risks.map((risk) => (
                  <article className={`risk-card ${risk.severity}`} key={risk.id}>
                    <AlertTriangle/>
                    <div className="risk-copy"><b>{risk.title}</b><p>{risk.detail}</p><div className="risk-solution"><strong>해결 방법</strong><span>{riskSolution(risk.category)}</span></div></div>
                    {risk.fontFamily && <select defaultValue="" onChange={(event) => { if (event.target.value) void fixFont(risk.fontFamily!, event.target.value, risk.slideNumber); }}>
                      <option value="" disabled>글꼴 교체</option>
                      {(recommendFonts(risk.fontFamily, installedFonts).length ? recommendFonts(risk.fontFamily, installedFonts) : replacementOptions.filter(f => f !== risk.fontFamily)).map(font => <option key={font}>{font}</option>)}
                    </select>}
                  </article>
                ))}
              </div>
            </div>

            <aside className="font-panel">
              <div className="section-title"><span>사용 글꼴</span><small>{analysis.fonts.length}</small></div>
              {fontSummary.map((font) => <div className="font-row" key={font.family}>
                <div><b>{font.family}</b><small>{font.slides.length}개 슬라이드 · {font.occurrences}회</small></div>
                <span className={font.installed ? "installed" : "missing"}>{analysis.sourceType === "public-link" ? "감지됨" : font.embedded ? "문서 포함" : font.installed ? "설치됨" : "누락"}</span>
              </div>)}
              {analysis.fonts.length > 5 && <p className="more-fonts">외 {analysis.fonts.length - 5}개 글꼴</p>}
              {analysis.sourceType === "public-link" ? <div className="save-area"><p>링크 검사는 위험 가능성을 빠르게 찾는 사전 검사입니다. 실제 PowerPoint 변환 결과는 PPTX 정밀 검사로 확인하세요.</p></div> : <div className="save-area">
                  <p>수정한 프레젠테이션을 새 파일로 저장합니다.</p>
                  <button className="primary-button" onClick={save}><Save size={17}/> 수정본 저장</button>
                  {savedMessage && <small className="saved-message">{savedMessage}</small>}
                </div>}
            </aside>
          </section>
        </main>
      )}
      {loading && analysis && <div className="loading-overlay"><LoaderCircle className="spin"/><span>변경 사항을 적용하고 있습니다</span></div>}
    </div>
  );
}
