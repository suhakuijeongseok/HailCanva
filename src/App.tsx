import { useEffect, useMemo, useRef, useState } from "react";
import { AlertTriangle, Check, ChevronRight, ExternalLink, FileUp, FileWarning, Link2, LoaderCircle, LogIn, Save, ShieldCheck, Type, X } from "lucide-react";
import { analyzePptx, replaceFont } from "./lib/pptxAnalyzer";
import type { PickedFile, PresentationAnalysis, Severity } from "./types";

const replacementOptions = ["맑은 고딕", "Noto Sans KR", "Arial", "Calibri"];

function badgeLabel(severity: Severity) {
  return severity === "safe" ? "안전" : severity === "warning" ? "주의" : "위험";
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
  const fileInput = useRef<HTMLInputElement>(null);

  const currentSlide = analysis?.slides.find((slide) => slide.number === selectedSlide);
  const riskCount = analysis?.slides.reduce((sum, slide) => sum + slide.risks.length, 0) ?? 0;
  const missingFonts = analysis?.fonts.filter((font) => !font.installed) ?? [];
  const level = analysis ? (analysis.score >= 75 ? "safe" : analysis.score >= 45 ? "warning" : "danger") : "safe";

  useEffect(() => {
    void refreshCanvaStatus();
  }, []);

  const analyze = async (file: PickedFile) => {
    setLoading(true);
    setSavedMessage("");
    try {
      const installedFonts = await window.hailCanva?.listFonts() ?? [];
      const result = await analyzePptx(file.bytes, file.name, installedFonts);
      setSource(file);
      setWorkingBytes(file.bytes);
      setAnalysis(result);
      setSelectedSlide(result.slides[0]?.number ?? 1);
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
        const file = await window.hailCanva.downloadCanvaPptx(canvaLink);
        await analyze(file);
        return;
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
        await window.hailCanva.openCanvaConsent(canvaLink);
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
        setNeedsConsent(false);
      }

      setConnectionMessage(publicResult.message);
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
        const file = await window.hailCanva.downloadCanvaPptx(canvaLink);
        await analyze(file);
        return;
      }

      setConnectionMessage("공개 Canva 슬라이드의 텍스트와 배치를 수집하고 있습니다.");
      const result = await window.hailCanva.analyzePublicCanva(canvaLink);
      setSource(null);
      setWorkingBytes(null);
      setAnalysis(result);
      setSelectedSlide(result.slides[0]?.number ?? 1);
      setConnectionMessage("");
    } catch (error) {
      setConnectionMessage(error instanceof Error ? error.message : "공개 Canva 디자인을 분석하지 못했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const completeCanvaConsent = async () => {
    if (!window.hailCanva || !canvaLink.trim()) return;
    setLoading(true);
    setConnectionMessage("열린 Canva 창에서 약관을 직접 확인한 뒤 창을 닫아 주세요.");
    try {
      await window.hailCanva.openCanvaConsent(canvaLink);
      setNeedsConsent(false);
      setLoading(false);
      await analyzeCanvaLink();
    } catch (error) {
      setConnectionMessage(error instanceof Error ? error.message : "Canva 동의 창을 열지 못했습니다.");
    } finally {
      setLoading(false);
    }
  };

  const fixFont = async (from: string, to: string) => {
    if (!workingBytes || !source) return;
    setLoading(true);
    try {
      const updated = await replaceFont(workingBytes, from, to);
      setWorkingBytes(updated);
      const installedFonts = await window.hailCanva?.listFonts() ?? [to];
      const result = await analyzePptx(updated, source.name, [...installedFonts, to]);
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
        {analysis && <button className="secondary-button" onClick={() => { setAnalysis(null); setSource(null); }}><X size={16}/> 다른 파일</button>}
      </header>

      {!analysis ? (
        <main className="landing">
          <section className="intro">
            <span className="eyebrow">CANVA → POWERPOINT</span>
            <h1>발표 전에,<br/><em>깨질 곳부터</em> 찾으세요.</h1>
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
                <input id="canva-link" value={canvaLink} onChange={(event) => setCanvaLink(event.target.value)} placeholder="https://www.canva.com/design/DAG.../edit" onFocus={() => void refreshCanvaStatus()}/>
              </div>
              <div className="link-actions">
                <button className="primary-button" onClick={() => void analyzePublicLink()} disabled={loading || !canvaLink.trim()}>링크 사전 검사 <ChevronRight size={17}/></button>
                <button className="secondary-button" onClick={() => void analyzeCanvaLink()} disabled={loading || !canvaLink.trim()}>PPTX 정밀 검사</button>
              </div>
              <div className="connection-row">
                <span className={`connection-dot ${canvaConnected ? "online" : ""}`}/>
                <span>{canvaConnected ? "Canva 연결됨 · 비공개 링크도 검사 가능" : "공개 링크는 계정 연결 없이 검사"}</span>
                {!canvaConnected && canvaConfigured && <button onClick={() => void connect()}><LogIn size={14}/> 비공개 링크 연결</button>}
              </div>
              {connectionMessage && <p className="connection-message">{connectionMessage}</p>}
              {needsConsent && <button className="consent-button" onClick={() => void completeCanvaConsent()} disabled={loading}><ExternalLink size={14}/> Canva 동의 창 열기</button>}
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
              <p>{analysis.slideCount}개 슬라이드 · {riskCount}개 잠재적 문제{analysis.sourceType === "public-link" ? " · 링크 기반 추정" : ""}</p>
              {analysis.scanNotice && <p className="scan-notice">{analysis.scanNotice}</p>}
            </div>
            <div className={`score-card ${level}`}>
              <div className="score-ring"><strong>{analysis.score}</strong><span>/ 100</span></div>
              <div><b>{analysis.sourceType === "public-link" ? "사전 위험 점수" : "호환성 점수"}</b><span>{level === "safe" ? "발표 준비가 양호합니다" : "위험 요소를 확인하세요"}</span></div>
            </div>
          </section>

          <section className="stat-grid">
            <article><ShieldCheck/><div><strong>{analysis.slides.filter(s => s.severity === "safe").length}</strong><span>안전한 슬라이드</span></div></article>
            <article><FileWarning/><div><strong>{riskCount}</strong><span>발견된 위험 요소</span></div></article>
            <article><Type/><div><strong>{analysis.sourceType === "public-link" ? analysis.fonts.length : missingFonts.length}</strong><span>{analysis.sourceType === "public-link" ? "감지된 글꼴" : "누락된 글꼴"}</span></div></article>
          </section>

          <section className="workspace">
            <aside className="slide-list">
              <div className="section-title"><span>슬라이드</span><small>{analysis.scannedSlideCount === undefined ? analysis.slideCount : `${analysis.scannedSlideCount}/${analysis.slideCount} 검사`}</small></div>
              {analysis.slides.map((slide) => (
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
              <div className="text-preview"><span>텍스트 미리보기</span><p>{currentSlide?.textPreview || "텍스트가 없는 슬라이드입니다."}</p></div>
              <div className="risks">
                <h3>위험 요소</h3>
                {!currentSlide?.risks.length && <div className="empty-state"><Check/><div><b>발견된 문제가 없습니다</b><span>기본 호환성 검사 결과 안전합니다.</span></div></div>}
                {currentSlide?.risks.map((risk) => (
                  <article className={`risk-card ${risk.severity}`} key={risk.id}>
                    <AlertTriangle/>
                    <div className="risk-copy"><b>{risk.title}</b><p>{risk.detail}</p></div>
                    {risk.fontFamily && <select defaultValue="" onChange={(event) => { if (event.target.value) void fixFont(risk.fontFamily!, event.target.value); }}>
                      <option value="" disabled>글꼴 교체</option>
                      {replacementOptions.filter(f => f !== risk.fontFamily).map(font => <option key={font}>{font}</option>)}
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
