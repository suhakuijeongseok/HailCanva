# HailCanva 제품·기술 설계

## 목표

Canva에서 만든 프레젠테이션이 Microsoft PowerPoint에서 열릴 때 발생할 수 있는 글꼴 대체, 줄바꿈 변화, 텍스트 넘침 및 효과 손실을 발표 전에 탐지하고 안전한 수정안을 제공한다.

## 제품 원칙

- 원본 PPTX를 덮어쓰지 않는다.
- 기본 분석은 네트워크 전송 없이 로컬에서 수행한다.
- 자동 수정 전 문제와 변경 내용을 사용자에게 보여준다.
- 편집 가능성을 우선하되 해결할 수 없는 요소만 선택적으로 이미지화한다.
- 휴리스틱 검사와 실제 렌더링 검사를 명확히 구분한다.

## 시스템 경계

```text
Chrome Extension (후속)
  └─ Canva 다운로드 감지 / 로컬 앱 실행
       ↓ Native Messaging
Electron Desktop
  ├─ Canva Link Import / OAuth PKCE
  ├─ Canva PPTX Export Client
  ├─ PPTX OOXML Analyzer
  ├─ Windows Font Inventory
  ├─ Compatibility Rules
  ├─ PPTX Patcher
  ├─ PowerPoint Renderer (후속)
  └─ Visual Diff Engine (후속)
```

Chrome 확장 프로그램은 Canva DOM을 분석 엔진으로 사용하지 않는다. Canva UI 변경에 영향을 적게 받도록 다운로드 감지와 로컬 앱 호출만 담당한다.

## Canva 링크 가져오기

1. 사용자가 `canva.com/design/{designId}/...` 링크를 입력한다.
2. 격리된 임시 Electron 세션에서 익명으로 링크를 연다.
3. 권한 요청·로그인 리디렉션·시각 콘텐츠 유무로 공개 접근 여부를 판정한다.
4. 쿠키 배너는 접근성 문구로 닫고, Canva 이용 약관 모달은 `consent_required` 상태로 별도 판정한다.
5. 약관 동의가 필요하면 앱 전용 브라우저 창을 표시하고 사용자가 직접 처리하도록 handoff한다.
6. 접근성 레이블로 공유 및 다운로드 메뉴를 찾아 PPTX 다운로드를 시도한다.
7. 다운로드가 성공하면 임시 파일을 읽어 기존 OOXML 분석기로 전달한다.
8. 공개 다운로드가 불가능하지만 OAuth가 연결되어 있으면 공식 Export API로 재시도한다.
9. OAuth 사용 시 `design:meta:read`, `design:content:read` 권한으로 디자인 접근 가능 여부를 검사한다.
10. access token은 메모리에만 보관하고 앱 종료 시 폐기한다.

익명 자동화는 Canva 내부 CSS나 비공개 API를 사용하지 않는다. 접근성 레이블과 실제 다운로드 이벤트만 이용하며, CAPTCHA나 UI 변경이 감지되면 우회하지 않고 다운로드 불가로 처리한다.

개발 환경에서는 Client Secret을 `.env`에서 읽는다. 배포판에 Secret을 포함하면 추출될 수 있으므로 정식 배포에서는 OAuth 코드 교환과 refresh token 관리를 별도의 HTTPS 인증 서버로 이전한다.

## 위험도 모델

| 조건 | 기본 감점 |
|---|---:|
| 현재 PC에 없는 글꼴 | 28 |
| 일반 폰트지만 현재 PC에 없음 | 12 |
| 고밀도 텍스트 개체 | 12~28 |
| 특수 효과 또는 영상 | 12 |

- 75~100: 안전
- 45~74: 주의
- 0~44: 위험

점수는 절대적인 호환성 보장이 아니라 우선 확인할 슬라이드를 찾기 위한 지표다.

## 다음 구현 단계

### PowerPoint 렌더링

Electron main process에서 PowerShell을 통해 로컬 PowerPoint COM 객체를 호출한다. 사용자 세션에서만 실행하고, PowerPoint 미설치 시 정적 분석만 제공한다.

### 시각적 비교

Canva에서 함께 내려받은 PDF 또는 PNG를 기준본으로 받는다. PowerPoint 렌더링 이미지와 페이지별로 정렬하고 SSIM, 윤곽선 및 OCR 영역 차이를 결합한다.

### Chrome 연결

Manifest V3 확장 프로그램이 `chrome.downloads`로 Canva의 PPTX 다운로드 완료를 감지한다. 사용자의 확인 후 Native Messaging host로 파일 경로를 전달하고 데스크톱 앱을 연다.
