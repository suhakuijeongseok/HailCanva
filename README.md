# HailCanva

Canva에서 내보낸 PowerPoint 파일을 발표 전에 검사하는 Windows 데스크톱 앱입니다. 모든 분석은 로컬에서 이루어집니다.

## 현재 구현된 MVP

- PPTX 드래그앤드롭 및 파일 선택
- Canva 디자인 링크 입력 및 OAuth 계정 연결
- 공개 Canva 링크의 익명 접근·PPTX 다운로드 가능 여부 확인
- 공개 Canva 링크를 다운로드하지 않고 현재 표시된 슬라이드의 텍스트·글꼴·배치·미디어 위험 사전 검사
- 다운로드가 허용된 공개 링크의 PPTX 자동 수집 및 검사
- Canva 링크에서 PPTX 자동 내보내기·다운로드
- Windows에 설치된 글꼴과 PPTX 사용 글꼴 비교
- 슬라이드별 누락 글꼴 탐지
- 텍스트 밀도가 높은 개체 탐지
- 광선·반사·변형 텍스트·비디오 효과 경고
- 슬라이드별 PowerPoint 호환성 점수
- 누락 글꼴 일괄 교체
- 수정한 PPTX를 별도 파일로 저장

## 실행

Node.js 20 이상이 필요합니다.

```powershell
npm.cmd install
npm.cmd run dev
```

### Canva 링크 연동 설정

공개 보기 링크는 별도 설정 없이 `링크 사전 검사`로 확인할 수 있습니다. 이 검사는 PPTX를 내려받지 않고 Canva가 브라우저에 표시한 요소를 읽어 텍스트 밀도, 작은 글자, 긴 문자열, 가장자리 배치, 미디어 포함 여부를 추정합니다. `PPTX 정밀 검사`는 다운로드가 허용된 링크에서 파일을 받아 내부 구조와 로컬 글꼴까지 확인합니다.

Canva가 자동 페이지 전환을 제한하면 앱은 실제로 읽은 슬라이드 수를 `검사 수/전체 수`로 표시합니다. 읽지 못한 페이지를 첫 페이지 결과로 복제하지 않으며, 링크 검사는 어디까지나 빠른 사전 검사입니다.

Canva가 익명 편집 링크에서 이용 약관 동의를 요구하면 앱은 이를 다운로드 금지로 오판하지 않고 사용자 확인 창을 자동으로 엽니다. 사용자가 직접 내용을 확인하고 동의한 뒤 창을 닫으면 같은 앱 전용 세션으로 검사를 자동 재개합니다. 동의를 완료하지 않고 닫았을 때만 `Canva 동의 창 열기` 재시도 버튼이 표시됩니다. HailCanva가 약관 체크박스나 동의 버튼을 대신 누르지는 않습니다.

공개 열람만 가능하고 다운로드가 금지된 링크도 링크 사전 검사는 사용할 수 있습니다. 로그인 또는 접근 권한이 필요한 링크의 PPTX를 정밀 검사하려면 아래 OAuth 설정을 추가합니다.

1. [Canva Developer Portal](https://www.canva.com/developers/)에서 앱을 생성합니다.
2. 앱의 **Outside Canva** 설정에서 REST API 통합을 구성합니다.
3. 다음 권한을 활성화합니다.
   - `design:meta:read`
   - `design:content:read`
4. Redirect URL에 아래 주소를 정확히 등록합니다.

```text
http://127.0.0.1:47821/oauth/callback
```

5. `.env.example`을 `.env`로 복사하고 발급받은 값을 입력합니다.

```powershell
Copy-Item .env.example .env
```

```dotenv
CANVA_CLIENT_ID=발급받은_client_id
CANVA_CLIENT_SECRET=발급받은_client_secret
CANVA_REDIRECT_URI=http://127.0.0.1:47821/oauth/callback
CANVA_OAUTH_PORT=47821
```

6. 앱을 다시 실행하고 **Canva 계정 연결**을 누릅니다. 브라우저에서 접근 권한을 허용한 뒤 디자인 링크를 입력합니다.

Client Secret이 들어 있는 `.env`는 Git에 포함되지 않습니다. 현재 방식은 개발·개인 사용을 위한 로컬 OAuth broker입니다. 일반 사용자에게 앱을 배포할 때는 Client Secret을 앱에 포함하지 말고 별도의 HTTPS 인증 서버로 옮겨야 합니다.

프로덕션 빌드:

```powershell
npm.cmd run build
npm.cmd start
```

## 구조

```text
electron/           Electron main/preload, 로컬 파일 및 Windows 글꼴 접근
src/lib/            PPTX 분석·수정 엔진
src/                React 사용자 인터페이스
docs/               제품 및 기술 설계
```

## 로드맵

1. Canva 링크에서 페이지별 PNG 기준본도 자동 다운로드
2. PowerPoint가 설치된 PC에서 슬라이드를 PNG로 렌더링
3. Canva PNG 기준본과 PowerPoint 렌더링 결과 비교
4. 텍스트 넘침의 실제 영역 탐지
5. 선택 요소 이미지 고정
6. Chrome 확장 프로그램 및 Native Messaging 연결

## 주의사항

링크 검사 점수는 Canva 웹 화면에서 관찰 가능한 요소만 이용한 추정치이며 누락 글꼴이나 실제 PowerPoint 변환 결과를 확정하지 않습니다. PPTX 검사는 파일 내부 구조와 로컬 글꼴을 이용한 사전 위험도입니다. PowerPoint 렌더링 기반 시각 비교는 다음 단계에 포함됩니다. 원본 파일은 덮어쓰지 않으며 수정본을 새 파일로 저장합니다.

Canva 공개 페이지 자동화는 CSS 클래스 대신 접근성 레이블을 사용하지만 Canva 화면이나 문구가 변경되면 다운로드 탐지가 실패할 수 있습니다. 이런 경우 앱은 파일 업로드 또는 공식 API 연결을 대안으로 안내합니다.
