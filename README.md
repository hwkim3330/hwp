# Office Agent Staff

`rhwp`를 브라우저 편집 엔진으로 활용하고, `claw-code`의 도구 지향 에이전트 발상을 차용한 한국어 문서 작성용 오피스 에이전트입니다.

현재 구현은 다음에 집중합니다.

- HWP/HWPX 파일 로드
- 빈 한글 문서 생성
- SVG 페이지 미리보기
- 한국어 업무 문서 생성/수정용 에이전트 플래너
- HWP 파일 다시 저장
- OpenAI 호환 API 또는 Ollama 연결
- LLM 미연결 시 오프라인 문서 초안 생성 fallback
- `Writer / Notes / Sheet / Slides` 워크스테이션
- 현재 탭 기준 모드별 에이전트 생성
- 브라우저 로컬 저장 복원

## 실행

```bash
cd /Users/parksik/office-agent-staff
python3 app.py
```

브라우저에서 `http://127.0.0.1:8765`를 엽니다.

## 데스크톱 앱 실행

```bash
cd /Users/parksik/office-agent-staff
npm install
npm run desktop
```

이 모드에서는 다음이 함께 동작합니다.

- 독립 macOS 앱 창으로 Office Agent Staff 실행
- 메뉴바에서 CPU / MEM / BAT 사용률 표시
- 클릭 시 작은 시스템 모니터 팝오버 표시
- GPU 모델 표시
- NPU/ANE 사용량은 현재 Electron 빌드에서 측정 불가

## macOS 패키징

```bash
cd /Users/parksik/office-agent-staff
npm run package:mac
```

빌드 결과물은 `dist-electron/`에 생성됩니다.

## GitHub Releases 배포

```bash
cd /Users/parksik/office-agent-staff
npm run package:mac
npm run release:github
```

현재 기준으로는 `GitHub Releases`에 `dmg`와 `zip`을 올리는 방식이 가장 간단합니다.
서명과 notarization이 없으면 Gatekeeper 경고가 날 수 있으니, 공개 배포 전에는 Apple Developer 인증서까지 붙이는 편이 안전합니다.

## 환경 변수

```bash
export LLM_BASE_URL="http://127.0.0.1:11434/v1"
export LLM_MODEL="gemma4:latest"
export LLM_API_KEY=""
```

기본값은 로컬 Ollama를 가정합니다. OpenAI 호환 서버라면 `LLM_BASE_URL`과 `LLM_API_KEY`만 바꾸면 됩니다.
모델 서버가 없어도 앱은 동작하며, 이 경우 내장 오프라인 플래너가 기본 문서 초안을 생성합니다.

## 구조

- `app.py`: 정적 파일 서버 + 에이전트 플래너 API
- `static/index.html`: UI
- `static/app.js`: rhwp 연동, 문서 스냅샷 추출, 에이전트 명령 실행
- `static/styles.css`: 레이아웃/스타일
- `electron/main.js`: Electron 데스크톱 셸 + 메뉴바 모니터
- `electron/monitor.html`: 시스템 모니터 팝오버 UI

## 구현 메모

- 문서 렌더링/편집은 CDN의 `@rhwp/core`를 직접 사용합니다.
- 에이전트는 문서 스냅샷을 바탕으로 JSON 작업 계획을 반환합니다.
- 지원 작업:
  - `set_document_html`
  - `append_html`
  - `replace_paragraph_text`
  - `replace_paragraph_html`
  - `set_note_text`
  - `set_sheet_data`
  - `set_slides`
  - `no_op`

## 한계

- 외부 HWP 문서를 부분 수정할 때 원본 복합 서식은 일부 잃을 수 있습니다.
- 현재는 업무 문서 초안 생성과 단락 단위 수정에 최적화되어 있습니다.
- `.docx/.xlsx/.pptx`는 아직 완전 파싱 편집이 아니라 가져오기 보조 수준입니다.
- GPU 실시간 사용률과 NPU/ANE 사용률은 현재 공개 Node/Electron 경로만으로는 안정적으로 수집하지 못합니다.
