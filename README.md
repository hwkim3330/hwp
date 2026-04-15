# hwp

`rhwp`를 브라우저 편집 엔진으로 활용하고, `claw-code`의 도구 지향 에이전트 발상을 차용한 한국어 문서 작성용 오피스 에이전트입니다.

현재 구현은 다음에 집중합니다.

- HWP/HWPX 파일 로드
- DOCX 파일을 Writer로 가져오기
- XLSX 파일을 Sheet로 가져오기
- PPTX 파일을 Slides로 가져오기
- Writer에서 HWP/DOCX 저장
- Sheet에서 CSV/XLSX 저장
- Slides에서 Markdown/PPTX 저장
- 빈 한글 문서 생성
- SVG 페이지 미리보기
- 한국어 업무 문서 생성/수정용 에이전트 플래너
- HWP 파일 다시 저장
- OpenAI 호환 API 또는 Ollama 연결
- LLM 미연결 시 오프라인 문서 초안 생성 fallback
- `Writer / Notes / Sheet / Slides` 워크스테이션
- 현재 탭 기준 모드별 에이전트 생성
- 브라우저 로컬 저장 복원
- `browser-use` 참조 기반 `Computer Use` 브라우저 계획/세션/단계 실행
- 작업 기억 자동 저장 / 고정 / 재사용 / 워크스페이스 패키지 포함
- 프로젝트 이름 / 목표 / 집중 모드
- 첫 실행 온보딩 / 시작 템플릿 / 기본 설정 시트

## 실행

```bash
cd /Users/parksik/hwp
python3 app.py
```

브라우저에서 `http://127.0.0.1:8765`를 엽니다.

## 데스크톱 앱 실행

```bash
cd /Users/parksik/hwp
npm install
npm run desktop
```

이 모드에서는 다음이 함께 동작합니다.

- 독립 macOS 앱 창으로 hwp 실행
- 메뉴바에서 CPU / MEM / BAT 사용률 표시
- 클릭 시 작은 시스템 모니터 팝오버 표시
- `Vision Lab` 창에서 Head Pointer / 외장 카메라 / lid angle 실험 구조 확인
- GPU 모델 표시
- NPU/ANE 사용량은 현재 Electron 빌드에서 측정 불가
- 첫 실행 시 권한/준비 상태 점검 시트 표시

## macOS 패키징

```bash
cd /Users/parksik/hwp
npm run package:mac
```

빌드 결과물은 `dist-electron/`에 생성됩니다.
현재 앱 식별자는 `com.hwkim.hwp`입니다.
패키지에는 전용 `hwp` macOS 아이콘이 포함됩니다.
패키징 전에 기존 `dist-electron/` 산출물은 자동으로 정리합니다.

## GitHub Releases 배포

```bash
cd /Users/parksik/hwp
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

### SuperGemma MLX 실험 런타임

`Jiunsong/supergemma4-26b-uncensored-mlx-4bit-v2` 같은 MLX 모델은 별도 실험 런타임으로 붙일 수 있습니다.

```bash
cd /Users/parksik/hwp
./scripts/start_supergemma4_mlx.sh
```

기본 포트는 `127.0.0.1:8081`이며, 앱의 모델 프로필에서 `SuperGemma MLX`를 선택하면 해당 런타임을 사용합니다.

### Gemini CLI 보조 런타임

로컬 `gemini` CLI가 설치되어 있으면 앱의 모델 프로필에서 `Gemini CLI 보조`를 선택할 수 있습니다.

```bash
export GEMINI_CLI_CMD="gemini"
export GEMINI_CLI_MODEL="gemini-3-flash-preview"
```

이 프로필은 기본 Ollama/Gemma 경로를 대체하는 것이 아니라, 보조 플래너로 `Gemini CLI`를 호출하는 용도입니다.

## 구조

- `app.py`: 정적 파일 서버 + 에이전트 플래너 API
- `static/index.html`: UI
- `static/app.js`: rhwp 연동, 문서 스냅샷 추출, 에이전트 명령 실행
- `static/styles.css`: 레이아웃/스타일
- `electron/main.js`: Electron 데스크톱 셸 + 메뉴바 모니터
- `electron/monitor.html`: 시스템 모니터 팝오버 UI
- `office-agent-sources/browser-use`: 브라우저 자동화 참조 저장소

## 구현 메모

- 문서 렌더링/편집은 CDN의 `@rhwp/core`를 직접 사용합니다.
- 에이전트는 문서 스냅샷을 바탕으로 JSON 작업 계획을 반환합니다.
- Computer Use는 `browser-use`의 세션/행동 분리 아이디어를 참고해 브라우저 계획과 단계 실행을 나눕니다.
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
- OOXML 저장은 현재 앱이 재구성한 내용 기준으로 생성되며, 원본 문서의 고급 레이아웃과 서식까지 그대로 보존하지는 않습니다.
- GPU 실시간 사용률과 NPU/ANE 사용률은 현재 공개 Node/Electron 경로만으로는 안정적으로 수집하지 못합니다.
External office engines:

```bash
# ONLYOFFICE
./scripts/start_onlyoffice_docker.sh

# Collabora
./scripts/start_collabora_docker.sh
```

Current environment note:
- This machine currently does not have `docker` or `podman`, so these engines are configured in the app but not runnable until a container runtime is installed.

CLI:

```bash
python3 scripts/hwp_cli.py runtime
python3 scripts/hwp_cli.py search "Gemma 4 official docs"
python3 scripts/hwp_cli.py gemini "현재 문서를 한국어 보고서 구조로 바꿔야 할 핵심만 JSON으로 정리해줘"
python3 scripts/hwp_cli.py plan "회의록 형태로 정리해줘" --mode writer
python3 scripts/hwp_cli.py browser-plan "Gemma 4 release notes 찾아줘"
python3 scripts/hwp_cli.py workspace-agent .runtime/sample.json "보고서 형식으로 정리해줘" --mode writer
```

MCP:

```bash
# stdio MCP server
python3 scripts/hwp_mcp.py
```

프로젝트 루트의 [.mcp.json](/Users/parksik/hwp/.mcp.json)은 `Claude Code`, `Gemini CLI` 같은 MCP 클라이언트가 `hwp` 도구를 직접 붙일 수 있도록 준비된 기본 설정입니다.

현재 MCP 도구:
- `hwp_runtime`
- `hwp_search`
- `hwp_plan`
- `hwp_browser_plan`
- `hwp_gemini`
