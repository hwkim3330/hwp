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

## 실행

```bash
cd /Users/parksik/office-agent-staff
python3 app.py
```

브라우저에서 `http://127.0.0.1:8765`를 엽니다.

## 환경 변수

```bash
export LLM_BASE_URL="http://127.0.0.1:11434/v1"
export LLM_MODEL="gemma3:27b"
export LLM_API_KEY=""
```

기본값은 로컬 Ollama를 가정합니다. OpenAI 호환 서버라면 `LLM_BASE_URL`과 `LLM_API_KEY`만 바꾸면 됩니다.
모델 서버가 없어도 앱은 동작하며, 이 경우 내장 오프라인 플래너가 기본 문서 초안을 생성합니다.

## 구조

- `app.py`: 정적 파일 서버 + 에이전트 플래너 API
- `static/index.html`: UI
- `static/app.js`: rhwp 연동, 문서 스냅샷 추출, 에이전트 명령 실행
- `static/styles.css`: 레이아웃/스타일

## 구현 메모

- 문서 렌더링/편집은 CDN의 `@rhwp/core`를 직접 사용합니다.
- 에이전트는 문서 스냅샷을 바탕으로 JSON 작업 계획을 반환합니다.
- 지원 작업:
  - `set_document_html`
  - `append_html`
  - `replace_paragraph_text`
  - `replace_paragraph_html`
  - `no_op`

## 한계

- 외부 HWP 문서를 부분 수정할 때 원본 복합 서식은 일부 잃을 수 있습니다.
- 현재는 업무 문서 초안 생성과 단락 단위 수정에 최적화되어 있습니다.
