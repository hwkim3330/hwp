# Geulbit Ecosystem Plan

## 목표

Geulbit을 단순 브라우저 편집기에서 벗어나, 한국어 문서 생태계를 실제로 다루는 로컬 오피스 워크스테이션으로 끌어올린다.

핵심 방향은 다음 두 축이다.

- `HwpForge`를 중심으로 `HWPX` 읽기/쓰기/변환 파이프라인 강화
- `hwp-rs`를 이용해 구형 `HWP` 읽기 경로를 보강

## 조사 결과

### 1. HwpForge

- 로컬 경로: `/Users/parksik/office-agent-sources/HwpForge`
- 최근 확인 커밋: `799d908` on `2026-03-24`
- 라이선스: `MIT OR Apache-2.0`
- 강점:
  - `HWPX` 읽기/쓰기 지원
  - `HWP5(.hwp)` 읽기 경로 제공
  - `Markdown`, `JSON`, `CLI`, `MCP` 흐름이 이미 있음
  - AI 편집 워크플로우를 전제로 설계됨
- 가져올 가치가 큰 부분:
  - `to-md`, `to-json`, `patch`, `from-json` 같은 구조화 편집 흐름
  - HWPX 섹션 단위 패치 전략
  - MCP/CLI 도구 설계

### 2. hwp-rs

- 로컬 경로: `/Users/parksik/office-agent-sources/hwp-rs`
- 최근 확인 커밋: `e41e7b4` on `2022-11-11`
- 라이선스: `Apache-2.0`
- 강점:
  - 구형 바이너리 `HWP` 파서
  - Python 바인딩 `libhwp` 제공
  - 문단, 표, 첨부 리소스 추출 경로 존재
- 가져올 가치가 큰 부분:
  - 레거시 `.hwp` 문서 텍스트/표 추출
  - Python 백엔드 보강용 브리지 후보

## 제품 관점 결론

당장 메인 엔진은 `rhwp` 브라우저 편집 + 현재 로컬 에이전트 구조를 유지한다.

그 위에 아래 순서로 붙인다.

1. `HwpForge`를 서버 측 변환/검증 엔진으로 도입
2. `HWPX <-> Markdown/JSON` 왕복 경로를 앱 기능으로 노출
3. `hwp-rs`는 구형 `.hwp` 읽기 보강용으로 제한 도입

## 통합 원칙

- 프런트는 계속 빠른 로컬 워크스테이션 역할을 유지한다.
- 포맷 정합성과 변환은 Rust 도구에 넘긴다.
- LLM은 문서를 직접 만지지 않고 구조화 명령 또는 JSON DSL만 생성한다.
- `.hwp` 레거시 포맷은 우선 읽기 중심으로 다룬다.

## 실제 작업 순서

### Phase 1. HWPX 서버 파이프라인

- Python 서버에서 Rust CLI를 호출할 수 있는 브리지 추가
- `HWPX -> Markdown`
- `HWPX -> JSON`
- `JSON -> HWPX`
- `HWPX patch`

완료 기준:

- 기존 문서를 열고 JSON 섹션 편집 후 새 `hwpx`로 저장 가능

### Phase 2. Writer 문서 DSL 정착

- 현재 `set_document_blocks`를 확장
- heading, paragraph, bullets, numbered, table 외에:
  - image
  - caption
  - page break
  - section break
  - style token

완료 기준:

- Writer 에이전트 출력이 HTML보다 DSL 중심으로 이동

### Phase 3. 레거시 HWP 읽기 강화

- `hwp-rs` 또는 `libhwp` 경로 테스트
- `.hwp` 파일에서:
  - 문단
  - 표
  - 첨부 바이너리
  - 메타데이터
  추출

완료 기준:

- `.hwp`를 열었을 때 단순 렌더에 그치지 않고 구조 요약과 편집용 변환을 제공

### Phase 4. 앱 품질

- 최근 문서
- 자동 복구
- 프로젝트 번들 저장
- 배포용 Electron 업데이트 경로
- 서명 / notarization

## 지금 바로 해야 할 일

1. `HwpForge` CLI 또는 라이브러리를 로컬에서 빌드한다.
2. 샘플 `hwpx`를 앱에서 JSON 편집 후 다시 `hwpx`로 재생성하는 최소 경로를 만든다.
3. 그 뒤에 `hwp-rs`로 구형 `hwp` 읽기 보강을 붙인다.

## 하지 않을 일

- 지금 단계에서 무리하게 네이티브 macOS 앱으로 전환하지 않는다.
- `.docx/.xlsx/.pptx` 원본 서식 완전 보존을 먼저 약속하지 않는다.
- LLM이 직접 바이너리 포맷을 조작하게 두지 않는다.
