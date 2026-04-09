# hwp + ONLYOFFICE Integration Plan

## 결론

`Google Docs`나 `Office Online` 코드를 가져오는 방향은 배제한다.

대신 `ONLYOFFICE Docs`를 외부 편집 엔진으로 두고, `hwp`는 다음 역할에 집중한다.

- 한국어 문서 에이전트 런타임
- HWP/HWPX 입출력과 변환 계층
- 검색, 비전, 시스템 액션, 세션 로그
- 연구노트/보고서/회의록/슬라이드 워크플로 GUI

## 근거

공식 자료 기준으로 `ONLYOFFICE Docs`는 다음을 제공한다.

- 웹 기반 문서/시트/프레젠테이션 편집기
- OOXML 중심 포맷 호환
- 임베드 가능한 `DocsAPI.DocEditor`
- 변환 API
- 오픈소스 `DocumentServer`

## 참고 공식 자료

- ONLYOFFICE DocumentServer
  - https://github.com/ONLYOFFICE/DocumentServer
- ONLYOFFICE Docs API
  - https://api.onlyoffice.com/docs
- Embedding
  - https://api.onlyoffice.com/docs/docs-api/more-information/faq/embedding/
- Conversion API
  - https://api.onlyoffice.com/docs/docs-api/using-wopi/conversion-api/
- How it works / converting
  - https://api.onlyoffice.com/docs/docs-api/get-started/how-it-works/converting-and-downloading-file/

## 목표 구조

### 1. hwp Frontend

현재 `static/index.html`, `static/app.js`, `static/styles.css`를 유지한다.

여기에 다음 GUI를 추가한다.

- 편집 엔진 선택
  - Native Writer
  - ONLYOFFICE Writer
- 문서 열기 방식 선택
  - HWP/HWPX
  - DOCX/XLSX/PPTX
- 에이전트 적용 범위
  - 현재 문서
  - 선택 영역
  - 새 초안

### 2. hwp Backend

현재 `app.py`를 오케스트레이터로 유지한다.

역할:

- 세션 로그
- 검색
- 시스템 액션
- Gemma 4 플래너
- HwpForge 감지 및 변환
- ONLYOFFICE 문서 설정 생성

추가할 API:

- `POST /api/onlyoffice/config`
- `POST /api/onlyoffice/callback`
- `POST /api/convert/docx-to-hwpx`
- `POST /api/convert/hwpx-to-docx`

### 3. ONLYOFFICE Server

별도 서비스로 둔다.

역할:

- DOCX/XLSX/PPTX 실제 편집
- 실시간 협업
- 저장 시 callback
- 문서 변환

즉 `hwp` 앱 안에 엔진을 복붙하는 게 아니라, 외부 문서 서버로 붙인다.

## 권장 1차 통합 범위

### Writer

- `DOCX` 문서는 ONLYOFFICE로 편집
- `HWP/HWPX`는 현재 native writer 유지
- 필요 시 `HWPX -> DOCX` 변환 후 ONLYOFFICE로 열기
- 저장 후 다시 `DOCX -> HWPX` 경로 제공

### Sheet

- `XLSX`는 ONLYOFFICE 우선
- `CSV`는 현재 가벼운 시트 유지

### Slides

- `PPTX`는 ONLYOFFICE 우선
- 현재 내부 슬라이드 초안 생성기는 유지

## 왜 이 구성이 맞는가

### 가져오지 말아야 할 것

- Google Docs 제품 코드
- Office Online 제품 코드

이 둘은 공개 오픈소스가 아니고, 직접 가져다 붙이는 전략이 성립하지 않는다.

### 가져와야 할 것

- ONLYOFFICE 편집기
- HwpForge
- hwp-rs
- rhwp
- claw-code 런타임 패턴

이 조합이면 `hwp`는 범용 오피스와 한국어 문서 에이전트를 동시에 가져갈 수 있다.

## 실행 단계

### Phase 1

- ONLYOFFICE를 iframe 기반으로 임베드
- `DOCX/XLSX/PPTX` 문서를 ONLYOFFICE에서 열기
- `hwp` 우측 패널의 에이전트와 세션 로그는 그대로 유지

### Phase 2

- 문서 저장 callback 구현
- 검색/에이전트 결과를 OOXML 편집 영역으로 반영
- 선택 영역 기준 수정 요청 지원

### Phase 3

- `HWPX <-> DOCX` 변환 자동화
- 연구노트/보고서/회의록 GUI 워크플로를 ONLYOFFICE와 연결
- 멀티 사용자 협업 상태를 `hwp` Command Center에 표시

## 현재 판단

지금 가장 먼저 할 일은 이것이다.

1. `ONLYOFFICE`를 별도 로컬 서비스로 붙인다.
2. `Writer/Sheet/Slides` 중 OOXML 계열은 그쪽으로 보낸다.
3. `hwp`는 에이전트와 한국어 문서 계층에 집중한다.

즉 `오피스 엔진`과 `에이전트 OS`를 분리하는 게 맞다.
