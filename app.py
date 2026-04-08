#!/usr/bin/env python3

import json
import os
import re
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib import error, request


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
HOST = os.environ.get("OFFICE_AGENT_HOST", "127.0.0.1")
PORT = int(os.environ.get("OFFICE_AGENT_PORT", "8765"))
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "http://127.0.0.1:11434/v1").rstrip("/")
LLM_MODEL = os.environ.get("LLM_MODEL", "gemma4:latest")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")


SYSTEM_PROMPT = """당신은 한국어 오피스 문서 편집 에이전트다.

역할:
- 공문, 보고서, 기획안, 회의록, 제안서, 인사 문서 같은 한국어 업무 문서를 다룬다.
- 출력은 반드시 JSON 객체 하나만 반환한다.
- 자연어 설명만 하지 말고 실제 편집 계획을 함께 반환한다.

편집 원칙:
- 표, 제목, 번호 목록, 강조가 필요하면 HTML로 표현한다.
- 문서를 전면 재작성할 필요가 있으면 set_document_html을 사용한다.
- 기존 문서의 일부만 고치면 replace_paragraph_text 또는 replace_paragraph_html을 사용한다.
- 별도 수정이 불필요하면 no_op를 사용한다.
- 사실을 지어내지 말고, 사용자의 요청이 불명확하면 보수적으로 문안을 만든다.

반환 스키마:
{
  "reply": "사용자에게 보여줄 짧은 설명",
  "operations": [
    {
      "type": "set_document_html",
      "html": "<h1>...</h1><p>...</p>"
    },
    {
      "type": "append_html",
      "html": "<p>...</p>"
    },
    {
      "type": "replace_paragraph_text",
      "section": 0,
      "paragraph": 1,
      "text": "새 문단 텍스트"
    },
    {
      "type": "replace_paragraph_html",
      "section": 0,
      "paragraph": 1,
      "html": "<p><strong>수정</strong> 문단</p>"
    },
    {
      "type": "no_op",
      "reason": "변경할 내용 없음"
    }
  ]
}

제약:
- JSON 외 텍스트 금지.
- operations는 반드시 배열.
- 문단 인덱스는 제공된 스냅샷 기준을 사용한다.
- HTML은 table, thead, tbody, tr, th, td, h1, h2, h3, p, ul, ol, li, strong, em, br만 사용한다.
"""


def compact_document(document):
    paragraphs = document.get("paragraphs", [])
    compact_paragraphs = []
    for item in paragraphs[:120]:
        text = (item.get("text") or "").strip()
        compact_paragraphs.append(
            {
                "section": item.get("section", 0),
                "paragraph": item.get("paragraph", 0),
                "text": text[:800],
            }
        )
    return {
        "fileName": document.get("fileName", ""),
        "pageCount": document.get("pageCount", 0),
        "sectionCount": document.get("sectionCount", 0),
        "paragraphCount": len(paragraphs),
        "paragraphs": compact_paragraphs,
    }


def extract_json_object(text):
    text = text.strip()
    if text.startswith("{") and text.endswith("}"):
        return text
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("model did not return a JSON object")
    return match.group(0)


def validate_plan(plan):
    if not isinstance(plan, dict):
        raise ValueError("plan must be an object")
    operations = plan.get("operations")
    if not isinstance(operations, list):
        raise ValueError("operations must be a list")

    cleaned = []
    for op in operations:
        if not isinstance(op, dict):
            continue
        op_type = op.get("type")
        if op_type == "set_document_html" and isinstance(op.get("html"), str):
            cleaned.append({"type": op_type, "html": op["html"]})
        elif op_type == "append_html" and isinstance(op.get("html"), str):
            cleaned.append({"type": op_type, "html": op["html"]})
        elif op_type == "replace_paragraph_text":
            cleaned.append(
                {
                    "type": op_type,
                    "section": int(op.get("section", 0)),
                    "paragraph": int(op.get("paragraph", 0)),
                    "text": str(op.get("text", "")),
                }
            )
        elif op_type == "replace_paragraph_html":
            cleaned.append(
                {
                    "type": op_type,
                    "section": int(op.get("section", 0)),
                    "paragraph": int(op.get("paragraph", 0)),
                    "html": str(op.get("html", "")),
                }
            )
        elif op_type == "no_op":
            cleaned.append({"type": op_type, "reason": str(op.get("reason", ""))})

    return {
        "reply": str(plan.get("reply", "")).strip(),
        "operations": cleaned,
    }


def paragraph_texts(document):
    paragraphs = document.get("paragraphs", [])
    return [str(item.get("text", "")).strip() for item in paragraphs if str(item.get("text", "")).strip()]


def html_paragraphs(lines):
    return "".join(f"<p>{line}</p>" for line in lines if line.strip())


def build_weekly_report_html():
    return (
        "<h1>주간업무보고</h1>"
        "<p><strong>보고 개요</strong><br>이번 주 핵심 진행 사항과 다음 주 계획을 정리합니다.</p>"
        "<h2>핵심 요약</h2>"
        "<ul>"
        "<li>우선순위 과제 진행 상황을 점검했습니다.</li>"
        "<li>협업 이슈와 일정 리스크를 정리했습니다.</li>"
        "<li>다음 주 실행 항목을 담당자 기준으로 배정했습니다.</li>"
        "</ul>"
        "<h2>진행 현황</h2>"
        "<table><thead><tr><th>과제</th><th>진행률</th><th>현황</th><th>비고</th></tr></thead>"
        "<tbody>"
        "<tr><td>핵심 과제 A</td><td>80%</td><td>주요 산출물 초안 완료</td><td>검토 예정</td></tr>"
        "<tr><td>핵심 과제 B</td><td>60%</td><td>관계 부서 의견 수렴 중</td><td>일정 확인 필요</td></tr>"
        "<tr><td>운영 과제</td><td>100%</td><td>정기 운영 완료</td><td>특이사항 없음</td></tr>"
        "</tbody></table>"
        "<h2>다음 주 계획</h2>"
        "<ol>"
        "<li>핵심 과제 A 최종본 제출</li>"
        "<li>핵심 과제 B 협의 결과 반영</li>"
        "<li>리스크 항목 재점검 및 일정 확정</li>"
        "</ol>"
    )


def build_minutes_html():
    return (
        "<h1>회의록</h1>"
        "<p><strong>회의 개요</strong><br>회의 목적, 논의 사항, 후속 조치를 정리합니다.</p>"
        "<h2>주요 논의 내용</h2>"
        "<ol>"
        "<li>현재 진행 현황과 주요 이슈를 공유했습니다.</li>"
        "<li>의사결정이 필요한 안건의 대안을 비교했습니다.</li>"
        "<li>실행 일정과 담당 부서를 확정했습니다.</li>"
        "</ol>"
        "<h2>결정 사항</h2>"
        "<table><thead><tr><th>항목</th><th>결정 내용</th><th>책임</th><th>기한</th></tr></thead>"
        "<tbody>"
        "<tr><td>안건 1</td><td>제안안 기준으로 추진</td><td>담당팀</td><td>이번 주 내</td></tr>"
        "<tr><td>안건 2</td><td>추가 검토 후 확정</td><td>협업부서</td><td>차주 초</td></tr>"
        "</tbody></table>"
        "<h2>액션 아이템</h2>"
        "<table><thead><tr><th>No</th><th>작업</th><th>담당자</th><th>마감일</th></tr></thead>"
        "<tbody>"
        "<tr><td>1</td><td>결정 사항 문서화</td><td>간사</td><td>D+1</td></tr>"
        "<tr><td>2</td><td>추가 자료 취합</td><td>실무 담당</td><td>D+3</td></tr>"
        "<tr><td>3</td><td>후속 회의 일정 공지</td><td>PM</td><td>D+5</td></tr>"
        "</tbody></table>"
    )


def build_official_notice_html():
    return (
        "<h1>공문 초안</h1>"
        "<p>1. 귀 기관의 무궁한 발전을 기원합니다.</p>"
        "<p>2. 아래와 같이 관련 사항을 안내드리오니 업무에 참고하여 주시기 바랍니다.</p>"
        "<h2>주요 내용</h2>"
        "<table><thead><tr><th>구분</th><th>내용</th></tr></thead>"
        "<tbody>"
        "<tr><td>목적</td><td>업무 추진 기준 및 일정 공유</td></tr>"
        "<tr><td>대상</td><td>관련 부서 및 담당자</td></tr>"
        "<tr><td>협조 요청</td><td>기한 내 자료 제출 및 검토 의견 회신</td></tr>"
        "</tbody></table>"
        "<h2>협조 사항</h2>"
        "<ol>"
        "<li>담당자는 세부 자료를 검토하여 제출해 주시기 바랍니다.</li>"
        "<li>이견이 있는 경우 별도 의견서를 함께 회신해 주시기 바랍니다.</li>"
        "<li>추가 문의는 주관 부서로 연락해 주시기 바랍니다.</li>"
        "</ol>"
        "<p>붙임: 관련 자료 1부. 끝.</p>"
    )


def build_proposal_html():
    return (
        "<h1>제안서 초안</h1>"
        "<h2>제안 배경</h2>"
        "<p>현행 업무의 비효율과 개선 필요성을 정리합니다.</p>"
        "<h2>제안 목표</h2>"
        "<ul>"
        "<li>업무 처리 시간을 단축합니다.</li>"
        "<li>의사결정 품질을 높입니다.</li>"
        "<li>반복 업무를 자동화합니다.</li>"
        "</ul>"
        "<h2>실행 방안</h2>"
        "<table><thead><tr><th>단계</th><th>내용</th><th>성과</th></tr></thead>"
        "<tbody>"
        "<tr><td>1</td><td>현황 진단</td><td>문제 지점 도출</td></tr>"
        "<tr><td>2</td><td>시범 적용</td><td>개선 효과 검증</td></tr>"
        "<tr><td>3</td><td>전사 확산</td><td>표준 프로세스 정착</td></tr>"
        "</tbody></table>"
        "<h2>기대 효과</h2>"
        "<ol>"
        "<li>문서 작성 및 검토 리드타임 감소</li>"
        "<li>중복 커뮤니케이션 최소화</li>"
        "<li>업무 가시성 향상</li>"
        "</ol>"
    )


def build_polish_html(document):
    paragraphs = paragraph_texts(document)
    if not paragraphs:
        return (
            "<h1>문서 정리본</h1>"
            "<p>정리할 본문이 없어 기본 문서 구조만 생성했습니다.</p>"
        )
    polished = []
    for index, line in enumerate(paragraphs[:12], start=1):
        cleaned = re.sub(r"\s+", " ", line).strip(" -")
        polished.append(f"{index}. {cleaned}")
    return (
        "<h1>정리본</h1>"
        "<p>기존 문서를 보다 정돈된 업무 문체로 재구성했습니다.</p>"
        + html_paragraphs(polished)
    )


def fallback_plan(user_prompt, document, reason):
    prompt = user_prompt.strip()
    prompt_lower = prompt.lower()

    if any(keyword in prompt for keyword in ("주간업무보고", "주간 보고", "업무보고")):
        html = build_weekly_report_html()
        reply = "주간업무보고 형식으로 초안을 작성했습니다."
    elif any(keyword in prompt for keyword in ("회의록", "결정사항", "액션아이템")):
        html = build_minutes_html()
        reply = "회의록과 액션 아이템 표 형태로 정리했습니다."
    elif any(keyword in prompt for keyword in ("공문", "시행문", "협조 요청")):
        html = build_official_notice_html()
        reply = "공문체 초안으로 재구성했습니다."
    elif any(keyword in prompt for keyword in ("제안서", "기획안", "보고서")):
        html = build_proposal_html()
        reply = "제안서형 문서 뼈대를 작성했습니다."
    elif any(keyword in prompt for keyword in ("다듬어", "격식", "정리", "교정", "공문체")) or "polish" in prompt_lower:
        html = build_polish_html(document)
        reply = "기존 문서를 정돈된 문체로 재구성했습니다."
    else:
        html = (
            "<h1>업무 문서 초안</h1>"
            f"<p><strong>요청 사항</strong><br>{prompt}</p>"
            "<h2>핵심 내용</h2>"
            "<ul>"
            "<li>요청 목적과 배경을 정리합니다.</li>"
            "<li>실행 항목과 일정을 구분합니다.</li>"
            "<li>후속 조치와 담당자를 명확히 합니다.</li>"
            "</ul>"
            "<h2>실행 계획</h2>"
            "<table><thead><tr><th>항목</th><th>내용</th><th>기한</th></tr></thead>"
            "<tbody>"
            "<tr><td>1</td><td>초안 검토</td><td>즉시</td></tr>"
            "<tr><td>2</td><td>의견 반영</td><td>협의 후</td></tr>"
            "<tr><td>3</td><td>최종 확정</td><td>승인 후</td></tr>"
            "</tbody></table>"
        )
        reply = "기본 업무 문서 초안을 생성했습니다."

    return {
        "reply": f"{reply} 오프라인 플래너를 사용했습니다.",
        "operations": [{"type": "set_document_html", "html": html}],
        "meta": {"planner": "fallback", "reason": reason},
    }


def call_llm(user_prompt, document):
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "user_request": user_prompt,
                        "document": compact_document(document),
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "temperature": 0.3,
    }

    headers = {"Content-Type": "application/json"}
    if LLM_API_KEY:
        headers["Authorization"] = f"Bearer {LLM_API_KEY}"

    req = request.Request(
        f"{LLM_BASE_URL}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    with request.urlopen(req, timeout=90) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["choices"][0]["message"]["content"]
    return validate_plan(json.loads(extract_json_object(content)))


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def do_GET(self):
        if self.path == "/healthz":
            self._send_json({"ok": True, "model": LLM_MODEL, "baseUrl": LLM_BASE_URL, "planner": "llm+fallback"})
            return
        if self.path == "/":
            self.path = "/index.html"
        return super().do_GET()

    def do_POST(self):
        if self.path != "/api/plan":
            self._send_json({"ok": False, "error": "not found"}, status=HTTPStatus.NOT_FOUND)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8")
            payload = json.loads(body)
            user_prompt = str(payload.get("prompt", "")).strip()
            document = payload.get("document", {})
            if not user_prompt:
                raise ValueError("prompt is required")
            try:
                plan = call_llm(user_prompt, document)
                plan = {**plan, "meta": {"planner": "llm"}}
            except (error.HTTPError, error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
                plan = fallback_plan(user_prompt, document, str(exc))
            self._send_json({"ok": True, "plan": plan})
        except Exception as exc:
            self._send_json(
                {"ok": False, "error": "request_failed", "detail": str(exc)},
                status=HTTPStatus.BAD_REQUEST,
            )

    def log_message(self, format, *args):
        return

    def _send_json(self, payload, status=HTTPStatus.OK):
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Office Agent Staff listening on http://{HOST}:{PORT}")
    print(f"LLM endpoint: {LLM_BASE_URL}/chat/completions")
    print(f"Model: {LLM_MODEL}")
    server.serve_forever()


if __name__ == "__main__":
    main()
