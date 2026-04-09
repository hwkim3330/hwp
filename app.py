#!/usr/bin/env python3

import json
import os
import re
import shutil
import subprocess
from html.parser import HTMLParser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from urllib import error, request


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
HOST = os.environ.get("OFFICE_AGENT_HOST", "127.0.0.1")
PORT = int(os.environ.get("OFFICE_AGENT_PORT", "8765"))
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "http://127.0.0.1:11434/v1").rstrip("/")
LLM_MODEL = os.environ.get("LLM_MODEL", "gemma4:latest")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
LLM_TIMEOUT_SECONDS = float(os.environ.get("LLM_TIMEOUT_SECONDS", "20"))
LLM_MAX_TOKENS = int(os.environ.get("LLM_MAX_TOKENS", "900"))
HWPFORGE_CMD = os.environ.get("HWPFORGE_CMD", "hwpforge")
ENABLE_WEB_SEARCH = os.environ.get("ENABLE_WEB_SEARCH", "1") != "0"
ENABLE_VISION_EXPERIMENTS = os.environ.get("ENABLE_VISION_EXPERIMENTS", "1") != "0"
ENABLE_SYSTEM_ACTIONS = os.environ.get("ENABLE_SYSTEM_ACTIONS", "1") != "0"
ENABLE_FILE_AUTOMATION = os.environ.get("ENABLE_FILE_AUTOMATION", "1") != "0"
USER_AGENT = "Geulbit/1.0 (+https://github.com/hwkim3330/geulbit)"


SYSTEM_PROMPT = """당신은 한국어 오피스 문서 편집 에이전트다.

역할:
- 공문, 보고서, 기획안, 회의록, 제안서, 인사 문서 같은 한국어 업무 문서를 다룬다.
- 출력은 반드시 JSON 객체 하나만 반환한다.
- 자연어 설명만 하지 말고 실제 편집 계획을 함께 반환한다.

편집 원칙:
- 현재 작업 모드(mode)에 맞는 출력을 우선한다.
- writer 모드에서는 가능하면 구조화된 blocks를 우선 사용한다.
- 표, 제목, 번호 목록이 필요하면 blocks 또는 제한된 HTML로 표현한다.
- 문서를 전면 재작성할 필요가 있으면 set_document_blocks 또는 set_document_html을 사용한다.
- 기존 문서의 일부만 고치면 replace_paragraph_text, replace_paragraph_blocks 또는 replace_paragraph_html을 사용한다.
- notes 모드에서는 set_note_text를 사용한다.
- sheet 모드에서는 set_sheet_data를 사용한다.
- slides 모드에서는 set_slides를 사용한다.
- 별도 수정이 불필요하면 no_op를 사용한다.
- 사실을 지어내지 말고, 사용자의 요청이 불명확하면 보수적으로 문안을 만든다.

반환 스키마:
{
  "reply": "사용자에게 보여줄 짧은 설명",
  "operations": [
    {
      "type": "set_document_blocks",
      "blocks": [
        {"kind": "heading", "level": 1, "text": "제목"},
        {"kind": "paragraph", "text": "문단 내용"},
        {"kind": "bullets", "items": ["항목 1", "항목 2"]},
        {"kind": "table", "headers": ["항목", "내용"], "rows": [["1", "설명"]]}
      ]
    },
    {
      "type": "set_document_html",
      "html": "<h1>...</h1><p>...</p>"
    },
    {
      "type": "append_blocks",
      "blocks": [
        {"kind": "paragraph", "text": "추가 문단"}
      ]
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
      "type": "replace_paragraph_blocks",
      "section": 0,
      "paragraph": 1,
      "blocks": [
        {"kind": "paragraph", "text": "새 문단 텍스트"}
      ]
    },
    {
      "type": "replace_paragraph_html",
      "section": 0,
      "paragraph": 1,
      "html": "<p><strong>수정</strong> 문단</p>"
    },
    {
      "type": "set_note_text",
      "text": "메모장 전체 내용"
    },
    {
      "type": "set_sheet_data",
      "columns": ["항목", "담당", "상태"],
      "rows": [
        {"항목": "업무 1", "담당": "홍길동", "상태": "진행중"}
      ]
    },
    {
      "type": "set_slides",
      "slides": [
        {"title": "제목", "bullets": ["핵심 1", "핵심 2"]}
      ]
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
- blocks의 kind는 heading, paragraph, bullets, numbered, table만 사용한다.
- HTML은 table, thead, tbody, tr, th, td, h1, h2, h3, p, ul, ol, li, strong, em, br만 사용한다.
- notes는 일반 텍스트만 사용한다.
- sheet rows는 columns 키와 맞는 객체 배열로 반환한다.
- slides는 title 문자열과 bullets 문자열 배열을 사용한다.
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


def compact_sheet(sheet):
    columns = [str(column) for column in sheet.get("columns", [])[:16]]
    rows = []
    for item in sheet.get("rows", [])[:50]:
        if not isinstance(item, dict):
            continue
        row = {}
        for column in columns[:16]:
            row[column] = str(item.get(column, ""))[:400]
        rows.append(row)
    return {"columns": columns, "rows": rows}


def compact_slides(slides):
    compact = []
    for item in slides[:20]:
        if not isinstance(item, dict):
            continue
        bullets = [str(bullet)[:300] for bullet in item.get("bullets", [])[:8]]
        compact.append({"title": str(item.get("title", ""))[:200], "bullets": bullets})
    return compact


def extract_json_object(text):
    text = text.strip()
    if text.startswith("{") and text.endswith("}"):
        return text
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("model did not return a JSON object")
    return match.group(0)


def is_ollama_base_url():
    return "127.0.0.1:11434" in LLM_BASE_URL or "localhost:11434" in LLM_BASE_URL


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
        if op_type == "set_document_blocks":
            cleaned.append({"type": op_type, "blocks": normalize_blocks(op.get("blocks", []))})
        elif op_type == "set_document_html" and isinstance(op.get("html"), str):
            cleaned.append({"type": op_type, "html": op["html"]})
        elif op_type == "append_blocks":
            cleaned.append({"type": op_type, "blocks": normalize_blocks(op.get("blocks", []))})
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
        elif op_type == "replace_paragraph_blocks":
            cleaned.append(
                {
                    "type": op_type,
                    "section": int(op.get("section", 0)),
                    "paragraph": int(op.get("paragraph", 0)),
                    "blocks": normalize_blocks(op.get("blocks", [])),
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
        elif op_type == "set_note_text":
            cleaned.append({"type": op_type, "text": str(op.get("text", ""))})
        elif op_type == "set_sheet_data":
            columns = [str(column) for column in op.get("columns", [])[:16]]
            rows = []
            for item in op.get("rows", [])[:100]:
                if not isinstance(item, dict):
                    continue
                row = {}
                for column in columns:
                    row[column] = str(item.get(column, ""))
                rows.append(row)
            cleaned.append({"type": op_type, "columns": columns, "rows": rows})
        elif op_type == "set_slides":
            slides = []
            for slide in op.get("slides", [])[:20]:
                if not isinstance(slide, dict):
                    continue
                slides.append(
                    {
                        "title": str(slide.get("title", "")),
                        "bullets": [str(bullet) for bullet in slide.get("bullets", [])[:10]],
                    }
                )
            cleaned.append({"type": op_type, "slides": slides})
        elif op_type == "no_op":
            cleaned.append({"type": op_type, "reason": str(op.get("reason", ""))})

    return {
        "reply": str(plan.get("reply", "")).strip(),
        "operations": cleaned,
    }


def normalize_blocks(blocks):
    normalized = []
    for block in blocks[:80]:
        if not isinstance(block, dict):
            continue
        kind = str(block.get("kind", "")).strip()
        if kind == "heading":
            level = int(block.get("level", 1))
            normalized.append(
                {
                    "kind": "heading",
                    "level": max(1, min(3, level)),
                    "text": str(block.get("text", ""))[:2000],
                }
            )
        elif kind == "paragraph":
            normalized.append({"kind": "paragraph", "text": str(block.get("text", ""))[:4000]})
        elif kind in {"bullets", "numbered"}:
            items = [str(item)[:800] for item in block.get("items", [])[:20] if str(item).strip()]
            normalized.append({"kind": kind, "items": items})
        elif kind == "table":
            headers = [str(item)[:200] for item in block.get("headers", [])[:12]]
            rows = []
            for row in block.get("rows", [])[:40]:
                if not isinstance(row, list):
                    continue
                rows.append([str(cell)[:400] for cell in row[:12]])
            normalized.append({"kind": "table", "headers": headers, "rows": rows})
    return normalized


def paragraph_texts(document):
    paragraphs = document.get("paragraphs", [])
    return [str(item.get("text", "")).strip() for item in paragraphs if str(item.get("text", "")).strip()]


def html_paragraphs(lines):
    return "".join(f"<p>{line}</p>" for line in lines if line.strip())


def blocks_to_html(blocks):
    parts = []
    for block in normalize_blocks(blocks):
        kind = block["kind"]
        if kind == "heading":
            level = block.get("level", 1)
            tag = f"h{max(1, min(3, level))}"
            parts.append(f"<{tag}>{escape_html(block.get('text', ''))}</{tag}>")
        elif kind == "paragraph":
            parts.append(f"<p>{escape_html(block.get('text', '')).replace(chr(10), '<br>')}</p>")
        elif kind == "bullets":
            items = "".join(f"<li>{escape_html(item)}</li>" for item in block.get("items", []))
            parts.append(f"<ul>{items}</ul>")
        elif kind == "numbered":
            items = "".join(f"<li>{escape_html(item)}</li>" for item in block.get("items", []))
            parts.append(f"<ol>{items}</ol>")
        elif kind == "table":
            headers = "".join(f"<th>{escape_html(item)}</th>" for item in block.get("headers", []))
            rows = []
            for row in block.get("rows", []):
                cells = "".join(f"<td>{escape_html(cell)}</td>" for cell in row)
                rows.append(f"<tr>{cells}</tr>")
            parts.append(f"<table><thead><tr>{headers}</tr></thead><tbody>{''.join(rows)}</tbody></table>")
    return "".join(parts)


def escape_html(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_generic_writer_blocks(prompt):
    return [
        {"kind": "heading", "level": 1, "text": "업무 문서 초안"},
        {"kind": "paragraph", "text": f"요청 사항\n{prompt}"},
        {
            "kind": "heading",
            "level": 2,
            "text": "핵심 내용",
        },
        {
            "kind": "bullets",
            "items": [
                "요청 목적과 배경을 정리합니다.",
                "실행 항목과 일정을 구분합니다.",
                "후속 조치와 담당자를 명확히 합니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "실행 계획"},
        {
            "kind": "table",
            "headers": ["항목", "내용", "기한"],
            "rows": [
                ["1", "초안 검토", "즉시"],
                ["2", "의견 반영", "협의 후"],
                ["3", "최종 확정", "승인 후"],
            ],
        },
    ]


def build_research_writer_blocks(prompt, search_results):
    bullets = []
    references = []
    for index, item in enumerate(search_results[:5], start=1):
        snippet = item.get("snippet", "").strip()
        summary = item["title"]
        if snippet:
            summary = f"{summary}: {snippet[:180]}"
        bullets.append(summary)
        references.append([str(index), item["title"], item["url"]])
    return [
        {"kind": "heading", "level": 1, "text": "조사 기반 초안"},
        {
            "kind": "paragraph",
            "text": f"요청 사항\n{prompt.strip()}\n\n아래 내용은 웹 검색 결과를 바탕으로 정리한 작업용 초안입니다.",
        },
        {"kind": "heading", "level": 2, "text": "검색 요약"},
        {"kind": "numbered", "items": bullets or ["유효한 검색 결과를 찾지 못했습니다."]},
        {"kind": "heading", "level": 2, "text": "활용 방향"},
        {
            "kind": "bullets",
            "items": [
                "핵심 주장과 배경 설명을 검색 결과와 현재 목적에 맞게 다시 정리합니다.",
                "출처가 필요한 문장은 참고 링크를 기반으로 수동 검증 후 반영합니다.",
                "최종 문서에서는 초안 문체를 목적에 맞는 보고서 또는 논문 문체로 다듬습니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "참고 링크"},
        {
            "kind": "table",
            "headers": ["No", "제목", "링크"],
            "rows": references or [["1", "검색 결과 없음", ""]],
        },
    ]


class BlockHtmlParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.blocks = []
        self.stack = []
        self.current_table = None
        self.current_row = None
        self.current_cell = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        self.stack.append(tag)
        if tag == "table":
            self.current_table = {"kind": "table", "headers": [], "rows": []}
        elif tag == "tr":
            self.current_row = []
        elif tag in {"th", "td"}:
            self.current_cell = []

    def handle_endtag(self, tag):
        tag = tag.lower()
        text = self._flush_text_for(tag)
        if tag in {"h1", "h2", "h3"} and text:
            self.blocks.append({"kind": "heading", "level": int(tag[1]), "text": text})
        elif tag == "p" and text:
            self.blocks.append({"kind": "paragraph", "text": text})
        elif tag == "ul" and text:
            self.blocks.append({"kind": "bullets", "items": [item for item in text.split("\n") if item]})
        elif tag == "ol" and text:
            self.blocks.append({"kind": "numbered", "items": [item for item in text.split("\n") if item]})
        elif tag in {"th", "td"}:
            self.current_row.append(text)
        elif tag == "tr" and self.current_table is not None and self.current_row is not None:
            if self.current_table["headers"]:
                self.current_table["rows"].append(self.current_row)
            else:
                self.current_table["headers"] = self.current_row
            self.current_row = None
        elif tag == "table" and self.current_table is not None:
            self.blocks.append(self.current_table)
            self.current_table = None
        if self.stack:
            self.stack.pop()

    def handle_data(self, data):
        if not self.stack:
            return
        tag = self.stack[-1]
        if tag in {"th", "td"}:
            self.current_cell.append(data)
        else:
            text = data.strip()
            if text:
                self.current_cell.append(text)

    def _flush_text_for(self, tag):
        text = " ".join(part.strip() for part in self.current_cell if part.strip()).strip()
        if tag in {"li"} and text:
            self.current_cell = []
            return text
        if tag in {"ul", "ol"}:
            text = "\n".join(part.strip() for part in self.current_cell if part.strip())
        self.current_cell = []
        return text


def html_to_blocks(html):
    parser = BlockHtmlParser()
    parser.feed(str(html or ""))
    blocks = normalize_blocks(parser.blocks)
    return blocks or build_generic_writer_blocks("내용 없음")


class DuckDuckGoHtmlParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.results = []
        self._capture_title = False
        self._capture_snippet = False
        self._current = None

    def handle_starttag(self, tag, attrs):
        attrs_map = dict(attrs)
        class_name = attrs_map.get("class", "")
        if tag == "a" and "result__a" in class_name:
            self._capture_title = True
            self._current = {"title": "", "url": attrs_map.get("href", ""), "snippet": ""}
        elif tag in {"a", "div"} and "result__snippet" in class_name and self._current:
            self._capture_snippet = True

    def handle_endtag(self, tag):
        if tag == "a" and self._capture_title:
            self._capture_title = False
            if self._current and self._current["title"]:
                self.results.append(self._current)
                self._current = None
        if tag in {"a", "div"} and self._capture_snippet:
            self._capture_snippet = False

    def handle_data(self, data):
        text = re.sub(r"\s+", " ", data or "").strip()
        if not text or not self._current:
            return
        if self._capture_title:
            self._current["title"] = f"{self._current['title']} {text}".strip()
        elif self._capture_snippet:
            self._current["snippet"] = f"{self._current['snippet']} {text}".strip()


def should_use_web_search(prompt):
    lowered = prompt.lower()
    return any(
        keyword in lowered
        for keyword in (
            "검색",
            "논문",
            "참고문헌",
            "레퍼런스",
            "자료 찾아",
            "웹 검색",
            "조사",
            "research",
            "citation",
            "source",
        )
    )


def web_search(query, max_results=5):
    if not ENABLE_WEB_SEARCH:
        return []
    url = f"https://html.duckduckgo.com/html/?q={quote(query)}"
    req = request.Request(url, headers={"User-Agent": USER_AGENT})
    with request.urlopen(req, timeout=10) as response:
        html = response.read().decode("utf-8", errors="ignore")
    parser = DuckDuckGoHtmlParser()
    parser.feed(html)
    unique = []
    seen = set()
    for item in parser.results:
        url_value = normalize_search_url(item.get("url", "").strip())
        title = item.get("title", "").strip()
        if not url_value or not title or url_value in seen:
            continue
        seen.add(url_value)
        unique.append({"title": title[:240], "url": url_value[:800], "snippet": item.get("snippet", "")[:400]})
        if len(unique) >= max_results:
            break
    return unique


def format_search_context(results):
    if not results:
        return ""
    lines = ["web_search_results:"]
    for index, item in enumerate(results, start=1):
        lines.append(f"{index}. {item['title']}")
        lines.append(f"   url: {item['url']}")
        if item.get("snippet"):
            lines.append(f"   snippet: {item['snippet']}")
    return "\n".join(lines)


def normalize_search_url(url_value):
    if not url_value:
        return ""
    if url_value.startswith("//"):
        url_value = f"https:{url_value}"
    parsed = urlparse(url_value)
    if "duckduckgo.com" in parsed.netloc:
        target = parse_qs(parsed.query).get("uddg", [""])[0]
        if target:
            return unquote(target)
    return url_value


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


def build_note_text(prompt, workspace):
    current = str(workspace.get("noteText", "")).strip()
    header = f"[메모] {prompt.strip()}"
    sections = [header, "", "핵심 메모", "- 배경 정리", "- 해야 할 일", "- 확인 필요 사항"]
    if current:
        sections.extend(["", "기존 메모 참고", current[:1500]])
    return "\n".join(sections).strip()


def build_sheet_data(prompt):
    columns = ["항목", "담당", "상태", "기한", "우선순위", "비고"]
    rows = [
        {"항목": "요청 분석", "담당": "본인", "상태": "완료", "기한": "즉시", "우선순위": "상", "비고": prompt[:80]},
        {"항목": "실행 계획 수립", "담당": "담당자 지정", "상태": "진행중", "기한": "오늘", "우선순위": "상", "비고": "세부 일정 확정 필요"},
        {"항목": "결과 검토", "담당": "검토자", "상태": "대기", "기한": "차주", "우선순위": "중", "비고": "리스크 재확인"},
    ]
    return {"columns": columns, "rows": rows}


def build_slides_data(prompt):
    base = prompt.strip() or "오피스 에이전트 업무 제안"
    return [
        {"title": "개요", "bullets": [base, "목적과 배경", "핵심 메시지"]},
        {"title": "현황", "bullets": ["현재 문제 정의", "주요 병목", "영향 범위"]},
        {"title": "해결안", "bullets": ["제안 방식", "기대 효과", "필요 자원"]},
        {"title": "실행 계획", "bullets": ["단계별 일정", "담당 역할", "다음 액션"]},
    ]


def is_mode_compatible(mode, operations):
    allowed_map = {
        "writer": {
            "set_document_blocks",
            "set_document_html",
            "append_blocks",
            "append_html",
            "replace_paragraph_text",
            "replace_paragraph_blocks",
            "replace_paragraph_html",
            "no_op",
        },
        "notes": {"set_note_text", "no_op"},
        "sheet": {"set_sheet_data", "no_op"},
        "slides": {"set_slides", "no_op"},
    }
    allowed = allowed_map.get(mode, allowed_map["writer"])
    return any(op.get("type") in allowed and op.get("type") != "no_op" for op in operations)


def finalize_plan(mode, user_prompt, document, workspace, plan):
    operations = plan.get("operations", [])
    if is_mode_compatible(mode, operations):
        return plan

    if mode == "notes" and all(op.get("type") == "no_op" for op in operations):
        return {
            "reply": "메모장 초안을 생성했습니다.",
            "operations": [{"type": "set_note_text", "text": build_note_text(user_prompt, workspace)}],
        }
    if mode == "sheet":
        sheet = build_sheet_data(user_prompt)
        return {
            "reply": "시트 초안을 생성했습니다.",
            "operations": [{"type": "set_sheet_data", "columns": sheet["columns"], "rows": sheet["rows"]}],
        }
    if mode == "slides":
        return {
            "reply": "슬라이드 초안을 생성했습니다.",
            "operations": [{"type": "set_slides", "slides": build_slides_data(user_prompt)}],
        }
    if mode == "writer" and all(op.get("type") == "no_op" for op in operations):
        return fallback_plan(user_prompt, document, workspace, "writer_noop_promoted")
    return plan


def fallback_plan(user_prompt, document, workspace, reason, search_results=None):
    prompt = user_prompt.strip()
    prompt_lower = prompt.lower()
    mode = str(workspace.get("mode", "writer"))

    if mode == "notes":
        return {
            "reply": "메모장 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_note_text", "text": build_note_text(prompt, workspace)}],
            "meta": {"planner": "fallback", "reason": reason},
        }

    if mode == "sheet":
        sheet = build_sheet_data(prompt)
        return {
            "reply": "시트 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_sheet_data", "columns": sheet["columns"], "rows": sheet["rows"]}],
            "meta": {"planner": "fallback", "reason": reason},
        }

    if mode == "slides":
        return {
            "reply": "슬라이드 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_slides", "slides": build_slides_data(prompt)}],
            "meta": {"planner": "fallback", "reason": reason},
        }

    if search_results and should_use_web_search(prompt):
        blocks = build_research_writer_blocks(prompt, search_results)
        return {
            "reply": "검색 결과를 바탕으로 연구용 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": blocks}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results)},
        }

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
        blocks = build_generic_writer_blocks(prompt)
        if search_results:
            blocks.extend(
                [
                    {"kind": "heading", "level": 2, "text": "검색 참고"},
                    {
                        "kind": "bullets",
                        "items": [f"{item['title']} | {item['url']}" for item in search_results[:5]],
                    },
                ]
            )
        reply = "기본 업무 문서 초안을 생성했습니다."
        return {
            "reply": f"{reply} 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": blocks}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
        }

    return {
        "reply": f"{reply} 오프라인 플래너를 사용했습니다.",
        "operations": [{"type": "set_document_html", "html": html}],
        "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
    }


def call_llm(user_prompt, document):
    search_results = []
    if should_use_web_search(user_prompt):
        try:
            search_results = web_search(user_prompt, max_results=5)
        except Exception:
            search_results = []
    user_content = json.dumps(
        {
            "user_request": user_prompt,
            "document": compact_document(document),
            "mode": document.get("mode", "writer"),
            "workspace": {
                "noteText": str(document.get("noteText", ""))[:3000],
                "sheet": compact_sheet(document.get("sheet", {})),
                "slides": compact_slides(document.get("slides", [])),
            },
            "web_search": search_results,
        },
        ensure_ascii=False,
    )
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]
    if is_ollama_base_url():
        return call_ollama_llm(messages), search_results

    payload = {
        "model": LLM_MODEL,
        "messages": messages,
        "temperature": 0.2,
        "max_tokens": LLM_MAX_TOKENS,
        "response_format": {"type": "json_object"},
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
    with request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["choices"][0]["message"]["content"]
    return validate_plan(json.loads(extract_json_object(content))), search_results


def call_ollama_llm(messages):
    payload = {
        "model": LLM_MODEL,
        "messages": messages,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.1,
            "num_predict": LLM_MAX_TOKENS,
        },
    }
    req = request.Request(
        "http://127.0.0.1:11434/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["message"]["content"]
    return validate_plan(json.loads(extract_json_object(content)))


def resolve_hwpforge_cmd():
    explicit = Path(HWPFORGE_CMD).expanduser()
    if explicit.is_file():
        return str(explicit)
    found = shutil.which(HWPFORGE_CMD)
    if found:
        return found
    local_candidates = [
        Path.home() / ".cargo" / "bin" / "hwpforge",
        Path("/Users/parksik/office-agent-sources/HwpForge/target/debug/hwpforge"),
        Path("/Users/parksik/office-agent-sources/HwpForge/target/debug/hwpforge-bindings-cli"),
    ]
    for candidate in local_candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def hwpforge_status():
    command = resolve_hwpforge_cmd()
    if not command:
        return {"available": False, "command": None, "detail": "hwpforge CLI not found"}
    try:
        result = subprocess.run(
            [command, "--help"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except Exception as exc:
        return {"available": False, "command": command, "detail": str(exc)}
    return {
        "available": result.returncode == 0,
        "command": command,
        "detail": (result.stdout or result.stderr).splitlines()[0] if (result.stdout or result.stderr) else "",
    }


def permission_registry():
    return [
        {
            "id": "documents.write",
            "label": "Document Write",
            "status": "granted",
            "detail": "Writer / Notes / Sheet / Slides 편집 허용",
        },
        {
            "id": "files.automation",
            "label": "File Automation",
            "status": "granted" if ENABLE_FILE_AUTOMATION else "disabled",
            "detail": "로컬 문서 로드 및 내보내기",
        },
        {
            "id": "vision.experiments",
            "label": "Vision Experiments",
            "status": "granted" if ENABLE_VISION_EXPERIMENTS else "disabled",
            "detail": "카메라 미리보기 및 얼굴 중심 정렬 실험",
        },
        {
            "id": "web.search",
            "label": "Web Search",
            "status": "disabled" if not ENABLE_WEB_SEARCH else "granted",
            "detail": "검색 기반 문서 보강 및 연구 컨텍스트 주입",
        },
        {
            "id": "system.actions",
            "label": "System Actions",
            "status": "disabled" if not ENABLE_SYSTEM_ACTIONS else "experimental",
            "detail": "open_url / reveal_path / open_app 제한 액션 허용",
        },
    ]


def tool_registry():
    hwpforge = hwpforge_status()
    return [
        {
            "id": "writer_blocks",
            "label": "Writer Blocks",
            "category": "documents",
            "status": "ready",
            "detail": "구조화 문서 블록 적용 엔진",
        },
        {
            "id": "hwpforge",
            "label": "HwpForge",
            "category": "format",
            "status": "ready" if hwpforge["available"] else "offline",
            "detail": hwpforge["detail"],
        },
        {
            "id": "gemma4_planner",
            "label": "Gemma 4 Planner",
            "category": "llm",
            "status": "ready" if is_ollama_base_url() else "partial",
            "detail": f"{LLM_MODEL} via {LLM_BASE_URL}",
        },
        {
            "id": "vision_lab",
            "label": "Vision Lab",
            "category": "vision",
            "status": "ready" if ENABLE_VISION_EXPERIMENTS else "disabled",
            "detail": "Electron 실험실 기반 카메라 / 정렬 테스트",
        },
        {
            "id": "web_search",
            "label": "Web Search",
            "category": "research",
            "status": "disabled" if not ENABLE_WEB_SEARCH else "ready",
            "detail": "DuckDuckGo HTML 검색 결과를 Gemma 4 컨텍스트에 주입",
        },
        {
            "id": "system_actions",
            "label": "System Actions",
            "category": "automation",
            "status": "disabled" if not ENABLE_SYSTEM_ACTIONS else "experimental",
            "detail": "권한 게이트 뒤 open_url / reveal_path / open_app 실행",
        },
    ]


def runtime_registry():
    return {
        "llm": {"model": LLM_MODEL, "baseUrl": LLM_BASE_URL, "ollama": is_ollama_base_url()},
        "permissions": permission_registry(),
        "tools": tool_registry(),
    }


def run_system_action(action, payload):
    if not ENABLE_SYSTEM_ACTIONS:
        raise PermissionError("system actions disabled")
    if action == "open_url":
        target = str(payload.get("url", "")).strip()
        if not target.startswith(("http://", "https://")):
            raise ValueError("invalid url")
        subprocess.run(["open", target], check=True)
        return {"action": action, "target": target}
    if action == "reveal_path":
        target = str(payload.get("path", "")).strip()
        if not target:
            raise ValueError("path required")
        subprocess.run(["open", "-R", target], check=True)
        return {"action": action, "target": target}
    if action == "open_app":
        target = str(payload.get("app", "")).strip()
        if not target:
            raise ValueError("app required")
        subprocess.run(["open", "-a", target], check=True)
        return {"action": action, "target": target}
    raise ValueError(f"unknown action: {action}")


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def do_GET(self):
        if self.path == "/healthz":
            self._send_json(
                {
                    "ok": True,
                    "model": LLM_MODEL,
                    "baseUrl": LLM_BASE_URL,
                    "planner": "llm+fallback",
                    "hwpforge": hwpforge_status(),
                }
            )
            return
        if self.path == "/api/runtime":
            self._send_json({"ok": True, "runtime": runtime_registry()})
            return
        if self.path == "/":
            self.path = "/index.html"
        return super().do_GET()

    def do_POST(self):
        if self.path == "/api/search":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                query = str(payload.get("query", "")).strip()
                if not query:
                    raise ValueError("query is required")
                self._send_json({"ok": True, "results": web_search(query, max_results=5)})
            except Exception as exc:
                self._send_json({"ok": False, "error": "search_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/system-action":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                action = str(payload.get("action", "")).strip()
                result = run_system_action(action, payload)
                self._send_json({"ok": True, "result": result})
            except Exception as exc:
                self._send_json({"ok": False, "error": "system_action_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/hwpforge/status":
            self._send_json({"ok": True, "hwpforge": hwpforge_status()})
            return
        if self.path != "/api/plan":
            self._send_json({"ok": False, "error": "not found"}, status=HTTPStatus.NOT_FOUND)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8")
            payload = json.loads(body)
            user_prompt = str(payload.get("prompt", "")).strip()
            document = payload.get("document", {})
            workspace = {
                "mode": payload.get("mode", "writer"),
                "noteText": payload.get("noteText", ""),
                "sheet": payload.get("sheet", {}),
                "slides": payload.get("slides", []),
            }
            if not user_prompt:
                raise ValueError("prompt is required")
            try:
                llm_document = {**document, **workspace}
                plan, search_results = call_llm(user_prompt, llm_document)
                plan = finalize_plan(workspace["mode"], user_prompt, document, workspace, plan)
                meta = {"planner": "llm"}
                if search_results:
                    meta["search_results"] = len(search_results)
                plan = {**plan, "meta": meta}
            except (error.HTTPError, error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
                fallback_results = []
                if should_use_web_search(user_prompt):
                    try:
                        fallback_results = web_search(user_prompt, max_results=5)
                    except Exception:
                        fallback_results = []
                plan = fallback_plan(user_prompt, document, workspace, str(exc), fallback_results)
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
    print(f"Geulbit listening on http://{HOST}:{PORT}")
    print(f"LLM endpoint: {LLM_BASE_URL}/chat/completions")
    print(f"Model: {LLM_MODEL}")
    print(f"LLM timeout: {LLM_TIMEOUT_SECONDS}s")
    server.serve_forever()


if __name__ == "__main__":
    main()
