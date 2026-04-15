#!/usr/bin/env python3

import base64
import json
import mimetypes
import os
import re
import shutil
import subprocess
import threading
import time
import uuid
from html.parser import HTMLParser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from urllib import error, request


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
ONLYOFFICE_RUNTIME_DIR = ROOT / ".runtime" / "onlyoffice"
MEMORY_RUNTIME_PATH = ROOT / ".runtime" / "memory.json"
HOST = os.environ.get("OFFICE_AGENT_HOST", "127.0.0.1")
PORT = int(os.environ.get("OFFICE_AGENT_PORT", "8765"))
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "http://127.0.0.1:11434/v1").rstrip("/")
LLM_MODEL = os.environ.get("LLM_MODEL", "gemma4:latest")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
LLM_TIMEOUT_SECONDS = float(os.environ.get("LLM_TIMEOUT_SECONDS", "20"))
LLM_MAX_TOKENS = int(os.environ.get("LLM_MAX_TOKENS", "900"))
MLX_EXPERIMENTAL_BASE_URL = os.environ.get("MLX_EXPERIMENTAL_BASE_URL", "http://127.0.0.1:8081/v1").rstrip("/")
MLX_EXPERIMENTAL_MODEL = os.environ.get(
    "MLX_EXPERIMENTAL_MODEL", "Jiunsong/supergemma4-26b-uncensored-mlx-4bit-v2"
)
MLX_EXPERIMENTAL_VENV = Path(os.environ.get("MLX_EXPERIMENTAL_VENV", str(ROOT / ".venv-mlx"))).expanduser()
MLX_EXPERIMENTAL_SERVER = Path(
    os.environ.get("MLX_EXPERIMENTAL_SERVER", str(MLX_EXPERIMENTAL_VENV / "bin" / "mlx_lm.server"))
).expanduser()
GEMINI_CLI_CMD = os.environ.get("GEMINI_CLI_CMD", "gemini")
GEMINI_CLI_MODEL = os.environ.get("GEMINI_CLI_MODEL", "gemini-3-flash-preview")
GEMINI_CLI_TIMEOUT_SECONDS = float(os.environ.get("GEMINI_CLI_TIMEOUT_SECONDS", "45"))
HWPFORGE_CMD = os.environ.get("HWPFORGE_CMD", "hwpforge")
ENABLE_WEB_SEARCH = os.environ.get("ENABLE_WEB_SEARCH", "1") != "0"
ENABLE_VISION_EXPERIMENTS = os.environ.get("ENABLE_VISION_EXPERIMENTS", "1") != "0"
ENABLE_SYSTEM_ACTIONS = os.environ.get("ENABLE_SYSTEM_ACTIONS", "1") != "0"
ENABLE_FILE_AUTOMATION = os.environ.get("ENABLE_FILE_AUTOMATION", "1") != "0"
ENABLE_BROWSER_AUTOMATION = os.environ.get("ENABLE_BROWSER_AUTOMATION", "1") != "0"
ONLYOFFICE_DOCS_URL = os.environ.get("ONLYOFFICE_DOCS_URL", "http://127.0.0.1:8080").rstrip("/")
COLLABORA_URL = os.environ.get("COLLABORA_URL", "http://127.0.0.1:9980").rstrip("/")
USER_AGENT = "hwp/1.0 (+https://github.com/hwkim3330/hwp)"
SESSION_ID = f"session-{int(time.time())}"
SESSION_EVENTS = []
MAX_SESSION_EVENTS = 120
AGENT_TASKS = {}
AGENT_TASK_LOCK = threading.Lock()
ONLYOFFICE_SESSIONS = {}
BROWSER_USE_REFERENCE_DIR = Path(os.environ.get("BROWSER_USE_REFERENCE_DIR", str(Path.home() / "office-agent-sources" / "browser-use"))).expanduser()
BROWSER_USE_SESSIONS = {}
MEMPALACE_REFERENCE_DIR = Path(os.environ.get("MEMPALACE_REFERENCE_DIR", str(Path.home() / "office-agent-sources" / "mempalace"))).expanduser()
MEMORY_ITEMS = []
MAX_MEMORY_ITEMS = 400
APP_INFO_CACHE = {"ts": 0.0, "data": None}
APP_INFO_TTL_SECONDS = 900
HWP_MCP_SERVER = ROOT / "scripts" / "hwp_mcp.py"


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

COMPUTER_USE_SYSTEM_PROMPT = """당신은 브라우저 중심 컴퓨터 유즈 플래너다.

역할:
- 사용자의 목표를 웹 탐색 단계로 나눈다.
- 응답은 반드시 JSON 객체 하나만 반환한다.
- 실제 실행 가능한 브라우저 단계만 제안한다.

반환 스키마:
{
  "reply": "짧은 설명",
  "summary": "무엇을 할지 한 줄 요약",
  "actions": [
    {"type": "open_url", "label": "사이트 열기", "url": "https://example.com"},
    {"type": "search_query", "label": "검색 실행", "query": "검색어"},
    {"type": "open_app", "label": "Safari 열기", "app": "Safari"},
    {"type": "note", "label": "확인 포인트", "text": "로그인 여부를 확인한다"}
  ]
}

제약:
- JSON 외 텍스트 금지.
- actions는 배열.
- type은 open_url, search_query, open_app, note만 사용한다.
- open_url은 http/https URL만 사용한다.
- search_query는 실제 검색어 문자열을 넣는다.
- open_app은 꼭 필요한 경우만 Safari 또는 Google Chrome만 사용한다.
- 최대 6단계까지만 반환한다.
"""

VISION_SYSTEM_PROMPT = """당신은 화면과 카메라 프레임을 분석하는 로컬 비전 어시스턴트다.

역할:
- 사용자가 준 이미지 한 장을 보고 현재 보이는 핵심 UI 요소나 대상을 요약한다.
- 응답은 반드시 JSON 객체 하나만 반환한다.
- 좌표는 이미지 전체 기준 퍼센트(0~100)로 반환한다.

반환 스키마:
{
  "reply": "짧은 설명",
  "summary": "현재 장면 한 줄 요약",
  "regions": [
    {"label": "버튼", "x": 12, "y": 18, "width": 30, "height": 10}
  ],
  "actions": [
    "다음으로 무엇을 확인하면 좋은지",
    "클릭/입력/확인 포인트"
  ]
}

제약:
- JSON 외 텍스트 금지.
- regions는 최대 6개.
- label은 짧고 명확하게.
- x, y, width, height는 0 이상 100 이하 숫자.
- 보이지 않는 것은 추정하지 말고, 실제 보이는 요소만 설명한다.
"""


def log_session_event(event_type, payload):
    SESSION_EVENTS.append(
        {
            "ts": int(time.time()),
            "type": event_type,
            "payload": payload,
        }
    )
    del SESSION_EVENTS[:-MAX_SESSION_EVENTS]


def _task_checkpoint(stage, attempt=0, detail="", payload=None):
    return {
        "ts": int(time.time()),
        "stage": stage,
        "attempt": attempt,
        "detail": detail,
        "payload": payload or {},
    }


def create_agent_task(prompt, model_profile, document, workspace):
    task_id = uuid.uuid4().hex
    task = {
        "id": task_id,
        "prompt": prompt[:800],
        "model_profile": model_profile,
        "mode": str(workspace.get("mode", "writer")),
        "status": "queued",
        "created_at": int(time.time()),
        "updated_at": int(time.time()),
        "attempts": 0,
        "max_attempts": 3,
        "checkpoints": [
            _task_checkpoint(
                "queued",
                detail="작업이 대기열에 등록되었습니다.",
                payload={
                    "paragraphs": len(document.get("paragraphs", []) or []),
                    "sheet_rows": len((workspace.get("sheet", {}) or {}).get("rows", []) or []),
                    "slides": len(workspace.get("slides", []) or []),
                },
            )
        ],
        "plan": None,
        "error": "",
    }
    with AGENT_TASK_LOCK:
        AGENT_TASKS[task_id] = task
        overflow = len(AGENT_TASKS) - 20
        if overflow > 0:
            for stale_id in list(AGENT_TASKS.keys())[:overflow]:
                if stale_id != task_id:
                    AGENT_TASKS.pop(stale_id, None)
    return task


def get_agent_task(task_id):
    with AGENT_TASK_LOCK:
        task = AGENT_TASKS.get(task_id)
        return json.loads(json.dumps(task)) if task else None


def update_agent_task(task_id, **changes):
    with AGENT_TASK_LOCK:
        task = AGENT_TASKS.get(task_id)
        if not task:
            return None
        task.update(changes)
        task["updated_at"] = int(time.time())
        return task


def append_agent_task_checkpoint(task_id, stage, attempt=0, detail="", payload=None):
    with AGENT_TASK_LOCK:
        task = AGENT_TASKS.get(task_id)
        if not task:
            return None
        task.setdefault("checkpoints", []).append(_task_checkpoint(stage, attempt=attempt, detail=detail, payload=payload))
        task["updated_at"] = int(time.time())
        return task


def run_agent_task(task_id, prompt, model_profile, document, workspace):
    update_agent_task(task_id, status="running")
    append_agent_task_checkpoint(task_id, "running", detail="계획 생성을 시작합니다.")
    plan = None
    last_error = ""
    for attempt in range(1, 4):
        update_agent_task(task_id, attempts=attempt)
        append_agent_task_checkpoint(task_id, "attempt_start", attempt=attempt, detail=f"{attempt}차 시도 시작")
        try:
            llm_document = {**document, **workspace}
            plan, search_results, memory_results = call_llm(prompt, llm_document, model_profile)
            plan = finalize_plan(workspace["mode"], prompt, document, workspace, plan)
            meta = {"planner": "llm", "model_profile": model_profile, "attempt": attempt}
            if search_results:
                meta["search_results"] = len(search_results)
            if memory_results:
                meta["memory_results"] = len(memory_results)
            plan = {**plan, "meta": meta}
            append_agent_task_checkpoint(
                task_id,
                "attempt_success",
                attempt=attempt,
                detail=f"{attempt}차 시도 성공",
                payload={"operations": len([item for item in plan.get("operations", []) if item.get("type") != "no_op"])},
            )
            break
        except (error.HTTPError, error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
            last_error = str(exc)
            append_agent_task_checkpoint(task_id, "attempt_error", attempt=attempt, detail=last_error[:280])
    if plan is None:
        fallback_results = []
        if should_use_web_search(prompt):
            try:
                fallback_results = web_search(prompt, max_results=5)
            except Exception:
                fallback_results = []
        plan = fallback_plan(prompt, document, workspace, last_error or "task_failed", fallback_results)
        append_agent_task_checkpoint(
            task_id,
            "fallback",
            attempt=min(3, int(get_agent_task(task_id).get("attempts", 0) or 0)),
            detail="LLM 실패 후 fallback 계획으로 전환",
            payload={"reason": last_error[:280], "search_results": len(fallback_results)},
        )
    remember_memory("agent", f"에이전트 요청 · {workspace['mode']}", prompt, {"mode": workspace["mode"], "task_id": task_id})
    log_session_event(
        "agent_task",
        {
            "task_id": task_id,
            "mode": workspace["mode"],
            "prompt": prompt[:400],
            "planner": plan.get("meta", {}).get("planner", "unknown"),
            "attempts": int(get_agent_task(task_id).get("attempts", 0) or 0),
        },
    )
    update_agent_task(task_id, status="completed", plan=plan, error=last_error)
    append_agent_task_checkpoint(
        task_id,
        "completed",
        attempt=int(get_agent_task(task_id).get("attempts", 0) or 0),
        detail="작업 계획 생성 완료",
        payload={"planner": plan.get("meta", {}).get("planner", "unknown")},
    )


def mempalace_reference_status():
    path = MEMPALACE_REFERENCE_DIR
    if not path.is_dir():
        return {"available": False, "path": str(path), "detail": "mempalace reference not found"}
    try:
        result = subprocess.run(
            ["git", "-C", str(path), "log", "--oneline", "-1"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        detail = (result.stdout or result.stderr).strip() or "reference available"
    except Exception as exc:
        detail = str(exc)
    return {"available": True, "path": str(path), "detail": detail}


def remember_memory(kind, title, text, metadata=None):
    title = re.sub(r"\s+", " ", str(title or "").strip())[:160]
    text = re.sub(r"\s+", " ", str(text or "").strip())[:4000]
    if not title and not text:
        return
    item = {
        "id": uuid.uuid4().hex,
        "ts": int(time.time()),
        "kind": str(kind or "note")[:40],
        "title": title or "Untitled memory",
        "text": text,
        "pinned": False,
        "metadata": metadata or {},
    }
    MEMORY_ITEMS.insert(0, item)
    del MEMORY_ITEMS[MAX_MEMORY_ITEMS:]
    save_memory_store()


def tokenize_memory_text(text):
    return [token for token in re.findall(r"[0-9A-Za-z가-힣]{2,}", str(text).lower()) if len(token) >= 2]


def search_memory(query, limit=6):
    q = re.sub(r"\s+", " ", str(query or "").strip())
    if not q:
        return MEMORY_ITEMS[:limit]
    query_tokens = tokenize_memory_text(q)
    scored = []
    for item in MEMORY_ITEMS:
        haystack = f"{item.get('title', '')} {item.get('text', '')}".lower()
        score = 0.0
        if item.get("pinned"):
            score += 4.0
        for token in query_tokens:
            if token in haystack:
                score += 1.0
        if q.lower() in haystack:
            score += 3.0
        age_penalty = max(0, int(time.time()) - int(item.get("ts", 0))) / 86400
        score += max(0, 2.5 - min(2.5, age_penalty / 3))
        if score > 0:
            scored.append((score, item))
    scored.sort(key=lambda pair: (-pair[0], -pair[1].get("ts", 0)))
    return [item for _, item in scored[:limit]]


def compact_memory_items(items):
    compact = []
    for item in items[:8]:
        compact.append(
            {
                "id": item.get("id", ""),
                "kind": item.get("kind", ""),
                "title": item.get("title", "")[:160],
                "text": item.get("text", "")[:700],
                "ts": item.get("ts", 0),
                "pinned": bool(item.get("pinned")),
            }
        )
    return compact


def set_memory_pinned(memory_id, pinned):
    for item in MEMORY_ITEMS:
        if item.get("id") == memory_id:
            item["pinned"] = bool(pinned)
            save_memory_store()
            return item
    raise KeyError("memory not found")


def delete_memory(memory_id):
    for index, item in enumerate(MEMORY_ITEMS):
        if item.get("id") == memory_id:
            removed = MEMORY_ITEMS.pop(index)
            save_memory_store()
            return removed
    raise KeyError("memory not found")


def save_memory_store():
    MEMORY_RUNTIME_PATH.parent.mkdir(parents=True, exist_ok=True)
    MEMORY_RUNTIME_PATH.write_text(
        json.dumps(MEMORY_ITEMS[:MAX_MEMORY_ITEMS], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_memory_store():
    if not MEMORY_RUNTIME_PATH.is_file():
        return
    try:
        items = json.loads(MEMORY_RUNTIME_PATH.read_text(encoding="utf-8"))
    except Exception:
        return
    MEMORY_ITEMS.clear()
    for item in items[:MAX_MEMORY_ITEMS]:
        if not isinstance(item, dict):
            continue
        MEMORY_ITEMS.append(
            {
                "id": str(item.get("id") or uuid.uuid4().hex),
                "ts": int(item.get("ts", int(time.time()))),
                "kind": str(item.get("kind", "note"))[:40],
                "title": str(item.get("title", "Untitled memory"))[:160],
                "text": str(item.get("text", ""))[:4000],
                "pinned": bool(item.get("pinned")),
                "metadata": item.get("metadata", {}) if isinstance(item.get("metadata", {}), dict) else {},
            }
        )


def import_memory_items(items, replace=False):
    if replace:
        MEMORY_ITEMS.clear()
    for item in items[:MAX_MEMORY_ITEMS]:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title", "")).strip()[:160]
        text = str(item.get("text", "")).strip()[:4000]
        if not text:
            continue
        MEMORY_ITEMS.append(
            {
                "id": str(item.get("id") or uuid.uuid4().hex),
                "ts": int(item.get("ts", int(time.time()))),
                "kind": str(item.get("kind", "note"))[:40],
                "title": title or text[:80],
                "text": text,
                "pinned": bool(item.get("pinned")),
                "metadata": item.get("metadata", {}) if isinstance(item.get("metadata", {}), dict) else {},
            }
        )
    del MEMORY_ITEMS[MAX_MEMORY_ITEMS:]
    save_memory_store()


def session_snapshot():
    return {
        "id": SESSION_ID,
        "events": SESSION_EVENTS[-60:],
    }


def list_onlyoffice_sessions():
    items = []
    for value in ONLYOFFICE_SESSIONS.values():
      items.append(
          {
              "id": value["id"],
              "title": value["title"],
              "extension": value["extension"],
              "mode": value["mode"],
              "created_at": value["created_at"],
              "launch_url": f"/onlyoffice.html?session={value['id']}",
              "share_url": f"http://{HOST}:{PORT}/onlyoffice.html?session={value['id']}",
          }
      )
    return sorted(items, key=lambda item: item["created_at"], reverse=True)


def browser_use_reference_status():
    path = BROWSER_USE_REFERENCE_DIR
    if not path.is_dir():
        return {"available": False, "path": str(path), "detail": "browser-use reference not found"}
    try:
        result = subprocess.run(
            ["git", "-C", str(path), "log", "--oneline", "-1"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        detail = (result.stdout or result.stderr).strip() or "reference available"
    except Exception as exc:
        detail = str(exc)
    return {"available": True, "path": str(path), "detail": detail}


def list_browser_use_sessions():
    items = []
    for value in BROWSER_USE_SESSIONS.values():
        items.append(
            {
                "id": value["id"],
                "goal": value["goal"],
                "status": value["status"],
                "created_at": value["created_at"],
                "executed_steps": len(value.get("history", [])),
                "summary": value.get("plan", {}).get("summary", ""),
                "actions": value.get("plan", {}).get("actions", []),
            }
        )
    return sorted(items, key=lambda item: item["created_at"], reverse=True)


def normalize_computer_use_actions(actions):
    normalized = []
    for action in actions or []:
        if not isinstance(action, dict):
            continue
        action_type = str(action.get("type", "")).strip()
        label = str(action.get("label", "")).strip() or action_type
        if action_type == "open_url":
            url = str(action.get("url", "")).strip()
            if url.startswith(("http://", "https://")):
                normalized.append({"type": action_type, "label": label[:120], "url": url[:800]})
        elif action_type == "search_query":
            query = re.sub(r"\s+", " ", str(action.get("query", "")).strip())
            if query:
                normalized.append({"type": action_type, "label": label[:120], "query": query[:240]})
        elif action_type == "open_app":
            app_name = str(action.get("app", "")).strip()
            if app_name in {"Safari", "Google Chrome"}:
                normalized.append({"type": action_type, "label": label[:120], "app": app_name})
        elif action_type == "note":
            text = re.sub(r"\s+", " ", str(action.get("text", "")).strip())
            if text:
                normalized.append({"type": action_type, "label": label[:120], "text": text[:400]})
        if len(normalized) >= 6:
            break
    return normalized


def validate_computer_use_plan(payload):
    if not isinstance(payload, dict):
        raise ValueError("computer use plan must be an object")
    reply = str(payload.get("reply", "")).strip() or "브라우저 작업 계획을 생성했습니다."
    summary = str(payload.get("summary", "")).strip() or reply
    actions = normalize_computer_use_actions(payload.get("actions", []))
    if not actions:
        raise ValueError("computer use plan requires actions")
    return {"reply": reply[:400], "summary": summary[:300], "actions": actions}


def validate_vision_result(payload):
    if not isinstance(payload, dict):
        raise ValueError("vision result must be an object")
    reply = str(payload.get("reply", "")).strip() or "장면 분석을 완료했습니다."
    summary = str(payload.get("summary", "")).strip() or reply
    regions = []
    for item in payload.get("regions", [])[:6]:
        if not isinstance(item, dict):
            continue
        try:
            x = max(0.0, min(100.0, float(item.get("x", 0))))
            y = max(0.0, min(100.0, float(item.get("y", 0))))
            width = max(1.0, min(100.0, float(item.get("width", 0))))
            height = max(1.0, min(100.0, float(item.get("height", 0))))
        except (TypeError, ValueError):
            continue
        label = re.sub(r"\s+", " ", str(item.get("label", "")).strip())[:80]
        if not label:
            continue
        regions.append({"label": label, "x": x, "y": y, "width": width, "height": height})
    actions = [re.sub(r"\s+", " ", str(item).strip())[:200] for item in payload.get("actions", [])[:6] if str(item).strip()]
    return {"reply": reply[:400], "summary": summary[:300], "regions": regions, "actions": actions}


def decode_data_url_image(image_data_url):
    raw = str(image_data_url or "").strip()
    match = re.match(r"^data:(image/[\w.+-]+);base64,(.+)$", raw, re.DOTALL)
    if not match:
        raise ValueError("invalid image data url")
    return match.group(1), match.group(2)


def call_vision_llm(prompt, image_data_url, source="screen"):
    if not ENABLE_VISION_EXPERIMENTS:
        raise PermissionError("vision experiments disabled")
    mime_type, image_b64 = decode_data_url_image(image_data_url)
    user_content = json.dumps(
        {
            "prompt": prompt,
            "source": source,
            "mime_type": mime_type,
        },
        ensure_ascii=False,
    )
    if is_ollama_base_url():
        payload = {
            "model": LLM_MODEL,
            "messages": [
                {"role": "system", "content": VISION_SYSTEM_PROMPT},
                {"role": "user", "content": user_content, "images": [image_b64]},
            ],
            "stream": False,
            "format": "json",
            "options": {
                "temperature": 0.1,
                "num_predict": min(LLM_MAX_TOKENS, 600),
            },
        }
        req = request.Request(
            "http://127.0.0.1:11434/api/chat",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with request.urlopen(req, timeout=max(LLM_TIMEOUT_SECONDS, 30)) as response:
            raw = response.read().decode("utf-8")
        data = json.loads(raw)
        content = data["message"]["content"]
        return validate_vision_result(json.loads(extract_json_object(content)))

    raise ValueError("vision analysis currently requires local Ollama multimodal runtime")


def build_computer_use_fallback_plan(goal, current_url="", search_results=None):
    topic = re.sub(r"\s+", " ", goal).strip() or "작업 목표"
    actions = []
    if current_url.startswith(("http://", "https://")):
        actions.append({"type": "open_url", "label": "현재 작업 페이지 열기", "url": current_url})
    else:
        actions.append({"type": "open_app", "label": "브라우저 준비", "app": "Safari"})
    actions.append({"type": "search_query", "label": "목표 기반 검색", "query": topic})
    if search_results:
        first = search_results[0]
        if first.get("url", "").startswith(("http://", "https://")):
            actions.append({"type": "open_url", "label": "첫 참고 링크 열기", "url": first["url"]})
            actions.append({"type": "note", "label": "확인 포인트", "text": f"{first.get('title', '참고 링크')}의 핵심 근거와 필요한 수치를 먼저 확인합니다."})
    else:
        actions.append({"type": "note", "label": "확인 포인트", "text": "검색 결과 상단 3개를 비교해 공식 문서나 원문 자료부터 확인합니다."})
    actions.append({"type": "note", "label": "문서 반영", "text": "확인한 근거를 Writer 또는 연구노트에 구조화된 블록으로 옮깁니다."})
    return {
        "reply": "브라우저 중심 컴퓨터 유즈 계획을 생성했습니다.",
        "summary": "브라우저를 열고 검색 결과를 검토한 뒤 문서로 옮기는 흐름",
        "actions": normalize_computer_use_actions(actions),
    }


def call_computer_use_llm(goal, current_url="", search_results=None):
    user_content = json.dumps(
        {
            "goal": goal,
            "current_url": current_url,
            "web_search": search_results or [],
        },
        ensure_ascii=False,
    )
    messages = [
        {"role": "system", "content": COMPUTER_USE_SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]
    llm_config = resolve_llm_config("balanced")
    if is_ollama_config(llm_config):
        return validate_computer_use_plan(call_ollama_json(messages, llm_config))

    payload = {
        "model": llm_config["model"],
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": min(LLM_MAX_TOKENS, 500),
        "response_format": {"type": "json_object"},
    }
    headers = {"Content-Type": "application/json"}
    if llm_config["api_key"]:
        headers["Authorization"] = f"Bearer {llm_config['api_key']}"
    req = request.Request(
        f"{llm_config['base_url']}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    with request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["choices"][0]["message"]["content"]
    return validate_computer_use_plan(json.loads(extract_json_object(content)))


def create_browser_use_session(goal, plan):
    session_id = uuid.uuid4().hex
    BROWSER_USE_SESSIONS[session_id] = {
        "id": session_id,
        "goal": goal,
        "created_at": int(time.time()),
        "status": "planned",
        "plan": plan,
        "history": [],
    }
    log_session_event("computer_use_plan", {"session_id": session_id, "goal": goal[:240], "steps": len(plan.get("actions", []))})
    remember_memory("browser_plan", f"브라우저 계획 · {goal[:80]}", f"{plan.get('summary', '')}\n{json.dumps(plan.get('actions', []), ensure_ascii=False)}", {"session_id": session_id})
    return session_id


def run_computer_use_step(session_id, step_index):
    if not ENABLE_BROWSER_AUTOMATION:
        raise PermissionError("browser automation disabled")
    session = BROWSER_USE_SESSIONS.get(session_id)
    if not session:
        raise KeyError("session not found")
    actions = session.get("plan", {}).get("actions", [])
    if not isinstance(step_index, int) or step_index < 0 or step_index >= len(actions):
        raise IndexError("invalid step index")
    action = actions[step_index]
    action_type = action.get("type")
    result = None
    if action_type == "open_url":
        result = run_system_action("open_url", {"url": action["url"]})
    elif action_type == "search_query":
        search_url = f"https://duckduckgo.com/?q={quote(action['query'])}"
        result = run_system_action("open_url", {"url": search_url})
    elif action_type == "open_app":
        result = run_system_action("open_app", {"app": action["app"]})
    elif action_type == "note":
        result = {"type": "note", "text": action.get("text", "")}
    else:
        raise ValueError(f"unsupported action: {action_type}")
    session["status"] = "running" if action_type != "note" else "review"
    session.setdefault("history", []).append({"ts": int(time.time()), "step_index": step_index, "action": action, "result": result})
    log_session_event("computer_use_step", {"session_id": session_id, "step_index": step_index, "type": action_type})
    remember_memory("browser_step", f"브라우저 단계 {step_index + 1}", json.dumps({"action": action, "result": result}, ensure_ascii=False), {"session_id": session_id})
    return {"session_id": session_id, "step_index": step_index, "action": action, "result": result}


def onlyoffice_doc_type(extension):
    return {
        "docx": "word",
        "xlsx": "cell",
        "pptx": "slide",
    }.get(extension, "word")


def onlyoffice_mime_type(extension):
    return {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }.get(extension, mimetypes.guess_type(f"file.{extension}")[0] or "application/octet-stream")


def create_onlyoffice_session(title, extension, content_base64, mode):
    ONLYOFFICE_RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    session_id = uuid.uuid4().hex
    file_name = f"{session_id}.{extension}"
    file_path = ONLYOFFICE_RUNTIME_DIR / file_name
    file_path.write_bytes(base64.b64decode(content_base64.encode("utf-8")))
    file_url = f"http://{HOST}:{PORT}/api/onlyoffice/file/{session_id}"
    config = {
        "document": {
            "fileType": extension,
            "key": session_id,
            "title": title,
            "url": file_url,
        },
        "documentType": onlyoffice_doc_type(extension),
        "editorConfig": {
            "mode": "edit",
            "callbackUrl": f"http://{HOST}:{PORT}/api/onlyoffice/callback?file_id={session_id}",
            "customization": {
                "autosave": True,
                "forcesave": True,
                "compactHeader": False,
                "toolbarNoTabs": False,
            },
        },
    }
    ONLYOFFICE_SESSIONS[session_id] = {
        "id": session_id,
        "title": title,
        "extension": extension,
        "mode": mode,
        "docs_url": ONLYOFFICE_DOCS_URL,
        "file_path": str(file_path),
        "config": config,
        "created_at": int(time.time()),
    }
    log_session_event("onlyoffice_session", {"id": session_id, "title": title, "extension": extension, "mode": mode})
    return session_id


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


def extract_last_json_object(text):
    raw = str(text or "").strip()
    if not raw:
        raise ValueError("no JSON object found")
    for index in range(len(raw) - 1, -1, -1):
        if raw[index] != "{":
            continue
        candidate = raw[index:]
        try:
            json.loads(candidate)
            return candidate
        except Exception:
            continue
    return extract_json_object(raw)


def is_ollama_base_url():
    return "127.0.0.1:11434" in LLM_BASE_URL or "localhost:11434" in LLM_BASE_URL


def resolve_llm_config(model_profile="balanced"):
    profile = str(model_profile or "balanced").strip().lower()
    if profile == "experimental":
        return {
            "profile": "experimental",
            "label": "supergemma4 mlx",
            "base_url": MLX_EXPERIMENTAL_BASE_URL,
            "model": MLX_EXPERIMENTAL_MODEL,
            "api_key": "",
            "provider": "mlx",
        }
    if profile == "gemini":
        return {
            "profile": "gemini",
            "label": "gemini cli",
            "provider": "gemini-cli",
            "model": GEMINI_CLI_MODEL,
            "command": resolve_gemini_cli_cmd(),
        }
    return {
        "profile": profile,
        "label": "default",
        "base_url": LLM_BASE_URL,
        "model": LLM_MODEL,
        "api_key": LLM_API_KEY,
        "provider": "ollama" if is_ollama_base_url() else "openai-compatible",
    }


def is_ollama_config(config):
    base_url = str(config.get("base_url", "")).rstrip("/")
    return "127.0.0.1:11434" in base_url or "localhost:11434" in base_url


def mlx_runtime_status():
    status = {
        "configured": MLX_EXPERIMENTAL_SERVER.is_file(),
        "server": str(MLX_EXPERIMENTAL_SERVER),
        "baseUrl": MLX_EXPERIMENTAL_BASE_URL,
        "model": MLX_EXPERIMENTAL_MODEL,
        "ready": False,
        "detail": "MLX experimental runtime not reachable",
    }
    try:
        req = request.Request(
            f"{MLX_EXPERIMENTAL_BASE_URL}/models",
            headers={"Content-Type": "application/json"},
            method="GET",
        )
        with request.urlopen(req, timeout=2) as response:
            raw = json.loads(response.read().decode("utf-8"))
        status["ready"] = True
        status["detail"] = f"experimental runtime ready · {(raw.get('data') or [{}])[0].get('id', MLX_EXPERIMENTAL_MODEL)}"
    except Exception as exc:
        status["detail"] = str(exc)
    return status


def resolve_gemini_cli_cmd():
    explicit = Path(GEMINI_CLI_CMD).expanduser()
    if explicit.is_file():
        return str(explicit)
    found = shutil.which(GEMINI_CLI_CMD)
    if found:
        return found
    return None


def gemini_cli_status():
    command = resolve_gemini_cli_cmd()
    if not command:
        return {"available": False, "command": None, "model": GEMINI_CLI_MODEL, "detail": "gemini CLI not found"}
    try:
        result = subprocess.run(
            [command, "--version"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
            cwd=str(ROOT),
        )
    except Exception as exc:
        return {"available": False, "command": command, "model": GEMINI_CLI_MODEL, "detail": str(exc)}
    detail = (result.stdout or result.stderr).strip().splitlines()
    return {
        "available": result.returncode == 0,
        "command": command,
        "model": GEMINI_CLI_MODEL,
        "detail": detail[0] if detail else "gemini CLI ready",
    }


def mcp_server_status():
    return {
        "available": HWP_MCP_SERVER.is_file(),
        "path": str(HWP_MCP_SERVER),
        "detail": "local stdio MCP server for runtime/search/plan/browser/gemini",
    }


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
        {"kind": "heading", "level": 2, "text": "문서 개요"},
        {
            "kind": "bullets",
            "items": [
                "문서 목적, 대상, 산출물을 먼저 한 줄씩 정리합니다.",
                "핵심 내용과 실행 항목은 분리해서 읽기 쉽게 배치합니다.",
                "최종 검토 시 일정, 담당, 후속 조치가 빠지지 않게 점검합니다.",
            ],
        },
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


def build_technical_report_blocks(prompt):
    topic = re.sub(r"\s+", " ", prompt).strip() or "기술 검토 주제"
    return [
        {"kind": "heading", "level": 1, "text": "기술문서 초안"},
        {"kind": "paragraph", "text": f"문서 주제\n{topic}"},
        {"kind": "heading", "level": 2, "text": "1. 목적"},
        {
            "kind": "paragraph",
            "text": "본 문서는 시험 또는 검토 목적, 대상 시스템, 기대 산출물을 명확히 정의하기 위한 초안입니다.",
        },
        {"kind": "heading", "level": 2, "text": "2. 범위"},
        {
            "kind": "bullets",
            "items": [
                "대상 장비, 소프트웨어 버전, 적용 조건을 구분합니다.",
                "문서에서 다루는 항목과 제외하는 항목을 함께 적습니다.",
                "재현 가능한 환경 정보는 별도 표로 남깁니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "3. 용어 및 약어"},
        {
            "kind": "table",
            "headers": ["용어", "설명"],
            "rows": [
                ["용어 1", "핵심 개념 설명"],
                ["약어 1", "약어의 전체 명칭과 의미"],
            ],
        },
        {"kind": "heading", "level": 2, "text": "4. 실험 환경 및 기본 설정"},
        {
            "kind": "table",
            "headers": ["구분", "내용", "비고"],
            "rows": [
                ["테스트 대상", "장비 또는 시스템명", "버전 기입"],
                ["환경", "네트워크/OS/설정 값", "재현 조건"],
                ["도구", "사용한 툴과 스크립트", "필수 여부"],
            ],
        },
        {"kind": "heading", "level": 2, "text": "5. 결과 및 해석"},
        {
            "kind": "numbered",
            "items": [
                "실험 결과를 정량 값과 함께 먼저 요약합니다.",
                "예상과 다른 결과가 나온 원인 후보를 분리해 적습니다.",
                "재시험 필요 여부와 다음 액션을 명확히 남깁니다.",
            ],
        },
    ]


def build_comparison_blocks(prompt):
    topic = re.sub(r"\s+", " ", prompt).strip() or "비교 대상"
    return [
        {"kind": "heading", "level": 1, "text": "비교표 초안"},
        {"kind": "paragraph", "text": f"비교 주제\n{topic}"},
        {"kind": "heading", "level": 2, "text": "요약"},
        {
            "kind": "bullets",
            "items": [
                "성공/실패 또는 우수/보통/미흡처럼 판정 기준을 먼저 적습니다.",
                "정량 지표와 한 줄 결론을 함께 제시합니다.",
                "최종 선택 시 runtime, 품질, 안정성을 따로 봅니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "정량 비교"},
        {
            "kind": "table",
            "headers": ["대상", "기준", "성능", "비고", "판정"],
            "rows": [
                ["안 1", "기준값", "측정값", "특징 요약", "검토"],
                ["안 2", "기준값", "측정값", "특징 요약", "검토"],
                ["안 3", "기준값", "측정값", "특징 요약", "검토"],
            ],
        },
        {"kind": "heading", "level": 2, "text": "한 줄 결론"},
        {
            "kind": "paragraph",
            "text": "이번 조건에서 가장 적합한 대안과 그 이유를 한 문단으로 정리합니다.",
        },
    ]


def build_paper_blocks(prompt):
    topic = re.sub(r"\s+", " ", prompt).strip() or "연구 주제"
    return [
        {"kind": "heading", "level": 1, "text": topic},
        {
            "kind": "paragraph",
            "text": "Key Words : 핵심 키워드 1, 핵심 키워드 2, 핵심 키워드 3",
        },
        {"kind": "heading", "level": 2, "text": "목차"},
        {
            "kind": "paragraph",
            "text": "Ⅰ. 서론   Ⅱ. 방법 또는 비교 대상   Ⅲ. 실험 결과 및 분석   Ⅳ. 결론",
        },
        {"kind": "heading", "level": 2, "text": "Ⅰ. 서론"},
        {
            "kind": "paragraph",
            "text": "연구 배경, 문제 정의, 기존 방법의 한계, 본 문서가 다루는 핵심 비교 관점을 서론에서 정리합니다.",
        },
        {"kind": "heading", "level": 2, "text": "Ⅱ. 방법 또는 비교 대상"},
        {
            "kind": "numbered",
            "items": [
                "대상 방법 1의 핵심 원리와 장단점을 정리합니다.",
                "대상 방법 2의 핵심 원리와 장단점을 정리합니다.",
                "비교 지표와 실험 조건을 명시합니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "Ⅲ. 실험 결과 및 분석"},
        {
            "kind": "table",
            "headers": ["항목", "조건", "결과", "해석"],
            "rows": [
                ["저부하", "조건 기입", "결과 기입", "의미 요약"],
                ["중부하", "조건 기입", "결과 기입", "의미 요약"],
                ["고부하", "조건 기입", "결과 기입", "의미 요약"],
            ],
        },
        {"kind": "heading", "level": 2, "text": "Ⅳ. 결론"},
        {
            "kind": "bullets",
            "items": [
                "핵심 결과를 다시 한 번 짧게 요약합니다.",
                "실무 또는 연구 관점에서의 의미를 분리해 적습니다.",
                "향후 검증할 항목과 남은 한계를 정리합니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "참고문헌"},
        {
            "kind": "numbered",
            "items": [
                "참고문헌 1",
                "참고문헌 2",
                "참고문헌 3",
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


def build_complete_research_note_blocks(prompt, search_results):
    topic = re.sub(r"\s+", " ", prompt).strip() or "연구 주제"
    summary_items = []
    reference_rows = []
    for index, item in enumerate(search_results[:6], start=1):
        title = item.get("title", "").strip()
        snippet = item.get("snippet", "").strip()
        summary_items.append(f"{title}{': ' + snippet[:160] if snippet else ''}")
        reference_rows.append([str(index), title, item.get("url", "")])
    if not summary_items:
        summary_items = [
            "관련 자료를 추가 조사해 핵심 배경과 최신 동향을 보강합니다.",
            "현재 초안은 구조 중심 템플릿이며, 사실 확인 후 세부 내용을 채웁니다.",
        ]
        reference_rows = [["1", "추가 조사 필요", ""]]

    return [
        {"kind": "heading", "level": 1, "text": "연구노트"},
        {"kind": "paragraph", "text": f"주제\n{topic}"},
        {"kind": "heading", "level": 2, "text": "연구 목적"},
        {
            "kind": "bullets",
            "items": [
                "이 주제가 왜 중요한지와 현재 문맥에서 필요한 이유를 정리합니다.",
                "이번 노트에서 확인해야 할 기술적·업무적 포인트를 분명히 적습니다.",
                "후속 문서나 발표 자료로 이어질 수 있도록 핵심 질문을 구조화합니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "핵심 질문"},
        {
            "kind": "numbered",
            "items": [
                "이 주제의 현재 상태와 핵심 변화는 무엇인가?",
                "실제 업무나 제품에 적용할 때 어떤 장점과 제약이 있는가?",
                "다음 단계에서 검증하거나 구현해야 할 항목은 무엇인가?",
            ],
        },
        {"kind": "heading", "level": 2, "text": "조사 요약"},
        {"kind": "numbered", "items": summary_items},
        {"kind": "heading", "level": 2, "text": "활용 인사이트"},
        {
            "kind": "bullets",
            "items": [
                "발견한 정보 중 바로 제품 설계나 문서 작업에 반영할 수 있는 항목을 고릅니다.",
                "성능, 비용, 유지보수성 같은 관점에서 실제 선택 기준을 따로 적습니다.",
                "실험용과 운영용 구성을 구분해 리스크를 줄이는 방향으로 정리합니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "다음 액션"},
        {
            "kind": "table",
            "headers": ["항목", "내용", "우선순위"],
            "rows": [
                ["1", "참고 링크 원문 검토 및 사실 확인", "상"],
                ["2", "핵심 인사이트를 문서 또는 발표 구조로 재정리", "상"],
                ["3", "실행 가능한 실험 항목 선정", "중"],
            ],
        },
        {"kind": "heading", "level": 2, "text": "참고 링크"},
        {
            "kind": "table",
            "headers": ["No", "제목", "링크"],
            "rows": reference_rows,
        },
    ]


def build_basic_research_note_blocks(prompt):
    topic = re.sub(r"\s+", " ", prompt).strip() or "연구 주제"
    return [
        {"kind": "heading", "level": 1, "text": "연구노트"},
        {"kind": "paragraph", "text": f"주제\n{topic}"},
        {"kind": "heading", "level": 2, "text": "배경"},
        {
            "kind": "paragraph",
            "text": "이 주제를 검토하는 이유와 현재 문맥에서의 중요성을 간단히 정리합니다. 이후 조사 결과가 들어오면 이 섹션을 구체화합니다.",
        },
        {"kind": "heading", "level": 2, "text": "핵심 질문"},
        {
            "kind": "numbered",
            "items": [
                "무엇을 확인해야 하는가?",
                "실제 적용 시 기대 효과와 제약은 무엇인가?",
                "다음으로 검증해야 할 항목은 무엇인가?",
            ],
        },
        {"kind": "heading", "level": 2, "text": "예비 인사이트"},
        {
            "kind": "bullets",
            "items": [
                "현재 초안은 조사 전 구조를 잡는 단계입니다.",
                "검색 결과와 원문 검토 후 사실 기반 내용으로 치환해야 합니다.",
                "최종 문서나 발표 자료로 이어질 수 있게 섹션을 분리했습니다.",
            ],
        },
        {"kind": "heading", "level": 2, "text": "다음 액션"},
        {
            "kind": "table",
            "headers": ["항목", "내용", "우선순위"],
            "rows": [
                ["1", "자료 조사 및 참고 링크 수집", "상"],
                ["2", "핵심 질문별 메모 보강", "상"],
                ["3", "최종 문서/발표 자료 구조로 재정리", "중"],
            ],
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

    if "연구노트" in prompt:
        blocks = build_complete_research_note_blocks(prompt, search_results) if search_results else build_basic_research_note_blocks(prompt)
        return {
            "reply": "연구노트 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": blocks}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
        }

    if any(keyword in prompt for keyword in ("기술문서", "기술 문서", "시험보고서", "시험 절차서", "설계서", "사용자 가이드", "user guide")):
        return {
            "reply": "기술문서 구조 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": build_technical_report_blocks(prompt)}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
        }

    if any(keyword in prompt for keyword in ("비교표", "비교 표", "비교분석", "비교 분석", "정량 비교")):
        return {
            "reply": "비교표 구조 초안을 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": build_comparison_blocks(prompt)}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
        }

    if any(keyword in prompt for keyword in ("논문", "학회", "저널", "초록", "abstract")):
        return {
            "reply": "논문형 문서 뼈대를 생성했습니다. 오프라인 플래너를 사용했습니다.",
            "operations": [{"type": "set_document_blocks", "blocks": build_paper_blocks(prompt)}],
            "meta": {"planner": "fallback", "reason": reason, "search_results": len(search_results or [])},
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


def call_llm(user_prompt, document, model_profile="balanced"):
    search_results = []
    memory_results = search_memory(user_prompt, limit=6)
    llm_config = resolve_llm_config(model_profile)
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
            "memory_recall": compact_memory_items(memory_results),
            "web_search": search_results,
        },
        ensure_ascii=False,
    )
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]
    if llm_config.get("provider") == "gemini-cli":
        return call_gemini_cli_llm(messages, llm_config), search_results, memory_results
    if is_ollama_config(llm_config):
        return call_ollama_llm(messages, llm_config), search_results, memory_results

    payload = {
        "model": llm_config["model"],
        "messages": messages,
        "temperature": 0.2,
        "max_tokens": LLM_MAX_TOKENS,
        "response_format": {"type": "json_object"},
    }

    headers = {"Content-Type": "application/json"}
    if llm_config["api_key"]:
        headers["Authorization"] = f"Bearer {llm_config['api_key']}"

    req = request.Request(
        f"{llm_config['base_url']}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    with request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["choices"][0]["message"]["content"]
    return validate_plan(json.loads(extract_json_object(content))), search_results, memory_results


def call_ollama_json(messages, llm_config):
    payload = {
        "model": llm_config["model"],
        "messages": messages,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.1,
            "num_predict": LLM_MAX_TOKENS,
        },
    }
    req = request.Request(
        f"{llm_config['base_url'].replace('/v1', '')}/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS) as response:
        raw = response.read().decode("utf-8")
    data = json.loads(raw)
    content = data["message"]["content"]
    return json.loads(extract_json_object(content))


def call_ollama_llm(messages, llm_config):
    return validate_plan(call_ollama_json(messages, llm_config))


def run_gemini_cli(prompt, model=None):
    status = gemini_cli_status()
    command = status.get("command")
    if not command:
        raise FileNotFoundError("gemini CLI not found")
    args = [
        command,
        "-p",
        str(prompt or "").strip(),
        "--output-format",
        "json",
        "--approval-mode",
        "plan",
    ]
    if model or status.get("model"):
        args.extend(["--model", str(model or status.get("model"))])
    result = subprocess.run(
        args,
        capture_output=True,
        text=True,
        timeout=GEMINI_CLI_TIMEOUT_SECONDS,
        check=False,
        cwd=str(ROOT),
    )
    stdout = (result.stdout or "").strip()
    stderr = (result.stderr or "").strip()
    raw = stdout or stderr
    if result.returncode != 0:
        raise RuntimeError(raw or stderr or f"gemini CLI exited with {result.returncode}")
    parse_candidates = [candidate for candidate in (stdout, stderr, f"{stdout}\n{stderr}".strip()) if candidate]
    outer = None
    for candidate in parse_candidates:
        try:
            outer = json.loads(extract_last_json_object(candidate))
            break
        except Exception:
            continue
    if outer is None:
        raise ValueError("gemini CLI returned no parseable JSON envelope")
    response_text = outer.get("response", "")
    if not response_text:
        raise ValueError("gemini CLI returned no response payload")
    return {
        "session_id": outer.get("session_id"),
        "response": response_text,
        "stats": outer.get("stats", {}),
        "raw": outer,
    }


def call_gemini_cli_llm(messages, llm_config):
    system_prompt = next((item.get("content", "") for item in messages if item.get("role") == "system"), SYSTEM_PROMPT)
    user_prompt = next((item.get("content", "") for item in messages if item.get("role") == "user"), "")
    prompt = (
        f"{system_prompt}\n\n"
        "반드시 JSON 객체 하나만 반환해. 설명, 코드블록, 추가 문장은 금지.\n\n"
        f"입력:\n{user_prompt}"
    )
    result = run_gemini_cli(prompt, model=llm_config.get("model"))
    return validate_plan(json.loads(extract_json_object(str(result.get("response", "")))))


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
        {
            "id": "browser.automation",
            "label": "Browser Automation",
            "status": "disabled" if not ENABLE_BROWSER_AUTOMATION else "experimental",
            "detail": "브라우저 계획 생성 및 단계별 실행",
        },
        {
            "id": "gemini.cli",
            "label": "Gemini CLI",
            "status": "granted" if gemini_cli_status()["available"] else "disabled",
            "detail": "로컬 Gemini CLI를 보조 플래너와 리서치 엔진으로 호출",
        },
    ]


def tool_registry():
    hwpforge = hwpforge_status()
    mlx_status = mlx_runtime_status()
    gemini_status = gemini_cli_status()
    mcp_status = mcp_server_status()
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
            "id": "mlx_experimental",
            "label": "MLX Experimental",
            "category": "llm",
            "status": "ready" if mlx_status["ready"] else "partial",
            "detail": mlx_status["detail"],
        },
        {
            "id": "gemini_cli",
            "label": "Gemini CLI",
            "category": "llm",
            "status": "ready" if gemini_status["available"] else "offline",
            "detail": f"{gemini_status['model']} via {gemini_status['command'] or 'unavailable'}",
        },
        {
            "id": "mcp_bridge",
            "label": "MCP Bridge",
            "category": "automation",
            "status": "ready" if mcp_status["available"] else "offline",
            "detail": mcp_status["detail"],
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
        {
            "id": "onlyoffice_bridge",
            "label": "ONLYOFFICE Bridge",
            "category": "documents",
            "status": "ready",
            "detail": f"OOXML 편집 세션 브리지 via {ONLYOFFICE_DOCS_URL}",
        },
        {
            "id": "collabora_bridge",
            "label": "Collabora Bridge",
            "category": "documents",
            "status": "partial",
            "detail": f"협업 편집 엔진 endpoint via {COLLABORA_URL}",
        },
        {
            "id": "computer_use",
            "label": "Computer Use",
            "category": "automation",
            "status": "disabled" if not ENABLE_BROWSER_AUTOMATION else "experimental",
            "detail": browser_use_reference_status()["detail"],
        },
        {
            "id": "memory_recall",
            "label": "Memory Recall",
            "category": "memory",
            "status": "ready",
            "detail": mempalace_reference_status()["detail"],
        },
    ]


def runtime_registry():
    mlx_status = mlx_runtime_status()
    gemini_status = gemini_cli_status()
    mcp_status = mcp_server_status()
    active_tasks = 0
    with AGENT_TASK_LOCK:
        active_tasks = len([item for item in AGENT_TASKS.values() if item.get("status") in {"queued", "running"}])
    return {
        "session": {"id": SESSION_ID, "eventCount": len(SESSION_EVENTS)},
        "agentTasks": {"active": active_tasks, "total": len(AGENT_TASKS)},
        "llm": {"model": LLM_MODEL, "baseUrl": LLM_BASE_URL, "ollama": is_ollama_base_url()},
        "mlxExperimental": mlx_status,
        "geminiCli": gemini_status,
        "mcp": mcp_status,
        "onlyoffice": {"docsUrl": ONLYOFFICE_DOCS_URL, "sessions": len(ONLYOFFICE_SESSIONS)},
        "collabora": {"url": COLLABORA_URL},
        "computerUse": {"sessions": len(BROWSER_USE_SESSIONS), "reference": browser_use_reference_status()},
        "memory": {"items": len(MEMORY_ITEMS), "reference": mempalace_reference_status()},
        "permissions": permission_registry(),
        "tools": tool_registry(),
    }


def app_version():
    try:
        package = json.loads((ROOT / "package.json").read_text(encoding="utf-8"))
        return str(package.get("version", "0.0.0"))
    except Exception:
        return "0.0.0"


def latest_release_info():
    now = time.time()
    cached = APP_INFO_CACHE.get("data")
    if cached and now - APP_INFO_CACHE.get("ts", 0.0) < APP_INFO_TTL_SECONDS:
        return cached

    release = {
        "checked_at": int(now),
        "available": False,
        "tag": None,
        "name": None,
        "url": "https://github.com/hwkim3330/hwp/releases",
        "published_at": None,
        "notes_url": "https://github.com/hwkim3330/hwp/releases",
        "status": "offline",
        "detail": "latest release unavailable",
    }
    req = request.Request(
        "https://api.github.com/repos/hwkim3330/hwp/releases/latest",
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": USER_AGENT,
        },
        method="GET",
    )
    try:
        with request.urlopen(req, timeout=4) as response:
            payload = json.loads(response.read().decode("utf-8"))
        release.update(
            {
                "available": True,
                "tag": payload.get("tag_name"),
                "name": payload.get("name"),
                "url": payload.get("html_url") or release["url"],
                "notes_url": payload.get("html_url") or release["url"],
                "published_at": payload.get("published_at"),
                "status": "ready",
                "detail": "latest release resolved",
            }
        )
    except Exception as exc:
        release["detail"] = str(exc)

    APP_INFO_CACHE["ts"] = now
    APP_INFO_CACHE["data"] = release
    return release


def app_info():
    current_version = app_version()
    release = latest_release_info()
    latest_tag = str(release.get("tag") or "")
    normalized_latest = latest_tag[1:] if latest_tag.startswith("v") else latest_tag
    update_available = bool(normalized_latest and normalized_latest != current_version)
    return {
        "name": "hwp",
        "version": current_version,
        "bundleId": "com.hwkim.hwp",
        "repo": "https://github.com/hwkim3330/hwp",
        "releasesUrl": "https://github.com/hwkim3330/hwp/releases",
        "latestRelease": release,
        "updateAvailable": update_available,
        "packaging": {
            "target": ["dmg", "zip"],
            "icon": (ROOT / "build-assets" / "hwp.icns").is_file(),
            "artifactPattern": "dist-electron/hwp-<version>-arm64.{dmg,zip}",
            "signed": False,
            "notarized": False,
        },
    }


def run_system_action(action, payload):
    if not ENABLE_SYSTEM_ACTIONS:
        raise PermissionError("system actions disabled")
    if action == "open_url":
        target = str(payload.get("url", "")).strip()
        if not target.startswith(("http://", "https://")):
            raise ValueError("invalid url")
        subprocess.run(["open", target], check=True)
        log_session_event("system_action", {"action": action, "target": target})
        remember_memory("system_action", "URL 열기", target, {"action": action})
        return {"action": action, "target": target}
    if action == "reveal_path":
        target = str(payload.get("path", "")).strip()
        if not target:
            raise ValueError("path required")
        subprocess.run(["open", "-R", target], check=True)
        log_session_event("system_action", {"action": action, "target": target})
        remember_memory("system_action", "파일 위치 열기", target, {"action": action})
        return {"action": action, "target": target}
    if action == "open_app":
        target = str(payload.get("app", "")).strip()
        if not target:
            raise ValueError("app required")
        subprocess.run(["open", "-a", target], check=True)
        log_session_event("system_action", {"action": action, "target": target})
        remember_memory("system_action", "앱 열기", target, {"action": action})
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
                    "version": app_version(),
                    "model": LLM_MODEL,
                    "baseUrl": LLM_BASE_URL,
                    "mlxExperimental": mlx_runtime_status(),
                    "planner": "llm+fallback",
                    "hwpforge": hwpforge_status(),
                    "onlyoffice": {"docsUrl": ONLYOFFICE_DOCS_URL, "sessions": len(ONLYOFFICE_SESSIONS)},
                    "collabora": {"url": COLLABORA_URL},
                    "computerUse": {"sessions": len(BROWSER_USE_SESSIONS), "reference": browser_use_reference_status()},
                    "memory": {"items": len(MEMORY_ITEMS), "reference": mempalace_reference_status()},
                }
            )
            return
        if self.path == "/api/runtime":
            self._send_json({"ok": True, "runtime": runtime_registry()})
            return
        if self.path == "/api/app-info":
            self._send_json({"ok": True, "app": app_info()})
            return
        if self.path.startswith("/api/memory/search"):
            query = parse_qs(urlparse(self.path).query).get("q", [""])[0]
            self._send_json({"ok": True, "items": compact_memory_items(search_memory(query, limit=8))})
            return
        if self.path == "/api/memory/export":
            self._send_json({"ok": True, "items": MEMORY_ITEMS[:MAX_MEMORY_ITEMS]})
            return
        if self.path == "/api/session":
            self._send_json({"ok": True, "session": session_snapshot()})
            return
        if self.path.startswith("/api/agent-task/"):
            task_id = self.path.rsplit("/", 1)[-1]
            task = get_agent_task(task_id)
            if not task:
                self._send_json({"ok": False, "error": "task_not_found"}, status=HTTPStatus.NOT_FOUND)
                return
            self._send_json({"ok": True, "task": task})
            return
        if self.path == "/api/computer-use/sessions":
            self._send_json({"ok": True, "sessions": list_browser_use_sessions()})
            return
        if self.path == "/api/onlyoffice/sessions":
            self._send_json({"ok": True, "sessions": list_onlyoffice_sessions()})
            return
        if self.path.startswith("/api/onlyoffice/file/"):
            session_id = self.path.rsplit("/", 1)[-1]
            item = ONLYOFFICE_SESSIONS.get(session_id)
            if not item:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            file_path = Path(item["file_path"])
            if not file_path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            raw = file_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", onlyoffice_mime_type(item["extension"]))
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(raw)
            return
        if self.path.startswith("/api/onlyoffice/config/"):
            session_id = self.path.rsplit("/", 1)[-1]
            item = ONLYOFFICE_SESSIONS.get(session_id)
            if not item:
                self._send_json({"ok": False, "error": "session_not_found"}, status=HTTPStatus.NOT_FOUND)
                return
            self._send_json({"ok": True, "session": item})
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
                results = web_search(query, max_results=5)
                log_session_event("search", {"query": query, "count": len(results)})
                remember_memory("search", f"검색 · {query[:80]}", json.dumps(results[:5], ensure_ascii=False), {"count": len(results)})
                self._send_json({"ok": True, "results": results})
            except Exception as exc:
                self._send_json({"ok": False, "error": "search_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/onlyoffice/config":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                title = str(payload.get("title", "untitled")).strip() or "untitled"
                extension = str(payload.get("extension", "")).strip().lower()
                content_base64 = str(payload.get("content_base64", "")).strip()
                mode = str(payload.get("mode", "writer")).strip()
                if extension not in {"docx", "xlsx", "pptx"}:
                    raise ValueError("unsupported extension")
                if not content_base64:
                    raise ValueError("content_base64 is required")
                session_id = create_onlyoffice_session(title, extension, content_base64, mode)
                self._send_json(
                    {
                        "ok": True,
                        "session_id": session_id,
                        "launch_url": f"/onlyoffice.html?session={session_id}",
                        "docs_url": ONLYOFFICE_DOCS_URL,
                    }
                )
            except Exception as exc:
                self._send_json({"ok": False, "error": "onlyoffice_config_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path.startswith("/api/onlyoffice/callback"):
            try:
                query = parse_qs(urlparse(self.path).query)
                session_id = (query.get("file_id") or [""])[0]
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                log_session_event("onlyoffice_callback", {"id": session_id, "status": payload.get("status"), "payload": payload})
                self._send_json({"error": 0})
            except Exception as exc:
                self._send_json({"error": 1, "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
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
        if self.path == "/api/memory/add":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                title = str(payload.get("title", "")).strip()
                text = str(payload.get("text", "")).strip()
                kind = str(payload.get("kind", "note")).strip() or "note"
                if not text:
                    raise ValueError("text is required")
                remember_memory(kind, title or text[:80], text, {"source": "manual"})
                self._send_json({"ok": True, "items": compact_memory_items(MEMORY_ITEMS)})
            except Exception as exc:
                self._send_json({"ok": False, "error": "memory_add_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/memory/pin":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                memory_id = str(payload.get("id", "")).strip()
                pinned = bool(payload.get("pinned", True))
                if not memory_id:
                    raise ValueError("id is required")
                item = set_memory_pinned(memory_id, pinned)
                self._send_json({"ok": True, "item": compact_memory_items([item])[0]})
            except Exception as exc:
                self._send_json({"ok": False, "error": "memory_pin_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/memory/delete":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                memory_id = str(payload.get("id", "")).strip()
                if not memory_id:
                    raise ValueError("id is required")
                item = delete_memory(memory_id)
                self._send_json({"ok": True, "item": compact_memory_items([item])[0]})
            except Exception as exc:
                self._send_json({"ok": False, "error": "memory_delete_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/memory/import":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                items = payload.get("items", [])
                replace = bool(payload.get("replace", False))
                if not isinstance(items, list):
                    raise ValueError("items must be a list")
                import_memory_items(items, replace=replace)
                self._send_json({"ok": True, "items": compact_memory_items(MEMORY_ITEMS)})
            except Exception as exc:
                self._send_json({"ok": False, "error": "memory_import_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/computer-use/plan":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                goal = re.sub(r"\s+", " ", str(payload.get("goal", "")).strip())
                current_url = str(payload.get("current_url", "")).strip()
                if not goal:
                    raise ValueError("goal is required")
                search_results = []
                if should_use_web_search(goal):
                    try:
                        search_results = web_search(goal, max_results=3)
                    except Exception:
                        search_results = []
                try:
                    plan = call_computer_use_llm(goal, current_url=current_url, search_results=search_results)
                    meta = {"planner": "llm"}
                except (error.HTTPError, error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
                    plan = build_computer_use_fallback_plan(goal, current_url=current_url, search_results=search_results)
                    meta = {"planner": "fallback", "reason": str(exc)}
                session_id = create_browser_use_session(goal, plan)
                self._send_json({"ok": True, "session_id": session_id, "plan": {**plan, "meta": meta}})
            except Exception as exc:
                self._send_json({"ok": False, "error": "computer_use_plan_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/vision/analyze":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                prompt = re.sub(r"\s+", " ", str(payload.get("prompt", "")).strip()) or "현재 장면에서 중요한 대상과 다음 액션을 알려줘."
                image_data_url = str(payload.get("image_data_url", "")).strip()
                source = str(payload.get("source", "screen")).strip() or "screen"
                if not image_data_url:
                    raise ValueError("image_data_url is required")
                result = call_vision_llm(prompt, image_data_url, source=source)
                log_session_event(
                    "vision_analyze",
                    {
                        "source": source,
                        "prompt": prompt[:240],
                        "regions": len(result.get("regions", [])),
                        "actions": len(result.get("actions", [])),
                    },
                )
                remember_memory("vision", f"비전 분석 · {source}", f"{result.get('summary', '')}\n{json.dumps(result.get('actions', []), ensure_ascii=False)}", {"source": source})
                self._send_json({"ok": True, "result": result})
            except Exception as exc:
                self._send_json({"ok": False, "error": "vision_analyze_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/computer-use/run":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                session_id = str(payload.get("session_id", "")).strip()
                step_index = int(payload.get("step_index", -1))
                result = run_computer_use_step(session_id, step_index)
                self._send_json({"ok": True, "result": result})
            except Exception as exc:
                self._send_json({"ok": False, "error": "computer_use_run_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/hwpforge/status":
            self._send_json({"ok": True, "hwpforge": hwpforge_status()})
            return
        if self.path == "/api/gemini-cli/run":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                prompt = str(payload.get("prompt", "")).strip()
                model = str(payload.get("model", GEMINI_CLI_MODEL)).strip() or GEMINI_CLI_MODEL
                if not prompt:
                    raise ValueError("prompt is required")
                result = run_gemini_cli(prompt, model=model)
                log_session_event(
                    "gemini_cli_run",
                    {
                        "prompt": prompt[:400],
                        "model": model,
                        "session_id": result.get("session_id", ""),
                    },
                )
                remember_memory("gemini", f"Gemini CLI · {model}", prompt, {"session_id": result.get("session_id", "")})
                self._send_json(
                    {
                        "ok": True,
                        "session_id": result.get("session_id"),
                        "response": result.get("response", ""),
                        "stats": result.get("stats", {}),
                    }
                )
            except Exception as exc:
                self._send_json({"ok": False, "error": "gemini_cli_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path == "/api/agent-task/start":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body or "{}")
                prompt = str(payload.get("prompt", "")).strip()
                model_profile = str(payload.get("modelProfile", "balanced")).strip() or "balanced"
                document = payload.get("document", {})
                workspace = {
                    "mode": payload.get("mode", "writer"),
                    "noteText": payload.get("noteText", ""),
                    "sheet": payload.get("sheet", {}),
                    "slides": payload.get("slides", []),
                }
                if not prompt:
                    raise ValueError("prompt is required")
                task = create_agent_task(prompt, model_profile, document, workspace)
                worker = threading.Thread(
                    target=run_agent_task,
                    args=(task["id"], prompt, model_profile, document, workspace),
                    daemon=True,
                )
                worker.start()
                self._send_json({"ok": True, "task": get_agent_task(task["id"])})
            except Exception as exc:
                self._send_json({"ok": False, "error": "agent_task_start_failed", "detail": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        if self.path != "/api/plan":
            self._send_json({"ok": False, "error": "not found"}, status=HTTPStatus.NOT_FOUND)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8")
            payload = json.loads(body)
            user_prompt = str(payload.get("prompt", "")).strip()
            model_profile = str(payload.get("modelProfile", "balanced")).strip() or "balanced"
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
                plan, search_results, memory_results = call_llm(user_prompt, llm_document, model_profile)
                plan = finalize_plan(workspace["mode"], user_prompt, document, workspace, plan)
                meta = {"planner": "llm", "model_profile": model_profile}
                if search_results:
                    meta["search_results"] = len(search_results)
                if memory_results:
                    meta["memory_results"] = len(memory_results)
                plan = {**plan, "meta": meta}
            except (error.HTTPError, error.URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
                fallback_results = []
                if should_use_web_search(user_prompt):
                    try:
                        fallback_results = web_search(user_prompt, max_results=5)
                    except Exception:
                        fallback_results = []
                plan = fallback_plan(user_prompt, document, workspace, str(exc), fallback_results)
                memory_results = search_memory(user_prompt, limit=6)
            remember_memory("agent", f"에이전트 요청 · {workspace['mode']}", user_prompt, {"mode": workspace["mode"]})
            log_session_event(
                "agent_run",
                {
                    "mode": workspace["mode"],
                    "prompt": user_prompt[:400],
                    "planner": plan.get("meta", {}).get("planner", "unknown"),
                    "operations": len([item for item in plan.get("operations", []) if item.get("type") != "no_op"]),
                },
            )
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
    load_memory_store()
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"hwp listening on http://{HOST}:{PORT}")
    print(f"LLM endpoint: {LLM_BASE_URL}/chat/completions")
    print(f"Model: {LLM_MODEL}")
    print(f"LLM timeout: {LLM_TIMEOUT_SECONDS}s")
    server.serve_forever()


if __name__ == "__main__":
    main()
