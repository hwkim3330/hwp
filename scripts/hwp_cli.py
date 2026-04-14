#!/usr/bin/env python3

import argparse
import json
import sys
from urllib import error, request


DEFAULT_BASE_URL = "http://127.0.0.1:8765"


def http_json(method, url, payload=None):
    data = None
    headers = {}
    if payload is not None:
      data = json.dumps(payload).encode("utf-8")
      headers["Content-Type"] = "application/json"
    req = request.Request(url, data=data, headers=headers, method=method)
    with request.urlopen(req, timeout=60) as response:
        raw = response.read().decode("utf-8")
    return json.loads(raw or "{}")


def print_output(data, pretty=True):
    if pretty:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(data, ensure_ascii=False))


def blocks_to_writer(blocks, anchor=-1, object_seed=0):
    paragraphs = []
    styles = []
    objects = []
    next_object_index = object_seed
    for block in blocks or []:
        kind = str((block or {}).get("kind", "")).strip()
        if kind == "heading":
            level = int(block.get("level", 1) or 1)
            text = str(block.get("text", "")).strip()
            if text:
                paragraphs.append(text)
                styles.append(f"heading{min(max(level, 1), 3)}")
        elif kind == "paragraph":
            text = str(block.get("text", "")).strip()
            if text:
                paragraphs.append(text)
                styles.append("body")
        elif kind in {"bullets", "numbered"}:
            items = [str(item).strip() for item in block.get("items", []) if str(item).strip()]
            objects.append(
                {
                    "id": f"cli-object-{next_object_index}",
                    "kind": kind,
                    "title": str(block.get("title", "")).strip() or ("번호 목록" if kind == "numbered" else "목록"),
                    "items": items or ["새 항목"],
                    "anchor": anchor,
                }
            )
            next_object_index += 1
        elif kind == "table":
            headers = [str(item).strip() for item in block.get("headers", [])]
            rows = [[str(cell).strip() for cell in row] for row in block.get("rows", []) if isinstance(row, list)]
            objects.append(
                {
                    "id": f"cli-object-{next_object_index}",
                    "kind": "table",
                    "title": str(block.get("title", "")).strip() or "표",
                    "headers": headers or ["항목", "내용"],
                    "rows": rows or [["", ""]],
                    "anchor": anchor,
                }
            )
            next_object_index += 1
    return paragraphs, styles, objects


def apply_plan_to_workspace(snapshot, plan):
    snapshot = dict(snapshot or {})
    snapshot.setdefault("version", 1)
    snapshot.setdefault("mode", "writer")
    snapshot.setdefault("writer", {"paragraphs": []})
    snapshot.setdefault("writerParagraphStyles", [])
    snapshot.setdefault("writerObjects", [])
    snapshot.setdefault("noteText", "")
    snapshot.setdefault("sheet", {"columns": [], "rows": []})
    snapshot.setdefault("slides", [])

    paragraphs = [str(item) for item in snapshot.get("writer", {}).get("paragraphs", [])]
    styles = [str(item or "body") for item in snapshot.get("writerParagraphStyles", [])]
    while len(styles) < len(paragraphs):
        styles.append("body")
    objects = list(snapshot.get("writerObjects", []))

    def normalize_objects():
        paragraph_count = len(paragraphs)
        normalized = []
        for item in objects:
            current = dict(item)
            anchor = current.get("anchor", -1)
            if not isinstance(anchor, int):
                anchor = -1
            if paragraph_count <= 0:
                current["anchor"] = -1
            else:
                current["anchor"] = max(-1, min(paragraph_count - 1, anchor))
            normalized.append(current)
        return normalized

    for operation in plan.get("operations", []):
        op_type = operation.get("type")
        if op_type == "set_document_blocks":
            paragraphs, styles, objects = blocks_to_writer(operation.get("blocks", []))
        elif op_type == "append_blocks":
            anchor = len(paragraphs) - 1
            add_paragraphs, add_styles, add_objects = blocks_to_writer(operation.get("blocks", []), anchor=anchor, object_seed=len(objects))
            paragraphs.extend(add_paragraphs)
            styles.extend(add_styles)
            objects.extend(add_objects)
        elif op_type == "replace_paragraph_text":
            index = int(operation.get("paragraph", -1) or -1)
            text = str(operation.get("text", "")).strip()
            if 0 <= index < len(paragraphs):
                paragraphs[index] = text
        elif op_type == "replace_paragraph_blocks":
            index = int(operation.get("paragraph", -1) or -1)
            if 0 <= index < len(paragraphs):
                repl_paragraphs, repl_styles, repl_objects = blocks_to_writer(operation.get("blocks", []), anchor=index, object_seed=len(objects))
                if repl_paragraphs:
                    paragraphs[index:index + 1] = repl_paragraphs
                    styles[index:index + 1] = repl_styles
                objects = [item for item in objects if int(item.get("anchor", -1)) != index]
                objects.extend(repl_objects)
        elif op_type == "set_note_text":
            snapshot["noteText"] = str(operation.get("text", ""))
            snapshot["mode"] = "notes"
        elif op_type == "set_sheet_data":
            snapshot["sheet"] = {
                "columns": [str(col) for col in operation.get("columns", [])],
                "rows": operation.get("rows", []),
            }
            snapshot["mode"] = "sheet"
        elif op_type == "set_slides":
            snapshot["slides"] = operation.get("slides", [])
            snapshot["mode"] = "slides"

    snapshot["writer"] = {"paragraphs": paragraphs}
    snapshot["writerParagraphStyles"] = styles[: len(paragraphs)]
    snapshot["writerObjects"] = normalize_objects()
    return snapshot


def synthesize_writer_plan(prompt, reply=""):
    title = "업무 문서 초안"
    prompt_text = str(prompt or "").strip()
    reply_text = str(reply or "").strip()
    return {
        "operations": [
            {
                "type": "set_document_blocks",
                "blocks": [
                    {"kind": "heading", "level": 1, "text": title},
                    {"kind": "paragraph", "text": f"요청 사항\n{prompt_text or '작업 내용을 정리합니다.'}"},
                    {"kind": "heading", "level": 2, "text": "핵심 내용"},
                    {
                        "kind": "bullets",
                        "items": [
                            reply_text or "요청 목적과 현재 상태를 먼저 요약합니다.",
                            "핵심 내용은 항목별로 분리해 정리합니다.",
                            "후속 조치와 일정은 표로 명확히 남깁니다.",
                        ],
                    },
                    {"kind": "heading", "level": 2, "text": "실행 계획"},
                    {
                        "kind": "table",
                        "headers": ["항목", "내용", "상태"],
                        "rows": [
                            ["1", "초안 정리", "진행"],
                            ["2", "검토 및 보강", "대기"],
                            ["3", "최종 확정", "예정"],
                        ],
                    },
                ],
            }
        ]
    }


def build_parser():
    parser = argparse.ArgumentParser(description="hwp CLI")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="hwp server base URL")
    parser.add_argument("--compact", action="store_true", help="print compact JSON")
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--compact", action="store_true", help=argparse.SUPPRESS)
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("runtime", help="show runtime registry", parents=[common])
    sub.add_parser("session", help="show session events", parents=[common])

    search = sub.add_parser("search", help="run web search", parents=[common])
    search.add_argument("query", help="search query")

    gemini = sub.add_parser("gemini", help="run local Gemini CLI through hwp server", parents=[common])
    gemini.add_argument("prompt", help="Gemini CLI prompt")
    gemini.add_argument("--model", default="", help="override Gemini CLI model")

    plan = sub.add_parser("plan", help="run document planner", parents=[common])
    plan.add_argument("prompt", help="planner prompt")
    plan.add_argument("--mode", default="writer", choices=["writer", "notes", "sheet", "slides"])
    plan.add_argument("--model-profile", default="balanced", choices=["balanced", "fast", "deep", "experimental", "gemini"])

    browser = sub.add_parser("browser-plan", help="create browser computer-use plan", parents=[common])
    browser.add_argument("goal", help="browser goal")
    browser.add_argument("--current-url", default="", help="current browser URL")

    workspace = sub.add_parser("workspace-agent", help="apply agent plan to a workspace snapshot", parents=[common])
    workspace.add_argument("workspace", help="input workspace json path")
    workspace.add_argument("prompt", help="prompt to apply")
    workspace.add_argument("--output", help="output workspace json path")
    workspace.add_argument("--mode", default="writer", choices=["writer", "notes", "sheet", "slides"])
    workspace.add_argument("--model-profile", default="balanced", choices=["balanced", "fast", "deep", "experimental", "gemini"])

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    base = args.base_url.rstrip("/")
    try:
        if args.command == "runtime":
            data = http_json("GET", f"{base}/api/runtime")
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "session":
            data = http_json("GET", f"{base}/api/session")
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "search":
            data = http_json("POST", f"{base}/api/search", {"query": args.query})
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "gemini":
            payload = {"prompt": args.prompt}
            if args.model:
                payload["model"] = args.model
            data = http_json("POST", f"{base}/api/gemini-cli/run", payload)
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "plan":
            payload = {
                "prompt": args.prompt,
                "modelProfile": args.model_profile,
                "mode": args.mode,
                "document": {"pageCount": 0, "sectionCount": 0, "paragraphs": []},
                "noteText": "",
                "sheet": {"columns": [], "rows": []},
                "slides": [],
            }
            data = http_json("POST", f"{base}/api/plan", payload)
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "browser-plan":
            data = http_json("POST", f"{base}/api/computer-use/plan", {"goal": args.goal, "current_url": args.current_url})
            print_output(data, pretty=not args.compact)
            return 0
        if args.command == "workspace-agent":
            with open(args.workspace, "r", encoding="utf-8") as handle:
                snapshot = json.load(handle)
            payload = {
                "prompt": args.prompt,
                "modelProfile": args.model_profile,
                "mode": args.mode,
                "document": {
                    "pageCount": 0,
                    "sectionCount": 0,
                    "paragraphs": [
                        {"text": str(item)}
                        for item in snapshot.get("writer", {}).get("paragraphs", [])
                    ],
                },
                "noteText": snapshot.get("noteText", ""),
                "sheet": snapshot.get("sheet", {"columns": [], "rows": []}),
                "slides": snapshot.get("slides", []),
            }
            data = http_json("POST", f"{base}/api/plan", payload)
            if not data.get("ok"):
                print_output(data, pretty=not args.compact)
                return 1
            plan = data.get("plan", {})
            if args.mode == "writer":
                operations = plan.get("operations", [])
                has_writer_blocks = any(item.get("type") in {"set_document_blocks", "append_blocks", "replace_paragraph_blocks"} for item in operations)
                if not has_writer_blocks:
                    plan = synthesize_writer_plan(args.prompt, plan.get("reply", ""))
            updated = apply_plan_to_workspace(snapshot, plan)
            output_path = args.output or args.workspace
            with open(output_path, "w", encoding="utf-8") as handle:
                json.dump(updated, handle, ensure_ascii=False, indent=2)
            print_output(
                {
                    "ok": True,
                    "output": output_path,
                    "reply": data.get("plan", {}).get("reply", ""),
                    "operations": len(plan.get("operations", [])),
                },
                pretty=not args.compact,
            )
            return 0
    except error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        sys.stderr.write(body + "\n")
        return 1
    except Exception as exc:
        sys.stderr.write(f"{exc}\n")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
