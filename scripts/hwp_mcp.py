#!/usr/bin/env python3

import json
import sys
from urllib import request


BASE_URL = "http://127.0.0.1:8765"
SERVER_INFO = {"name": "hwp-mcp", "version": "0.1.0"}
PROTOCOL_VERSION = "2024-11-05"


TOOLS = [
    {
        "name": "hwp_runtime",
        "description": "Get current hwp runtime state, tools, permissions, models, and sessions.",
        "inputSchema": {
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        },
    },
    {
        "name": "hwp_search",
        "description": "Run web search through hwp and return normalized results.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Search query"},
            },
            "required": ["query"],
            "additionalProperties": False,
        },
    },
    {
        "name": "hwp_plan",
        "description": "Generate a structured office/document plan for writer, notes, sheet, or slides.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt": {"type": "string"},
                "mode": {"type": "string", "enum": ["writer", "notes", "sheet", "slides"]},
                "modelProfile": {"type": "string", "enum": ["balanced", "fast", "deep", "experimental", "gemini"]},
            },
            "required": ["prompt"],
            "additionalProperties": False,
        },
    },
    {
        "name": "hwp_browser_plan",
        "description": "Generate a browser/computer-use plan for a web task.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "goal": {"type": "string"},
                "currentUrl": {"type": "string"},
            },
            "required": ["goal"],
            "additionalProperties": False,
        },
    },
    {
        "name": "hwp_gemini",
        "description": "Call local Gemini CLI through hwp for auxiliary reasoning.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt": {"type": "string"},
                "model": {"type": "string"},
            },
            "required": ["prompt"],
            "additionalProperties": False,
        },
    },
]


def http_json(method, path, payload=None):
    body = None
    headers = {}
    if payload is not None:
      body = json.dumps(payload).encode("utf-8")
      headers["Content-Type"] = "application/json"
    req = request.Request(f"{BASE_URL}{path}", data=body, headers=headers, method=method)
    with request.urlopen(req, timeout=120) as response:
        return json.loads(response.read().decode("utf-8") or "{}")


def mcp_response(result, request_id):
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def mcp_error(code, message, request_id=None):
    return {"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}}


def write_message(message):
    raw = json.dumps(message, ensure_ascii=False).encode("utf-8")
    sys.stdout.write(f"Content-Length: {len(raw)}\r\n\r\n")
    sys.stdout.flush()
    sys.stdout.buffer.write(raw)
    sys.stdout.buffer.flush()


def read_message():
    headers = {}
    while True:
        line = sys.stdin.buffer.readline()
        if not line:
            return None
        if line in (b"\r\n", b"\n"):
            break
        text = line.decode("utf-8", errors="replace").strip()
        if ":" in text:
            key, value = text.split(":", 1)
            headers[key.strip().lower()] = value.strip()
    length = int(headers.get("content-length", "0"))
    if length <= 0:
        return None
    body = sys.stdin.buffer.read(length)
    if not body:
        return None
    return json.loads(body.decode("utf-8"))


def tool_text(payload):
    return {"content": [{"type": "text", "text": json.dumps(payload, ensure_ascii=False, indent=2)}], "isError": False}


def call_tool(name, arguments):
    args = arguments or {}
    if name == "hwp_runtime":
        return tool_text(http_json("GET", "/api/runtime"))
    if name == "hwp_search":
        return tool_text(http_json("POST", "/api/search", {"query": str(args.get("query", ""))}))
    if name == "hwp_plan":
        payload = {
            "prompt": str(args.get("prompt", "")),
            "modelProfile": str(args.get("modelProfile", "balanced") or "balanced"),
            "mode": str(args.get("mode", "writer") or "writer"),
            "document": {"pageCount": 0, "sectionCount": 0, "paragraphs": []},
            "noteText": "",
            "sheet": {"columns": [], "rows": []},
            "slides": [],
        }
        return tool_text(http_json("POST", "/api/plan", payload))
    if name == "hwp_browser_plan":
        payload = {"goal": str(args.get("goal", "")), "current_url": str(args.get("currentUrl", ""))}
        return tool_text(http_json("POST", "/api/computer-use/plan", payload))
    if name == "hwp_gemini":
        payload = {"prompt": str(args.get("prompt", ""))}
        if args.get("model"):
            payload["model"] = str(args.get("model"))
        return tool_text(http_json("POST", "/api/gemini-cli/run", payload))
    raise ValueError(f"unknown tool: {name}")


def handle_request(message):
    method = message.get("method")
    request_id = message.get("id")
    if method == "initialize":
        return mcp_response(
            {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": SERVER_INFO,
            },
            request_id,
        )
    if method == "notifications/initialized":
        return None
    if method == "ping":
        return mcp_response({}, request_id)
    if method == "tools/list":
        return mcp_response({"tools": TOOLS}, request_id)
    if method == "tools/call":
        params = message.get("params", {})
        name = params.get("name")
        arguments = params.get("arguments", {})
        try:
            return mcp_response(call_tool(name, arguments), request_id)
        except Exception as exc:
            return mcp_response(
                {
                    "content": [{"type": "text", "text": str(exc)}],
                    "isError": True,
                },
                request_id,
            )
    return mcp_error(-32601, f"method not found: {method}", request_id)


def main():
    while True:
        message = read_message()
        if message is None:
            return 0
        response = handle_request(message)
        if response is not None:
            write_message(response)


if __name__ == "__main__":
    raise SystemExit(main())
