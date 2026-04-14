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

    plan = sub.add_parser("plan", help="run document planner", parents=[common])
    plan.add_argument("prompt", help="planner prompt")
    plan.add_argument("--mode", default="writer", choices=["writer", "notes", "sheet", "slides"])
    plan.add_argument("--model-profile", default="balanced", choices=["balanced", "fast", "deep", "experimental"])

    browser = sub.add_parser("browser-plan", help="create browser computer-use plan", parents=[common])
    browser.add_argument("goal", help="browser goal")
    browser.add_argument("--current-url", default="", help="current browser URL")

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
