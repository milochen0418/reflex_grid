"""A stand-in for the AI Hub, for the suites: it speaks the hub's wire as reflex_ddns_auth.ai
does (``POST /api/ai/stream`` answering lines of JSON, ``GET /api/ai/status``) and answers AI
Fill's batches, formula suggestions and the VBA editor by fixed rules, so the browser side can be
checked without a model. The server under test must be started with ``DDNS_AI_HUB_URL`` pointing here
and a ``DDNS_AUTH_SECRET``."""

from __future__ import annotations

import json
import re
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, List
from urllib.parse import urlparse

MODEL = "Stub (rules)"


class Stub:
    asks: List[Dict[str, Any]] = []
    fill_delay = 0.6
    suggest_delay = 0.8
    vba_line_delay = 0.08


def fill_answer(prompt: str) -> str:
    """One line per numbered row: Negative, Positive or Neutral by a few words; a title first
    when asked for one."""
    head, _, rows = prompt.partition("Rows:")
    lines = ["Title: Sentiment"] if "Title:" in head else []
    for m in re.finditer(r"^(\d+)\. (.*)$", rows, re.M):
        text = m.group(2).lower()
        if any(w in text for w in ("leak", "broke", "bad", "壞")):
            word = "Negative"
        elif any(w in text for w in ("great", "love", "good", "好")):
            word = "Positive"
        else:
            word = "Neutral"
        lines.append(f"{m.group(1)}. {word}")
    return "\n".join(lines)


def suggest_answer(prompt: str) -> str:
    """The sum of the column above the cell."""
    m = re.search(r"cell ([A-Z]+)(\d+)\.", prompt)
    if not m:
        return "I can't tell."
    col, row = m.group(1), int(m.group(2))
    return f"FORMULA: =SUM({col}2:{col}{row - 1})\nWHY: Adds the numbers above."


def vba_answer(messages: List[Dict[str, Any]]) -> str:
    """The VBA editor's requests: an answer in words to a question; the fix of a known error; code
    that doesn't compile when told "broken first", and right after the compiler's word; else a new
    Sub. "rename it" goes on from the last answer. Selected lines come back in upper case."""
    last = str(messages[-1].get("content") or "")
    if last.startswith("That doesn't compile"):
        return '```vba\nSub Total()\n    Range("B1").Value = 3\nEnd Sub\n```'
    if "rename it" in last:
        answer = str(messages[-2].get("content") or "").replace("Sub Added", "Sub Renamed")
        return answer.replace("```vba\n", "```vba\n' REMOVE Added\n")
    request = last.rpartition("<request>")[2]
    module = re.search(r"<module [^>]*>\n(.*?)\n?</module>", last, re.S)
    code = module.group(1) if module else ""
    target = re.search(r"<target [^>]*>\n(.*?)\n?</target>", last, re.S)
    if "what does" in request:
        return "It writes 42 into cell A1."
    if "leave it" in request:
        return f"```vba\n{code}\n```"
    if "Fix this error" in request:
        fixed = (target.group(1) if target else code).replace('Worksheets("Nope")', 'Worksheets("Sheet1")')
        if "End Sub" not in fixed:
            fixed += "\nEnd Sub"
        return f"```vba\n{fixed}\n```"
    if "broken first" in request:
        return '```vba\nSub Total()\n    Dim x As\nEnd Sub\n```'
    if target:
        return f"```vba\n{target.group(1).upper()}\n```"
    return '```vba\nSub Added()\n    Range("B1").Value = "hi"\nEnd Sub\n```'


def inline_answer(prompt: str) -> str:
    """A VBA suggestion: the rest of a line that ends in Range("A1")."""
    return 'Value = 1' if 'Range("A1").<CURSOR>' in prompt else ""


class _Handler(BaseHTTPRequestHandler):
    def log_message(self, *args: Any) -> None:
        pass

    def _json(self, data: Dict[str, Any], status: int = 200) -> None:
        raw = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        if urlparse(self.path).path == "/api/ai/status":
            self._json({"ok": True, "model": MODEL, "provider": "stub", "local": True})
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/ai/stream":
            self._json({"error": "not found"}, 404)
            return
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
        system = str(body.get("system") or "")
        messages = body.get("messages") or [{}]
        prompt = str(messages[-1].get("content") or "")
        Stub.asks.append({"role": body.get("role"), "system": system[:40], "prompt": prompt, "auth": bool(self.headers.get("Authorization"))})
        between = 0.0
        if system.startswith("You fill one column"):
            text, delay = fill_answer(prompt), Stub.fill_delay
        elif system.startswith("You suggest one Excel formula"):
            text, delay = suggest_answer(prompt), Stub.suggest_delay
        elif system.startswith("You change one module of Excel VBA"):
            text, delay, between = vba_answer(messages), 0.2, Stub.vba_line_delay
        elif system.startswith("You are a code completion engine inside a VBA"):
            text, delay = inline_answer(prompt), 0.1
        else:
            text, delay = "", 0.0
        time.sleep(delay)
        self.send_response(200)
        self.send_header("Content-Type", "application/x-ndjson")
        self.end_headers()
        pieces = text.splitlines(keepends=True) if between else [text]
        try:
            for event in ({"model": MODEL}, *({"text": piece} for piece in pieces), {"done": True, "stop": "end_turn"}):
                self.wfile.write((json.dumps(event) + "\n").encode())
                self.wfile.flush()
                if between and "text" in event:
                    time.sleep(between)
        except (BrokenPipeError, ConnectionResetError):
            pass  # the person pressed Stop


def start(url: str) -> ThreadingHTTPServer:
    """Serves the stand-in at ``url``'s port (in a thread of this process)."""
    port = urlparse(url).port or 80
    server = ThreadingHTTPServer(("127.0.0.1", port), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server
