"""A stand-in for the AI Hub, for the suites: it speaks the hub's wire as reflex_ddns_auth.ai
does (``POST /api/ai/stream`` answering lines of JSON, ``GET /api/ai/status``) and answers AI
Fill's batches and formula suggestions by fixed rules, so the browser side can be checked
without a model. The server under test must be started with ``DDNS_AI_HUB_URL`` pointing here
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
        prompt = str(((body.get("messages") or [{}])[-1]).get("content") or "")
        Stub.asks.append({"role": body.get("role"), "system": system[:40], "prompt": prompt, "auth": bool(self.headers.get("Authorization"))})
        if system.startswith("You fill one column"):
            text, delay = fill_answer(prompt), Stub.fill_delay
        elif system.startswith("You suggest one Excel formula"):
            text, delay = suggest_answer(prompt), Stub.suggest_delay
        else:
            text, delay = "", 0.0
        time.sleep(delay)
        self.send_response(200)
        self.send_header("Content-Type", "application/x-ndjson")
        self.end_headers()
        for event in ({"model": MODEL}, {"text": text}, {"done": True, "stop": "end_turn"}):
            self.wfile.write((json.dumps(event) + "\n").encode())
            self.wfile.flush()


def start(url: str) -> ThreadingHTTPServer:
    """Serves the stand-in at ``url``'s port (in a thread of this process)."""
    port = urlparse(url).port or 80
    server = ThreadingHTTPServer(("127.0.0.1", port), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server
