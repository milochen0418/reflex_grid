"""One tab's AI in the VBA editor (``vba_ai`` words the requests): the change it is writing, which
the editor shows as a diff to accept or reject, and the grey suggestion at the cursor. Each person
has their own; what they accept the editor puts in as their own edit (``assets/grid/vba_ai.jsx``).

A state of its own, so that an answer streaming in doesn't redraw what hangs on ``GridState``."""

from __future__ import annotations

import asyncio
import time
from typing import Any, Dict, List, Optional

import reflex as rx

from . import ai, vba_ai
from .state import GridState

# How often an answer is shown as it grows.
STREAM_EVERY = 0.15
TALK_KEPT = 6  # earlier turns of a conversation about one change


class VbaAiState(rx.State):
    # The change being written: {"seq", "docKey": the editor's module it is for, "kind": "write" |
    # "fix", "status": "writing" | "checking" | "done" | "answer" | "error", "text": what the AI has
    # written so far (the code, or an answer in words), "code": the module with the change in it,
    # "model", "problem": a compile error it still has, "error"}.
    draft: Dict[str, Any] = {}
    # The grey suggestion: {"req", "text"}.
    ghost: Dict[str, Any] = {}
    suggest: str = rx.LocalStorage("on", name="rg_vba_suggest")
    _job: int = 0
    _target: Optional[vba_ai.Target] = None
    _talk: List[Dict[str, str]] = []

    async def _begin(self, kind: str, request: str, code: str, line: int = 0, first: int = 0, last: int = 0, refine: bool = False) -> Optional[Dict[str, Any]]:
        """Starts a change (inside ``async with self``): what ``_write`` needs. With ``refine``, the
        request goes on from the change on screen."""
        grid = await self.get_state(GridState)
        _, book, _ = grid._bound()
        if book is None or not request or grid.vba_locked or not ai.ENABLED:
            return None
        with book.lock:
            wb = book.doc.wb
        name = grid.vba_module or "Module1"
        if refine and self._target is not None and self._talk:
            target, messages = self._target, [*self._talk, {"role": "user", "content": request}]
        else:
            target = vba_ai.fix_target(code, line, not grid.vba_run_error) if kind == "fix" else vba_ai.target_of(code, line, first, last)
            messages = [{"role": "user", "content": vba_ai.write_prompt(vba_ai.workbook_brief(wb, name), name, grid.vba_kind or "standard", target, request)}]
        self._job += 1
        self._target, self._talk = target, []
        self.ghost = {}
        self.draft = {"seq": self._job, "docKey": grid.vba_key, "kind": kind, "status": "writing", "text": "", "code": "", "model": "", "problem": "", "error": ""}
        return {"seq": self._job, "kind": kind, "target": target, "messages": messages, "project": wb.vba, "name": name, "user": (grid.me or {}).get("user", "")}

    async def _fail(self, seq: int, why: str) -> None:
        async with self:
            if self._job == seq:
                self.draft = {**self.draft, "status": "error", "error": why}

    async def _write(self, job: Dict[str, Any]) -> None:
        """Asks until an answer compiles, or the rounds are over: the person then sees the last
        answer, with its error."""
        seq, target, messages, name = job["seq"], job["target"], job["messages"], job["name"]
        for attempt in range(vba_ai.ROUNDS):
            reply = ai.write(vba_ai.SYSTEM, messages, job["user"])
            answer = ""
            shown = 0.0
            try:
                async for piece in reply:
                    answer += piece
                    if time.monotonic() - shown >= STREAM_EVERY:
                        async with self:
                            if self._job != seq:
                                return
                            self.draft = {**self.draft, "status": "writing", "text": vba_ai.code_of(answer) or answer, "model": reply.model}
                        shown = time.monotonic()
            except ai.FAILED as ex:
                await self._fail(seq, str(ex) or "The AI took too long to answer.")
                return
            code = vba_ai.code_of(answer)
            if reply.stop == "refusal" or not answer.strip():
                await self._fail(seq, "The AI gave no answer: ask again, or in other words." if reply.stop != "refusal" else "The model declined this request.")
                return
            if reply.stop == "max_tokens":
                await self._fail(seq, "The answer was cut off: it got too long. Select the lines to change, or ask for less.")
                return
            talk = [*messages, {"role": "assistant", "content": answer}]
            if not code:
                async with self:
                    if self._job == seq:
                        self._talk = talk
                        self.draft = {**self.draft, "status": "answer", "text": answer.strip(), "model": reply.model}
                return
            module = target.result(code)
            if module.rstrip() == target.code.rstrip():
                async with self:
                    if self._job == seq:
                        self.draft = {**self.draft, "status": "answer", "text": "The AI's answer changes nothing in this code.", "model": reply.model}
                return
            problems, ours = await asyncio.to_thread(vba_ai.new_problems, job["project"], name, target.code, module, job["kind"] == "fix")
            if not ours or attempt == vba_ai.ROUNDS - 1:
                problem = f"Line {problems[0]['line']}: {problems[0]['message']}" if problems else ""
                async with self:
                    if self._job == seq:
                        self._target = target.after(code)
                        self._talk = talk if len(talk) <= TALK_KEPT + 2 else talk[:2] + talk[-TALK_KEPT:]
                        self.draft = {**self.draft, "status": "done", "text": code, "code": module, "model": reply.model, "problem": problem}
                return
            messages = [*talk, {"role": "user", "content": vba_ai.retry_request(ours, module, target)}]
            async with self:
                if self._job != seq:
                    return
                self.draft = {**self.draft, "status": "checking", "text": code, "model": reply.model}

    @rx.event(background=True)
    async def write(self, e: Dict[str, Any]):
        """Writes what the person asked: ``prompt`` for the code ``code``, with the cursor on
        ``line`` and the lines ``first`` to ``last`` selected."""
        async with self:
            job = await self._begin("write", str(e.get("prompt") or "").strip(), str(e.get("code") or ""), int(e.get("line") or 0), int(e.get("first") or 0), int(e.get("last") or 0), bool(e.get("refine")))
        if job:
            await self._write(job)

    @rx.event(background=True)
    async def fix(self):
        """Fixes the compile or run-time error on screen."""
        async with self:
            grid = await self.get_state(GridState)
            message, code, line = grid.vba_fix_text, grid.vba_code, grid.vba_error_line
            job = await self._begin("fix", vba_ai.fix_request(message, line, code) if message else "", code, line) if message else None
        if job:
            await self._write(job)

    @rx.event(background=True)
    async def suggest_run(self, e: Dict[str, Any]):
        """A grey suggestion for the cursor between ``prefix`` and ``suffix``."""
        req = int(e.get("req") or 0)
        prefix, suffix = str(e.get("prefix") or ""), str(e.get("suffix") or "")
        async with self:
            grid = await self.get_state(GridState)
            _, book, _ = grid._bound()
            if book is None or self.suggest == "off" or self.draft or not ai.ENABLED:
                self.ghost = {"req": req, "text": ""}
                return
            with book.lock:
                wb = book.doc.wb
            name = grid.vba_module or "Module1"
            brief = vba_ai.workbook_brief(wb, name)[: vba_ai.INLINE_BRIEF]
            user = (grid.me or {}).get("user", "")
        text = ""
        try:
            answer, _ = await ai.ask(vba_ai.INLINE_SYSTEM, vba_ai.inline_prompt(prefix, suffix, brief), max_tokens=1024, user=user, timeout=25)
            text = vba_ai.clean_inline(answer, prefix, suffix)
            if text and await asyncio.to_thread(vba_ai.breaks, wb.vba, name, prefix, text, suffix):
                text = ""
        except Exception:  # noqa: BLE001 — no suggestion, nothing else
            pass
        async with self:
            self.ghost = {"req": req, "text": text}

    @rx.event
    def on_event(self, e: Dict[str, Any]):
        """What the editor's AI bar does."""
        kind = str(e.get("type") or "")
        if kind == "ask":
            return VbaAiState.write(e)
        if kind == "fix":
            return VbaAiState.fix
        if kind == "suggest":
            return VbaAiState.suggest_run(e)
        if kind == "toggle":
            self.suggest = "off" if self.suggest != "off" else "on"
            self.ghost = {}
        elif kind == "close":  # stopped, rejected or accepted: the editor did what follows
            self._job += 1
            self._target, self._talk = None, []
            self.draft, self.ghost = {}, {}
        return None
