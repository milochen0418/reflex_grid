"""Running macros: to completion (scripts, tests) or in steps a UI can drive.

A :class:`MacroRun` executes a macro on a worker thread that only works while its controller —
the thread holding the workbook's lock and calling :meth:`MacroRun.step` — waits for it. The run
pauses (``step`` returns) when the macro needs an answer (MsgBox with choices, InputBox) or has
used up a time slice; the controller then releases the workbook, shows the dialog or lets others
work, and calls ``step`` again later. Nothing ever holds the workbook while waiting for a person,
and what the macro did before a pause is already in the workbook (and shown to everyone).

Without a controller that can show dialogs (``interactive=False``) a run goes to completion in
the calling thread; dialogs then get scripted answers (``answers``) or their default button.
"""

from __future__ import annotations

import itertools
import threading
import time
from typing import Any, Callable, Dict, List, Optional

from ...vba.errors import MacroStopped, VbaError, VbaSyntaxError
from ...vba.runtime import Execution
from ...vba.stdlib import flush_print, msgbox_default
from ..session import EngineError, Session
from .host import MacroHost
from .project import runtime_for

DEFAULT_TIME_LIMIT = 60.0
SLICE = 0.25
PROMPT_TIMEOUT = 30 * 60
MAX_WORKERS = 32
STACK_SIZE = 256 * 1024 * 1024

_ids = itertools.count(1)
_workers = threading.BoundedSemaphore(MAX_WORKERS)


class MacroResult(dict):
    """What a run did: ``status`` (done, error, stopped, prompt, busy), ``output`` (Debug.Print),
    ``messages`` (MsgBox texts shown along the way), ``error``, ``prompt``."""


def _error_info(e: Exception) -> Dict[str, Any]:
    if isinstance(e, VbaSyntaxError):
        return {"kind": "compile", "number": 0, "description": e.msg, "module": e.module, "line": e.line, "text": str(e)}
    if isinstance(e, VbaError):
        return {"kind": "runtime", "number": e.number, "description": e.description, "module": e.module, "proc": e.proc, "line": e.line, "text": f"Run-time error '{e.number}': {e.description}"}
    return {"kind": "internal", "number": 51, "description": str(e), "text": str(e)}


class MacroRun:
    """One macro (or event handler) run. ``job(ex, host)`` is what it executes."""

    def __init__(self, session: Session, job: Callable[[Execution, MacroHost], Any], label: str, answers: Optional[List[Any]] = None, interactive: bool = False, time_limit: float = DEFAULT_TIME_LIMIT, user_name: str = ""):
        self.id = next(_ids)
        self.s = session
        self.job = job
        self.label = label
        self.answers = list(answers or [])
        self.interactive = interactive
        self.time_limit = time_limit
        self.user_name = user_name
        self.state = "new"
        self.prompt: Optional[Dict[str, Any]] = None
        self.result: Optional[MacroResult] = None
        self.messages: List[Dict[str, Any]] = []
        self.output: List[str] = []
        self.started = time.time()
        self._answer: Any = None
        self._cancel = False
        self._go = threading.Event()
        self._paused = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._slice_start = 0.0
        self._abandoned = False
        self.ex: Optional[Execution] = None
        self.host: Optional[MacroHost] = None
        self.value: Any = None

    # ------------------------------------------------------------ inside the run

    def _ask(self, ex: Execution, kind: str, prompt: str, title: str, buttons: int, default: Any = None, **extra: Any) -> Any:
        """MsgBox/InputBox: a scripted answer, the default, or a pause for the person."""
        if self.answers:
            return self.answers.pop(0)
        if not self.interactive:
            if kind == "msgbox":
                self.messages.append({"prompt": prompt, "title": title, "buttons": buttons})
                return msgbox_default(buttons)
            return default if default is not None else ""
        self.prompt = {"kind": kind, "prompt": prompt, "title": title, "buttons": buttons, "default": default if default is not None else "", **extra}
        answer = self._pause("prompt")
        self.prompt = None
        return answer

    def _checkpoint(self, ex: Execution) -> None:
        if self._cancel:
            raise MacroStopped("cancelled", "The macro was stopped.")
        if self.interactive and time.monotonic() - self._slice_start > SLICE:
            self._pause("busy")

    def _pause(self, why: str) -> Any:
        """Gives the workbook back to the controller and waits to be resumed."""
        host = self.host
        if host is not None:
            host.flush()
        # The time spent waiting doesn't count against the macro's time limit.
        t0 = time.monotonic()
        self.state = why
        self._go.clear()
        self._paused.set()
        if not self._go.wait(PROMPT_TIMEOUT if why == "prompt" else 600):
            self._cancel = True
            self._abandoned = True
            raise MacroStopped("cancelled", "The macro waited too long and was stopped.")
        if self.ex is not None:
            self.ex.extend_deadline(time.monotonic() - t0)
        if self._cancel:
            raise MacroStopped("cancelled", "The macro was stopped.")
        self.state = "running"
        self._slice_start = time.monotonic()
        return self._answer

    def _body(self) -> None:
        rt = runtime_for(self.s.doc().wb.vba)
        host = MacroHost(self.s, ask=self._ask, user_name=self.user_name)
        host.runtime = rt
        host.on_checkpoint = self._checkpoint
        self.host = host
        ex = Execution(rt, host, time_limit=self.time_limit)
        self.ex = ex
        status, error = "done", None
        with ex:
            try:
                self.value = self.job(ex, host)
            except MacroStopped as m:
                status = "stopped" if m.reason in ("end", "stop") else ("cancelled" if m.reason == "cancelled" else "error")
                if m.reason == "end":
                    rt.reset_state()
                if m.reason in ("timeout", "steps", "cancelled", "stop"):
                    error = {"kind": m.reason, "number": 18 if m.reason == "cancelled" else 0, "description": m.msg or m.reason, "text": m.msg or m.reason}
            except (VbaError, VbaSyntaxError) as e:
                status, error = "error", _error_info(e)
            finally:
                flush_print(ex)
        if not self._abandoned:
            try:
                host.finish()
            except VbaError as e:
                status, error = "error", _error_info(e)
        self.messages.extend(host.messages)
        self.output = list(ex.output)
        self.result = MacroResult(status=status, error=error, output=self.output, messages=self.messages, label=self.label, status_text=getattr(host, "status_text", ""))

    # ------------------------------------------------------------ the controller's side

    def step(self, answer: Any = None) -> MacroResult:
        """Runs (or resumes, with ``answer``) until the macro finishes or pauses. Call it while
        holding the workbook (inside ``book.act``); it returns with the worker parked."""
        if self.state in ("done", "error", "stopped", "cancelled"):
            return self.result or MacroResult(status=self.state)
        if not self.interactive:
            self._slice_start = time.monotonic()
            self._run_here()
            return self.result  # type: ignore[return-value]
        self._answer = answer
        self._paused.clear()
        self._slice_start = time.monotonic()
        if self._thread is None:
            if not _workers.acquire(blocking=False):
                return MacroResult(status="error", error={"kind": "internal", "number": 7, "description": "Too many macros are running; try again in a moment.", "text": "Too many macros are running"}, output=[], messages=[])
            old = threading.stack_size()
            try:
                threading.stack_size(STACK_SIZE)
                self._thread = threading.Thread(target=self._thread_main, name=f"macro-{self.id}", daemon=True)
                self._thread.start()
            finally:
                threading.stack_size(old)
        else:
            self._go.set()
        self._paused.wait()
        if self.result is not None:
            return self.result
        return MacroResult(status=self.state, prompt=self.prompt, output=list(self.ex.output if self.ex else []), messages=list(self.messages + (self.host.messages if self.host else [])), label=self.label, status_text=getattr(self.host, "status_text", ""))

    def _thread_main(self) -> None:
        try:
            self._run_here()
        finally:
            _workers.release()
            self._paused.set()

    def _run_here(self) -> None:
        self.state = "running"
        try:
            self._body()
        except Exception as e:  # a bug outside VBA code: report it, never leave a thread behind
            self.result = MacroResult(status="error", error=_error_info(e), output=self.output, messages=self.messages, label=self.label)
        self.state = (self.result or {}).get("status", "done")

    def cancel(self) -> None:
        """Stops the run at its next step (a paused run ends when stepped, without running on)."""
        self._cancel = True
        if self.ex is not None:
            self.ex.cancelled = True

    def abandon(self) -> None:
        """Ends a paused run whose controller is gone (a closed tab), without touching the
        workbook: its worker stops now instead of waiting for an answer."""
        if not self.paused:
            return
        self._cancel = True
        self._abandoned = True
        if self.ex is not None:
            self.ex.cancelled = True
        self._go.set()

    @property
    def paused(self) -> bool:
        return self.state in ("prompt", "busy")

    @property
    def finished(self) -> bool:
        return self.state in ("done", "error", "stopped", "cancelled")


def call_macro_job(name: str, args: Optional[List[Any]] = None) -> Callable[[Execution, MacroHost], Any]:
    def job(ex: Execution, host: MacroHost) -> Any:
        rt = ex.rt
        p = rt.find(name)
        if p is None:
            raise VbaError(1004, f"Cannot run the macro '{name}'. The macro may not be available in this workbook.")
        mod_err = rt.errors.get(p.module.name.lower())
        if mod_err is not None:
            raise mod_err
        return rt.call(ex, p, list(args or []))

    return job
