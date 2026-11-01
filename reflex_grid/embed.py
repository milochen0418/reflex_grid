"""A workbook on another app's page (a document in CoDoc in MD, say): its cells to read, and its
UserForms to use. The forms are the only way to change anything from there.

Cells: ``GET /_reflex_grid/embed/table?book=<id>&range=<reference>`` is a small page of its own,
not a Reflex page. It draws the cells (see :mod:`gridcore.engine.embed` for the references), loads
at once however many a document holds, and has nothing to edit with: no state, no events, no view
of the book, so the person isn't among the people on it either. ``/_reflex_grid/embed/live`` is
the websocket that carries the cells again whenever they change.

Forms: the intent ``grid.form`` (``GridState.open_embed_form``). Choosing what to show: the intent
``grid.pick`` (:mod:`.pick`), or Insert › Embed on the workbook. Both end in :func:`directive`,
the line a document says it with. Neither ever creates a workbook: an address that leads nowhere
stays one (a workbook's own page makes one, as in CoDoc)."""

from __future__ import annotations

import asyncio
import html
from collections import OrderedDict
from typing import Any, Dict, List, Optional

from starlette.requests import Request
from starlette.responses import HTMLResponse
from starlette.websockets import WebSocket, WebSocketDisconnect

from gridcore.engine import EngineError
from gridcore.engine.embed import table_view
from gridcore.engine.html_export import TABLE_CSS

from .books import STORE, Book, valid_id

TABLE_PATH = "/_reflex_grid/embed/table"
LIVE_PATH = "/_reflex_grid/embed/live"
GONE = "This workbook isn't there any more."
WENT = {"error": GONE, "gone": True}
# How much of a burst of changes (a macro at work) runs together.
BURST = 0.05

_DRAWN: "OrderedDict[tuple, tuple]" = OrderedDict()
_DRAWN_MAX = 64

PAGE_CSS = TABLE_CSS + """
html { color-scheme: light; }
body { margin: 0; background: #FFFFFF; color: #1F2937; font-family: Calibri, Carlito, 'Segoe UI', Arial, sans-serif; }
.sheet { border-collapse: collapse; table-layout: fixed; background: #FFFFFF; }
.sheet tr:first-child td { position: sticky; top: 0; z-index: 1; background: #FFFFFF; }
.sheet tr.fresh td { animation: fresh 1.6s ease-out; }
@keyframes fresh { from { box-shadow: inset 0 0 0 99px rgba(34, 197, 94, .28); } to { box-shadow: inset 0 0 0 99px transparent; } }
#foot { position: sticky; left: 0; bottom: 0; display: flex; align-items: center; gap: 6px; padding: 4px 2px; background: #FFFFFF; font: 11px system-ui, sans-serif; color: #6B7280; }
#foot[hidden] { display: none; }
#dot { width: 7px; height: 7px; border-radius: 50%; background: #22C55E; }
#dot.off { background: #9CA3AF; }
#error { padding: 12px; font: 13px system-ui, sans-serif; color: #B91C1C; }
"""

# Says how tall the page is to the page around it (it holds a number, for whoever embeds it): when
# it draws too, because a browser doesn't run a ResizeObserver in a frame out of sight. Draws the
# cells again each time they come: rows that are new or differ light up, and the latest one is
# scrolled to when the table is taller than its frame.
PAGE_JS = """
(() => {
  const box = document.getElementById("table");
  const tell = () => window.parent !== window && window.parent.postMessage(
    {type: "reflex-grid-embed", height: Math.ceil(document.body.getBoundingClientRect().height)}, "*");
  new ResizeObserver(tell).observe(document.body);
  tell();
  const rowsOf = () => Array.from(box.querySelectorAll("tr"));
  const foot = document.getElementById("foot");
  let gone = document.body.dataset.gone === "1";
  const show = (m) => {
    foot.hidden = !!m.error;
    gone = !!m.gone;
    if (m.error) {
      box.innerHTML = "";
      const p = document.createElement("div");
      p.id = "error";
      p.textContent = m.error;
      box.appendChild(p);
    } else {
      const was = rowsOf().map((r) => r.textContent);
      box.innerHTML = m.html;
      const rows = rowsOf();
      rows.forEach((r, i) => { if (was.length && r.textContent !== was[i]) r.classList.add("fresh"); });
      if (was.length && rows.length > was.length) window.scrollTo(0, document.documentElement.scrollHeight);
    }
    document.getElementById("rows").textContent = m.note || "";
    tell();
  };
  const live = () => {
    const scheme = location.protocol === "https:" ? "wss:" : "ws:";
    const ws = new WebSocket(scheme + "//" + location.host + LIVE_PATH + location.search);
    const dot = document.getElementById("dot");
    ws.onopen = () => dot.classList.remove("off");
    ws.onmessage = (ev) => show(JSON.parse(ev.data));
    ws.onclose = () => {
      dot.classList.add("off");
      if (!gone) setTimeout(live, 2000);
    };
  };
  if (!gone) live();
})();
""".replace("LIVE_PATH", repr(LIVE_PATH))


def directive(book_id: str, kind: str, value: str) -> str:
    """The line that shows a workbook's form (``kind`` "form") or cells ("range") in a document."""
    return f"{{%grid {book_id} {kind}={value} %}}"


def embed_choices(book_id: str, info: Dict[str, Any]) -> List[Dict[str, str]]:
    """What ``embed.info`` offers, as rows to pick from: ``label``, ``note``, ``code``."""
    sel = info["selection"]
    out = []
    if sel["kind"] == "table":
        out.append({"label": f"Table {sel['ref']}", "note": f"The selected table, on {sel['sheet']}. It grows with the table.", "code": directive(book_id, "range", sel["ref"])})
    else:
        out.append({"label": f"Cells {sel['area']}", "note": f"The selected cells, on {sel['sheet']}. " + ("New rows join in." if sel["grows"] else "Just these cells."), "code": directive(book_id, "range", sel["ref"])})
    for t in info["tables"]:
        if t["name"] != sel["ref"]:
            out.append({"label": f"Table {t['name']}", "note": f"{t['area']} on {t['sheet']}. It grows with the table.", "code": directive(book_id, "range", t["name"])})
    for n in info["names"]:
        out.append({"label": f"Name {n['name']}", "note": f"{n['area']} on {n['sheet']}.", "code": directive(book_id, "range", n["name"])})
    for f in info["forms"]:
        out.append({"label": f"UserForm {f['name']}", "note": f"“{f['caption']}”: people use it right in the document.", "code": directive(book_id, "form", f["name"])})
    return out


async def open_book(book_id: str) -> Optional[Book]:
    """The workbook, read from its file when nobody has it open; None when there is none."""
    if not valid_id(book_id) or not STORE.exists(book_id):
        return None
    return STORE.get(book_id) or await asyncio.to_thread(STORE.open, book_id)


def _note(rows: int, shown: int) -> str:
    if rows != shown:
        return f"The latest {shown:,} of {rows:,} rows"
    return "1 row" if rows == 1 else f"{rows:,} rows"


def _drawn(book: Book, ref: str) -> Dict[str, Any]:
    """The cells as they are now. A workbook is a snapshot nobody changes (each edit makes a new
    one), so everyone showing the same cells of it shares one drawing, made outside its lock."""
    with book.lock:
        wb = book.doc.wb
    key = (book.key, ref)
    hit = _DRAWN.get(key)
    if hit is not None and hit[0] is wb:
        return hit[1]
    try:
        t = table_view(wb, ref)
        out = {"html": t["html"], "note": _note(t["rows"], t["shown"])}
    except EngineError as ex:
        out = {"error": str(ex).split("\n", 1)[0]}
    _DRAWN[key] = (wb, out)
    _DRAWN.move_to_end(key)
    while len(_DRAWN) > _DRAWN_MAX:
        _DRAWN.popitem(last=False)
    return out


def _page(shown: Dict[str, Any], title: str) -> str:
    wrong = "error" in shown
    body = f'<div id="error">{html.escape(shown["error"])}</div>' if wrong else shown["html"]
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>{html.escape(title)}</title><style>{PAGE_CSS}</style></head>
<body{' data-gone="1"' if shown.get("gone") else ""}><div id="table">{body}</div>
<div id="foot"{" hidden" if wrong else ""}><span id="dot"></span><span>Live · Read only</span><span id="rows">{html.escape(shown.get("note", ""))}</span></div>
<script>{PAGE_JS}</script></body></html>"""


async def table_page(request: Request) -> HTMLResponse:
    ref = request.query_params.get("range", "")
    book = await open_book(request.query_params.get("book", ""))
    shown = await asyncio.to_thread(_drawn, book, ref) if book is not None else WENT
    return HTMLResponse(_page(shown, ref), status_code=200 if book is not None else 404, headers={"Cache-Control": "no-store"})


async def live(ws: WebSocket) -> None:
    """Sends the cells when they change, for as long as the page is there. Watching keeps the
    workbook open but puts nobody on it."""
    await ws.accept()
    ref = ws.query_params.get("range", "")
    book = await open_book(ws.query_params.get("book", ""))
    if book is None:
        await ws.send_json(WENT)
        await ws.close()
        return
    loop = asyncio.get_running_loop()
    changed = asyncio.Event()

    def wake() -> None:
        loop.call_soon_threadsafe(changed.set)

    async def closing() -> None:
        while (await ws.receive())["type"] != "websocket.disconnect":
            pass

    gone = asyncio.ensure_future(closing())
    book.watchers.add(wake)
    last = None
    try:
        while not gone.done():
            changed.clear()
            shown = WENT if book.deleted else await asyncio.to_thread(_drawn, book, ref)
            if shown != last:
                await ws.send_json(shown)
                last = shown
            if book.deleted:
                break
            waiting = asyncio.ensure_future(changed.wait())
            await asyncio.wait({gone, waiting}, return_when=asyncio.FIRST_COMPLETED)
            waiting.cancel()
            await asyncio.sleep(BURST)
    except (WebSocketDisconnect, RuntimeError):
        pass
    finally:
        book.watchers.discard(wake)
        gone.cancel()
