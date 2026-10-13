"""The backend's own HTTP routes (next to Reflex's).

``POST /_reflex_grid/bye``: a tab that is closing (or reloading) says goodbye with
``navigator.sendBeacon`` — a websocket message wouldn't get out of a page that is going away.
It names the tab by its Reflex client token, which only that tab knows, so nobody can make
someone else leave."""

from __future__ import annotations

import asyncio
import json

from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import Response
from starlette.routing import Route

from . import sessions
from .books import STORE
from .state import push

BYE_PATH = "/_reflex_grid/bye"
_TASKS: set = set()


async def bye(request: Request) -> Response:
    try:
        data = json.loads((await request.body())[:4096] or b"{}")
    except ValueError:
        return Response(status_code=400)
    token = str(data.get("token") or "")
    peer = str(data.get("peer") or "")
    entry = sessions.peek(token) if token else None
    if entry is not None and entry.view is not None and entry.view.key == peer and entry.book_id:
        book = STORE.get(entry.book_id)
        if book is not None:
            book.leave(peer)
            # Tell the others now, through one of them (outside this request).
            task = asyncio.create_task(push(entry.book_id))
            _TASKS.add(task)
            task.add_done_callback(_TASKS.discard)
    return Response(status_code=204)


api = Starlette(routes=[Route(BYE_PATH, bye, methods=["POST"])])
