"""Calling everyone on a workbook (the DDNS Intent ``call.join``, as in CoDoc in MD).

Needs a ``call.join`` provider named for the server (there is no re-ddns registry locally):

    DDNS_INTENT_PROVIDER_CALL_JOIN=livekit DDNS_INTENT_URL_LIVEKIT=http://localhost:3200 \\
    poetry run ./run_test_suite.sh collab_call

The call app doesn't have to run: the suite checks ReflexGrid's side — who is rung, declining,
accepting, joining later. When it does run there (e.g. reflex_ddns_livekit_audio_chat started
with ``reflex run --frontend-port 3200 --backend-port 8200``), Amy and Ben also connect and hear
each other.

1. Amy, Ben and Cat are on the workbook; Amy calls everyone: her call dialog opens, Ben and Cat
   are rung (Amy isn't).
2. Cat declines: no more ringing for her, the title bar offers "Join call · 1".
3. Ben accepts: he is in the call (2), and Cat sees "Join call · 2".
4. Dan opens the workbook later: he isn't rung, and sees "Join call · 2".
"""

import os
import sys
import time
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _lib.grid import Grid, run_people  # noqa: E402


_INBOUND_AUDIO_BYTES = """async () => {
    const room = window.livekitClient && window.livekitClient.room;
    if (!room) return -1;
    let total = 0;
    for (const p of room.remoteParticipants.values()) {
        for (const pub of p.audioTrackPublications.values()) {
            const track = pub.track;
            if (!track || !track.receiver) continue;
            const stats = await track.receiver.getStats();
            stats.forEach((r) => { if (r.type === 'inbound-rtp') total += r.bytesReceived || 0; });
        }
    }
    return total;
}"""


def popup(g: Grid):
    return g.page.get_by_role("alertdialog", name="Incoming call")


def call_app_running() -> bool:
    app = (os.getenv("DDNS_INTENT_PROVIDER_CALL_JOIN") or "").split(",")[0].strip()
    url = os.getenv("DDNS_INTENT_URL_" + app.upper().replace("-", "_"), "")
    if not url:
        return False
    try:
        urllib.request.urlopen(url, timeout=3)
        return True
    except Exception:
        return False


def in_call(g: Grid, count: int) -> None:
    frame = g.page.frame_locator("#ddns-intent-frame")
    frame.locator("#connection-status", has_text="Connected").wait_for(timeout=60000)
    end = time.time() + 30
    while frame.locator(".participant-card").count() != count:
        if time.time() > end:
            raise AssertionError(f"{frame.locator('.participant-card').count()} in the call, expected {count}")
        time.sleep(0.3)


def hears(g: Grid) -> None:
    call = g.page.query_selector('iframe[data-intent-action="call.join"]').content_frame()
    first = call.evaluate(_INBOUND_AUDIO_BYTES)
    for _ in range(30):
        if first > 0:
            break
        time.sleep(0.5)
        first = call.evaluate(_INBOUND_AUDIO_BYTES)
    time.sleep(3)
    second = call.evaluate(_INBOUND_AUDIO_BYTES)
    assert second > first > 0, f"no audio received ({first} -> {second})"


def body(grids) -> None:
    if not os.getenv("DDNS_INTENT_PROVIDER_CALL_JOIN"):
        raise AssertionError("set DDNS_INTENT_PROVIDER_CALL_JOIN (and DDNS_INTENT_URL_<APP>) for this suite; see its docstring")
    amy, ben, cat = grids
    for g in grids:
        g.page.locator("#people-count", has_text="3 here").wait_for(timeout=15000)

    amy.page.get_by_role("button", name="Call everyone").click()
    amy.page.locator("#ddns-intent-frame").wait_for(state="attached", timeout=15000)
    src = amy.page.locator("#ddns-intent-frame").get_attribute("src") or ""
    assert "room=" not in src, f"the call id is in the iframe URL: {src}"
    for g in (ben, cat):
        popup(g).wait_for(state="visible", timeout=10000)
        assert "is calling everyone" in popup(g).inner_text()
    assert popup(amy).count() == 0, "the caller is rung"
    ben.shot("incoming_call")

    popup(cat).get_by_role("button", name="Decline").click()
    popup(cat).wait_for(state="detached", timeout=10000)
    cat.page.get_by_role("button", name="Join call · 1").wait_for(timeout=10000)

    popup(ben).get_by_role("button", name="Accept").click()
    popup(ben).wait_for(state="detached", timeout=10000)
    ben.page.locator("#ddns-intent-frame").wait_for(state="attached", timeout=15000)
    cat.page.get_by_role("button", name="Join call · 2").wait_for(timeout=10000)
    time.sleep(2)
    assert popup(cat).count() == 0, "declined once: not rung again"
    if call_app_running():
        print("call app running: checking the call itself")
        in_call(amy, 2)
        in_call(ben, 2)
        hears(amy)
        hears(ben)
        amy.shot("in_call")
    else:
        print("call app not running: checked ringing, declining and joining only")

    # Dan comes later: the call is offered, not rung.
    ctx = amy.page.context.browser.new_context(viewport={"width": 1280, "height": 800})
    dan = Grid(ctx.new_page(), amy.out)
    dan.open(amy.base_url, amy.book)
    dan.page.get_by_role("button", name="Join call · 2").wait_for(timeout=15000)
    time.sleep(1.5)
    assert popup(dan).count() == 0, "a newcomer is rung"
    dan.shot("join_later")


if __name__ == "__main__":
    sys.exit(run_people("collab_call", body, count=3, media=True, ignore=lambda e: "ERR_CONNECTION_REFUSED" in e or ":3200" in e))
