"""Reflex state: the bridge between the browser (grid component, ribbon, dialogs) and the
gridcore engine session that holds the workbook."""

from __future__ import annotations

import base64
import os
from typing import Any, Dict, List, Optional

import reflex as rx

from gridcore.core.addr import MAX_COLS, MAX_ROWS, RangeRef, a1, col_to_letters
from gridcore.engine import EngineError, Session, geometry, objects, ui_state, viewport
from gridcore.engine.cmd_format import CELL_STYLE_NAMES, FORMAT_PRESETS
from gridcore.engine.samples import SAMPLES, new_sample
from gridcore.functions import all_functions
from gridcore.numfmt import parse as parse_fmt

from . import sessions

FUNCTIONS = sorted(
    ({"name": f.name, "signature": f.signature, "description": f.description, "category": f.category} for f in all_functions() if not f.name.startswith("_")),
    key=lambda x: x["name"],
)
CATEGORIES = sorted({f["category"] for f in FUNCTIONS})
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _rng(g) -> str:
    r0, c0, r1, c1 = (int(x) for x in g)
    return RangeRef(r0, c0, r1, c1).a1()


def _num(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


class GridState(rx.State):
    # Grid component props.
    geometry: Dict[str, Any] = {}
    data: Dict[str, Any] = {}
    objects: Dict[str, Any] = {}
    ui: Dict[str, Any] = {}
    functions: List[Dict[str, str]] = []
    names: List[str] = []
    focus_key: int = 0
    edit_request: Dict[str, Any] = {}
    window: List[int] = [0, 0, 80, 30]
    # Chrome.
    tab: str = "Home"
    backstage: bool = False
    selected_object: int = -1
    selected_kind: str = ""
    # Dialogs: which one is open and its fields.
    dialog: str = ""
    dlg: Dict[str, Any] = {}
    dlg_error: str = ""

    # ------------------------------------------------------------ computed

    @rx.var
    def ribbon(self) -> Dict[str, Any]:
        return self.ui.get("ribbon", {}) if self.ui else {}

    @rx.var
    def sheet_tabs(self) -> List[Dict[str, Any]]:
        out = []
        for i, s in enumerate(self.ui.get("sheets", []) if self.ui else []):
            if s.get("visible", True):
                out.append({"index": i, "name": s["name"], "color": s.get("color") or "", "active": i == self.ui.get("sheet", 0), "protected": s.get("protected", False)})
        return out

    @rx.var
    def hidden_sheets(self) -> List[str]:
        return [s["name"] for s in (self.ui.get("sheets", []) if self.ui else []) if not s.get("visible", True)]

    @rx.var
    def status_text(self) -> str:
        st = (self.ui or {}).get("stats") or {}
        if not st or st.get("count", 0) < 2:
            return ""
        parts = []
        if st.get("numeric"):
            avg = st.get("average")
            if avg is not None:
                parts.append(f"Average: {_fmt_stat(avg)}")
        parts.append(f"Count: {int(st.get('count', 0))}")
        if st.get("numeric"):
            parts.append(f"Sum: {_fmt_stat(st.get('sum', 0))}")
        return "    ".join(parts)

    @rx.var
    def doc_title(self) -> str:
        ui = self.ui or {}
        name = ui.get("docName", "Book1")
        return name + (" •" if ui.get("dirty") else "")

    @rx.var
    def circular_ref(self) -> str:
        c = (self.ui or {}).get("circular") or []
        return c[0] if c else ""

    @rx.var
    def open_docs(self) -> List[Dict[str, Any]]:
        return [{"name": d.get("name", ""), "index": i} for i, d in enumerate((self.ui or {}).get("docs") or [])]

    @rx.var
    def zoom(self) -> int:
        return int((self.geometry or {}).get("zoom", 100))

    @rx.var
    def table_selected(self) -> bool:
        return bool((self.ui or {}).get("table"))

    @rx.var
    def dlg_list(self) -> List[Dict[str, Any]]:
        """The list a dialog shows (filter values, names, rules, find results, columns), indexed."""
        d = self.dlg or {}
        key = {"filter": "values", "nameManager": "names", "cfManage": "rules", "find": "results", "replace": "results", "removeDuplicates": "cols"}.get(self.dialog)
        if key is None:
            return []
        items = [{**x, "i": i} for i, x in enumerate(d.get(key, []) or [])]
        if self.dialog == "filter":
            q = str(d.get("search", "") or "").lower()
            if q:
                items = [x for x in items if q in str(x.get("text", "")).lower()]
        return items[:2000]

    @rx.var
    def fn_list(self) -> List[Dict[str, str]]:
        if self.dialog != "insertFunction":
            return []
        d = self.dlg or {}
        q = str(d.get("search", "") or "").lower()
        cat = d.get("category", "Most Used")
        most = ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "MIN", "XLOOKUP", "VLOOKUP", "SUMIFS", "COUNTIF", "INDEX", "MATCH", "TEXT", "ROUND", "IFERROR", "FILTER", "SORT", "UNIQUE", "LET", "TODAY"]
        out = []
        for f in FUNCTIONS:
            if q and q not in f["name"].lower() and q not in f["description"].lower():
                continue
            if not q and cat == "Most Used" and f["name"] not in most:
                continue
            if not q and cat not in ("Most Used", "All") and f["category"] != cat:
                continue
            out.append(f)
        return out[:400]

    @rx.var
    def fn_selected(self) -> Dict[str, str]:
        if self.dialog != "insertFunction":
            return {}
        name = (self.dlg or {}).get("selected", "SUM")
        shown = self.fn_list
        if shown and not any(f["name"] == name for f in shown):
            return shown[0]
        return next((f for f in FUNCTIONS if f["name"] == name), {"name": name, "signature": "", "description": ""})

    @rx.var
    def dlg_sheets(self) -> List[str]:
        return list((self.dlg or {}).get("sheets", []) or [])

    # ------------------------------------------------------------ plumbing

    def _entry(self) -> sessions.Entry:
        return sessions.get(self.router.session.client_token)

    def _refresh(self, entry: sessions.Entry, seq: Optional[int] = None, extra_ui: Optional[Dict[str, Any]] = None) -> None:
        s = entry.session
        d = s.doc()
        wb = d.wb
        si = wb.active_sheet
        if entry.sheet != si or entry.sig[:1] != (d.uid,):
            # Another sheet or workbook: start at its top-left.
            if entry.sheet != -1:
                self.window = [0, 0, 80, 30]
            entry.sheet = si
        w = self.window or [0, 0, 80, 30]
        vp = viewport(wb, si, RangeRef(int(w[0]), int(w[1]), int(w[2]), int(w[3])))
        if vp != self.data:
            self.data = vp
        sig = (d.uid, id(wb), si)
        if sig != entry.sig:
            entry.sig = sig
            geo = geometry(wb.sheet(si))
            if geo != self.geometry:
                self.geometry = geo
            ob = objects(wb, si, with_images=False)
            for im in ob["images"]:
                img = next((x for x in wb.sheet(si).images if x.id == im["id"]), None)
                if img is None:
                    continue
                k = (d.uid, img.id, len(img.data))
                src = entry.images.get(k)
                if src is None:
                    src = f"data:{img.mime};base64," + base64.b64encode(img.data).decode("ascii")
                    entry.images[k] = src
                im["src"] = src
            if ob != self.objects:
                self.objects = ob
            names = sorted({n.name for n in wb.names if not n.hidden} | {t.name for sh in wb.sheets for t in sh.tables}, key=str.lower)
            if names != self.names:
                self.names = names
        ui = ui_state(s)
        ui["seq"] = seq if seq is not None else (self.ui or {}).get("seq", 0)
        if extra_ui:
            ui.update(extra_ui)
        if ui != self.ui:
            self.ui = ui
        if self.selected_object >= 0 and not any(o["id"] == self.selected_object for o in (self.objects.get("charts", []) + self.objects.get("images", []))):
            self.selected_object = -1
            self.selected_kind = ""

    def _run(self, s: Session, cid: str, params: Optional[Dict[str, Any]] = None):
        return s.execute(cid, params or {})

    def _focus(self) -> None:
        self.focus_key += 1

    # ------------------------------------------------------------ lifecycle

    @rx.event
    def on_load(self):
        entry = self._entry()
        if not self.functions:
            self.functions = FUNCTIONS
        entry.sig = ()
        entry.sheet = -1
        self._refresh(entry)
        self._focus()

    # ------------------------------------------------------------ grid events

    @rx.event
    def on_event(self, e: Dict[str, Any]):
        entry = self._entry()
        s = entry.session
        t = e.get("type")
        seq = int(e.get("seq") or 0)
        out = None
        extra: Dict[str, Any] = {}
        try:
            if t == "viewport":
                w = e.get("window") or [0, 0, 80, 30]
                self.window = [max(0, int(w[0])), max(0, int(w[1])), min(MAX_ROWS - 1, int(w[2])), min(MAX_COLS - 1, int(w[3]))]
            elif t == "select":
                sel = e.get("sel") or {}
                ranges = sel.get("ranges") or []
                if ranges:
                    act = sel.get("active") or ranges[-1][:2]
                    anc = sel.get("anchor") or act
                    self._run(s, "selection.set", {"range": ",".join(_rng(g) for g in ranges[:256]), "active": a1(int(act[0]), int(act[1])), "anchor": a1(int(anc[0]), int(anc[1]))})
            elif t == "commit":
                out = self._commit(s, e, seq, extra)
            elif t == "command":
                out = self._command(entry, str(e.get("id")), e.get("params") or {})
            elif t == "copy":
                r = self._run(s, "edit.cut" if e.get("cut") else "edit.copy", {"text": e.get("text")} if isinstance(e.get("text"), str) else {})
                if not isinstance(e.get("text"), str):
                    extra["clipboardText"] = {"seq": seq, "text": r["text"]}
            elif t == "paste":
                text = e.get("text")
                self._run(s, "edit.paste", {"text": text} if isinstance(text, str) else {})
            elif t == "resize":
                axis = e.get("axis")
                idx = [int(i) for i in (e.get("indices") or [])][:16384]
                if idx:
                    spans = _spans(idx, axis)
                    if e.get("autofit"):
                        self._run(s, "home.autofitColumnWidth" if axis == "cols" else "home.autofitRowHeight", {axis: spans})
                    else:
                        size = float(e.get("size") or 0)
                        if axis == "cols":
                            self._run(s, "home.columnWidth", {"cols": spans, "width": size})
                        else:
                            self._run(s, "home.rowHeight", {"rows": spans, "height": size})
            elif t == "fill":
                self._run(s, "edit.autoFill", {"source": e["source"], "target": e["target"], "mode": e.get("mode", "series")})
            elif t == "move":
                self._run(s, "edit.moveCells", {"from": e["from"], "to": e["to"], "copy": bool(e.get("copy"))})
            elif t == "filter":
                self._open_filter(s, int(e["r"]), int(e["c"]), e.get("table"))
            elif t == "link":
                r = self._run(s, "hyperlink.follow", {"cell": a1(int(e["r"]), int(e["c"]))})
                if r and r.get("url"):
                    out = rx.call_script(f"window.open({_js_str(r['url'])}, '_blank', 'noopener')")
            elif t == "object":
                self._run(s, "object.move", {k: e[k] for k in ("id", "x", "y", "width", "height") if k in e})
            elif t == "objectSelect":
                oid = e.get("id")
                self.selected_object = int(oid) if oid is not None else -1
                self.selected_kind = str(e.get("kind") or "")
                if self.selected_kind == "chart":
                    self.tab = "Chart Design"
                elif self.tab == "Chart Design":
                    self.tab = "Home"
            elif t == "dialog":
                self._open_dialog(entry, str(e.get("name")), e.get("params") or {})
            elif t == "namebox":
                out = self._namebox(s, str(e.get("text", "")))
        except EngineError as ex:
            msg = str(ex).split("\n", 1)[0]
            extra["lastError"] = {"seq": seq, "message": msg, "kind": t, **({"r": e.get("r"), "c": e.get("c"), "text": e.get("text")} if t == "commit" else {})}
            out = rx.toast.error(msg, position="bottom-center")
        except Exception as ex:  # never let a UI event kill the session
            extra["lastError"] = {"seq": seq, "message": str(ex), "kind": t}
            out = rx.toast.error(f"Something went wrong: {ex}", position="bottom-center")
        self._refresh(entry, seq, extra)
        return out

    def _commit(self, s: Session, e: Dict[str, Any], seq: int, extra: Dict[str, Any]):
        r, c = int(e["r"]), int(e["c"])
        text = str(e.get("text") or "")
        mode = e.get("mode", "normal")
        if mode == "fill":
            self._run(s, "range.fill", {"input": text})
            return None
        if mode == "array":
            self._run(s, "cell.set", {"cell": a1(r, c), "input": text, "array": True})
            return None
        self._run(s, "cell.set", {"cell": a1(r, c), "input": text})
        move = e.get("move", "down")
        if move != "none":
            self._run(s, "selection.set", {"range": ",".join(_rng(g) for g in s.doc().selection.ranges), "active": a1(r, c)})
            self._run(s, "selection.next", {"forward": move in ("down", "right"), "byRow": move in ("left", "right")})
        msgs = s.messages[:]
        s.messages.clear()
        if msgs:
            return rx.toast.warning(msgs[-1], position="bottom-center")
        return None

    def _command(self, entry: sessions.Entry, cid: str, params: Dict[str, Any]):
        s = entry.session
        if cid == "ui.save":
            return self._download(s, "xlsx")
        if cid == "ui.saveCsv":
            return self._download(s, "csv")
        if cid == "ui.print":
            return self._print(s)
        if cid.startswith("ui."):
            return None
        r = self._run(s, cid, params)
        msgs = s.messages[:]
        s.messages.clear()
        if cid == "data.removeDuplicates" and isinstance(r, dict):
            return rx.toast.info(f"{r.get('removed', 0)} duplicate values found and removed; {r.get('remaining', 0)} unique values remain.", position="bottom-center")
        if cid == "edit.find" and isinstance(r, dict) and "count" in r:
            return None
        if msgs:
            return rx.toast.warning(msgs[-1], position="bottom-center")
        return None

    def _download(self, s: Session, fmt: str):
        r = self._run(s, "file.saveBytes", {"format": fmt})
        data = base64.b64decode(r["data"])
        return rx.download(data=data, filename=r["name"], mime_type=r["mime"])

    def _print(self, s: Session):
        r = self._run(s, "file.exportHtml", {"print": True})
        page = _js_str(r["html"])
        js = (
            "(() => { let f = document.getElementById('rg-print-frame'); if (f) f.remove();"
            " f = document.createElement('iframe'); f.id = 'rg-print-frame';"
            " Object.assign(f.style, {position: 'fixed', right: '0', bottom: '0', width: '0', height: '0', border: '0'});"
            " document.body.appendChild(f); const d = f.contentWindow.document; d.open(); d.write(" + page + "); d.close();"
            " setTimeout(() => { f.contentWindow.focus(); f.contentWindow.print(); }, 300); })()"
        )
        return rx.call_script(js)

    def _namebox(self, s: Session, text: str):
        if not text:
            return None
        try:
            self._run(s, "edit.goTo", {"reference": text})
            return None
        except EngineError:
            pass
        from gridcore.engine.cmd_formulas import check_name

        if check_name(text) is None:
            self._run(s, "formulas.defineName", {"name": text})
            return rx.toast.success(f"Defined the name {text}.", position="bottom-center")
        raise EngineError("Reference isn't valid.")

    # ------------------------------------------------------------ ribbon actions

    @rx.event
    def cmd(self, cid: str, params: Dict[str, Any] = {}):
        entry = self._entry()
        out = None
        try:
            out = self._command(entry, cid, dict(params or {}))
        except EngineError as ex:
            out = rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")
        self._refresh(entry)
        self._focus()
        return out

    @rx.event
    def set_tab(self, tab: str):
        self.tab = tab
        self.backstage = tab == "File"
        if tab == "File":
            self.tab = "Home"

    @rx.event
    def close_backstage(self):
        self.backstage = False
        self._focus()

    @rx.event
    def zoom_to(self, z):
        entry = self._entry()
        try:
            self._run(entry.session, "view.zoom", {"percent": int(float(z if not isinstance(z, list) else z[0]))})
        except EngineError:
            pass
        entry.sig = ()
        self._refresh(entry)

    @rx.event
    def zoom_step(self, delta: int):
        z = max(10, min(400, int(self.zoom) + int(delta)))
        return GridState.zoom_to(z)

    @rx.event
    def activate_sheet(self, index: int):
        entry = self._entry()
        try:
            self._run(entry.session, "sheet.activate", {"sheet": int(index)})
        except EngineError as ex:
            return rx.toast.error(str(ex))
        self._refresh(entry)
        self._focus()

    @rx.event
    def new_sheet(self):
        entry = self._entry()
        s = entry.session
        try:
            self._run(s, "home.insertSheet", {"index": len(s.doc().wb.sheets)})
        except EngineError as ex:
            return rx.toast.error(str(ex))
        self._refresh(entry)
        self._focus()

    @rx.event
    def new_workbook(self, sample: str = ""):
        entry = self._entry()
        s = entry.session
        if sample and sample in SAMPLES:
            new_sample(s, sample)
        else:
            self._run(s, "file.new", {})
        self.backstage = False
        self.window = [0, 0, 80, 30]
        entry.sig = ()
        self._refresh(entry)
        self._focus()

    @rx.event
    def switch_window(self, index: int):
        entry = self._entry()
        self._run(entry.session, "window.activate", {"index": int(index)})
        self.backstage = False
        self.window = [0, 0, 80, 30]
        entry.sig = ()
        self._refresh(entry)
        self._focus()

    @rx.event
    def close_workbook(self):
        entry = self._entry()
        self._run(entry.session, "file.close", {})
        self.backstage = False
        entry.sig = ()
        self._refresh(entry)
        self._focus()

    @rx.event
    def save_as(self, fmt: str):
        entry = self._entry()
        self.backstage = False
        try:
            out = self._download(entry.session, fmt)
        except EngineError as ex:
            return rx.toast.error(str(ex))
        self._refresh(entry)
        return out

    @rx.event
    def print_sheet(self):
        entry = self._entry()
        self.backstage = False
        try:
            return self._print(entry.session)
        except EngineError as ex:
            return rx.toast.error(str(ex))

    @rx.event
    def save_html(self):
        entry = self._entry()
        self.backstage = False
        try:
            r = self._run(entry.session, "file.exportHtml", {"print": False})
        except EngineError as ex:
            return rx.toast.error(str(ex))
        return rx.download(data=r["html"].encode("utf-8"), filename=r["name"], mime_type="text/html")

    @rx.event
    async def open_upload(self, files: List[rx.UploadFile]):
        entry = self._entry()
        s = entry.session
        out = None
        for f in files[:1]:
            data = await f.read()
            name = os.path.basename(f.name or f.filename or "Book.xlsx")
            try:
                r = self._run(s, "file.open", {"data": data, "name": name})
                w = r.get("warnings") or []
                out = rx.toast.success(f"Opened {name}" + (f" ({len(w)} warnings)" if w else ""), position="bottom-center")
            except EngineError as ex:
                out = rx.toast.error(str(ex), position="bottom-center")
        self.backstage = False
        self.dialog = ""
        self.window = [0, 0, 80, 30]
        entry.sig = ()
        self._refresh(entry)
        self._focus()
        return out

    @rx.event
    async def picture_upload(self, files: List[rx.UploadFile]):
        entry = self._entry()
        out = None
        for f in files[:1]:
            data = await f.read()
            try:
                self._run(entry.session, "insert.picture", {"data": data, "name": f.name or f.filename or "picture"})
            except EngineError as ex:
                out = rx.toast.error(str(ex), position="bottom-center")
        self.dialog = ""
        self._refresh(entry)
        self._focus()
        return out

    # ------------------------------------------------------------ dialogs

    @rx.event
    def open_dialog(self, name: str, params: Dict[str, Any] = {}):
        entry = self._entry()
        try:
            self._open_dialog(entry, name, dict(params or {}))
        except EngineError as ex:
            return rx.toast.error(str(ex).split("\n", 1)[0])

    @rx.event
    def dialog_open_change(self, is_open: bool):
        if not is_open:
            self.dialog = ""
            self.dlg = {}
            self.dlg_error = ""
            self._focus()

    @rx.event
    def toggle_list_item(self, field: str, index: int):
        d = dict(self.dlg)
        items = [dict(x) for x in d.get(field, [])]
        if 0 <= index < len(items):
            items[index]["checked"] = not items[index].get("checked", False)
        d[field] = items
        self.dlg = d

    @rx.event
    def close_dialog(self):
        self.dialog = ""
        self.dlg = {}
        self.dlg_error = ""
        self._focus()

    @rx.event
    def set_dlg(self, field: str, value: Any):
        d = dict(self.dlg)
        d[field] = value
        self.dlg = d

    @rx.event
    def toggle_filter_value(self, index: int):
        d = dict(self.dlg)
        vals = [dict(v) for v in d.get("values", [])]
        if 0 <= index < len(vals):
            vals[index]["checked"] = not vals[index]["checked"]
        d["values"] = vals
        self.dlg = d

    @rx.event
    def filter_select_all(self, on: bool):
        d = dict(self.dlg)
        d["values"] = [{**v, "checked": bool(on)} for v in d.get("values", [])]
        d["blanksChecked"] = bool(on)
        self.dlg = d

    def _open_dialog(self, entry: sessions.Entry, name: str, params: Dict[str, Any]) -> None:
        s = entry.session
        d = s.doc()
        sh = d.sheet
        sel = d.selection
        ui = self.ui or {}
        rb = ui.get("ribbon", {})
        dlg: Dict[str, Any] = {}
        if name == "formatCells":
            st = d.wb.styles.get(sh.style_id(sel.active))
            nf = parse_fmt(st.num_fmt)
            b = st.borders
            dlg = {
                "tab": params.get("tab", "number"),
                "numFmt": st.num_fmt,
                "category": nf.kind(),
                "fontName": st.font.name,
                "fontSize": str(st.font.size).rstrip("0").rstrip("."),
                "bold": st.font.bold,
                "italic": st.font.italic,
                "underline": st.font.underline,
                "strike": st.font.strike,
                "fontColor": st.font.color.hex(d.wb.theme) or "#000000",
                "fillColor": (st.fill.fg.hex(d.wb.theme) or "") if st.fill.pattern != "none" else "",
                "h": st.align.h,
                "v": st.align.v,
                "wrap": st.align.wrap,
                "shrink": st.align.shrink,
                "indent": str(st.align.indent),
                "rotation": str(st.align.rotation),
                "locked": st.protection.locked,
                "hidden": st.protection.hidden,
                "borderStyle": "thin",
                "borderColor": "#000000",
                "borderPreset": "",
                "sample": "",
            }
            dlg["orig"] = dict(dlg)
        elif name in ("find", "replace"):
            last = s.last_find or {}
            dlg = {"mode": name, "what": last.get("what", ""), "with": "", "matchCase": bool(last.get("matchCase", False)), "wholeCell": bool(last.get("wholeCell", False)), "within": last.get("within", "sheet"), "lookIn": last.get("lookIn", "formulas"), "results": [], "status": ""}
        elif name == "goTo":
            dlg = {"reference": "", "special": "blanks"}
        elif name == "pasteSpecial":
            if s.clipboard is None:
                raise EngineError("Copy something first: Paste Special pastes what you copied in ReflexGrid.")
            dlg = {"what": "all", "operation": "none", "skipBlanks": False, "transpose": False}
        elif name == "defineName":
            from gridcore.formula.printer import quote_sheet

            dlg = {"name": "", "refersTo": "=" + f"{quote_sheet(sh.name)}!{sel.current().abs_a1()}", "scope": "Workbook", "comment": ""}
        elif name == "nameManager":
            dlg = {"names": self._run(s, "formulas.nameManager")["names"]}
        elif name == "cfHighlight":
            dlg = {"kind": params.get("kind", "greaterThan"), "value": "", "value2": "", "preset": "lightRed", "period": "today"}
        elif name == "cfTop":
            dlg = {"kind": params.get("kind", "top"), "rank": "10", "preset": "lightRed"}
        elif name == "cfRule":
            dlg = {"formula": "=", "fill": "#FFC7CE", "font": "#9C0006", "bold": False}
        elif name == "cfManage":
            dlg = {"rules": self._run(s, "home.manageRules", {"scope": "sheet"})["rules"]}
        elif name == "validation":
            v = s.execute("data.validationAt", {})
            dlg = {
                "kind": (v or {}).get("kind", "any"),
                "operator": (v or {}).get("operator", "between"),
                "f1": ((v or {}).get("f1") or "").strip('"') if (v or {}).get("kind") == "list" and ((v or {}).get("f1") or "").startswith('"') else ("=" + v["f1"] if v and v.get("f1") and v.get("kind") == "list" else (v or {}).get("f1", "") or ""),
                "f2": (v or {}).get("f2") or "",
                "allowBlank": (v or {}).get("allowBlank", True),
                "inCellDropdown": (v or {}).get("inCellDropdown", True),
                "inputTitle": (v or {}).get("inputTitle", ""),
                "inputMessage": (v or {}).get("inputMessage", ""),
                "errorStyle": (v or {}).get("errorStyle", "stop"),
                "errorTitle": (v or {}).get("errorTitle", ""),
                "errorMessage": (v or {}).get("errorMessage", ""),
                "tab": "settings",
            }
        elif name == "sort":
            from gridcore.engine.cmd_data import _data_region, guess_header

            si, rng, owner = _data_region(s, {})
            header = owner.header_row if hasattr(owner, "header_row") else (True if owner is not None else guess_header(sh, rng, d.wb))
            cols = []
            for c in range(rng.c0, min(rng.c1, rng.c0 + 200) + 1):
                label = f"Column {col_to_letters(c)}"
                if header:
                    v = sh.value((rng.r0, c))
                    if v is not None:
                        label = str(v) if not isinstance(v, float) else f"{v:g}"
                cols.append({"value": col_to_letters(c), "label": label})
            dlg = {"range": rng.a1(), "header": header, "cols": cols, "k1": cols[0]["value"] if cols else "A", "o1": "asc", "k2": "", "o2": "asc", "k3": "", "o3": "asc", "matchCase": False}
        elif name == "insertFunction":
            dlg = {"search": "", "category": "Most Used", "selected": "SUM"}
        elif name == "hyperlink":
            link = ui.get("link")
            dlg = {"target": link or "", "text": "", "tooltip": ""}
        elif name == "note":
            note = ui.get("note")
            dlg = {"text": (note or {}).get("text", ""), "threaded": bool(params.get("threaded", False))}
        elif name in ("rowHeight", "columnWidth", "defaultWidth"):
            if name == "rowHeight":
                v = sh.row_height(sel.active[0])
            elif name == "columnWidth":
                v = sh.col_width(sel.active[1])
            else:
                v = sh.default_col_width
            dlg = {"value": f"{v:g}"}
        elif name in ("insertCells", "deleteCells"):
            dlg = {"shift": "down" if name == "insertCells" else "up"}
        elif name == "renameSheet":
            idx = int(params.get("index", d.wb.active_sheet))
            dlg = {"index": idx, "name": d.wb.sheets[idx].name}
        elif name == "moveSheet":
            idx = int(params.get("index", d.wb.active_sheet))
            dlg = {"index": idx, "to": str(idx), "copy": False, "sheets": [x.name for x in d.wb.sheets]}
        elif name == "tabColor":
            dlg = {"index": int(params.get("index", d.wb.active_sheet))}
        elif name == "unhideSheet":
            hidden = [x.name for x in d.wb.sheets if x.visibility != "visible"]
            if not hidden:
                raise EngineError("No sheets are hidden.")
            dlg = {"sheet": hidden[0], "sheets": hidden}
        elif name == "protectSheet":
            dlg = {"password": "", "formatCells": False, "formatColumns": False, "formatRows": False, "insertRows": False, "insertColumns": False, "deleteRows": False, "deleteColumns": False, "sort": False, "autoFilter": False}
        elif name in ("unprotectSheet", "protectWorkbook", "unprotectWorkbook"):
            dlg = {"password": ""}
        elif name == "insertChart":
            dlg = {"kind": params.get("kind", "columnClustered"), "title": ""}
        elif name == "insertTable":
            from gridcore.engine.cmd_data import guess_header
            from gridcore.engine.selection import current_region

            rng = sel.current() if not sel.current().is_single() else current_region(sh, sel.active)
            dlg = {"range": rng.a1(), "header": guess_header(sh, rng, d.wb) or all(isinstance(sh.value((rng.r0, c)), str) for c in range(rng.c0, min(rng.c1, rng.c0 + 50) + 1)), "style": "TableStyleMedium2"}
        elif name == "removeDuplicates":
            from gridcore.engine.cmd_data import _data_region, guess_header

            si, rng, owner = _data_region(s, {})
            header = owner.header_row if hasattr(owner, "header_row") else guess_header(sh, rng, d.wb)
            cols = []
            for i, c in enumerate(range(rng.c0, min(rng.c1, rng.c0 + 100) + 1)):
                v = sh.value((rng.r0, c)) if header else None
                cols.append({"index": i, "label": (str(v) if v is not None else f"Column {col_to_letters(c)}"), "checked": True})
            dlg = {"range": rng.a1(), "header": header, "cols": cols}
        elif name == "textToColumns":
            dlg = {"tab": True, "comma": False, "semicolon": False, "space": False, "other": "", "consecutive": False}
        elif name == "properties":
            p = d.wb.props
            dlg = {"title": p.title, "subject": p.subject, "author": p.author, "company": p.company, "keywords": p.keywords, "description": p.description}
        elif name == "open":
            dlg = {}
        elif name == "insertPicture":
            dlg = {}
        elif name == "chartTitle":
            ch = next((c for c in sh.charts if c.id == self.selected_object), None)
            if ch is None:
                raise EngineError("Select a chart first.")
            dlg = {"id": ch.id, "title": ch.title or "", "xTitle": ch.x_title or "", "yTitle": ch.y_title or ""}
        elif name == "shortcuts":
            dlg = {}
        elif name == "goalSeek":
            dlg = {"setCell": a1(*sel.active), "toValue": "", "byChangingCell": ""}
        elif name == "sparkline":
            r = sel.current()
            loc = RangeRef(r.r0, r.c1 + 1, r.r1, r.c1 + 1) if r.c1 + 1 < MAX_COLS else r
            dlg = {"kind": params.get("kind", "line"), "data": r.a1(), "location": loc.a1()}
        elif name == "selectData":
            ch = next((c for c in sh.charts if c.id == self.selected_object), None)
            if ch is None:
                raise EngineError("Select a chart first.")
            dlg = {"id": ch.id, "range": ch.source or ""}
        else:
            raise EngineError(f"unknown dialog `{name}`")
        self.dlg = dlg
        self.dlg_error = ""
        self.dialog = name

    def _open_filter(self, s: Session, r: int, c: int, table: Optional[str]) -> None:
        sel_cell = a1(r, c)
        s.execute("selection.set", {"range": sel_cell})
        p = {"col": col_to_letters(c)}
        if table:
            p["table"] = table
        res = s.execute("data.filterValues", p)
        vals = [{"text": v["text"], "checked": v["checked"]} for v in res["values"]]
        self.dlg = {
            "col": col_to_letters(c),
            "table": table or "",
            "values": vals,
            "blanks": res["blanks"],
            "blanksChecked": res["blanksChecked"],
            "active": res["active"],
            "numeric": res["numeric"],
            "search": "",
            "op": "",
            "value": "",
            "op2": "",
            "value2": "",
            "and": True,
            "mode": "values",
        }
        self.dlg_error = ""
        self.dialog = "filter"

    @rx.event
    def dialog_ok(self, action: str = "ok"):
        entry = self._entry()
        s = entry.session
        name = self.dialog
        d = dict(self.dlg)
        out = None
        keep_open = False
        try:
            out, keep_open = self._dialog_action(s, name, action, d)
        except EngineError as ex:
            self.dlg_error = str(ex).split("\n", 1)[0]
            self._refresh(entry)
            return None
        if not keep_open:
            self.dialog = ""
            self.dlg = {}
            self.dlg_error = ""
            self._focus()
        self._refresh(entry)
        return out

    def _dialog_action(self, s: Session, name: str, action: str, d: Dict[str, Any]):
        run = lambda cid, **p: s.execute(cid, p)
        if name == "formatCells":
            o = d.get("orig", {})
            patch: Dict[str, Any] = {}
            if d.get("numFmt") != o.get("numFmt"):
                patch["numFmt"] = d.get("numFmt") or "General"
            font = {}
            for k_ui, k in (("fontName", "name"), ("bold", "bold"), ("italic", "italic"), ("underline", "underline"), ("strike", "strike"), ("fontColor", "color")):
                if d.get(k_ui) != o.get(k_ui):
                    font[k] = d.get(k_ui)
            if d.get("fontSize") != o.get("fontSize"):
                size = _num(d.get("fontSize"))
                if size is None or not 1 <= size <= 409:
                    raise EngineError("The font size must be a number between 1 and 409.")
                font["size"] = size
            if font:
                patch["font"] = font
            if d.get("fillColor") != o.get("fillColor"):
                patch["fill"] = {"color": d.get("fillColor") or None}
            align = {}
            for k in ("h", "v", "wrap", "shrink"):
                if d.get(k) != o.get(k):
                    align[k] = d.get(k)
            for k in ("indent", "rotation"):
                if d.get(k) != o.get(k):
                    v = _num(d.get(k))
                    if v is None:
                        raise EngineError(f"{k.title()} must be a number.")
                    align[k] = int(v)
            if align:
                patch["align"] = align
            prot = {k: d.get(k) for k in ("locked", "hidden") if d.get(k) != o.get(k)}
            if prot:
                patch["protection"] = prot
            if patch:
                run("home.formatCells", **patch)
            if d.get("borderPreset"):
                run("home.borders", preset=d["borderPreset"], style=d.get("borderStyle", "thin"), color=d.get("borderColor") or "#000000")
            return None, False
        if name in ("find", "replace"):
            common = {"what": d.get("what", ""), "matchCase": d.get("matchCase"), "wholeCell": d.get("wholeCell"), "within": d.get("within", "sheet"), "lookIn": d.get("lookIn", "formulas")}
            if action == "findAll":
                r = s.execute("edit.find", {**common, "all": True})
                d["results"] = r["matches"][:500]
                d["status"] = f"{len(r['matches'])} cell(s) found"
                self.dlg = d
                return None, True
            if action == "findNext":
                r = s.execute("edit.find", common)
                d["status"] = f"{r['sheet']}!{r['cell']}"
                self.dlg = d
                return None, True
            if action == "replace":
                s.execute("edit.replace", {**common, "with": d.get("with", "")})
                return None, True
            if action == "replaceAll":
                r = s.execute("edit.replace", {**common, "with": d.get("with", ""), "all": True})
                d["status"] = f"All done. We made {r['replaced']} replacements."
                self.dlg = d
                return None, True
            if action.startswith("goto:"):
                _, sheet, cell = action.split(":", 2)
                s.execute("edit.goTo", {"reference": f"'{sheet}'!{cell}"})
                return None, True
            return None, False
        if name == "goTo":
            if action == "special":
                run("edit.goToSpecial", kind=d.get("special", "blanks"))
            else:
                ref = (d.get("reference") or "").strip()
                if not ref:
                    raise EngineError("Enter a reference.")
                run("edit.goTo", reference=ref)
            return None, False
        if name == "pasteSpecial":
            run("edit.pasteSpecial", what=d.get("what", "all"), operation=d.get("operation", "none"), skipBlanks=bool(d.get("skipBlanks")), transpose=bool(d.get("transpose")))
            return None, False
        if name == "defineName":
            p = {"name": d.get("name", ""), "refersTo": d.get("refersTo", ""), "comment": d.get("comment", "")}
            if d.get("scope") and d["scope"] != "Workbook":
                p["scope"] = d["scope"]
            run("formulas.defineName", **p)
            return None, False
        if name == "nameManager":
            if action.startswith("delete:"):
                nm = action.split(":", 1)[1]
                entry = next((x for x in d.get("names", []) if x["name"] == nm), None)
                if entry and entry.get("table"):
                    raise EngineError("A table name can't be deleted here; use Convert to Range.")
                p = {"name": nm}
                if entry and entry.get("scope") not in (None, "Workbook"):
                    p["scope"] = entry["scope"]
                s.execute("formulas.deleteName", p)
                d["names"] = s.execute("formulas.nameManager", {})["names"]
                self.dlg = d
                return None, True
            return None, False
        if name == "cfHighlight":
            k = d.get("kind")
            rule: Dict[str, Any] = {"preset": d.get("preset", "lightRed")}
            if k in ("greaterThan", "lessThan", "equal"):
                rule.update({"type": "cellIs", "operator": k, "value": d.get("value", "")})
            elif k == "between":
                rule.update({"type": "cellIs", "operator": "between", "value": d.get("value", ""), "value2": d.get("value2", "")})
            elif k == "text":
                rule.update({"type": "containsText", "text": d.get("value", "")})
            elif k == "date":
                rule.update({"type": "timePeriod", "period": d.get("period", "today")})
            elif k == "duplicate":
                rule.update({"type": "duplicateValues" if d.get("value", "duplicate") != "unique" else "uniqueValues"})
            run("home.conditionalFormat", rule=rule)
            return None, False
        if name == "cfTop":
            k = d.get("kind", "top")
            rank = int(_num(d.get("rank"), 10) or 10)
            if k in ("top", "bottom", "topPercent", "bottomPercent"):
                rule = {"type": "top10", "rank": rank, "bottom": k.startswith("bottom"), "percent": k.endswith("Percent"), "preset": d.get("preset", "lightRed")}
            else:
                rule = {"type": "aboveAverage", "below": k == "below", "preset": d.get("preset", "lightRed")}
            run("home.conditionalFormat", rule=rule)
            return None, False
        if name == "cfRule":
            style = {"fill": {"color": d.get("fill") or None}, "font": {"color": d.get("font") or None, "bold": bool(d.get("bold"))}}
            run("home.conditionalFormat", rule={"type": "expression", "formula": d.get("formula", ""), "style": style})
            return None, False
        if name == "cfManage":
            if action.startswith("delete:"):
                s.execute("home.deleteRule", {"index": int(action.split(":", 1)[1])})
            elif action.startswith("up:") or action.startswith("down:"):
                kind, idx = action.split(":", 1)
                s.execute("home.moveRule", {"index": int(idx), "up": kind == "up"})
            elif action == "clearSheet":
                s.execute("home.clearRules", {"scope": "sheet"})
            else:
                return None, False
            d["rules"] = s.execute("home.manageRules", {"scope": "sheet"})["rules"]
            self.dlg = d
            return None, True
        if name == "validation":
            if action == "clear":
                run("data.clearValidation")
                return None, False
            p = {k: d.get(k) for k in ("kind", "operator", "f1", "f2", "allowBlank", "inCellDropdown", "inputTitle", "inputMessage", "errorStyle", "errorTitle", "errorMessage")}
            if p["kind"] == "any" and not (p.get("inputTitle") or p.get("inputMessage")):
                run("data.clearValidation")
                return None, False
            run("data.validation", **p)
            return None, False
        if name == "sort":
            keys = []
            for i in (1, 2, 3):
                col = d.get(f"k{i}")
                if col:
                    keys.append({"col": col, "order": d.get(f"o{i}", "asc")})
            if not keys:
                raise EngineError("Choose a column to sort by.")
            run("data.sort", range=d.get("range"), keys=keys, header=bool(d.get("header")), matchCase=bool(d.get("matchCase")))
            return None, False
        if name == "filter":
            p: Dict[str, Any] = {"col": d.get("col")}
            if d.get("table"):
                p["table"] = d["table"]
            if action in ("sortAsc", "sortDesc"):
                s.execute("selection.set", {"range": f"{d.get('col')}1"} if False else {"range": s.doc().selection.current().a1()})
                s.execute("data.sortAscending" if action == "sortAsc" else "data.sortDescending", {"col": d.get("col")})
                return None, False
            if action == "clear":
                p["criterion"] = None
                s.execute("data.filterColumn", p)
                return None, False
            if d.get("mode") == "custom" and d.get("op"):
                conds = [{"op": d["op"], "value": d.get("value", "")}]
                if d.get("op2"):
                    conds.append({"op": d["op2"], "value": d.get("value2", "")})
                p["criterion"] = {"kind": "custom", "conditions": conds, "and": bool(d.get("and", True))}
            else:
                vals = d.get("values", [])
                if all(v["checked"] for v in vals) and (d.get("blanksChecked") or not d.get("blanks")):
                    p["criterion"] = None
                else:
                    p["criterion"] = {"kind": "values", "values": [v["text"] for v in vals if v["checked"]], "blanks": bool(d.get("blanksChecked"))}
            s.execute("data.filterColumn", p)
            return None, False
        if name == "insertFunction":
            fn = self.fn_selected.get("name") or d.get("selected") or "SUM"
            self.edit_request = {"seq": (self.edit_request or {}).get("seq", 0) + 1, "text": f"={fn}("}
            return None, False
        if name == "hyperlink":
            if action == "remove":
                run("insert.removeHyperlink")
                return None, False
            p = {"target": d.get("target", "")}
            if d.get("text"):
                p["text"] = d["text"]
            if d.get("tooltip"):
                p["tooltip"] = d["tooltip"]
            run("insert.hyperlink", **p)
            return None, False
        if name == "note":
            if action == "delete":
                run("review.deleteComment")
                return None, False
            if d.get("threaded"):
                run("review.newComment", text=d.get("text", ""))
            else:
                run("review.newNote", text=d.get("text", ""))
            return None, False
        if name == "rowHeight":
            run("home.rowHeight", height=_num(d.get("value"), -1))
            return None, False
        if name == "columnWidth":
            run("home.columnWidth", width=_num(d.get("value"), -1))
            return None, False
        if name == "defaultWidth":
            run("home.defaultWidth", width=_num(d.get("value"), -1))
            return None, False
        if name == "insertCells":
            run("home.insertCells", shift=d.get("shift", "down"))
            return None, False
        if name == "deleteCells":
            run("home.deleteCells", shift=d.get("shift", "up"))
            return None, False
        if name == "renameSheet":
            run("sheet.rename", sheet=int(d["index"]), name=d.get("name", ""))
            return None, False
        if name == "moveSheet":
            run("sheet.move", sheet=int(d["index"]), to=int(_num(d.get("to"), 0)), copy=bool(d.get("copy")))
            return None, False
        if name == "tabColor":
            run("sheet.tabColor", sheet=int(d["index"]), color=action if action != "ok" else None)
            return None, False
        if name == "unhideSheet":
            run("sheet.unhide", sheet=d.get("sheet"))
            return None, False
        if name == "protectSheet":
            p = {k: bool(d.get(k)) for k in ("formatCells", "formatColumns", "formatRows", "insertRows", "insertColumns", "deleteRows", "deleteColumns", "sort", "autoFilter")}
            if d.get("password"):
                p["password"] = d["password"]
            run("review.protectSheet", **p)
            return None, False
        if name == "unprotectSheet":
            run("review.unprotectSheet", password=d.get("password", ""))
            return None, False
        if name == "protectWorkbook":
            run("review.protectWorkbook", password=d.get("password", ""))
            return None, False
        if name == "unprotectWorkbook":
            run("review.unprotectWorkbook", password=d.get("password", ""))
            return None, False
        if name == "insertChart":
            p = {"kind": d.get("kind", "columnClustered")}
            if d.get("title"):
                p["title"] = d["title"]
            run("insert.chart", **p)
            return None, False
        if name == "insertTable":
            run("insert.table", range=d.get("range"), header=bool(d.get("header")), style=d.get("style", "TableStyleMedium2"))
            return None, False
        if name == "removeDuplicates":
            cols = [c["index"] for c in d.get("cols", []) if c.get("checked")]
            if not cols:
                raise EngineError("Select at least one column.")
            r = s.execute("data.removeDuplicates", {"range": d.get("range"), "columns": cols, "header": bool(d.get("header"))})
            return rx.toast.info(f"{r.get('removed', 0)} duplicate values found and removed; {r.get('remaining', 0)} unique values remain.", position="bottom-center"), False
        if name == "textToColumns":
            delims = [ch for k, ch in (("tab", "\t"), ("comma", ","), ("semicolon", ";"), ("space", " ")) if d.get(k)]
            if not delims and not d.get("other"):
                raise EngineError("Choose at least one delimiter.")
            run("data.textToColumns", delimiters=delims or [str(d.get("other"))[:1]], other=d.get("other") or None, consecutive=bool(d.get("consecutive")))
            return None, False
        if name == "properties":
            run("file.properties", **{k: d.get(k, "") for k in ("title", "subject", "author", "company", "keywords", "description")})
            return None, False
        if name == "chartTitle":
            run("chart.update", id=int(d["id"]), title=d.get("title") or None, xTitle=d.get("xTitle") or None, yTitle=d.get("yTitle") or None)
            return None, False
        if name == "goalSeek":
            goal = _num(d.get("toValue"))
            if goal is None:
                raise EngineError("To value must be a number.")
            r = s.execute("data.goalSeek", {"setCell": d.get("setCell", ""), "toValue": goal, "byChangingCell": d.get("byChangingCell", "")})
            msg = f"Goal Seeking with Cell {r['target']} " + ("found a solution." if r["found"] else "may not have found a solution.") + f" Target value: {goal:g}. Current value: {r['result']}"
            return (rx.toast.success(msg, position="bottom-center") if r["found"] else rx.toast.warning(msg, position="bottom-center")), False
        if name == "sparkline":
            run("insert.sparkline", kind=d.get("kind", "line"), data=d.get("data", ""), location=d.get("location", ""))
            return None, False
        if name == "selectData":
            run("chart.selectData", id=int(d["id"]), range=d.get("range", ""))
            return None, False
        return None, False

    @rx.event
    def number_format(self, value: str):
        if value == "__more__":
            return GridState.open_dialog("formatCells", {"tab": "number"})
        return GridState.cmd("home.numberFormat", {"format": value})

    @rx.event
    def set_edit_request(self, text: str):
        self.edit_request = {"seq": (self.edit_request or {}).get("seq", 0) + 1, "text": text}

    @rx.event
    def audit(self, what: str):
        entry = self._entry()
        s = entry.session
        try:
            if what == "precedents":
                r = s.execute("formulas.tracePrecedents", {})
                items = [f"{p['sheet']}!{p['range']}" for p in r["precedents"]]
                return rx.toast.info("Precedents: " + (", ".join(items) if items else "none"), position="bottom-center")
            if what == "dependents":
                r = s.execute("formulas.traceDependents", {})
                items = [f"{p['sheet']}!{p['cell']}" for p in r["dependents"]]
                return rx.toast.info("Dependents: " + (", ".join(items[:30]) if items else "none"), position="bottom-center")
            if what == "errors":
                r = s.execute("formulas.errorChecking", {})
                items = [f"{e['sheet']}!{e['cell']} {e['error']}" for e in r["errors"]]
                circ = r.get("circular") or []
                msg = ("Errors: " + ", ".join(items[:20])) if items else "No formula errors were found."
                if circ:
                    msg += " Circular references: " + ", ".join(circ)
                return rx.toast.info(msg, position="bottom-center")
            if what == "invalid":
                r = s.execute("data.circleInvalid", {})
                items = r["invalid"]
                return rx.toast.info(("Invalid data in: " + ", ".join(items[:40])) if items else "All data is valid.", position="bottom-center")
        except EngineError as ex:
            return rx.toast.error(str(ex).split("\n", 1)[0], position="bottom-center")

    # ------------------------------------------------------------ chart ribbon

    @rx.event
    def chart_cmd(self, field: str, value: Any):
        entry = self._entry()
        oid = self.selected_object
        if oid < 0:
            return rx.toast.error("Select a chart first.")
        try:
            if field == "switch":
                self._run(entry.session, "chart.switchRowColumn", {"id": oid})
            elif field == "delete":
                self._run(entry.session, "object.delete", {"id": oid})
                self.selected_object = -1
                self.tab = "Home"
            else:
                self._run(entry.session, "chart.update", {"id": oid, field: value})
        except EngineError as ex:
            return rx.toast.error(str(ex).split("\n", 1)[0])
        self._refresh(entry)


def _spans(idx: List[int], axis: str) -> str:
    idx = sorted(set(idx))
    out = []
    start = prev = idx[0]
    for i in idx[1:]:
        if i == prev + 1:
            prev = i
            continue
        out.append((start, prev))
        start = prev = i
    out.append((start, prev))
    if axis == "cols":
        return ",".join(f"{col_to_letters(a)}:{col_to_letters(b)}" for a, b in out)
    return ",".join(f"{a + 1}:{b + 1}" for a, b in out)


def _fmt_stat(v: float) -> str:
    if v is None:
        return ""
    if abs(v) >= 1e15 or (v != 0 and abs(v) < 1e-9):
        return f"{v:.6g}"
    s = f"{v:,.10f}".rstrip("0").rstrip(".")
    return s


def _js_str(s: str) -> str:
    import json

    return json.dumps(s)
