"""Formatting commands: font, fill, borders, alignment, merging, number formats, cell styles,
Format Cells, row heights and column widths, hiding and unhiding."""

from __future__ import annotations

from dataclasses import replace
from typing import Any, Callable, Dict, List, Optional

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.value import number_to_text
from ..model.cell import Cell
from ..model.features import LineInfo
from ..model.style import (
    AUTO,
    Alignment,
    BorderLine,
    Borders,
    Color,
    Fill,
    Font,
    Protection,
    Style,
    Theme,
)
from ..numfmt import parse as parse_fmt
from ..numfmt.parse import validate as valid_format
from .layout import autofit_width
from .selection import Selection
from .session import EngineError, Session, command, edit, need, p_bool, p_float, p_int, p_str, param_sheet

ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
COMMA = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'
FORMAT_PRESETS = {
    "general": "General",
    "number": "0.00",
    "currency": '"$"#,##0.00',
    "accounting": ACCOUNTING,
    "shortDate": "m/d/yyyy",
    "longDate": "dddd, mmmm d, yyyy",
    "time": "h:mm:ss AM/PM",
    "percent": "0.00%",
    "fraction": "# ?/?",
    "scientific": "0.00E+00",
    "text": "@",
}
FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
MAX_STYLE_CELLS = 2_000_000


def _ranges(s: Session, p: Dict[str, Any]) -> List[RangeRef]:
    t = p.get("range")
    if t:
        out = []
        for part in str(t).rsplit("!", 1)[-1].split(","):
            r = RangeRef.parse(part)
            if r is None:
                raise EngineError(f"not a range: `{part}`")
            out.append(r)
        return out
    return list(s.doc().selection.ranges)


def check_format_allowed(sh) -> None:
    if sh.protection is not None and not sh.protection.format_cells:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")


def apply_style_ctx(ctx, si: int, ranges: List[RangeRef], f: Callable[[Style], Style]) -> None:
    wb = ctx.wb
    sm = ctx.sheet_mut(si)
    cache: Dict[int, int] = {}

    def derive(sid: int) -> int:
        n = cache.get(sid)
        if n is None:
            n = wb.styles.intern(f(wb.styles.get(sid)))
            cache[sid] = n
        return n

    fit = set()
    for rng in ranges:
        if rng.is_full_cols():
            for c in range(rng.c0, rng.c1 + 1):
                info = sm.col_info(c)
                sm.set_col(c, replace(info, style=derive(info.style or 0)))
            for r in list(sm.rows):
                info = sm.rows[r]
                if info.style is not None:
                    # Cells where a styled row crosses the columns get their own entry.
                    for c in range(rng.c0, min(rng.c1, rng.c0 + 16383) + 1):
                        if sm.cell((r, c)) is None:
                            sid = derive(info.style)
                            if sid != 0:
                                sm.set_cell((r, c), Cell(None, None, sid))
            for rc, cell in list(sm.cells.iter_range(rng)):
                sm.set_cell(rc, cell.with_style(derive(cell.style)))
                fit.add(rc[0])
        elif rng.is_full_rows():
            for r in range(rng.r0, rng.r1 + 1):
                info = sm.row_info(r)
                sm.set_row(r, replace(info, style=derive(info.style or 0)))
                fit.add(r)
            for rc, cell in list(sm.cells.iter_range(rng)):
                sm.set_cell(rc, cell.with_style(derive(cell.style)))
        else:
            if rng.count() > MAX_STYLE_CELLS:
                raise EngineError("The selection is too large.")
            for rc in rng.cells():
                cell = sm.cell(rc)
                if cell is not None:
                    n = derive(cell.style)
                    if n != cell.style:
                        sm.set_cell(rc, cell.with_style(n))
                else:
                    n = derive(sm.style_id(rc))
                    if n != 0:
                        sm.set_cell(rc, Cell(None, None, n))
            for r in range(rng.r0, min(rng.r1, rng.r0 + 5000) + 1):
                fit.add(r)
    for r in list(fit)[:20000]:
        ctx.fit_rows.append((si, r))


def apply_style(s: Session, p: Dict[str, Any], f: Callable[[Style], Style], widen: bool = False):
    d = s.doc()
    si = param_sheet(s, p)
    check_format_allowed(d.wb.sheet(si))
    ranges = _ranges(s, p)

    def run(ctx):
        apply_style_ctx(ctx, si, ranges, f)
        if widen:
            widen_for_numbers(ctx, si, ranges)

    edit(s, run)
    return None


def widen_for_numbers(ctx, si: int, ranges: List[RangeRef]) -> None:
    """Numbers that no longer fit a standard-width column widen it (as Excel does when a number
    format or font change would show ####)."""
    from .cmd_edit import auto_widen

    sh = ctx.wb.sheet(si)
    used = sh.used_range()
    if used is None:
        return
    n = 0
    for rng in ranges:
        inter = rng.intersection(used)
        if inter is None:
            continue
        for rc, cell in list(sh.cells.iter_range(inter)):
            if isinstance(cell.value, float):
                auto_widen(ctx, si, rc)
                n += 1
                if n > 100_000:
                    return


def _active_style(s: Session) -> Style:
    d = s.doc()
    return d.wb.styles.get(d.sheet.style_id(d.selection.active))


def _toggle(s: Session, p: Dict[str, Any], get: Callable[[Style], bool], set_: Callable[[Style, bool], Style]):
    on = p.get("on")
    if on is None:
        on = not get(_active_style(s))
    on = bool(on)
    return apply_style(s, p, lambda st: set_(st, on))


def _font(st: Style, **kw) -> Style:
    return replace(st, font=replace(st.font, **kw))


def _align(st: Style, **kw) -> Style:
    return replace(st, align=replace(st.align, **kw))


@command("home.bold", "Bold", menu=("Home", "Font"), shortcut="Ctrl+B", params="{on?: bool, range?}")
def bold(s, p):
    return _toggle(s, p, lambda st: st.font.bold, lambda st, on: _font(st, bold=on))


@command("home.italic", "Italic", menu=("Home", "Font"), shortcut="Ctrl+I", params="{on?: bool, range?}")
def italic(s, p):
    return _toggle(s, p, lambda st: st.font.italic, lambda st, on: _font(st, italic=on))


@command("home.underline", "Underline", menu=("Home", "Font"), shortcut="Ctrl+U", params="{on?: bool, range?}")
def underline(s, p):
    return _toggle(s, p, lambda st: st.font.underline == "single", lambda st, on: _font(st, underline="single" if on else "none"))


@command("home.doubleUnderline", "Double Underline", menu=("Home", "Font"), params="{on?: bool, range?}")
def double_underline(s, p):
    return _toggle(s, p, lambda st: st.font.underline == "double", lambda st, on: _font(st, underline="double" if on else "none"))


@command("home.strikethrough", "Strikethrough", menu=("Home", "Font"), shortcut="Ctrl+5", params="{on?: bool, range?}")
def strikethrough(s, p):
    return _toggle(s, p, lambda st: st.font.strike, lambda st, on: _font(st, strike=on))


@command("home.superscript", "Superscript", menu=("Home", "Font"), params="{on?: bool, range?}")
def superscript(s, p):
    return _toggle(s, p, lambda st: st.font.vert == "superscript", lambda st, on: _font(st, vert="superscript" if on else "baseline"))


@command("home.subscript", "Subscript", menu=("Home", "Font"), params="{on?: bool, range?}")
def subscript(s, p):
    return _toggle(s, p, lambda st: st.font.vert == "subscript", lambda st, on: _font(st, vert="subscript" if on else "baseline"))


@command("home.fontName", "Font", menu=("Home", "Font"), params='{name: "Arial", range?}')
def font_name(s, p):
    name = str(need(p, "name", "home.fontName")).strip()
    if not name:
        raise EngineError("home.fontName: empty font name")
    return apply_style(s, p, lambda st: _font(st, name=name))


@command("home.fontSize", "Font Size", menu=("Home", "Font"), params="{size: 12, range?}")
def font_size(s, p):
    size = p_float(p, "size")
    if size is None or not 1 <= size <= 409:
        raise EngineError("The font size must be a number between 1 and 409.")
    return apply_style(s, p, lambda st: _font(st, size=round(size * 2) / 2), widen=True)


def _step_size(cur: float, up: bool) -> float:
    if up:
        for x in FONT_SIZES:
            if x > cur:
                return float(x)
        return min(cur + 10, 409.0)
    for x in reversed(FONT_SIZES):
        if x < cur:
            return float(x)
    return max(cur - 1, 1.0)


@command("home.increaseFontSize", "Increase Font Size", menu=("Home", "Font"), shortcut="Ctrl+Shift+>", params="{range?}")
def inc_font(s, p):
    return apply_style(s, p, lambda st: _font(st, size=_step_size(st.font.size, True)), widen=True)


@command("home.decreaseFontSize", "Decrease Font Size", menu=("Home", "Font"), shortcut="Ctrl+Shift+<", params="{range?}")
def dec_font(s, p):
    return apply_style(s, p, lambda st: _font(st, size=_step_size(st.font.size, False)))


def parse_color(v: Any) -> Color:
    if v is None or v == "" or v == "auto" or v == "none":
        return AUTO
    if isinstance(v, dict):
        if "theme" in v:
            return Color.theme(int(v["theme"]), int(v.get("tint", 0)))
        if "rgb" in v:
            c = Color.from_hex(str(v["rgb"]))
            if c is None:
                raise EngineError(f"invalid colour `{v}`")
            return c
    c = Color.from_json(str(v))
    if c.is_auto and str(v) not in ("", "auto"):
        raise EngineError(f"invalid colour `{v}`")
    return c


@command("home.fontColor", "Font Color", menu=("Home", "Font"), params='{color: "#C00000" | "theme:4:0" | null, range?}')
def font_color(s, p):
    c = parse_color(p.get("color"))
    return apply_style(s, p, lambda st: _font(st, color=c))


@command("home.fillColor", "Fill Color", menu=("Home", "Font"), params='{color: "#FFFF00" | null (No Fill), range?}')
def fill_color(s, p):
    c = parse_color(p.get("color"))
    f = Fill() if c.is_auto else Fill.solid(c)
    return apply_style(s, p, lambda st: replace(st, fill=f))


BORDER_PRESETS = (
    "all", "outside", "thickOutside", "none", "bottom", "top", "left", "right", "doubleBottom", "thickBottom",
    "topBottom", "topThickBottom", "topDoubleBottom", "inside", "insideH", "insideV", "diagDown", "diagUp",
)


@command("home.borders", "Borders", menu=("Home", "Font"), params="{preset: all|outside|thickOutside|none|bottom|top|left|right|doubleBottom|thickBottom|topBottom|topThickBottom|topDoubleBottom|inside|insideH|insideV|diagDown|diagUp, style?: thin, color?, range?}")
def borders(s: Session, p: Dict[str, Any]):
    preset = p_str(p, "preset", "all")
    if preset not in BORDER_PRESETS:
        raise EngineError(f"home.borders: unknown preset `{preset}`")
    style = p_str(p, "style", "thin")
    color = parse_color(p.get("color"))
    line = BorderLine(style, color)
    thick = BorderLine("medium", color)
    double = BorderLine("double", color)
    none = BorderLine()
    d = s.doc()
    si = param_sheet(s, p)
    check_format_allowed(d.wb.sheet(si))
    ranges = _ranges(s, p)

    def f(ctx):
        for rng in ranges:
            if rng.is_full_cols() or rng.is_full_rows():
                used = ctx.wb.sheet(si).used_range()
                if used is None:
                    continue
                rng = RangeRef(rng.r0, rng.c0, min(rng.r1, used.r1), min(rng.c1, used.c1))
                if rng.r0 > rng.r1 or rng.c0 > rng.c1:
                    continue
            if rng.count() > MAX_STYLE_CELLS:
                raise EngineError("The selection is too large.")
            groups: Dict[tuple, List[RangeRef]] = {}

            def put(cells: RangeRef, **sides):
                key = tuple(sorted(sides.items()))
                groups.setdefault(key, []).append(cells)

            top_row = RangeRef(rng.r0, rng.c0, rng.r0, rng.c1)
            bottom_row = RangeRef(rng.r1, rng.c0, rng.r1, rng.c1)
            left_col = RangeRef(rng.r0, rng.c0, rng.r1, rng.c0)
            right_col = RangeRef(rng.r0, rng.c1, rng.r1, rng.c1)
            if preset == "all":
                put(rng, left=line, right=line, top=line, bottom=line)
            elif preset == "none":
                put(rng, left=none, right=none, top=none, bottom=none, diag_down=none, diag_up=none)
            elif preset in ("outside", "thickOutside"):
                ln = thick if preset == "thickOutside" else line
                put(top_row, top=ln)
                put(bottom_row, bottom=ln)
                put(left_col, left=ln)
                put(right_col, right=ln)
            elif preset == "bottom":
                put(bottom_row, bottom=line)
            elif preset == "top":
                put(top_row, top=line)
            elif preset == "left":
                put(left_col, left=line)
            elif preset == "right":
                put(right_col, right=line)
            elif preset == "doubleBottom":
                put(bottom_row, bottom=double)
            elif preset == "thickBottom":
                put(bottom_row, bottom=thick)
            elif preset == "topBottom":
                put(top_row, top=line)
                put(bottom_row, bottom=line)
            elif preset == "topThickBottom":
                put(top_row, top=line)
                put(bottom_row, bottom=thick)
            elif preset == "topDoubleBottom":
                put(top_row, top=line)
                put(bottom_row, bottom=double)
            elif preset in ("inside", "insideH", "insideV"):
                if preset in ("inside", "insideH") and rng.height > 1:
                    put(RangeRef(rng.r0, rng.c0, rng.r1 - 1, rng.c1), bottom=line)
                    put(RangeRef(rng.r0 + 1, rng.c0, rng.r1, rng.c1), top=line)
                if preset in ("inside", "insideV") and rng.width > 1:
                    put(RangeRef(rng.r0, rng.c0, rng.r1, rng.c1 - 1), right=line)
                    put(RangeRef(rng.r0, rng.c0 + 1, rng.r1, rng.c1), left=line)
            elif preset == "diagDown":
                put(rng, diag_down=line)
            elif preset == "diagUp":
                put(rng, diag_up=line)
            for key, cells in groups.items():
                sides = dict(key)
                apply_style_ctx(ctx, si, cells, lambda st, sides=sides: replace(st, borders=replace(st.borders, **sides)))

    edit(s, f)
    return None


def _set_h(h: str):
    def run(s, p):
        cur = _active_style(s).align.h
        target = "general" if cur == h and p.get("on") is None else h
        if p.get("on") is False:
            target = "general"
        return apply_style(s, p, lambda st: _align(st, h=target))

    return run


def _set_v(v: str):
    def run(s, p):
        return apply_style(s, p, lambda st: _align(st, v=v))

    return run


command("home.alignLeft", "Align Left", menu=("Home", "Alignment"), params="{range?}")(_set_h("left"))
command("home.alignCenter", "Center", menu=("Home", "Alignment"), shortcut="Ctrl+E", params="{range?}")(_set_h("center"))
command("home.alignRight", "Align Right", menu=("Home", "Alignment"), params="{range?}")(_set_h("right"))
command("home.justify", "Justify", menu=("Home", "Alignment"), params="{range?}")(_set_h("justify"))
command("home.alignTop", "Top Align", menu=("Home", "Alignment"), params="{range?}")(_set_v("top"))
command("home.alignMiddle", "Middle Align", menu=("Home", "Alignment"), params="{range?}")(_set_v("center"))
command("home.alignBottom", "Bottom Align", menu=("Home", "Alignment"), params="{range?}")(_set_v("bottom"))


@command("home.horizontalAlign", "Horizontal Alignment", params="{value: general|left|center|right|fill|justify|centerContinuous|distributed, range?}")
def horizontal_align(s, p):
    v = p_str(p, "value", "general")
    if v not in ("general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"):
        raise EngineError(f"unknown alignment `{v}`")
    return apply_style(s, p, lambda st: _align(st, h=v))


@command("home.wrapText", "Wrap Text", menu=("Home", "Alignment"), params="{on?: bool, range?}")
def wrap_text(s, p):
    return _toggle(s, p, lambda st: st.align.wrap, lambda st, on: _align(st, wrap=on))


@command("home.shrinkToFit", "Shrink to Fit", menu=("Home", "Alignment"), params="{on?: bool, range?}")
def shrink(s, p):
    return _toggle(s, p, lambda st: st.align.shrink, lambda st, on: _align(st, shrink=on))


@command("home.increaseIndent", "Increase Indent", menu=("Home", "Alignment"), params="{range?}")
def inc_indent(s, p):
    def f(st):
        h = st.align.h if st.align.h in ("left", "right", "distributed") else "left"
        return _align(st, indent=min(st.align.indent + 1, 250), h=h)

    return apply_style(s, p, f)


@command("home.decreaseIndent", "Decrease Indent", menu=("Home", "Alignment"), params="{range?}")
def dec_indent(s, p):
    return apply_style(s, p, lambda st: _align(st, indent=max(st.align.indent - 1, 0)))


@command("home.orientation", "Orientation", menu=("Home", "Alignment"), params="{rotation: -90..90 | 255 (vertical text), range?}")
def orientation(s, p):
    rot = p_int(p, "rotation", 0)
    if rot != 255 and not -90 <= rot <= 90:
        raise EngineError("The rotation must be between -90 and 90 degrees.")
    return apply_style(s, p, lambda st: _align(st, rotation=rot))


# ---------------------------------------------------------------- merge


def _merge(s: Session, p: Dict[str, Any], across: bool, center: bool):
    d = s.doc()
    si = param_sheet(s, p)
    sh = d.wb.sheet(si)
    check_format_allowed(sh)
    ranges = _ranges(s, p)
    if center and not across and len(ranges) == 1 and sh.merge_at(ranges[0].start) == ranges[0] and not ranges[0].is_single():
        # Merge & Center on a merged cell unmerges it.
        return unmerge(s, p)
    targets: List[RangeRef] = []
    for rng in ranges:
        if rng.is_full_cols() or rng.is_full_rows():
            raise EngineError("That selection is too large to merge.")
        if across:
            targets.extend(RangeRef(r, rng.c0, r, rng.c1) for r in range(rng.r0, rng.r1 + 1))
        else:
            targets.append(rng)
    targets = [t for t in targets if not t.is_single()]
    if not targets:
        return None
    for t in targets:
        if any(tb.range.intersects(t) for tb in sh.tables):
            raise EngineError("You can't merge cells in a table.")
    lost = False
    for t in targets:
        n = 0
        for rc, cell in sh.cells.iter_range(t):
            if cell.value is not None or cell.formula is not None:
                n += 1
        if n > 1 or (n == 1 and (sh.cell(t.start) is None or (sh.cell(t.start).value is None and sh.cell(t.start).formula is None))):
            lost = True
    if lost and not p_bool(p, "force", True):
        raise EngineError("Merging cells only keeps the upper-left value and discards other values.")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        wb = ctx.wb
        for t in targets:
            # The top-left value is kept (the first non-empty value moves there, like Excel's
            # merge of a range whose top-left is blank keeps the upper-left-most value).
            first = None
            for rc in sorted(rc for rc, _ in sm.cells.iter_range(t)):
                c = sm.cell(rc)
                if c is not None and (c.value is not None or c.formula is not None):
                    first = (rc, c)
                    break
            top_style = sm.style_id(t.start)
            for rc, cell in list(sm.cells.iter_range(t)):
                if rc != t.start and (cell.value is not None or cell.formula is not None):
                    sm.set_cell(rc, Cell(None, None, cell.style))
                    ctx.touch(si, rc)
            if first is not None and first[0] != t.start:
                sm.set_cell(t.start, Cell(first[1].value, first[1].formula, top_style))
                ctx.touch(si, t.start)
            sm.merges = [m for m in sm.merges if not m.intersects(t)] + [t]
            st = wb.styles.get(top_style)
            if center:
                st = _align(st, h="center")
            sid = wb.styles.intern(st)
            # Every cell of the merge carries the top-left format (borders draw around it).
            for rc in t.cells():
                c = sm.cell(rc)
                if c is None:
                    if sid != 0:
                        sm.set_cell(rc, Cell(None, None, sid))
                elif c.style != sid:
                    sm.set_cell(rc, c.with_style(sid))
        ctx.sel = Selection(ctx.sel.active, ctx.sel.active, list(ctx.sel.ranges))
        ctx.sel.expand_merges(sm)
        if not any(r.contains(ctx.sel.active) for r in ctx.sel.ranges):
            ctx.sel.active = ctx.sel.ranges[-1].start
        m = sm.merge_at(ctx.sel.active)
        if m is not None:
            ctx.sel.active = m.start

    edit(s, f)
    return {"merged": [t.a1() for t in targets]}


@command("home.mergeCenter", "Merge & Center", menu=("Home", "Alignment"), params="{range?}")
def merge_center(s, p):
    return _merge(s, p, False, True)


@command("home.mergeAcross", "Merge Across", menu=("Home", "Alignment", "Merge & Center"), params="{range?}")
def merge_across(s, p):
    return _merge(s, p, True, False)


@command("home.mergeCells", "Merge Cells", menu=("Home", "Alignment", "Merge & Center"), params="{range?}")
def merge_cells(s, p):
    return _merge(s, p, False, False)


@command("home.unmergeCells", "Unmerge Cells", menu=("Home", "Alignment", "Merge & Center"), params="{range?}")
def unmerge(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    check_format_allowed(d.wb.sheet(si))
    ranges = _ranges(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.merges = [m for m in sm.merges if not any(m.intersects(r) for r in ranges)]

    edit(s, f)
    return None


# ---------------------------------------------------------------- number formats


def adjust_decimals(code: str, value, delta: int) -> str:
    """Adds or removes one decimal place in a number format (Increase/Decrease Decimal)."""
    if code == "General":
        text = number_to_text(value) if isinstance(value, float) else ""
        dec = len(text.split(".", 1)[1]) if "." in text and "E" not in text else 0
        n = max(dec + delta, 0)
        return "0" if n == 0 else "0." + "0" * n
    out = []
    for sec in code.split(";"):
        chars = list(sec)
        in_quote = in_bracket = False
        last_digit = None
        point = None
        solid = False  # the section has 0 or # placeholders (not only ? alignment)
        i = 0
        while i < len(chars):
            c = chars[i]
            if c == '"':
                in_quote = not in_quote
            elif c == "[" and not in_quote:
                in_bracket = True
            elif c == "]" and not in_quote:
                in_bracket = False
                i += 1
                continue
            elif c in "\\_*" and not in_quote:
                # An escaped character, a padding width or a repeat-fill character follows.
                i += 2
                continue
            if in_quote or in_bracket:
                i += 1
                continue
            if c in "0#?":
                last_digit = i
                solid = solid or c in "0#"
            if c == "." and point is None:
                point = i
            if c in "Ee":
                break
            i += 1
        if last_digit is None:
            out.append(sec)
            continue
        if not solid:
            # Accounting zero sections ("-"??) align with ? placeholders: one per decimal.
            if delta > 0:
                chars.insert(last_digit + 1, "?")
            else:
                del chars[last_digit]
        elif delta > 0:
            if point is not None:
                chars.insert(last_digit + 1, "0")
            else:
                chars.insert(last_digit + 1, "0")
                chars.insert(last_digit + 1, ".")
        elif point is not None and last_digit > point:
            del chars[last_digit]
            if last_digit == point + 1:
                del chars[point]
        out.append("".join(chars))
    return ";".join(out)


def check_format(code: str) -> None:
    """Refuses a code Excel's Format Cells refuses (``[0;0``, 256 characters, ``0.00e+00``…)."""
    if not valid_format(code):
        raise EngineError("ReflexGrid can't use the number format you typed. Try one of the built-in number formats.")


def set_format(s: Session, p: Dict[str, Any], code: str):
    check_format(code)
    return apply_style(s, p, lambda st: replace(st, num_fmt=code), widen=True)


@command("home.numberFormat", "Number Format", menu=("Home", "Number"), params='{format: general|number|currency|accounting|shortDate|longDate|time|percent|fraction|scientific|text, or code: "0.00", range?}')
def number_format(s, p):
    code = p.get("code")
    if code is None:
        name = p_str(p, "format")
        if name is None:
            raise EngineError("home.numberFormat: missing `format` or `code`")
        code = FORMAT_PRESETS.get(name, name)
    return set_format(s, p, str(code))


@command("home.accounting", "Accounting Number Format", menu=("Home", "Number"), shortcut="Ctrl+Shift+$", params="{range?}")
def accounting(s, p):
    return set_format(s, p, ACCOUNTING)


@command("home.percent", "Percent Style", menu=("Home", "Number"), shortcut="Ctrl+Shift+%", params="{range?}")
def percent(s, p):
    return set_format(s, p, "0%")


@command("home.comma", "Comma Style", menu=("Home", "Number"), params="{range?}")
def comma(s, p):
    return set_format(s, p, COMMA)


def _decimals(s: Session, p: Dict[str, Any], delta: int):
    d = s.doc()
    sh = d.sheet
    a = d.selection.active
    code = d.wb.styles.get(sh.style_id(a)).num_fmt
    return set_format(s, p, adjust_decimals(code, sh.value(a), delta))


@command("home.increaseDecimal", "Increase Decimal", menu=("Home", "Number"), params="{range?}")
def inc_decimal(s, p):
    return _decimals(s, p, 1)


@command("home.decreaseDecimal", "Decrease Decimal", menu=("Home", "Number"), params="{range?}")
def dec_decimal(s, p):
    return _decimals(s, p, -1)


# ---------------------------------------------------------------- Format Cells


def _border_line(v: Any) -> BorderLine:
    if not v:
        return BorderLine()
    if isinstance(v, str):
        return BorderLine(v)
    return BorderLine(str(v.get("style", "thin")), parse_color(v.get("color")))


def style_from_patch(st: Style, patch: Dict[str, Any]) -> Style:
    """Applies a Format Cells patch (only the keys given change)."""
    if "numFmt" in patch and patch["numFmt"] is not None:
        st = replace(st, num_fmt=str(patch["numFmt"]))
    f = patch.get("font")
    if isinstance(f, dict):
        kw: Dict[str, Any] = {}
        for k in ("name", "bold", "italic", "strike"):
            if k in f:
                kw[k] = f[k] if k == "name" else bool(f[k])
        if "size" in f:
            kw["size"] = float(f["size"])
        if "underline" in f:
            u = f["underline"]
            kw["underline"] = "single" if u is True else ("none" if not u else str(u))
        if "color" in f:
            kw["color"] = parse_color(f["color"])
        if "vert" in f:
            kw["vert"] = str(f["vert"])
        st = replace(st, font=replace(st.font, **kw))
    fl = patch.get("fill")
    if isinstance(fl, dict):
        if "color" in fl:
            c = parse_color(fl["color"])
            st = replace(st, fill=Fill() if c.is_auto else Fill.solid(c))
        else:
            st = replace(st, fill=Fill(str(fl.get("pattern", "solid")), parse_color(fl.get("fg")), parse_color(fl.get("bg"))))
    elif fl is None and "fill" in patch:
        st = replace(st, fill=Fill())
    b = patch.get("borders")
    if isinstance(b, dict):
        kw = {}
        for k, attr in (("left", "left"), ("right", "right"), ("top", "top"), ("bottom", "bottom"), ("diagDown", "diag_down"), ("diagUp", "diag_up")):
            if k in b:
                kw[attr] = _border_line(b[k])
        st = replace(st, borders=replace(st.borders, **kw))
    a = patch.get("align")
    if isinstance(a, dict):
        kw = {}
        if "h" in a:
            kw["h"] = str(a["h"])
        if "v" in a:
            kw["v"] = str(a["v"])
        for k in ("wrap", "shrink"):
            if k in a:
                kw[k] = bool(a[k])
        if "indent" in a:
            kw["indent"] = max(0, min(int(a["indent"]), 250))
        if "rotation" in a:
            kw["rotation"] = int(a["rotation"])
        st = replace(st, align=replace(st.align, **kw))
    pr = patch.get("protection")
    if isinstance(pr, dict):
        st = replace(st, protection=replace(st.protection, **{k: bool(v) for k, v in pr.items() if k in ("locked", "hidden")}))
    return st


@command("home.formatCells", "Format Cells…", menu=("Home", "Cells", "Format"), shortcut="Ctrl+1",
         params='{numFmt?, font?: {name,size,bold,italic,underline,strike,color,vert}, fill?: {color} | {pattern,fg,bg}, borders?: {left,right,top,bottom,diagDown,diagUp: {style,color}}, align?: {h,v,wrap,shrink,indent,rotation}, protection?: {locked,hidden}, range?}')
def format_cells(s, p):
    patch = {k: p[k] for k in ("numFmt", "font", "fill", "borders", "align", "protection") if k in p}
    if patch.get("numFmt") is not None:
        check_format(str(patch["numFmt"]))
    d = s.doc()
    sh = d.sheet
    if sh.protection is not None and set(patch) == {"protection"}:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")
    return apply_style(s, p, lambda st: style_from_patch(st, patch))


@command("home.lockCell", "Lock Cell", menu=("Home", "Cells", "Format"), params="{on?: bool, range?}")
def lock_cell(s, p):
    return _toggle(s, p, lambda st: st.protection.locked, lambda st, on: replace(st, protection=replace(st.protection, locked=on)))


# ---------------------------------------------------------------- cell styles


def builtin_cell_style(name: str, theme: Theme) -> Optional[Style]:
    """The Cell Styles gallery (our own palette)."""
    s = Style()

    def solid(rgb: int) -> Fill:
        return Fill.solid(Color.rgb(rgb))

    def box(style: str, rgb: int) -> Borders:
        l = BorderLine(style, Color.rgb(rgb))
        return Borders(left=l, right=l, top=l, bottom=l)

    dark2 = Color.theme(3)
    if name == "Normal":
        return s
    table = {
        "Good": lambda: replace(s, fill=solid(0xC6EFCE), font=replace(s.font, color=Color.rgb(0x006100))),
        "Bad": lambda: replace(s, fill=solid(0xFFC7CE), font=replace(s.font, color=Color.rgb(0x9C0006))),
        "Neutral": lambda: replace(s, fill=solid(0xFFEB9C), font=replace(s.font, color=Color.rgb(0x9C5700))),
        "Calculation": lambda: replace(s, fill=solid(0xF2F2F2), font=replace(s.font, color=Color.rgb(0xFA7D00), bold=True), borders=box("thin", 0x7F7F7F)),
        "Check Cell": lambda: replace(s, fill=solid(0xA5A5A5), font=replace(s.font, color=Color.rgb(0xFFFFFF), bold=True), borders=box("double", 0x3F3F3F)),
        "Explanatory Text": lambda: replace(s, font=replace(s.font, italic=True, color=Color.rgb(0x7F7F7F))),
        "Input": lambda: replace(s, fill=solid(0xFFCC99), font=replace(s.font, color=Color.rgb(0x3F3F76)), borders=box("thin", 0x7F7F7F)),
        "Linked Cell": lambda: replace(s, font=replace(s.font, color=Color.rgb(0xFA7D00)), borders=Borders(bottom=BorderLine("double", Color.rgb(0xFF8001)))),
        "Note": lambda: replace(s, fill=solid(0xFFFFCC), borders=box("thin", 0xB2B2B2)),
        "Output": lambda: replace(s, fill=solid(0xF2F2F2), font=replace(s.font, color=Color.rgb(0x3F3F3F), bold=True), borders=box("thin", 0x3F3F3F)),
        "Warning Text": lambda: replace(s, font=replace(s.font, color=Color.rgb(0xFF0000))),
        "Title": lambda: replace(s, font=replace(s.font, size=18.0, color=dark2, name=theme.major_font)),
        "Heading 1": lambda: replace(s, font=replace(s.font, size=15.0, bold=True, color=dark2), borders=Borders(bottom=BorderLine("thick", Color.theme(4)))),
        "Heading 2": lambda: replace(s, font=replace(s.font, size=13.0, bold=True, color=dark2), borders=Borders(bottom=BorderLine("thick", Color.theme(4, 500)))),
        "Heading 3": lambda: replace(s, font=replace(s.font, bold=True, color=dark2), borders=Borders(bottom=BorderLine("medium", Color.theme(4, 400)))),
        "Heading 4": lambda: replace(s, font=replace(s.font, bold=True, color=dark2)),
        "Total": lambda: replace(s, font=replace(s.font, bold=True), borders=Borders(top=BorderLine("thin", Color.theme(4)), bottom=BorderLine("double", Color.theme(4)))),
        "Currency": lambda: replace(s, num_fmt=ACCOUNTING),
        "Comma": lambda: replace(s, num_fmt=COMMA),
        "Percent": lambda: replace(s, num_fmt="0%"),
        "Currency [0]": lambda: replace(s, num_fmt='_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)'),
        "Comma [0]": lambda: replace(s, num_fmt='_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'),
    }
    if name in table:
        return table[name]()
    pct = None
    rest = name
    if "% - " in name:
        a, rest = name.split("% - ", 1)
        try:
            pct = int(a)
        except ValueError:
            return None
    if not rest.startswith("Accent"):
        return None
    try:
        i = int(rest[6:])
    except ValueError:
        return None
    if not 1 <= i <= 6:
        return None
    if pct is None:
        return replace(s, fill=Fill.solid(Color.theme(3 + i)), font=replace(s.font, color=Color.theme(0)))
    tint = max(0, min(1000 - pct * 10, 1000))
    return replace(s, fill=Fill.solid(Color.theme(3 + i, tint)), font=replace(s.font, color=Color.theme(0) if pct >= 60 else Color.theme(1)))


CELL_STYLE_NAMES = [
    "Normal", "Bad", "Good", "Neutral", "Calculation", "Check Cell", "Explanatory Text", "Input", "Linked Cell",
    "Note", "Output", "Warning Text", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Title", "Total",
] + [f"{p}% - Accent{i}" for p in (20, 40, 60) for i in range(1, 7)] + [f"Accent{i}" for i in range(1, 7)] + [
    "Comma", "Comma [0]", "Currency", "Currency [0]", "Percent",
]


@command("home.cellStyle", "Cell Styles", menu=("Home", "Styles"), params='{name: "Good" | "Heading 1" | "40% - Accent1" | …, range?}')
def cell_style(s, p):
    name = p_str(p, "name", "Normal")
    d = s.doc()
    st = None
    for n, custom in d.wb.cell_styles:
        if n == name and isinstance(custom, Style):
            st = custom
    if st is None:
        st = builtin_cell_style(name, d.wb.theme)
    if st is None:
        raise EngineError(f"unknown cell style `{name}`")
    num_only = name in ("Currency", "Comma", "Percent", "Currency [0]", "Comma [0]")

    def f(old: Style) -> Style:
        if num_only:
            return replace(old, num_fmt=st.num_fmt)
        return replace(st, protection=old.protection)

    return apply_style(s, p, f)


# ---------------------------------------------------------------- rows & columns


def _lines(s: Session, p: Dict[str, Any], axis: str) -> List[int]:
    t = p.get("rows" if axis == "rows" else "cols")
    out: List[int] = []
    if t is not None:
        if isinstance(t, list):
            return [int(x) for x in t]
        for part in str(t).split(","):
            r = RangeRef.parse(part if ":" in part else (f"{part}:{part}"))
            if r is None:
                raise EngineError(f"not a {axis[:-1]} reference: `{part}`")
            if axis == "rows":
                out.extend(range(r.r0, r.r1 + 1))
            else:
                out.extend(range(r.c0, r.c1 + 1))
        return out
    for r in _ranges(s, p):
        if axis == "rows":
            out.extend(range(r.r0, min(r.r1, r.r0 + 1_048_575) + 1))
        else:
            out.extend(range(r.c0, r.c1 + 1))
    return sorted(set(out))


def check_rows_cols_allowed(sh, axis: str) -> None:
    if sh.protection is None:
        return
    ok = sh.protection.format_rows if axis == "rows" else sh.protection.format_columns
    if not ok:
        raise EngineError("The cell or chart you're trying to change is on a protected sheet.")


@command("home.rowHeight", "Row Height…", menu=("Home", "Cells", "Format"), params="{height: px, rows?: \"2:5\"}")
def row_height(s, p):
    h = p_float(p, "height")
    if h is None or not 0 <= h <= 546:
        raise EngineError("Row height must be between 0 and 546 pixels.")
    d = s.doc()
    si = param_sheet(s, p)
    check_rows_cols_allowed(d.wb.sheet(si), "rows")
    rows = _lines(s, p, "rows")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for r in rows:
            info = sm.row_info(r)
            if h == 0:
                sm.set_row(r, replace(info, hidden=True))
            else:
                sm.set_row(r, replace(info, size=float(h), hidden=False, custom=True))
            ctx.rows_shown = ctx.rows_shown or info.hidden != (h == 0)

    edit(s, f)
    return None


@command("home.columnWidth", "Column Width…", menu=("Home", "Cells", "Format"), params="{width: px, cols?: \"B:D\"}")
def column_width(s, p):
    w = p_float(p, "width")
    if w is None or not 0 <= w <= 1785:
        raise EngineError("Column width must be between 0 and 1785 pixels.")
    d = s.doc()
    si = param_sheet(s, p)
    check_rows_cols_allowed(d.wb.sheet(si), "cols")
    cols = _lines(s, p, "cols")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for c in cols:
            info = sm.col_info(c)
            if w == 0:
                sm.set_col(c, replace(info, hidden=True))
            else:
                sm.set_col(c, replace(info, size=float(w), hidden=False, custom=True))

    edit(s, f)
    return None


@command("home.autofitRowHeight", "AutoFit Row Height", menu=("Home", "Cells", "Format"), params="{rows?}")
def autofit_rows(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    check_rows_cols_allowed(d.wb.sheet(si), "rows")
    rows = _lines(s, p, "rows")
    used = d.wb.sheet(si).used_range()
    if used is not None:
        rows = [r for r in rows if r <= used.r1 or r in d.wb.sheet(si).rows]

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for r in rows:
            info = sm.row_info(r)
            # AutoFit shows a hidden row too, an empty one as well (Excel; an empty hidden column
            # stays hidden, see autofit_cols).
            if info.custom or info.size is not None or info.hidden:
                sm.set_row(r, replace(info, custom=False, size=None, hidden=False))
                ctx.rows_shown = ctx.rows_shown or info.hidden
            ctx.fit_rows.append((si, r))

    edit(s, f)
    return None


@command("home.autofitColumnWidth", "AutoFit Column Width", menu=("Home", "Cells", "Format"), params="{cols?}")
def autofit_cols(s, p):
    d = s.doc()
    si = param_sheet(s, p)
    check_rows_cols_allowed(d.wb.sheet(si), "cols")
    cols = _lines(s, p, "cols")
    used = d.wb.sheet(si).used_range()
    if used is not None:
        cols = [c for c in cols if c <= used.c1]
    else:
        cols = []

    def f(ctx):
        sm = ctx.sheet_mut(si)
        for c in cols:
            w = autofit_width(ctx.wb, sm, c)
            info = sm.col_info(c)
            if w > 0:
                sm.set_col(c, replace(info, size=round(w), hidden=False, custom=False))
            elif info.size is not None:
                sm.set_col(c, replace(info, size=None, custom=False))

    edit(s, f)
    return None


@command("home.defaultWidth", "Default Width…", menu=("Home", "Cells", "Format"), params="{width: px}")
def default_width(s, p):
    w = p_float(p, "width")
    if w is None or not 0 < w <= 1785:
        raise EngineError("Column width must be between 0 and 1785 pixels.")
    si = param_sheet(s, p)

    def f(ctx):
        sm = ctx.sheet_mut(si)
        sm.default_col_width = float(w)
        sm._invalidate_geometry()

    edit(s, f)
    return None


def _hide(s: Session, p: Dict[str, Any], axis: str, hidden: bool):
    d = s.doc()
    si = param_sheet(s, p)
    check_rows_cols_allowed(d.wb.sheet(si), axis)
    lines = _lines(s, p, axis)
    if hidden and axis == "cols" and len(lines) >= MAX_COLS:
        raise EngineError("You can't hide every column.")
    if hidden and axis == "rows" and len(lines) >= MAX_ROWS:
        raise EngineError("You can't hide every row.")

    def f(ctx):
        sm = ctx.sheet_mut(si)
        if axis == "rows":
            if not hidden:
                # Unhide every hidden row inside the span, not only the selected ones.
                lo, hi = min(lines), max(lines)
                lines2 = [r for r in sm.rows if lo <= r <= hi and sm.rows[r].hidden]
            else:
                lines2 = lines
            for r in lines2:
                info = sm.row_info(r)
                size = info.size
                if not hidden and size == 0:
                    size = None
                sm.set_row(r, replace(info, hidden=hidden, size=size))
            ctx.rows_shown = bool(lines2)
        else:
            if not hidden:
                lo, hi = min(lines), max(lines)
                lines2 = [c for c in sm.cols if lo <= c <= hi and sm.cols[c].hidden]
            else:
                lines2 = lines
            for c in lines2:
                info = sm.col_info(c)
                size = info.size
                if not hidden and size == 0:
                    size = None
                sm.set_col(c, replace(info, hidden=hidden, size=size))

    edit(s, f)
    return None


@command("home.hideRows", "Hide Rows", menu=("Home", "Cells", "Format", "Hide & Unhide"), shortcut="Ctrl+9", params="{rows?}")
def hide_rows(s, p):
    return _hide(s, p, "rows", True)


@command("home.unhideRows", "Unhide Rows", menu=("Home", "Cells", "Format", "Hide & Unhide"), shortcut="Ctrl+Shift+9", params="{rows?}")
def unhide_rows(s, p):
    return _hide(s, p, "rows", False)


@command("home.hideColumns", "Hide Columns", menu=("Home", "Cells", "Format", "Hide & Unhide"), shortcut="Ctrl+0", params="{cols?}")
def hide_cols(s, p):
    return _hide(s, p, "cols", True)


@command("home.unhideColumns", "Unhide Columns", menu=("Home", "Cells", "Format", "Hide & Unhide"), params="{cols?}")
def unhide_cols(s, p):
    return _hide(s, p, "cols", False)
