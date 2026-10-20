"""Format code parsing: sections, bracket tags and tokens.

Tokens are tuples whose first element is the kind:

* ``("lit", text)`` literal text
* ``("digit", "0" | "#" | "?")`` digit placeholder
* ``("point",)`` the decimal point (the first one in a number section)
* ``("comma",)`` an unresolved comma (resolved into thousands/scale markers or a literal)
* ``("percent",)``
* ``("exp", upper, plus)`` the ``E+``/``e-`` exponent marker
* ``("slash",)`` the fraction bar, ``("fixed", "8")`` a fixed denominator
* ``("fill", ch)`` the ``*x`` repeat fill, ``("at",)`` the text placeholder, ``("general",)``
* ``("date", part, n)`` with part one of ``y mo mi d h s sub eh em es ampm ap``
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple

MAX_SECTIONS = 4

COLORS = {"black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"}


def is_currency_char(c: str) -> bool:
    return c in "$€£¥₹₩₽¢₪₫฿"


@dataclass
class Cond:
    op: str  # = <> < <= > >=
    value: float

    def test(self, v: float) -> bool:
        op = self.op
        if op == "=":
            return v == self.value
        if op == "<>":
            return v != self.value
        if op == "<":
            return v < self.value
        if op == "<=":
            return v <= self.value
        if op == ">":
            return v > self.value
        return v >= self.value


@dataclass
class Section:
    toks: List[tuple] = field(default_factory=list)
    color: Optional[str] = None  # colour name or "color<N>"
    cond: Optional[Cond] = None
    kind: str = "number"  # number | date | text | general
    thousands: bool = False
    scale: int = 0
    percent: int = 0
    has_at: bool = False
    has_digits: bool = False
    exp: bool = False
    fraction: bool = False
    currency: bool = False
    fill: bool = False


def literal_text(t: tuple) -> str:
    k = t[0]
    if k == "lit" or k == "fixed":
        return t[1]
    if k == "digit":
        return t[1]
    if k == "point":
        return "."
    if k == "comma":
        return ","
    if k == "percent":
        return "%"
    if k == "exp":
        return ("E" if t[1] else "e") + ("+" if t[2] else "-")
    if k == "slash":
        return "/"
    return ""


def split_sections(code: str) -> List[str]:
    """Splits a code into sections at unquoted, unescaped, unbracketed semicolons."""
    out = [""]
    quoted = bracket = escape = False
    for c in code:
        if escape:
            out[-1] += c
            escape = False
            continue
        if c == '"':
            quoted = not quoted
            out[-1] += c
        elif c in "\\_*" and not quoted and not bracket:
            escape = True
            out[-1] += c
        elif c == "[" and not quoted:
            bracket = True
            out[-1] += c
        elif c == "]" and not quoted:
            bracket = False
            out[-1] += c
        elif c == ";" and not quoted and not bracket:
            if len(out) >= MAX_SECTIONS:
                break
            out.append("")
        else:
            out[-1] += c
    return out


def _push_lit(toks: List[tuple], s: str) -> None:
    if not s:
        return
    if toks and toks[-1][0] == "lit":
        toks[-1] = ("lit", toks[-1][1] + s)
    else:
        toks.append(("lit", s))


def _parse_color(lc: str) -> Optional[str]:
    if lc in COLORS:
        return lc
    if lc.startswith("color"):
        try:
            n = int(lc[5:].strip())
        except ValueError:
            return None
        if 1 <= n <= 56:
            return f"color{n}"
    return None


def _parse_cond(s: str) -> Optional[Cond]:
    s = s.strip()
    for op in ("<=", ">=", "<>", "<", ">", "="):
        if s.startswith(op):
            rest = s[len(op):]
            break
    else:
        return None
    try:
        v = float(rest.strip())
    except ValueError:
        return None
    if v != v or v in (float("inf"), float("-inf")):
        return None
    return Cond(op, v)


def _starts_with_ci(cs: str, i: int, pat: str) -> bool:
    return cs[i:i + len(pat)].lower() == pat


# ---------------------------------------------------------------- what Excel accepts

MAX_CODE_LEN = 255


def _bracket_kind(content: str) -> Optional[str]:
    """What a ``[…]`` tag is to Excel: colour, condition, currency/locale, DBNum, elapsed time, or
    None when Excel refuses it ([xyz], [0], [Purple], [Color57], [ENG], [d], [])."""
    lc = content.lower()
    if lc in COLORS or (lc.startswith("color") and lc[5:].strip().isdigit() and 1 <= int(lc[5:].strip()) <= 56):
        return "color"
    if lc[:1] in ("<", ">", "="):
        return "cond" if _parse_cond(lc) is not None else None
    if lc.startswith("$"):
        return "locale"
    if lc in ("dbnum1", "dbnum2", "dbnum3", "dbnum4"):
        return "dbnum"
    if lc and lc[0] in "hms" and all(ch == lc[0] for ch in lc):
        return "elapsed"
    return None


def _section_ok(toks: List[tuple]) -> bool:
    kinds = [t[0] for t in toks]
    if kinds.count("color") > 1 or kinds.count("elapsed") > 1 or kinds.count("exp") > 1:
        return False
    digits = [i for i, k in enumerate(kinds) if k == "digit"]
    is_date = "date" in kinds or "elapsed" in kinds
    if "exp" in kinds:
        e = kinds.index("exp")
        if "digit" not in kinds[:e] or "digit" not in kinds[e + 1:]:
            return False
    if "at" in kinds and (digits or is_date or "exp" in kinds):
        return False
    if "general" in kinds and digits:
        return False
    if is_date:
        if "exp" in kinds or "percent" in kinds:
            return False
        for i in digits:
            # Only 0s straight after a point: the fraction of a second (ss.00, hh.0).
            j = i - 1
            while j >= 0 and toks[j] == ("digit", "0"):
                j -= 1
            if toks[i][1] != "0" or j < 0 or kinds[j] != "point":
                return False
    return True


def validate(code: str) -> bool:
    """Whether Excel takes ``code`` as a number format (Format Cells, TEXT), as Excel 16 for Mac
    answers: at most 255 characters and 4 sections; every ``[`` closed and a tag it knows; nothing
    left dangling after ``\\ _ *`` or an open quote; a lowercase ``e``, or an ``E`` without its
    sign, is the era year, a date code; digit placeholders don't mix with dates (but for 0s after a
    point, the fraction of a second), nor with ``@`` or ``General``; one exponent, with digits on
    both sides. (Displaying a cell never fails: :func:`tokenize` reads any code.)"""
    if len(code) > MAX_CODE_LEN:
        return False
    sections: List[List[tuple]] = []
    toks: List[tuple] = []
    i, n = 0, len(code)
    while i < n:
        c = code[i]
        if c == '"':
            j = code.find('"', i + 1)
            if j < 0:
                return False
            toks.append(("lit",))
            i = j + 1
            continue
        if c in "\\_*":
            if i + 1 >= n:
                return False
            toks.append(("lit",))
            i += 2
            continue
        if c == "[":
            j = code.find("]", i + 1)
            if j < 0:
                return False
            kind = _bracket_kind(code[i + 1:j])
            if kind is None:
                return False
            toks.append((kind,))
            i = j + 1
            continue
        if c == ";":
            sections.append(toks)
            toks = []
            if len(sections) >= MAX_SECTIONS:
                return False
            i += 1
            continue
        if c in "0#?":
            toks.append(("digit", c))
        elif c == ".":
            toks.append(("point",))
        elif c == "%":
            toks.append(("percent",))
        elif c == "@":
            toks.append(("at",))
        elif c == "E" and i + 1 < n and code[i + 1] in "+-":
            toks.append(("exp",))
            i += 2
            continue
        elif c in "gG" and _starts_with_ci(code, i, "general"):
            toks.append(("general",))
            i += 7
            continue
        elif c in "aA" and (_starts_with_ci(code, i, "am/pm") or _starts_with_ci(code, i, "a/p")):
            toks.append(("lit",))  # AM/PM goes with numbers too ("0 AM/PM")
            i += 5 if _starts_with_ci(code, i, "am/pm") else 3
            continue
        elif c in "yYmMdDhHsSeE":
            toks.append(("date",))
        else:
            toks.append(("lit",))
        i += 1
    sections.append(toks)
    return all(_section_ok(t) for t in sections)


class _Raw:
    __slots__ = ("toks", "color", "cond", "currency")

    def __init__(self) -> None:
        self.toks: List[tuple] = []
        self.color: Optional[str] = None
        self.cond: Optional[Cond] = None
        self.currency = False


def _bracket(raw: _Raw, content: str) -> None:
    lc = content.strip().lower()
    color = _parse_color(lc)
    if color is not None:
        raw.color = color
        return
    if lc[:1] in ("<", ">", "="):
        c = _parse_cond(lc)
        if c is not None:
            raw.cond = c
        return
    if content.startswith("$"):
        sym = content[1:].split("-", 1)[0]
        if sym:
            raw.currency = True
            _push_lit(raw.toks, sym)
        return
    if lc and lc[0] in "hms" and all(ch == lc[0] for ch in lc):
        n = min(len(lc), 255)
        raw.toks.append(("date", {"h": "eh", "m": "em", "s": "es"}[lc[0]], n))
    # Anything else ([DBNum1], [natural], locale tags…) is ignored.


def tokenize(src: str) -> _Raw:
    raw = _Raw()
    toks = raw.toks
    cs = src
    n = len(cs)
    i = 0
    while i < n:
        c = cs[i]
        if c == '"':
            j = i + 1
            s = []
            while j < n:
                d = cs[j]
                j += 1
                if d == '"':
                    break
                s.append(d)
            text = "".join(s)
            if any(is_currency_char(ch) for ch in text):
                raw.currency = True
            _push_lit(toks, text)
            i = j
            continue
        if c == "\\":
            if i + 1 < n:
                d = cs[i + 1]
                if is_currency_char(d):
                    raw.currency = True
                _push_lit(toks, d)
            i += 2
            continue
        if c == "_":
            if i + 1 < n:
                _push_lit(toks, " ")
            i += 2
            continue
        if c == "*":
            if i + 1 < n:
                toks.append(("fill", cs[i + 1]))
            i += 2
            continue
        if c == "[":
            close = cs.find("]", i + 1)
            if close >= 0:
                _bracket(raw, cs[i + 1:close])
                i = close + 1
            else:
                _push_lit(toks, "[")
                i += 1
            continue
        if c in "0#?":
            toks.append(("digit", c))
        elif c == ".":
            toks.append(("point",))
        elif c == ",":
            toks.append(("comma",))
        elif c == "%":
            toks.append(("percent",))
        elif c == "@":
            toks.append(("at",))
        elif c == "/":
            toks.append(("slash",))
            i += 1
            if i < n and cs[i] in "123456789":
                j = i
                while j < n and cs[j].isdigit():
                    j += 1
                toks.append(("fixed", cs[i:j]))
                i = j
            continue
        elif c in "eE" and i + 1 < n and cs[i + 1] in "+-":
            toks.append(("exp", c == "E", cs[i + 1] == "+"))
            i += 2
            continue
        elif c in "gG" and _starts_with_ci(cs, i, "general"):
            toks.append(("general",))
            i += 7
            continue
        elif c in "aA" and _starts_with_ci(cs, i, "am/pm"):
            toks.append(("date", "ampm", 0))
            i += 5
            continue
        elif c in "aA" and _starts_with_ci(cs, i, "a/p"):
            p = cs[i + 2] if i + 2 < n else "P"
            toks.append(("date", "ap", (c, p)))
            i += 3
            continue
        elif c in "yYmMdDhHsSeE":
            lc = c.lower()
            k = 0
            while i + k < n and cs[i + k].lower() == lc:
                k += 1
            nn = min(k, 255)
            if lc == "y":
                toks.append(("date", "y", nn))
            elif lc == "e":
                toks.append(("date", "y", 4))
            elif lc == "m":
                toks.append(("date", "mo", nn))
            elif lc == "d":
                toks.append(("date", "d", nn))
            elif lc == "h":
                toks.append(("date", "h", nn))
            else:
                toks.append(("date", "s", nn))
            i += max(k, 1)
            continue
        else:
            if is_currency_char(c):
                raw.currency = True
            _push_lit(toks, c)
        i += 1
    return raw


def _date_part(t: tuple) -> Optional[str]:
    return t[1] if t[0] == "date" else None


def is_calendar(part: str) -> bool:
    return part in ("y", "mo", "d")


def analyze(raw: _Raw) -> Section:
    toks = list(raw.toks)
    is_date = any(t[0] == "date" for t in toks)
    has_general = any(t[0] == "general" for t in toks)
    sec = Section(
        color=raw.color,
        cond=raw.cond,
        has_at=any(t[0] == "at" for t in toks),
        currency=raw.currency,
        fill=any(t[0] == "fill" for t in toks),
    )
    if is_date:
        sec.kind = "date"
        # m/mm after hours or before seconds means minutes.
        for i, t in enumerate(toks):
            if t[0] != "date" or t[1] != "mo" or t[2] > 2:
                continue
            prev = next((_date_part(x) for x in reversed(toks[:i]) if x[0] == "date"), None)
            nxt = next((_date_part(x) for x in toks[i + 1:] if x[0] == "date"), None)
            if prev in ("h", "eh") or nxt in ("s", "es"):
                toks[i] = ("date", "mi", t[2])
        out: List[tuple] = []
        i = 0
        while i < len(toks):
            t = toks[i]
            if t[0] == "point" and i + 1 < len(toks) and toks[i + 1] == ("digit", "0"):
                k = 0
                while i + 1 + k < len(toks) and toks[i + 1 + k] == ("digit", "0"):
                    k += 1
                out.append(("date", "sub", min(k, 255)))
                i += 1 + k
                continue
            if t[0] in ("date", "fill"):
                out.append(t)
            else:
                _push_lit(out, literal_text(t))
            i += 1
        sec.toks = out
        return sec
    if has_general:
        sec.kind = "general"
        out = []
        for t in toks:
            if t[0] in ("general", "fill"):
                out.append(t)
            else:
                _push_lit(out, literal_text(t))
        sec.toks = out
        return sec
    sec.has_digits = any(t[0] == "digit" for t in toks)
    if sec.has_at and not sec.has_digits:
        sec.kind = "text"
        out = []
        for t in toks:
            if t[0] in ("at", "fill"):
                out.append(t)
            else:
                _push_lit(out, literal_text(t))
        sec.toks = out
        return sec
    # Number section.
    seen_point = seen_exp = seen_digit = fraction_slash = False
    for i in range(len(toks)):
        t = toks[i]
        k = t[0]
        replace: Optional[tuple] = None
        if k == "digit":
            seen_digit = True
        elif k == "point":
            if seen_point or seen_exp:
                replace = ("lit", ".")
            else:
                seen_point = True
        elif k == "percent":
            sec.percent += 1
        elif k == "exp":
            if seen_exp or not seen_digit:
                replace = ("lit", literal_text(t))
            else:
                seen_exp = True
                sec.exp = True
        elif k == "comma":
            nxt = next((x for x in toks[i + 1:] if x[0] != "comma"), None)
            next_digit = nxt is not None and nxt[0] == "digit"
            if seen_digit and next_digit and not seen_point and not seen_exp:
                sec.thousands = True
                replace = ("lit", "")
            elif seen_digit and not next_digit:
                sec.scale += 1
                replace = ("lit", "")
            else:
                replace = ("lit", ",")
        elif k == "slash":
            # Literal text may sit between the bar and the denominator (`#/ #`, `# ?=/=?`).
            nxt1 = toks[i + 1] if i + 1 < len(toks) else None
            first_non_lit = next((x for x in toks[i + 1:] if x[0] != "lit"), None)
            next_ok = (nxt1 is not None and nxt1[0] == "fixed") or (first_non_lit is not None and first_non_lit[0] == "digit")
            has_exp_tok = any(x[0] == "exp" for x in toks)
            if not fraction_slash and seen_digit and next_ok and not seen_point and not has_exp_tok:
                fraction_slash = True
                sec.fraction = True
            else:
                replace = ("lit", "/")
        elif k == "fixed":
            if i > 0 and toks[i - 1][0] == "slash" and sec.fraction:
                pass
            else:
                replace = ("lit", t[1])
        if replace is not None:
            toks[i] = replace
    if sec.fraction:
        toks = [("lit", ".") if t[0] == "point" else t for t in toks]
    out = []
    for t in toks:
        if t[0] == "lit":
            _push_lit(out, t[1])
        else:
            out.append(t)
    sec.toks = out
    return sec


def parse_sections(code: str) -> List[Section]:
    return [analyze(tokenize(s)) for s in split_sections(code)]
