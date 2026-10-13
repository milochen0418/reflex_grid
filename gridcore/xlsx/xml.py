"""XML helpers: escaping, ECMA-376 ``_xHHHH_`` string escapes, number formatting and a
namespace-agnostic view of ElementTree elements (strict and transitional files both work)."""

from __future__ import annotations

import math
import re
from typing import Iterator, List, Optional
from xml.etree import ElementTree as ET

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_ESC_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def local(tag: str) -> str:
    """Local name of a (possibly namespaced) tag."""
    i = tag.rfind("}")
    return tag[i + 1:] if i >= 0 else tag


def attr(e: ET.Element, name: str) -> Optional[str]:
    """Attribute by local name (``r:id`` and plain ``id`` are both found as ``id``)."""
    v = e.attrib.get(name)
    if v is not None:
        return v
    for k, v in e.attrib.items():
        if local(k) == name:
            return v
    return None


def attr_int(e: ET.Element, name: str) -> Optional[int]:
    v = attr(e, name)
    if v is None:
        return None
    try:
        return int(float(v.strip()))
    except ValueError:
        return None


def attr_float(e: ET.Element, name: str) -> Optional[float]:
    v = attr(e, name)
    if v is None:
        return None
    try:
        f = float(v.strip())
    except ValueError:
        return None
    return f if math.isfinite(f) else None


def flag(e: ET.Element, name: str, default: bool) -> bool:
    v = attr(e, name)
    if v is None:
        return default
    return v.strip().lower() in ("1", "true", "on")


def child(e: ET.Element, name: str) -> Optional[ET.Element]:
    for c in e:
        if local(c.tag) == name:
            return c
    return None


def first_child(e: ET.Element, *names: str) -> Optional[ET.Element]:
    """The first child matching any of ``names`` (in that preference order). Never use ``a or b``
    with elements: an element without children is falsy."""
    for n in names:
        c = child(e, n)
        if c is not None:
            return c
    return None


def kids(e: ET.Element, name: str) -> Iterator[ET.Element]:
    for c in e:
        if local(c.tag) == name:
            yield c


def path(e: ET.Element, *names: str) -> Optional[ET.Element]:
    cur: Optional[ET.Element] = e
    for n in names:
        if cur is None:
            return None
        cur = child(cur, n)
    return cur


def child_val(e: ET.Element, name: str) -> Optional[str]:
    c = child(e, name)
    return attr(c, "val") if c is not None else None


def text_of(e: Optional[ET.Element]) -> str:
    if e is None:
        return ""
    return e.text or ""


def rich_text(e: ET.Element) -> str:
    """Text of a shared-string item or inline string: ``<t>`` plus the ``<t>`` of each run
    (phonetic runs are skipped)."""
    out: List[str] = []
    for c in e:
        n = local(c.tag)
        if n == "t":
            out.append(c.text or "")
        elif n == "r":
            for t in c:
                if local(t.tag) == "t":
                    out.append(t.text or "")
    return decode_xstring("".join(out))


def decode_xstring(s: str) -> str:
    if "_x" not in s:
        return s

    def rep(m: re.Match) -> str:
        try:
            return chr(int(m.group(1), 16))
        except ValueError:
            return m.group(0)

    return _ESC_RE.sub(rep, s)


_XML_OK = re.compile("[^\u0009\u000A\u000D -퟿-�\U00010000-\U0010FFFF]")


def encode_xstring(s: str) -> str:
    """Escapes characters XML cannot hold as ``_xHHHH_`` (and a literal ``_xHHHH_`` as
    ``_x005F_xHHHH_``)."""
    if "_x" in s:
        s = _ESC_RE.sub(lambda m: "_x005F_" + m.group(0)[1:], s)
    if _XML_OK.search(s):
        s = _XML_OK.sub(lambda m: f"_x{ord(m.group(0)) & 0xFFFF:04X}_", s)
    return s


def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def esc_attr(s: str) -> str:
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    return s.replace("\n", "&#10;").replace("\r", "&#13;").replace("\t", "&#9;")


def clean(s: str) -> str:
    """Drops characters XML 1.0 cannot represent at all (for attributes)."""
    return _XML_OK.sub("", s)


def t_el(s: str) -> str:
    enc = esc(encode_xstring(s))
    if s[:1] in (" ", "\t", "\n", "\r") or s[-1:] in (" ", "\t", "\n", "\r") or "\n" in s:
        return f'<t xml:space="preserve">{enc}</t>'
    return f"<t>{enc}</t>"


def num(n: float) -> str:
    """A number as XML text: integers without a decimal point, others with full precision."""
    if not math.isfinite(n):
        return "0"
    if n == int(n) and abs(n) < 1e15:
        return str(int(n))
    return repr(float(n)).replace("e", "E")


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)
