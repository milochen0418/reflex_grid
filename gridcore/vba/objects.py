"""Objects of the VBA library and the common COM classes macros create: Collection,
Scripting.Dictionary, VBScript.RegExp, System.Collections.ArrayList and Debug. Classes that
reach outside the workbook (FileSystemObject, WScript.Shell, ADODB, XMLHTTP…) aren't available
in the sandbox."""

from __future__ import annotations

import decimal
import re
from typing import Any, Dict, Iterator, List, Optional, Tuple

from . import values as V
from .errors import MacroStopped, NotSupported, VbaError
from .runtime import Execution, HostObject
from .values import EMPTY, MISSING, NOTHING, NULL, VArray, is_object


def _plain(v: Any) -> Any:
    return V.value_of(v) if is_object(v) and v is not NOTHING else v


class Enumerator(HostObject):
    """What ``[_NewEnum]`` returns: something For Each can go through."""

    TYPE_NAME = "Unknown"

    def __init__(self, items: List[Any]):
        self.items = items

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter(list(self.items))


# ---------------------------------------------------------------- Collection


class Collection(HostObject):
    TYPE_NAME = "Collection"
    DEFAULT = "item"
    PARAMS = {"add": ("Item", "Key", "Before", "After"), "item": ("Index",), "remove": ("Index",)}

    def __init__(self) -> None:
        self.items: List[Any] = []
        self.keys: List[Optional[str]] = []

    def _index(self, ex: Execution, index: Any) -> int:
        index = _plain(index)
        if isinstance(index, str):
            k = index.lower()
            for i, kk in enumerate(self.keys):
                if kk == k:
                    return i
            raise VbaError(5)
        n = V.to_long(index)
        if not 1 <= n <= len(self.items):
            raise VbaError(5)
        return n - 1

    def g_add(self, ex, item=MISSING, key=MISSING, before=MISSING, after=MISSING):
        if item is MISSING:
            raise VbaError(449)
        k = None
        if key is not MISSING:
            kv = _plain(key)
            if not isinstance(kv, str):
                kv = V.to_string(kv)
            k = kv.lower()
            if k in self.keys:
                raise VbaError(457)
        if before is not MISSING and after is not MISSING:
            raise VbaError(5)
        pos = len(self.items)
        if before is not MISSING:
            pos = self._index(ex, before)
        elif after is not MISSING:
            pos = self._index(ex, after) + 1
        self.items.insert(pos, item)
        self.keys.insert(pos, k)
        return EMPTY

    def g_count(self, ex):
        return len(self.items)

    def g_item(self, ex, index=MISSING):
        if index is MISSING:
            raise VbaError(449)
        return self.items[self._index(ex, index)]

    def g_remove(self, ex, index=MISSING):
        if index is MISSING:
            raise VbaError(449)
        i = self._index(ex, index)
        del self.items[i]
        del self.keys[i]
        return EMPTY

    def g__newenum(self, ex):
        return Enumerator(self.items)

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter(list(self.items))


# ---------------------------------------------------------------- Scripting.Dictionary


def _dict_key(k: Any, text: bool) -> Tuple:
    t = type(k)
    if t is str:
        return ("s", k.lower() if text else k)
    if t is bool:
        return ("b", k)
    if k is EMPTY:
        return ("e",)
    if k is NULL:
        return ("null",)
    if t in (int, V.Integer, V.Byte, V.LongLong, float, V.Single, V.Date):
        f = float(k)
        return ("n", f)
    if t is V.Currency:
        return ("n", float(k))
    if t is decimal.Decimal:
        return ("n", float(k))
    if t is VArray:
        raise VbaError(13)
    return ("o", id(k))


class Dictionary(HostObject):
    TYPE_NAME = "Dictionary"
    DEFAULT = "item"
    PARAMS = {"add": ("Key", "Item"), "item": ("Key",), "exists": ("Key",), "remove": ("Key",), "key": ("Key",)}

    def __init__(self) -> None:
        self.data: Dict[Tuple, List[Any]] = {}
        self.text = False

    def g_add(self, ex, key=MISSING, item=MISSING):
        if key is MISSING or item is MISSING:
            raise VbaError(449)
        k = _dict_key(_plain_key(key), self.text)
        if k in self.data:
            raise VbaError(457)
        self.data[k] = [_plain_key(key), item]
        return EMPTY

    def g_item(self, ex, key=MISSING):
        if key is MISSING:
            raise VbaError(449)
        k = _dict_key(_plain_key(key), self.text)
        e = self.data.get(k)
        if e is None:
            # Reading a missing key adds it, as Scripting.Dictionary does.
            e = self.data[k] = [_plain_key(key), EMPTY]
        return e[1]

    def l_item(self, ex, value, key=MISSING):
        if key is MISSING:
            raise VbaError(449)
        k = _dict_key(_plain_key(key), self.text)
        e = self.data.get(k)
        if e is None:
            self.data[k] = [_plain_key(key), V.copy_value(value)]
        else:
            e[1] = V.copy_value(value)

    s_item = l_item

    def l_key(self, ex, new, key=MISSING):
        k = _dict_key(_plain_key(key), self.text)
        e = self.data.get(k)
        if e is None:
            raise VbaError(32811, "Element not found")
        nk = _dict_key(_plain_key(new), self.text)
        if nk in self.data and nk != k:
            raise VbaError(457)
        del self.data[k]
        self.data[nk] = [_plain_key(new), e[1]]

    def g_exists(self, ex, key=MISSING):
        return _dict_key(_plain_key(key), self.text) in self.data

    def g_count(self, ex):
        return len(self.data)

    def g_keys(self, ex):
        ks = [e[0] for e in self.data.values()]
        return VArray([(0, len(ks) - 1)], ks)

    def g_items(self, ex):
        vs = [e[1] for e in self.data.values()]
        return VArray([(0, len(vs) - 1)], vs)

    def g_remove(self, ex, key=MISSING):
        k = _dict_key(_plain_key(key), self.text)
        if k not in self.data:
            raise VbaError(32811, "Element not found")
        del self.data[k]
        return EMPTY

    def g_removeall(self, ex):
        self.data.clear()
        return EMPTY

    def g_comparemode(self, ex):
        return 1 if self.text else 0

    def l_comparemode(self, ex, v):
        if self.data:
            raise VbaError(5)
        self.text = V.to_long(_plain(v)) == 1

    def g_hashval(self, ex, key=MISSING):
        return hash(_dict_key(_plain_key(key), self.text)) & 0x7FFFFFFF

    def g__newenum(self, ex):
        return Enumerator([e[0] for e in self.data.values()])

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter([e[0] for e in self.data.values()])


def _plain_key(k: Any) -> Any:
    # An object key stays the object (keys may be objects).
    return k


# ---------------------------------------------------------------- VBScript.RegExp


MAX_REGEX_INPUT = 2_000_000


class RegExp(HostObject):
    TYPE_NAME = "IRegExp2"
    TYPE_NAMES = {"regexp", "iregexp2", "iregexp"}
    PARAMS = {"test": ("sourceString",), "execute": ("sourceString",), "replace": ("sourceString", "replaceVar")}

    def __init__(self) -> None:
        self.pattern = ""
        self.global_ = False
        self.ignorecase = False
        self.multiline = False

    def g_pattern(self, ex):
        return self.pattern

    def l_pattern(self, ex, v):
        self.pattern = V.to_string(_plain(v))

    def g_global(self, ex):
        return self.global_

    def l_global(self, ex, v):
        self.global_ = V.to_boolean(_plain(v))

    def g_ignorecase(self, ex):
        return self.ignorecase

    def l_ignorecase(self, ex, v):
        self.ignorecase = V.to_boolean(_plain(v))

    def g_multiline(self, ex):
        return self.multiline

    def l_multiline(self, ex, v):
        self.multiline = V.to_boolean(_plain(v))

    def _rx(self):
        flags = 0
        if self.ignorecase:
            flags |= re.IGNORECASE
        if self.multiline:
            flags |= re.MULTILINE
        try:
            return re.compile(_js_to_py(self.pattern), flags)
        except re.error as e:
            raise VbaError(5017, f"Syntax error in regular expression ({e})")

    def _subject(self, s: Any) -> str:
        t = V.to_string(_plain(s))
        if len(t) > MAX_REGEX_INPUT:
            raise VbaError(5, "The text is too long for a regular expression")
        return t

    def g_test(self, ex, s=MISSING):
        return self._rx().search(self._subject(s)) is not None

    def g_execute(self, ex, s=MISSING):
        rx = self._rx()
        text = self._subject(s)
        out = []
        if self.global_:
            for m in rx.finditer(text):
                out.append(Match(m))
                if len(out) > 1_000_000:
                    break
        else:
            m = rx.search(text)
            if m is not None:
                out.append(Match(m))
        return MatchCollection(out)

    def g_replace(self, ex, s=MISSING, rep=MISSING):
        rx = self._rx()
        text = self._subject(s)
        r = V.to_string(_plain(rep))

        def sub(m: "re.Match") -> str:
            return _js_replacement(r, m, text)

        return rx.sub(sub, text, count=0 if self.global_ else 1)


def _js_to_py(p: str) -> str:
    # VBScript patterns are ECMAScript ones; Python differs in a few corners.
    out = []
    i = 0
    in_class = False
    while i < len(p):
        c = p[i]
        if c == "\\" and i + 1 < len(p):
            out.append(p[i:i + 2])
            i += 2
            continue
        if c == "[":
            in_class = True
        elif c == "]":
            in_class = False
        elif c == "$" and not in_class:
            # ECMAScript `$` doesn't match before a final newline.
            out.append(r"(?!\n)$" if False else "$")
            i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _js_replacement(r: str, m: "re.Match", text: str) -> str:
    out = []
    i = 0
    while i < len(r):
        c = r[i]
        if c == "$" and i + 1 < len(r):
            n = r[i + 1]
            if n == "$":
                out.append("$")
                i += 2
                continue
            if n == "&":
                out.append(m.group(0))
                i += 2
                continue
            if n == "`":
                out.append(text[:m.start()])
                i += 2
                continue
            if n == "'":
                out.append(text[m.end():])
                i += 2
                continue
            if n.isdigit():
                j = i + 2
                if j < len(r) and r[j].isdigit() and int(r[i + 1:j + 1]) <= (m.re.groups or 0):
                    j += 1
                g = int(r[i + 1:j])
                if 1 <= g <= (m.re.groups or 0):
                    out.append(m.group(g) or "")
                    i = j
                    continue
        out.append(c)
        i += 1
    return "".join(out)


class Match(HostObject):
    TYPE_NAME = "IMatch2"
    TYPE_NAMES = {"match", "imatch2", "imatch"}
    DEFAULT = "value"

    def __init__(self, m: "re.Match"):
        self.value = m.group(0)
        self.first = m.start()
        self.subs = [g if g is not None else EMPTY for g in m.groups()]

    def g_value(self, ex):
        return self.value

    def g_firstindex(self, ex):
        return self.first

    def g_length(self, ex):
        return len(self.value)

    def g_submatches(self, ex):
        return SubMatches(self.subs)


class SubMatches(HostObject):
    TYPE_NAME = "ISubMatches"
    DEFAULT = "item"

    def __init__(self, items: List[Any]):
        self.items = items

    def g_count(self, ex):
        return len(self.items)

    def g_item(self, ex, i=MISSING):
        n = V.to_long(_plain(i))
        if not 0 <= n < len(self.items):
            raise VbaError(5)
        return self.items[n]

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter(list(self.items))


class MatchCollection(HostObject):
    TYPE_NAME = "IMatchCollection2"
    DEFAULT = "item"

    def __init__(self, items: List[Match]):
        self.items = items

    def g_count(self, ex):
        return len(self.items)

    def g_item(self, ex, i=MISSING):
        n = V.to_long(_plain(i))
        if not 0 <= n < len(self.items):
            raise VbaError(5)
        return self.items[n]

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter(list(self.items))


# ---------------------------------------------------------------- ArrayList


class ArrayList(HostObject):
    TYPE_NAME = "ArrayList"
    DEFAULT = "item"

    def __init__(self) -> None:
        self.items: List[Any] = []

    def _i(self, i: Any) -> int:
        n = V.to_long(_plain(i))
        if not 0 <= n < len(self.items):
            raise VbaError(5, "Index was out of range")
        return n

    def g_add(self, ex, v=MISSING):
        self.items.append(v)
        return len(self.items) - 1

    def g_count(self, ex):
        return len(self.items)

    def g_item(self, ex, i=MISSING):
        return self.items[self._i(i)]

    def l_item(self, ex, v, i=MISSING):
        self.items[self._i(i)] = v

    def g_insert(self, ex, i=MISSING, v=MISSING):
        n = V.to_long(_plain(i))
        if not 0 <= n <= len(self.items):
            raise VbaError(5)
        self.items.insert(n, v)
        return EMPTY

    def g_contains(self, ex, v=MISSING):
        return any(V.compare(x, v) == 0 for x in self.items if not is_object(x))

    def g_indexof(self, ex, v=MISSING, start=MISSING):
        for i, x in enumerate(self.items):
            if not is_object(x) and V.compare(x, v) == 0:
                return i
        return -1

    def g_remove(self, ex, v=MISSING):
        for i, x in enumerate(self.items):
            if not is_object(x) and V.compare(x, v) == 0:
                del self.items[i]
                break
        return EMPTY

    def g_removeat(self, ex, i=MISSING):
        del self.items[self._i(i)]
        return EMPTY

    def g_clear(self, ex):
        self.items.clear()
        return EMPTY

    def g_sort(self, ex):
        import functools

        self.items.sort(key=functools.cmp_to_key(lambda a, b: V.compare(a, b)))
        return EMPTY

    def g_reverse(self, ex):
        self.items.reverse()
        return EMPTY

    def g_toarray(self, ex):
        return VArray([(0, len(self.items) - 1)], list(self.items))

    def vba_enum(self, ex: Execution) -> Iterator[Any]:
        return iter(list(self.items))


# ---------------------------------------------------------------- Debug


class DebugObject(HostObject):
    TYPE_NAME = "Debug"

    def g_print(self, ex, *args):
        from .stdlib import debug_print

        debug_print(ex, [(a, ";" if i < len(args) - 1 else "", None) for i, a in enumerate(args)])
        return EMPTY

    def g_assert(self, ex, cond=MISSING):
        if cond is not MISSING and not V.is_true(_plain(cond)):
            raise MacroStopped("stop", "Debug.Assert failed")
        return EMPTY


DEBUG = DebugObject()


LIBRARY_CLASSES = {"collection", "dictionary", "regexp", "arraylist", "dataobject"}


def new_library_object(ex: Optional[Execution], name: str) -> Any:
    n = name.lower()
    if n == "collection":
        return Collection()
    if n == "dictionary":
        return Dictionary()
    if n == "regexp":
        return RegExp()
    if n == "arraylist":
        return ArrayList()
    if n == "dataobject":
        from .msforms import DataObject

        return DataObject()
    return None


_UNAVAILABLE = (
    ("scripting.filesystemobject", "the file system (Scripting.FileSystemObject)"),
    ("wscript.shell", "running programs (WScript.Shell)"),
    ("wscript.network", "the network (WScript.Network)"),
    ("shell.application", "the Windows shell (Shell.Application)"),
    ("adodb.", "databases (ADODB)"),
    ("dao.", "databases (DAO)"),
    ("msxml2.", "the network (MSXML2)"),
    ("microsoft.xmlhttp", "the network (XMLHTTP)"),
    ("winhttp.", "the network (WinHttp)"),
    ("outlook.application", "Outlook"),
    ("word.application", "Word"),
    ("powerpoint.application", "PowerPoint"),
    ("access.application", "Access"),
    ("internetexplorer.application", "Internet Explorer"),
    ("sapi.", "speech"),
)


def create_object(ex: Execution, progid: str) -> Any:
    p = progid.strip().lower()
    if p in ("scripting.dictionary",):
        return Dictionary()
    if p in ("vbscript.regexp", "vbscript.regexp.5.5"):
        return RegExp()
    if p in ("system.collections.arraylist",):
        return ArrayList()
    if p in ("new:{1c3b4210-f441-11ce-b9ea-00aa006b1a69}", "forms.dataobject", "forms.dataobject.1"):
        from .msforms import DataObject

        return DataObject()
    if p in ("excel.application", "excel.application.16"):
        return ex.host.global_get(ex, "Application", [], None)
    for prefix, what in _UNAVAILABLE:
        if p.startswith(prefix):
            raise NotSupported(what, 429)
    raise VbaError(429, f"ActiveX component can't create object: {progid}")
