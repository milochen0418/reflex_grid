"""Fixed lookup tables: built-in number formats, the legacy indexed colour palette, paper sizes,
functions that need the ``_xlfn.`` prefix in files, and unit conversions."""

from __future__ import annotations

import math
from typing import Optional

from ..numfmt.builtin import BUILTINS

_BUILTIN_BY_ID = dict(BUILTINS)


def builtin_format(fid: int) -> Optional[str]:
    return _BUILTIN_BY_ID.get(fid)


def builtin_id(code: str) -> Optional[int]:
    """Id to write for a format code when it is a locale-independent built-in (currency formats
    5–8 are written as custom formats because their meaning depends on the reader's locale)."""
    for i, c in BUILTINS:
        if c == code and not 5 <= i <= 8:
            return i
    return None


# The 64-entry legacy palette (`indexed` colours), RRGGBB.
INDEXED = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
    0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
    0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
    0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
    0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
    0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
]

PAPERS = [
    (1, "Letter"), (3, "Tabloid"), (4, "Ledger"), (5, "Legal"), (7, "Executive"), (8, "A3"), (9, "A4"),
    (11, "A5"), (12, "B4"), (13, "B5"), (20, "Envelope #10"), (27, "Envelope DL"), (34, "Envelope B5"),
]


def paper_name(pid: int) -> Optional[str]:
    for i, n in PAPERS:
        if i == pid:
            return n
    return None


def paper_id(name: str) -> Optional[int]:
    n = name.strip().lower()
    for i, p in PAPERS:
        if p.lower() == n:
            return i
    return None


# Functions introduced after Excel 2007 that files store with the `_xlfn.` prefix. Entries with
# `_xlws.` get `_xlfn._xlws.`.
FUTURE_FUNCTIONS = [
    "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST",
    "BINOM.DIST.RANGE", "BINOM.INV", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW",
    "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
    "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH",
    "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP", "ENCODEURL", "ERF.PRECISE",
    "ERFC.PRECISE", "EXPAND", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FIELDVALUE",
    "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
    "GAUSS", "GROUPBY", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMAGE", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH",
    "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA", "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET", "LOGNORM.DIST",
    "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT", "NEGBINOM.DIST",
    "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION",
    "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTOF", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI",
    "PIVOTBY", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE",
    "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST", "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS",
    "SINGLE", "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV",
    "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW", "TRIMRANGE",
    "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S", "VSTACK", "WEBSERVICE", "WEIBULL.DIST",
    "WORKDAY.INTL", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR", "Z.TEST", "_xlws.FILTER", "_xlws.SORT",
]

_FUTURE = {}
for _f in FUTURE_FUNCTIONS:
    if _f.startswith("_xlws."):
        _FUTURE[_f[6:]] = "_xlfn._xlws." + _f[6:]
    else:
        _FUTURE[_f] = "_xlfn." + _f


def file_function_name(name: str) -> Optional[str]:
    """``XLOOKUP`` → ``_xlfn.XLOOKUP``, ``FILTER`` → ``_xlfn._xlws.FILTER``."""
    return _FUTURE.get(name.upper())


# ---------------------------------------------------------------- units


def col_width_to_px(w: float) -> float:
    """Column width as stored in a file (characters including padding) → px."""
    if not math.isfinite(w) or w <= 0:
        return 0.0
    w = min(w, 255.0)
    return float(math.trunc(((256.0 * w + math.trunc(128.0 / 7.0)) / 256.0) * 7.0))


def px_to_col_width(px: float) -> float:
    """px → stored column width (rounds back to the same px)."""
    px = min(float(round(max(px, 0.0))), 1790.0)
    if px > 12.0:
        chars = math.trunc((px - 5.0) / 7.0 * 100.0 + 0.5) / 100.0
    else:
        chars = px / 12.0
    w = math.trunc((chars * 7.0 + 5.0) / 7.0 * 256.0) / 256.0
    if col_width_to_px(w) == px:
        return w
    s = max(math.ceil(px * 256.0 / 7.0 - 18.0), 0.0)
    return s / 256.0


def pt_to_px(pt: float) -> float:
    if not math.isfinite(pt) or pt <= 0:
        return 0.0
    return min(pt, 409.5) * 4.0 / 3.0


def px_to_pt(px: float) -> float:
    return round(max(px, 0.0) * 0.75 * 100.0) / 100.0


EMU_PER_PX = 9525.0
