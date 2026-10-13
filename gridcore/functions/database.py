"""Database functions (DSUM, DAVERAGE, DCOUNT…): aggregate a field of the records of a list that
match a criteria range."""

from __future__ import annotations

import math
from typing import List

from ..core.value import Array, CellError, Value, XlError, scalar, to_text
from .criteria import Criterion
from .registry import A, Arg, Ctx, fn
from .stat import mean, var_p, var_s

DB = "Database"


def _field_index(db: Array, field: Value) -> int:
    field = scalar(field)
    if isinstance(field, CellError):
        raise XlError(field)
    if isinstance(field, float):
        i = int(field)
        if not 1 <= i <= db.cols:
            raise XlError(CellError.VALUE)
        return i - 1
    name = to_text(field).strip().lower()
    for c in range(db.cols):
        h = db.get(0, c)
        if h is not None and to_text(h).strip().lower() == name:
            return c
    raise XlError(CellError.VALUE)


def _matching(db: Array, crit: Array) -> List[int]:
    """Database rows (1-based within the list, header excluded) that match any criteria row."""
    heads = [to_text(crit.get(0, c)).strip().lower() if crit.get(0, c) is not None else "" for c in range(crit.cols)]
    cols = []
    for h in heads:
        idx = None
        for c in range(db.cols):
            v = db.get(0, c)
            if v is not None and to_text(v).strip().lower() == h:
                idx = c
                break
        cols.append(idx)
    conds = []
    for r in range(1, crit.rows):
        row = []
        for c in range(crit.cols):
            v = crit.get(r, c)
            if v is None:
                continue
            if cols[c] is None:
                row.append((None, None))
                continue
            row.append((cols[c], Criterion.parse(v)))
        conds.append(row)
    out = []
    for r in range(1, db.rows):
        if not conds:
            out.append(r)
            continue
        for row in conds:
            ok = True
            for c, cr in row:
                if c is None or not cr.matches(db.get(r, c)):
                    ok = False
                    break
            if ok:
                out.append(r)
                break
    return out


def _values(args) -> List[Value]:
    db = args[0].value
    crit = args[2].value
    for v in (db, crit):
        if isinstance(v, CellError):
            raise XlError(v)
    db = db if isinstance(db, Array) else Array.scalar(db)
    crit = crit if isinstance(crit, Array) else Array.scalar(crit)
    if db.rows < 1 or crit.rows < 1:
        raise XlError(CellError.VALUE)
    c = _field_index(db, args[1].value)
    return [db.get(r, c) for r in _matching(db, crit)]


def _nums(vals: List[Value]) -> List[float]:
    out = []
    for v in vals:
        if isinstance(v, CellError):
            raise XlError(v)
        if type(v) is float:
            out.append(v)
    return out


def _db(name: str, f, desc: str) -> None:
    @fn(name, 3, 3, DB, A, f"{name}(database, field, criteria)", desc)
    def _impl(args, ctx, f=f):
        return f(_values(args))


_db("DSUM", lambda v: math.fsum(_nums(v)), "Adds the numbers in the field column of records in the database that match the criteria.")
_db("DAVERAGE", lambda v: mean(_nums(v)), "Returns the average of selected database entries.")
_db("DCOUNT", lambda v: float(len(_nums(v))), "Counts the cells that contain numbers in a database.")
_db("DCOUNTA", lambda v: float(sum(1 for x in v if x is not None)), "Counts nonblank cells in a database.")
_db("DMAX", lambda v: max(_nums(v), default=0.0), "Returns the maximum value from selected database entries.")
_db("DMIN", lambda v: min(_nums(v), default=0.0), "Returns the minimum value from selected database entries.")
_db("DSTDEV", lambda v: math.sqrt(var_s(_nums(v))), "Estimates the standard deviation based on a sample of selected database entries.")
_db("DSTDEVP", lambda v: math.sqrt(var_p(_nums(v))), "Calculates the standard deviation based on the entire population of selected database entries.")
_db("DVAR", lambda v: var_s(_nums(v)), "Estimates variance based on a sample from selected database entries.")
_db("DVARP", lambda v: var_p(_nums(v)), "Calculates variance based on the entire population of selected database entries.")


def _dproduct(v: List[Value]) -> float:
    ns = _nums(v)
    p = 1.0
    for x in ns:
        p *= x
    return p if ns else 0.0


_db("DPRODUCT", _dproduct, "Multiplies the values in a particular field of records that match the criteria in a database.")


def _dget(v: List[Value]) -> Value:
    if not v:
        raise XlError(CellError.VALUE)
    if len(v) > 1:
        raise XlError(CellError.NUM)
    return 0.0 if v[0] is None else v[0]


_db("DGET", _dget, "Extracts from a database a single record that matches the specified criteria.")
