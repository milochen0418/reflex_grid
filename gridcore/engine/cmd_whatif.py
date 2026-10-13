"""What-If Analysis: Goal Seek."""

from __future__ import annotations

import math
from typing import Any, Dict, Optional

from ..calc import Calc
from ..core.addr import a1, parse_cell
from ..model.cell import Cell
from .session import EngineError, Session, command, edit, need, p_float, param_sheet


@command("data.goalSeek", "Goal Seek…", menu=("Data", "Forecast", "What-If Analysis"),
         params='{setCell: "B5", toValue: 100, byChangingCell: "B1", maxIterations?: 100, maxChange?: 0.001}')
def goal_seek(s: Session, p: Dict[str, Any]):
    d = s.doc()
    si = param_sheet(s, p)
    target = parse_cell(str(need(p, "setCell", "data.goalSeek")).rsplit("!", 1)[-1])
    change = parse_cell(str(need(p, "byChangingCell", "data.goalSeek")).rsplit("!", 1)[-1])
    goal = p_float(p, "toValue")
    if target is None or change is None or goal is None:
        raise EngineError("Goal Seek needs a cell to set, a value and a cell to change.")
    sh = d.wb.sheet(si)
    tc = sh.cell(target)
    if tc is None or tc.formula is None:
        raise EngineError("Cell must contain a formula.")
    cc = sh.cell(change)
    if cc is not None and (cc.formula is not None or (cc.value is not None and not isinstance(cc.value, float))):
        raise EngineError("Cell must contain a value.")
    max_iter = int(p_float(p, "maxIterations", 100) or 100)
    tol = p_float(p, "maxChange", 0.001) or 0.001
    style = cc.style if cc is not None else sh.style_id(change)
    scratch = d.wb.clone()
    calc = Calc()
    calc.rebuild(scratch)

    def f(x: float) -> Optional[float]:
        sm = scratch.sheet_mut(si)
        sm.set_cell(change, Cell(x, None, style))
        calc.cells_changed(scratch, [(si, change[0], change[1])])
        v = scratch.sheet(si).value(target)
        if isinstance(v, bool) or not isinstance(v, float) or not math.isfinite(v):
            return None
        return v - goal

    x0 = cc.value if cc is not None and isinstance(cc.value, float) else 0.0
    best_x, best_err = x0, f(x0)
    if best_err is None:
        raise EngineError("Goal Seek can't evaluate the formula at the starting value.")
    found = abs(best_err) <= tol
    x_prev, e_prev = x0, best_err
    x = x0 + (abs(x0) * 0.01 if x0 else 0.01)
    it = 0
    lo = hi = None
    while not found and it < max_iter:
        it += 1
        e = f(x)
        if e is None:
            x = (x + x_prev) / 2
            continue
        if abs(e) < abs(best_err):
            best_x, best_err = x, e
        if abs(e) <= tol:
            found = True
            break
        # Keep a bracket when the sign changes: bisection is the fallback when secant stalls.
        if (e > 0) != (e_prev > 0):
            lo, hi = (x_prev, x) if x_prev < x else (x, x_prev)
        denom = e - e_prev
        if denom != 0 and math.isfinite(denom):
            nx = x - e * (x - x_prev) / denom
        else:
            nx = x * 2 if x else 1.0
        if lo is not None and not (lo < nx < hi):
            nx = (lo + hi) / 2
        if not math.isfinite(nx) or abs(nx) > 1e300:
            break
        x_prev, e_prev, x = x, e, nx

    value = best_x
    if abs(value - round(value)) < 1e-9:
        value = float(round(value))

    def commit(ctx):
        sm = ctx.sheet_mut(si)
        sm.set_cell(change, Cell(value, None, style))
        ctx.touch(si, change)

    edit(s, commit)
    result = s.doc().wb.sheet(si).value(target)
    return {"found": found, "cell": a1(*change), "value": value, "target": a1(*target), "result": result, "iterations": it}
