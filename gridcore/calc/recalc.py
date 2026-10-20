"""Dependency graph and recalculation.

The graph maps every formula cell to its static precedents (resolved areas). A reverse index
answers "which formulas read this cell?": single-cell precedents in a dict, ranges bucketed by
(sheet, column, 256-row block), very tall ranges per column and very wide ranges per sheet.
After an edit the dirty set is the transitive closure of dependents; it is evaluated in
topological order (Kahn's algorithm over the dirty subgraph), with on-demand evaluation as a
safety net for dynamic references (INDIRECT, OFFSET, names). Cycles produce ``#CIRC!`` unless
iterative calculation is on: then a formula read while it is being evaluated gives its previous
value, and the loop is evaluated again until it settles.
"""

from __future__ import annotations

import heapq
import sys
from collections import deque
from typing import Dict, Iterable, List, Optional, Set, Tuple

from ..core.addr import MAX_COLS, MAX_ROWS, RangeRef
from ..core.date import now_serial
from ..core.value import Array, CellError, Value
from ..formula import Call, Expr, ParseError, parse
from .eval import Area, Evaluator, Host, precedents

if sys.getrecursionlimit() < 20000:
    sys.setrecursionlimit(20000)

Key = Tuple[int, int, int]  # (sheet, row, col)

BLOCK = 256
TALL_BLOCKS = 64
WIDE = 64
MAX_COMPUTE_DEPTH = 1500
MAX_ITERATION_WORK = 2_000_000


def same_value(a: Value, b: Value) -> bool:
    return type(a) is type(b) and a == b


class Graph:
    def __init__(self) -> None:
        self.nodes: Dict[Key, Tuple[List[Area], bool]] = {}
        self.cell_deps: Dict[Key, Set[Key]] = {}
        self.block_deps: Dict[Tuple[int, int, int], Set[Tuple[int, int, Key]]] = {}
        self.tall_deps: Dict[Tuple[int, int], Set[Tuple[int, int, Key]]] = {}
        self.wide_deps: Dict[int, Set[Tuple[int, int, int, int, Key]]] = {}
        self.dynamic: Set[Key] = set()

    def __len__(self) -> int:
        return len(self.nodes)

    @staticmethod
    def _entries(areas: List[Area]):
        cells, blocks, talls, wides = set(), set(), set(), set()
        for a in areas:
            r = a.rng
            if r.r0 == r.r1 and r.c0 == r.c1:
                cells.add((a.sheet, r.r0, r.c0))
            elif r.c1 - r.c0 + 1 > WIDE:
                wides.add((a.sheet, r.r0, r.r1, r.c0, r.c1))
            else:
                b0, b1 = r.r0 // BLOCK, r.r1 // BLOCK
                if b1 - b0 + 1 > TALL_BLOCKS:
                    for c in range(r.c0, r.c1 + 1):
                        talls.add((a.sheet, c, r.r0, r.r1))
                else:
                    for c in range(r.c0, r.c1 + 1):
                        for b in range(b0, b1 + 1):
                            blocks.add((a.sheet, c, b, r.r0, r.r1))
        return cells, blocks, talls, wides

    def insert(self, k: Key, areas: List[Area], dynamic: bool) -> None:
        self.remove(k)
        cells, blocks, talls, wides = self._entries(areas)
        for c in cells:
            self.cell_deps.setdefault(c, set()).add(k)
        for s, c, b, r0, r1 in blocks:
            self.block_deps.setdefault((s, c, b), set()).add((r0, r1, k))
        for s, c, r0, r1 in talls:
            self.tall_deps.setdefault((s, c), set()).add((r0, r1, k))
        for s, r0, r1, c0, c1 in wides:
            self.wide_deps.setdefault(s, set()).add((r0, r1, c0, c1, k))
        if dynamic:
            self.dynamic.add(k)
        self.nodes[k] = (areas, dynamic)

    def remove(self, k: Key) -> None:
        old = self.nodes.pop(k, None)
        if old is None:
            return
        cells, blocks, talls, wides = self._entries(old[0])
        for c in cells:
            s = self.cell_deps.get(c)
            if s is not None:
                s.discard(k)
                if not s:
                    del self.cell_deps[c]
        for s, c, b, r0, r1 in blocks:
            st = self.block_deps.get((s, c, b))
            if st is not None:
                st.discard((r0, r1, k))
                if not st:
                    del self.block_deps[(s, c, b)]
        for s, c, r0, r1 in talls:
            st = self.tall_deps.get((s, c))
            if st is not None:
                st.discard((r0, r1, k))
                if not st:
                    del self.tall_deps[(s, c)]
        for s, r0, r1, c0, c1 in wides:
            st = self.wide_deps.get(s)
            if st is not None:
                st.discard((r0, r1, c0, c1, k))
                if not st:
                    del self.wide_deps[s]
        self.dynamic.discard(k)

    def dependents(self, sheet: int, r: int, c: int) -> List[Key]:
        """Formulas that read cell (r, c) of ``sheet``."""
        out: List[Key] = []
        s = self.cell_deps.get((sheet, r, c))
        if s:
            out.extend(s)
        b = self.block_deps.get((sheet, c, r // BLOCK))
        if b:
            out.extend(k for r0, r1, k in b if r0 <= r <= r1)
        t = self.tall_deps.get((sheet, c))
        if t:
            out.extend(k for r0, r1, k in t if r0 <= r <= r1)
        w = self.wide_deps.get(sheet)
        if w:
            out.extend(k for r0, r1, c0, c1, k in w if r0 <= r <= r1 and c0 <= c <= c1)
        return out

    def dependents_of_range(self, sheet: int, rng: RangeRef) -> Set[Key]:
        """Formulas that read any cell of ``rng``."""
        out: Set[Key] = set()
        if rng.count() <= 4096:
            for (r, c) in rng.cells():
                s = self.cell_deps.get((sheet, r, c))
                if s:
                    out.update(s)
        else:
            for (s, r, c), ks in self.cell_deps.items():
                if s == sheet and rng.contains((r, c)):
                    out.update(ks)
        b0, b1 = rng.r0 // BLOCK, rng.r1 // BLOCK
        if rng.width * (b1 - b0 + 1) <= len(self.block_deps):
            for c in range(rng.c0, rng.c1 + 1):
                for b in range(b0, b1 + 1):
                    st = self.block_deps.get((sheet, c, b))
                    if st:
                        out.update(k for r0, r1, k in st if r0 <= rng.r1 and r1 >= rng.r0)
        else:
            for (s, c, b), st in self.block_deps.items():
                if s == sheet and rng.c0 <= c <= rng.c1 and b0 <= b <= b1:
                    out.update(k for r0, r1, k in st if r0 <= rng.r1 and r1 >= rng.r0)
        for (s, c), st in self.tall_deps.items():
            if s == sheet and rng.c0 <= c <= rng.c1:
                out.update(k for r0, r1, k in st if r0 <= rng.r1 and r1 >= rng.r0)
        w = self.wide_deps.get(sheet)
        if w:
            out.update(k for r0, r1, c0, c1, k in w if r0 <= rng.r1 and r1 >= rng.r0 and c0 <= rng.c1 and c1 >= rng.c0)
        return out

    def precedents_of(self, k: Key) -> List[Area]:
        n = self.nodes.get(k)
        return list(n[0]) if n else []


class _PassHost(Host):
    """One recalculation pass: dirty formulas are computed on first read."""

    def __init__(self, wb, exprs: Dict[Key, Expr], calc: "Calc", now: float):
        self.wb = wb
        self.exprs = exprs
        self.calc = calc
        self.now = now
        self.results: Dict[Key, Value] = {}
        self.pending: Set[Key] = set(exprs)
        self.in_progress: Set[Key] = set()
        # New spills found during the pass: anchor key -> array.
        self.spills: Dict[Key, Array] = {}
        self.used: Dict[int, Optional[RangeRef]] = {}
        self.cycles: List[Key] = []
        self.depth = 0
        self.iterate = False
        self.prev: Dict[Key, Value] = {}
        self.stack: List[Key] = []
        self.members: Set[Key] = set()
        self.tainted: Set[Key] = set()
        self.taint = False
        self.volatile_hit = False

    # ------------------------------------------------------------ compute

    def compute(self, k: Key) -> Value:
        if k in self.results:
            if self.iterate and k in self.tainted:
                self.taint = True
            return self.results[k]
        expr = self.exprs.get(k)
        if expr is None:
            s = self.wb.sheet(k[0])
            return s.value((k[1], k[2])) if s is not None else None
        if k in self.in_progress or self.depth > MAX_COMPUTE_DEPTH:
            if self.iterate:
                return self._circular_read(k)
            self.cycles.append(k)
            return CellError.CIRC
        self.in_progress.add(k)
        self.depth += 1
        outer_taint = self.taint
        self.taint = False
        if self.iterate:
            self.stack.append(k)
        try:
            v = Evaluator(self, k[0], (k[1], k[2])).value(expr)
        except RecursionError:
            v = CellError.CALC
        sheet = self.wb.sheet(k[0])
        cell = sheet.cell((k[1], k[2])) if sheet is not None else None
        legacy = cell.formula.array if cell is not None and cell.formula is not None else None
        if legacy is not None:
            v = fit_to_range(v, legacy)
        self.depth -= 1
        self.in_progress.discard(k)
        self.pending.discard(k)
        if self.iterate:
            self.stack.pop()
            if self.taint:
                self.tainted.add(k)
            else:
                self.tainted.discard(k)
        self.taint = self.taint or outer_taint
        # A blank shows as 0 in a formula's result, inside an array as well.
        keep_blanks = isinstance(expr, Call) and expr.name in ("GROUPBY", "PIVOTBY")
        if isinstance(v, Array):
            a = v
            if not keep_blanks and any(x is None for x in a.data):
                a = Array(a.rows, a.cols, [0.0 if x is None else x for x in a.data])
            if a.rows > 1 or a.cols > 1:
                self.spills[k] = a
            v = a.data[0] if a.data else None
            if isinstance(v, Array):
                v = v.data[0] if v.data else None
        elif v is None:
            v = 0.0
        a = self.spills.get(k)
        if a is not None and self._spill_blocked(k, a):
            del self.spills[k]
            v = CellError.SPILL
        self.results[k] = v
        return v

    def _spill_blocked(self, k: Key, a: Array) -> bool:
        sh = self.wb.sheet(k[0])
        if sh is None:
            return False
        anchor = (k[1], k[2])
        end_r, end_c = k[1] + a.rows - 1, k[2] + a.cols - 1
        if end_r >= MAX_ROWS or end_c >= MAX_COLS:
            return True
        rng = RangeRef(k[1], k[2], end_r, end_c)
        if any(m.intersects(rng) for m in sh.merges):
            return True
        if any(t.range.contains(anchor) for t in sh.tables):
            return True
        for rc, cell in sh.cells.iter_range(rng):
            if rc != anchor and (cell.value is not None or cell.formula is not None):
                return True
        own = sh.spill_ranges.get(anchor)
        for o, r in sh.spill_ranges.items():
            if o != anchor and r.intersects(rng) and (k[0], o[0], o[1]) not in self.exprs:
                inter = r.intersection(rng)
                if inter is not None and not (own is not None and own.contains_range(inter)):
                    return True
        for o, arr in self.spills.items():
            if o != k and o[0] == k[0]:
                orng = RangeRef(o[1], o[2], o[1] + arr.rows - 1, o[2] + arr.cols - 1)
                if orng.intersects(rng):
                    return True
        return False

    def _previous(self, k: Key) -> Value:
        if k in self.prev:
            v = self.prev[k]
        else:
            s = self.wb.sheet(k[0])
            v = s.value((k[1], k[2])) if s is not None else None
        if v is None or v is CellError.CIRC:
            return 0.0
        return v

    def _circular_read(self, k: Key) -> Value:
        self.taint = True
        if k in self.stack:
            i = len(self.stack) - 1 - self.stack[::-1].index(k)
            self.members.update(self.stack[i:])
        else:
            self.members.add(k)
        return self._previous(k)

    def iterate_loops(self, order: List[Key], max_iterations: int, max_change: float) -> int:
        again = [k for k in order if k in self.tainted]
        st = set(again)
        max_passes = max(MAX_ITERATION_WORK // max(len(again), 1), 1)
        passes = 1
        while passes < min(max_iterations, max_passes) and not self._converged(max_change):
            self.prev = {k: self.results.pop(k) for k in again if k in self.results}
            self.pending = set(st)
            for k in list(self.spills):
                if k in st:
                    del self.spills[k]
            for k in again:
                self.compute(k)
            passes += 1
        return passes

    def _converged(self, max_change: float) -> bool:
        for k in self.members:
            new = self.results.get(k)
            if k not in self.results:
                continue
            old = self._previous(k)
            if isinstance(new, float) and isinstance(old, float):
                d = abs(new - old)
                if not (d < max_change or d == 0.0):
                    return False
            elif not same_value(new, old):
                return False
        return True

    # ------------------------------------------------------------ Host

    def note_volatile(self) -> None:
        self.volatile_hit = True

    def now_serial(self) -> float:
        return self.now

    def random(self) -> float:
        return self.calc.next_random()

    def spill_range(self, sheet: int, anchor: Tuple[int, int]) -> Optional[RangeRef]:
        k = (sheet, anchor[0], anchor[1])
        if k not in self.pending and k not in self.results:
            s = self.wb.sheet(sheet)
            return s.spill_ranges.get(anchor) if s is not None else None
        self.compute(k)
        arr = self.spills.get(k)
        if arr is None:
            return None
        end = (anchor[0] + arr.rows - 1, anchor[1] + arr.cols - 1)
        if end[0] >= MAX_ROWS or end[1] >= MAX_COLS:
            return None
        return RangeRef(anchor[0], anchor[1], end[0], end[1])

    def cell_value(self, sheet: int, rc: Tuple[int, int]) -> Value:
        k = (sheet, rc[0], rc[1])
        if k in self.results:
            if self.iterate and k in self.tainted:
                self.taint = True
            return self.results[k]
        if k in self.pending:
            return self.compute(k)
        if self.spills:
            for a, arr in self.spills.items():
                if a[0] == sheet and rc[0] >= a[1] and rc[1] >= a[2]:
                    dr, dc = rc[0] - a[1], rc[1] - a[2]
                    if dr < arr.rows and dc < arr.cols:
                        if self.iterate and a in self.tainted:
                            self.taint = True
                        return arr.data[dr * arr.cols + dc]
        s = self.wb.sheet(sheet)
        return s.value(rc) if s is not None else None

    def range_values(self, sheet: int, rng: RangeRef) -> List[Value]:
        h, w = rng.height, rng.width
        out: List[Value] = [None] * (h * w)
        sh = self.wb.sheet(sheet)
        if sh is None:
            return out
        r0, c0 = rng.r0, rng.c0
        formulas = []
        pending = self.pending
        results = self.results
        for (r, c), cell in sh.cells.iter_range(rng):
            i = (r - r0) * w + (c - c0)
            if cell.formula is not None:
                formulas.append((i, (r, c)))
            else:
                out[i] = cell.value
        if sh.spill:
            if len(sh.spill) <= h * w:
                for (r, c), v in sh.spill.items():
                    if r0 <= r <= rng.r1 and c0 <= c <= rng.c1:
                        i = (r - r0) * w + (c - c0)
                        if out[i] is None:
                            out[i] = v
            else:
                for (r, c) in rng.cells():
                    v = sh.spill.get((r, c))
                    if v is not None:
                        i = (r - r0) * w + (c - c0)
                        if out[i] is None:
                            out[i] = v
        for i, rc in formulas:
            k = (sheet, rc[0], rc[1])
            if k in results:
                if self.iterate and k in self.tainted:
                    self.taint = True
                out[i] = results[k]
            elif k in pending:
                out[i] = self.compute(k)
            else:
                out[i] = sh.value(rc)
        if self.spills:
            for a, arr in self.spills.items():
                if a[0] != sheet:
                    continue
                arng = RangeRef(a[1], a[2], a[1] + arr.rows - 1, a[2] + arr.cols - 1)
                inter = arng.intersection(rng)
                if inter is None:
                    continue
                if self.iterate and a in self.tainted:
                    self.taint = True
                for (r, c) in inter.cells():
                    if (r, c) == (a[1], a[2]):
                        continue
                    out[(r - r0) * w + (c - c0)] = arr.data[(r - a[1]) * arr.cols + (c - a[2])]
        return out

    def visit_range(self, sheet: int, rng: RangeRef, f) -> bool:
        sh = self.wb.sheet(sheet)
        if sh is None:
            return False
        if any(a[0] == sheet for a in self.spills):
            return False
        if sh.spill:
            if len(sh.spill_ranges) and any(r.intersects(rng) for r in sh.spill_ranges.values()):
                return False
        items = list(sh.cells.iter_range(rng))
        for rc, cell in items:
            if cell.formula is not None:
                self.cell_value(sheet, rc)
        if any(a[0] == sheet for a in self.spills):
            return False
        items.sort(key=lambda kv: kv[0])
        for rc, cell in items:
            v = self.cell_value(sheet, rc) if cell.formula is not None else cell.value
            if not f(v):
                break
        return True

    def used_range(self, sheet: int) -> Optional[RangeRef]:
        if sheet not in self.used:
            s = self.wb.sheet(sheet)
            self.used[sheet] = s.used_range() if s is not None else None
        return self.used[sheet]


def fit_to_range(v: Value, rng: RangeRef) -> Value:
    """A legacy (Ctrl+Shift+Enter) array formula's result fitted to the range it was entered in."""
    if rng.count() > 16_000_000:
        return CellError.NUM
    h, w = rng.height, rng.width
    if not isinstance(v, Array):
        return v if h * w == 1 else Array(h, w, [v] * (h * w))
    if h * w == 1:
        return v.data[0] if v.data else None
    data = []
    for r in range(h):
        for c in range(w):
            rr = 0 if v.rows == 1 else r
            cc = 0 if v.cols == 1 else c
            data.append(v.get(rr, cc, CellError.NA))
    return Array(h, w, data)


# A #SPILL! anchor this many rows and columns up/left of an edited cell may have been blocked or
# unblocked by the edit.
SPILL_REACH = 64


def spill_anchors_near(sh, rcs: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    """The spill anchors to evaluate again after the cells ``rcs`` of sheet ``sh`` changed: the
    anchors whose spill range holds a changed cell, and the ``#SPILL!`` anchors up to
    ``SPILL_REACH`` rows and columns up/left of one. The look for those is one pass over the
    cells around all the changed cells, or one window per changed cell when few cells changed
    on a big sheet; never a window per cell for every cell of a paste (a paste of 100,000 cells
    scanned 400 million entries that way)."""
    changed = set(rcs)
    out: List[Tuple[int, int]] = []
    for anchor, r in sh.spill_ranges.items():
        if r.count() <= len(rcs):
            hit = any(rc in changed and rc != anchor for rc in r.cells())
        else:
            hit = any(r.contains(rc) and rc != anchor for rc in rcs)
        if hit:
            out.append(anchor)
    reach = SPILL_REACH
    if len(rcs) * (reach + 1) ** 2 < len(sh.cells):
        seen: Set[Tuple[int, int]] = set()
        for rr, cc in rcs:
            look = RangeRef(max(rr - reach, 0), max(cc - reach, 0), rr, cc)
            for rc, cell in sh.cells.iter_range(look):
                if cell.value is CellError.SPILL and rc not in changed and rc not in seen:
                    seen.add(rc)
                    out.append(rc)
        return out
    r0 = min(r for r, _ in rcs)
    c0 = min(c for _, c in rcs)
    r1 = max(r for r, _ in rcs)
    c1 = max(c for _, c in rcs)
    few = len(rcs) <= (reach + 1) ** 2
    for (r, c), cell in sh.cells.iter_range(RangeRef(max(r0 - reach, 0), max(c0 - reach, 0), r1, c1)):
        if cell.value is not CellError.SPILL or (r, c) in changed:
            continue
        if few:
            near = any(r <= rr <= r + reach and c <= cc <= c + reach for rr, cc in rcs)
        else:
            near = any((rr, cc) in changed for rr in range(r, r + reach + 1) for cc in range(c, c + reach + 1))
        if near:
            out.append((r, c))
    return out


def _formula_expr(wb, k: Key) -> Optional[Expr]:
    s = wb.sheet(k[0])
    if s is None:
        return None
    cell = s.cell((k[1], k[2]))
    if cell is None or cell.formula is None:
        return None
    return cell.formula.get_expr()


class Calc:
    """Calculation state kept beside a workbook."""

    def __init__(self) -> None:
        self.graph = Graph()
        self.rng_state = 0x9E3779B97F4A7C15
        # Fixed "now" for deterministic tests; None = system clock.
        self.fixed_now: Optional[float] = None
        self.circular: List[Key] = []
        self.last_recalc_cells = 0

    def next_random(self) -> float:
        x = self.rng_state
        x ^= x >> 12
        x ^= (x << 25) & 0xFFFFFFFFFFFFFFFF
        x ^= x >> 27
        self.rng_state = x & 0xFFFFFFFFFFFFFFFF
        return (((x * 0x2545F4914F6CDD1D) & 0xFFFFFFFFFFFFFFFF) >> 11) / float(1 << 53)

    def now(self) -> float:
        return self.fixed_now if self.fixed_now is not None else now_serial()

    # ------------------------------------------------------------ graph

    def rebuild(self, wb) -> None:
        self.graph = Graph()
        for si, sheet in enumerate(wb.sheets):
            for (r, c), cell in sheet.cells.items():
                if cell.formula is not None:
                    e = cell.formula.get_expr()
                    if e is not None:
                        areas, dynamic = precedents(wb, si, e)
                        self.graph.insert((si, r, c), areas, dynamic)

    def recalc_all(self, wb) -> None:
        """Recalculates every formula (F9, and after loading)."""
        self.rebuild(wb)
        for i in range(len(wb.sheets)):
            s = wb.sheets[i]
            if s.spill_ranges or s.spill:
                sm = wb.sheet_mut(i)
                sm.spill = {}
                sm.spill_ranges = {}
        self._run(wb, list(self.graph.nodes), full=True)

    def cells_changed(self, wb, changed: Iterable[Key]) -> None:
        """Updates the graph for cells whose content changed and recalculates their dependents."""
        changed = list(dict.fromkeys(changed))
        cleared: List[Tuple[int, RangeRef]] = []
        for k in changed:
            e = _formula_expr(wb, k)
            if e is not None:
                areas, dynamic = precedents(wb, k[0], e)
                self.graph.insert(k, areas, dynamic)
            else:
                self.graph.remove(k)
                sh = wb.sheet(k[0])
                if sh is not None and (k[1], k[2]) in sh.spill_ranges:
                    sm = wb.sheet_mut(k[0])
                    r = sm.spill_ranges.pop((k[1], k[2]))
                    for rc in r.cells():
                        sm.spill.pop(rc, None)
                    cleared.append((k[0], r))
                sh = wb.sheet(k[0])
                cell = sh.cell((k[1], k[2])) if sh is not None else None
                if cell is not None and cell.formula is not None:
                    # A formula that doesn't parse evaluates to #NAME?.
                    wb.sheet_mut(k[0]).cells.set((k[1], k[2]), cell.with_value(CellError.NAME))
        if wb.calc.mode == "manual":
            own = [k for k in changed if k in self.graph.nodes]
            self._run(wb, own, full=False, closure=False)
            return
        seeds: List[Key] = list(changed)
        emptied: Dict[int, Tuple[int, int]] = {}
        by_sheet: Dict[int, List[Tuple[int, int]]] = {}
        for k in changed:
            by_sheet.setdefault(k[0], []).append((k[1], k[2]))
        for si, rcs in by_sheet.items():
            sh = wb.sheet(si)
            if sh is None:
                continue
            seeds.extend((si, r, c) for r, c in spill_anchors_near(sh, rcs))
            for rc in rcs:
                if not sh.cells.has(rc):
                    e0 = emptied.get(si)
                    emptied[si] = rc if e0 is None else (max(e0[0], rc[0]), max(e0[1], rc[1]))
        for si, end in emptied.items():
            for (s, r, c) in self.graph.nodes:
                if s == si and r <= end[0] and c <= end[1]:
                    sh = wb.sheet(s)
                    if sh is not None and sh.value((r, c)) is CellError.SPILL:
                        seeds.append((s, r, c))
        for k in changed:
            sh = wb.sheet(k[0])
            if sh is not None:
                r = sh.spill_ranges.get((k[1], k[2]))
                if r is not None:
                    seeds.extend((k[0], rr, cc) for i, (rr, cc) in enumerate(r.cells()) if i < 65536)
        for si, r in cleared:
            seeds.extend((si, rr, cc) for i, (rr, cc) in enumerate(r.cells()) if i < 65536)
        self._run(wb, seeds, full=False)

    def _closure(self, seeds: Iterable[Key]) -> Tuple[List[Key], Dict[Key, Set[Key]]]:
        """Dirty formulas reachable from ``seeds`` and the dependency edges among them."""
        g = self.graph
        seen: Set[Key] = set()
        edges: Dict[Key, Set[Key]] = {}
        queue: deque = deque()
        for k in seeds:
            if k in g.nodes and k not in seen:
                seen.add(k)
                queue.append(k)
            if k not in g.nodes:
                for d in g.dependents(*k):
                    if d not in seen:
                        seen.add(d)
                        queue.append(d)
        for k in g.dynamic:
            if k not in seen:
                seen.add(k)
                queue.append(k)
        while queue:
            k = queue.popleft()
            for d in g.dependents(*k):
                if d == k:
                    continue
                edges.setdefault(k, set()).add(d)
                if d not in seen:
                    seen.add(d)
                    queue.append(d)
        return list(seen), edges

    @staticmethod
    def _topo(nodes: List[Key], edges: Dict[Key, Set[Key]]) -> List[Key]:
        indeg: Dict[Key, int] = {k: 0 for k in nodes}
        for k, ds in edges.items():
            for d in ds:
                if d in indeg:
                    indeg[d] += 1
        heap = [k for k, n in indeg.items() if n == 0]
        heapq.heapify(heap)
        order: List[Key] = []
        while heap:
            k = heapq.heappop(heap)
            order.append(k)
            for d in edges.get(k, ()):
                if d in indeg:
                    indeg[d] -= 1
                    if indeg[d] == 0:
                        heapq.heappush(heap, d)
        if len(order) < len(nodes):
            done = set(order)
            order.extend(sorted(k for k in nodes if k not in done))
        return order

    def _run(self, wb, seeds: List[Key], full: bool, closure: bool = True) -> None:
        dynamic_again = bool(self.graph.dynamic)
        if closure:
            dirty, edges = self._closure(seeds)
        else:
            dirty, edges = list(seeds), {}
        for _round in range(8):
            if not dirty:
                break
            order = self._topo(dirty, edges)
            exprs: Dict[Key, Expr] = {}
            for k in order:
                e = _formula_expr(wb, k)
                if e is not None:
                    exprs[k] = e
            host = _PassHost(wb, exprs, self, self.now())
            host.iterate = wb.calc.iterative
            for k in order:
                if k in exprs:
                    host.compute(k)
            if host.iterate and host.members:
                host.iterate_loops(order, wb.calc.max_iterations, wb.calc.max_change)
            self.last_recalc_cells = len(host.results)
            self.circular = host.cycles
            # Write back.
            for k, v in host.results.items():
                sh = wb.sheet(k[0])
                if sh is None:
                    continue
                cell = sh.cell((k[1], k[2]))
                if cell is not None and not same_value(cell.value, v):
                    wb.sheet_mut(k[0]).cells.set((k[1], k[2]), cell.with_value(v))
            spill_changes: Set[Key] = set()
            freed: List[Tuple[int, RangeRef]] = []
            for k in order:
                new = host.spills.get(k)
                sh = wb.sheet(k[0])
                if sh is None:
                    continue
                anchor = (k[1], k[2])
                old = sh.spill_ranges.get(anchor)
                if old is None and new is None:
                    continue
                sm = wb.sheet_mut(k[0])
                if old is not None:
                    for rc in old.cells():
                        sm.spill.pop(rc, None)
                    del sm.spill_ranges[anchor]
                changed_range = old
                if new is not None:
                    end = (anchor[0] + new.rows - 1, anchor[1] + new.cols - 1)
                    rng = RangeRef(anchor[0], anchor[1], end[0], end[1])
                    blocked = (
                        end[0] >= MAX_ROWS
                        or end[1] >= MAX_COLS
                        or any(rc != anchor for rc, _ in sm.cells.iter_range(rng))
                        or any(rc in sm.spill for rc in rng.cells() if rc != anchor)
                        or any(m.intersects(rng) for m in sm.merges)
                        or any(t.range.contains(anchor) for t in sm.tables)
                    )
                    if blocked:
                        cell = sm.cell(anchor)
                        if cell is not None:
                            sm.cells.set(anchor, cell.with_value(CellError.SPILL))
                    else:
                        for i, rc in enumerate(rng.cells()):
                            if rc != anchor:
                                sm.spill[rc] = new.data[i]
                        sm.spill_ranges[anchor] = rng
                        changed_range = rng if changed_range is None else changed_range.union(rng)
                if changed_range is not None:
                    deps = self.graph.dependents_of_range(k[0], changed_range)
                    deps.discard(k)
                    spill_changes |= deps
                if old is not None:
                    nr = sm.spill_ranges.get(anchor)
                    if not (nr is not None and nr.contains(old.start) and nr.contains(old.end)):
                        freed.append((k[0], old))
            # Formulas with references only known while evaluating may have read a spill area
            # before it was laid out: they are evaluated once more.
            if dynamic_again and host.spills:
                dynamic_again = False
                spill_changes |= self.graph.dynamic
            for si, o in freed:
                for (s, r, c) in self.graph.nodes:
                    if s == si and r <= o.r1 and c <= o.c1:
                        sh = wb.sheet(s)
                        if sh is not None and sh.value((r, c)) is CellError.SPILL:
                            spill_changes.add((s, r, c))
            if not spill_changes:
                break
            dirty, edges = self._closure(sorted(spill_changes))


def evaluate(wb, sheet: int, at: Tuple[int, int], formula: str) -> Value:
    """Evaluates a formula in the context of a cell without storing it (conditional formats,
    data validation, the CLI ``eval``)."""
    try:
        expr = parse(formula[1:] if formula.startswith("=") else formula)
    except ParseError:
        return CellError.NAME
    return evaluate_expr(wb, sheet, at, expr)


def evaluate_expr(wb, sheet: int, at: Tuple[int, int], expr: Expr) -> Value:
    calc = Calc()
    host = _PassHost(wb, {}, calc, now_serial())
    return Evaluator(host, sheet, at).value(expr)
