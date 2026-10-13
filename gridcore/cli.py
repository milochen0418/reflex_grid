"""Command-line access to the engine.

    python -m gridcore.cli eval "=SUM(1,2)*2"
    python -m gridcore.cli cat book.xlsx [--sheet NAME] [--formulas] [--range A1:D20]
    python -m gridcore.cli run book.xlsx script.json [-o out.xlsx]
    python -m gridcore.cli convert book.xlsx book.csv
    python -m gridcore.cli commands [--json]
    python -m gridcore.cli functions [search]

A script is a JSON list of commands: ``[{"id": "cell.set", "params": {"cell": "A1", "input": "5"}}, …]``
(or ``["cell.set", {...}]`` pairs). Every command the UI uses is available.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, List

from .core.addr import RangeRef
from .core.value import Array, display


def _session():
    from .engine import Session

    return Session()


def cmd_eval(args) -> int:
    from .calc import evaluate
    from .model.workbook import Workbook

    wb = Workbook()
    if args.file:
        s = _session()
        s.run("file.open", path=args.file)
        wb = s.doc().wb
    v = evaluate(wb, wb.active_sheet, (0, 0), args.formula)
    if isinstance(v, Array):
        for r in range(v.rows):
            print("\t".join(display(v.get(r, c)) for c in range(v.cols)))
    else:
        print(display(v))
    return 0


def cmd_cat(args) -> int:
    from .engine.display import cell_text

    s = _session()
    s.run("file.open", path=args.file)
    d = s.doc()
    si = d.wb.active_sheet
    if args.sheet:
        i = d.wb.sheet_index(args.sheet)
        if i is None:
            print(f"no sheet named {args.sheet!r}; sheets: {[x.name for x in d.wb.sheets]}", file=sys.stderr)
            return 2
        si = i
    sh = d.wb.sheets[si]
    rng = RangeRef.parse(args.range) if args.range else sh.used_range()
    if rng is None:
        return 0
    for r in range(rng.r0, rng.r1 + 1):
        row = []
        for c in range(rng.c0, rng.c1 + 1):
            cell = sh.cell((r, c))
            if args.formulas and cell is not None and cell.formula is not None:
                row.append("=" + cell.formula.text)
            else:
                row.append(cell_text(d.wb, sh, (r, c)))
        print("\t".join(row))
    for w in d.warnings:
        print(f"warning: {w}", file=sys.stderr)
    return 0


def _script(path: str) -> List[Any]:
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, list):
        raise SystemExit("a script is a JSON list of commands")
    return data


def cmd_run(args) -> int:
    from .engine import EngineError

    s = _session()
    if args.file and os.path.exists(args.file):
        s.run("file.open", path=args.file)
    else:
        s.new_workbook()
    for i, step in enumerate(_script(args.script)):
        if isinstance(step, dict):
            cid, params = step.get("id"), step.get("params") or {}
        else:
            cid, params = step[0], (step[1] if len(step) > 1 else {})
        try:
            r = s.execute(str(cid), dict(params))
        except EngineError as e:
            print(f"step {i + 1} ({cid}) failed: {str(e).splitlines()[0]}", file=sys.stderr)
            return 1
        if r is not None and args.verbose:
            print(json.dumps({"step": i + 1, "id": cid, "result": r}, default=str))
    out = args.output or args.file
    if out:
        s.run("file.saveAs", path=out)
        print(f"saved {out}")
    return 0


def cmd_convert(args) -> int:
    s = _session()
    s.run("file.open", path=args.input)
    s.run("file.saveAs", path=args.output)
    print(f"{args.input} -> {args.output}")
    return 0


def cmd_commands(args) -> int:
    from .engine import COMMANDS

    if args.json:
        print(json.dumps([{"id": c.id, "label": c.label, "menu": list(c.menu), "shortcut": c.shortcut, "params": c.params} for c in COMMANDS.values()], indent=1))
        return 0
    for c in sorted(COMMANDS.values(), key=lambda c: c.id):
        where = " › ".join(c.menu)
        print(f"{c.id:32} {c.label:34} {c.shortcut or '':14} {where}")
    return 0


def cmd_functions(args) -> int:
    from .functions import all_functions

    q = (args.search or "").lower()
    for f in sorted(all_functions(), key=lambda f: f.name):
        if q and q not in f.name.lower() and q not in f.description.lower():
            continue
        print(f"{f.signature:60} {f.description}")
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="gridcore", description="ReflexGrid's spreadsheet engine from the command line.")
    sub = p.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("eval", help="evaluate a formula")
    e.add_argument("formula")
    e.add_argument("--file", help="evaluate in the context of a workbook")
    e.set_defaults(fn=cmd_eval)
    c = sub.add_parser("cat", help="print a sheet as tab-separated text")
    c.add_argument("file")
    c.add_argument("--sheet")
    c.add_argument("--range")
    c.add_argument("--formulas", action="store_true")
    c.set_defaults(fn=cmd_cat)
    r = sub.add_parser("run", help="run a JSON command script on a workbook")
    r.add_argument("file", help="workbook to open (created when missing)")
    r.add_argument("script")
    r.add_argument("-o", "--output")
    r.add_argument("-v", "--verbose", action="store_true")
    r.set_defaults(fn=cmd_run)
    v = sub.add_parser("convert", help="convert between .xlsx, .csv and .tsv")
    v.add_argument("input")
    v.add_argument("output")
    v.set_defaults(fn=cmd_convert)
    m = sub.add_parser("commands", help="list engine commands")
    m.add_argument("--json", action="store_true")
    m.set_defaults(fn=cmd_commands)
    f = sub.add_parser("functions", help="list worksheet functions")
    f.add_argument("search", nargs="?")
    f.set_defaults(fn=cmd_functions)
    args = p.parse_args(argv)
    try:
        return args.fn(args)
    except Exception as ex:  # a clean message instead of a traceback
        from .engine import EngineError

        if isinstance(ex, EngineError):
            print(str(ex).splitlines()[0], file=sys.stderr)
            return 1
        raise


if __name__ == "__main__":
    sys.exit(main())
