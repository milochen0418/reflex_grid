"""Every function survives hostile arguments."""

import random

from gridcore.core.value import Array, CellError
from gridcore.functions import Arg, Ctx, all_functions, call


class FuzzCtx(Ctx):
    def __init__(self):
        self.state = 0x2545F4914F6CDD1D

    def now_serial(self):
        return 46302.5

    def random(self):
        self.state = (self.state * 6364136223846793005 + 1) & 0xFFFFFFFFFFFFFFFF
        return (self.state >> 11) / float(1 << 53)


def test_hostile_inputs_never_crash():
    mixed = Array(2, 3, [1.0, "x", None, True, -1e300, 0.5])
    col = Array(3, 1, [3.0, 1.0, 2.0])
    patterns = [
        Arg(0.0), Arg(-1.0), Arg(1e308), Arg(-1e308), Arg(1e10), Arg(2.5), Arg(""), Arg("abc"), Arg("*?~"),
        Arg(True), Arg(CellError.NA), Arg(None), Arg(mixed), Arg(mixed, True), Arg(col, True), Arg(None, True),
    ]
    ctx = FuzzCtx()
    rng = random.Random(7)
    for spec in all_functions():
        mx = min(spec.max_args, spec.min_args + 3)
        for count in range(spec.min_args, mx + 1):
            for p in patterns:
                call(spec, [p] * count, ctx)
            args = [patterns[(i * 7 + count) % len(patterns)] for i in range(count)]
            call(spec, args, ctx)
            args = [rng.choice(patterns) for _ in range(count)]
            call(spec, args, ctx)
