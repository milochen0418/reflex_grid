"""Calculation: the formula evaluator and dependency-graph recalculation with dynamic arrays."""

from .. import functions as _functions  # noqa: F401  (registers the function library)
from .eval import SPECIAL_FUNCTIONS, Area, Evaluator, Host, RefVal, is_known_function, precedents
from .recalc import Calc, Graph, evaluate, evaluate_expr
