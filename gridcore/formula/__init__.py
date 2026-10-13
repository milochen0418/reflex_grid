"""The Excel formula language: tokenizer, parser, syntax tree, printer (A1 and R1C1) and
reference adjustment for copy/fill and structural edits."""

from . import adjust
from .ast import (
    ArrayLit,
    Binary,
    Bool,
    Call,
    Err,
    Expr,
    Invoke,
    Missing,
    Name,
    Num,
    Paren,
    Ref,
    Str,
    StructRef,
    Unary,
)
from .lexer import LexError, tokenize
from .parser import ParseError, normalize_function_name, parse, parse_struct, try_parse
from .printer import quote_sheet, reference_a1, sheet_prefix, to_text


def print_formula(e: Expr) -> str:
    """Formula text (A1, without ``=``)."""
    return to_text(e)


def print_r1c1(e: Expr, at) -> str:
    return to_text(e, at)
