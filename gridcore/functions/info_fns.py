"""Information and logical functions."""

from __future__ import annotations

import math
import platform

from ..core.value import Array, CellError, Value, XlError, scalar, to_bool, to_number
from .registry import A, Arg, Ctx, S, VAR, fn
from .util import as_array, has, num, text, to_int

I = "Information"
L = "Logical"


def _v(args, i=0) -> Value:
    return scalar(args[i].value)


@fn("ISBLANK", 1, 1, I, S, "ISBLANK(value)", "Returns TRUE if the value is blank.")
def _isblank(args, ctx):
    return _v(args) is None and args[0].from_ref


@fn("ISERR", 1, 1, I, S, "ISERR(value)", "Returns TRUE if the value is any error value except #N/A.")
def _iserr(args, ctx):
    v = _v(args)
    return isinstance(v, CellError) and v is not CellError.NA


@fn("ISERROR", 1, 1, I, S, "ISERROR(value)", "Returns TRUE if the value is any error value.")
def _iserror(args, ctx):
    return isinstance(_v(args), CellError)


@fn("ISNA", 1, 1, I, S, "ISNA(value)", "Returns TRUE if the value is the #N/A error value.")
def _isna(args, ctx):
    return _v(args) is CellError.NA


@fn("ISNUMBER", 1, 1, I, S, "ISNUMBER(value)", "Returns TRUE if the value is a number.")
def _isnumber(args, ctx):
    return type(_v(args)) is float


@fn("ISTEXT", 1, 1, I, S, "ISTEXT(value)", "Returns TRUE if the value is text.")
def _istext(args, ctx):
    return isinstance(_v(args), str)


@fn("ISNONTEXT", 1, 1, I, S, "ISNONTEXT(value)", "Returns TRUE if the value is not text.")
def _isnontext(args, ctx):
    return not isinstance(_v(args), str)


@fn("ISLOGICAL", 1, 1, I, S, "ISLOGICAL(value)", "Returns TRUE if the value is a logical value.")
def _islogical(args, ctx):
    return isinstance(_v(args), bool)


@fn("ISEVEN", 1, 1, I, S, "ISEVEN(number)", "Returns TRUE if the number is even.")
def _iseven(args, ctx):
    v = _v(args)
    if isinstance(v, bool):
        raise XlError(CellError.VALUE)
    return math.trunc(to_number(v)) % 2 == 0


@fn("ISODD", 1, 1, I, S, "ISODD(number)", "Returns TRUE if the number is odd.")
def _isodd(args, ctx):
    v = _v(args)
    if isinstance(v, bool):
        raise XlError(CellError.VALUE)
    return math.trunc(to_number(v)) % 2 == 1


@fn("NA", 0, 0, I, S, "NA()", "Returns the error value #N/A.")
def _na(args, ctx):
    return CellError.NA


@fn("N", 1, 1, I, S, "N(value)", "Returns a value converted to a number.")
def _n(args, ctx):
    v = _v(args)
    if isinstance(v, CellError):
        raise XlError(v)
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if type(v) is float:
        return v
    return 0.0


@fn("TYPE", 1, 1, I, A, "TYPE(value)", "Returns a number indicating the data type of a value.")
def _type(args, ctx):
    v = args[0].value
    if isinstance(v, Array):
        if args[0].from_ref and v.rows * v.cols == 1:
            v = v.data[0]
        else:
            return 64.0
    if v is None or type(v) is float:
        return 1.0
    if isinstance(v, str):
        return 2.0
    if isinstance(v, bool):
        return 4.0
    if isinstance(v, CellError):
        return 16.0
    return 1.0


@fn("ERROR.TYPE", 1, 1, I, S, "ERROR.TYPE(error_val)", "Returns a number corresponding to an error type.")
def _error_type(args, ctx):
    v = _v(args)
    if isinstance(v, CellError):
        return float(v.code)
    raise XlError(CellError.NA)


@fn("INFO", 1, 1, I, S, "INFO(type_text)", "Returns information about the current operating environment.", volatile=True)
def _info(args, ctx):
    t = text(args, 0).lower()
    if t == "osversion":
        return f"{platform.system()} {platform.release()}"
    if t == "release":
        return "16.0"
    if t == "system":
        return "mac" if platform.system() == "Darwin" else "pcdos"
    if t == "recalc":
        return "Automatic"
    if t == "numfile":
        return 1.0
    if t == "origin":
        return "$A:$A$1"
    if t == "directory":
        return ""
    raise XlError(CellError.VALUE)


def _logicals(args) -> list:
    out = []
    for a in args:
        v = a.value
        if isinstance(v, Array):
            for x in v.data:
                if isinstance(x, CellError):
                    raise XlError(x)
                if isinstance(x, bool):
                    out.append(x)
                elif type(x) is float:
                    out.append(x != 0)
            continue
        if isinstance(v, CellError):
            raise XlError(v)
        if v is None:
            if not a.from_ref:
                out.append(False)
            continue
        if isinstance(v, str):
            if a.from_ref:
                continue
            out.append(to_bool(v))
            continue
        out.append(to_bool(v))
    if not out:
        raise XlError(CellError.VALUE)
    return out


@fn("AND", 1, VAR, L, A, "AND(logical1, [logical2], …)", "Returns TRUE if all of its arguments are TRUE.")
def _and(args, ctx):
    return all(_logicals(args))


@fn("OR", 1, VAR, L, A, "OR(logical1, [logical2], …)", "Returns TRUE if any argument is TRUE.")
def _or(args, ctx):
    return any(_logicals(args))


@fn("XOR", 1, VAR, L, A, "XOR(logical1, [logical2], …)", "Returns a logical exclusive OR of all arguments.")
def _xor(args, ctx):
    return sum(1 for b in _logicals(args) if b) % 2 == 1


@fn("NOT", 1, 1, L, S, "NOT(logical)", "Reverses the logic of its argument.")
def _not(args, ctx):
    return not to_bool(_v(args))


@fn("TRUE", 0, 0, L, S, "TRUE()", "Returns the logical value TRUE.")
def _true(args, ctx):
    return True


@fn("FALSE", 0, 0, L, S, "FALSE()", "Returns the logical value FALSE.")
def _false(args, ctx):
    return False


@fn("CHOOSE", 2, VAR, "Lookup & Reference", (True, False), "CHOOSE(index_num, value1, [value2], …)", "Chooses a value from a list of values.")
def _choose(args, ctx):
    i = to_int(num(args, 0))
    if i < 1 or i >= len(args):
        raise XlError(CellError.VALUE)
    return args[i].value


# The evaluator implements these special forms itself; registering them gives the function
# list (Insert Function, autocomplete) their signatures and descriptions.
_SPECIAL_DOCS = [
    ("IF", 2, 3, L, "IF(logical_test, [value_if_true], [value_if_false])", "Specifies a logical test to perform."),
    ("IFS", 2, VAR, L, "IFS(logical_test1, value_if_true1, …)", "Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition."),
    ("IFERROR", 2, 2, L, "IFERROR(value, value_if_error)", "Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula."),
    ("IFNA", 2, 2, L, "IFNA(value, value_if_na)", "Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression."),
    ("SWITCH", 3, VAR, L, "SWITCH(expression, value1, result1, [default_or_value2, result2], …)", "Evaluates an expression against a list of values and returns the result corresponding to the first matching value."),
    ("LET", 3, VAR, L, "LET(name1, name_value1, calculation_or_name2, …)", "Assigns names to calculation results."),
    ("LAMBDA", 1, VAR, L, "LAMBDA([parameter1, parameter2, …], calculation)", "Create custom, reusable functions and call them by a friendly name."),
    ("MAP", 2, VAR, L, "MAP(array1, lambda_or_array<#>)", "Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA."),
    ("REDUCE", 3, 3, L, "REDUCE([initial_value], array, lambda)", "Reduces an array to an accumulated value by applying a LAMBDA to each value."),
    ("SCAN", 3, 3, L, "SCAN([initial_value], array, lambda)", "Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value."),
    ("BYROW", 2, 2, L, "BYROW(array, lambda)", "Applies a LAMBDA to each row and returns an array of the results."),
    ("BYCOL", 2, 2, L, "BYCOL(array, lambda)", "Applies a LAMBDA to each column and returns an array of the results."),
    ("MAKEARRAY", 3, 3, L, "MAKEARRAY(rows, cols, lambda)", "Returns a calculated array of a specified row and column size, by applying a LAMBDA."),
    ("ISOMITTED", 1, 1, I, "ISOMITTED(argument)", "Checks whether the value in a LAMBDA is missing and returns TRUE or FALSE."),
    ("ROW", 0, 1, "Lookup & Reference", "ROW([reference])", "Returns the row number of a reference."),
    ("COLUMN", 0, 1, "Lookup & Reference", "COLUMN([reference])", "Returns the column number of a reference."),
    ("ROWS", 1, 1, "Lookup & Reference", "ROWS(array)", "Returns the number of rows in a reference."),
    ("COLUMNS", 1, 1, "Lookup & Reference", "COLUMNS(array)", "Returns the number of columns in a reference."),
    ("AREAS", 1, 1, "Lookup & Reference", "AREAS(reference)", "Returns the number of areas in a reference."),
    ("ISREF", 1, 1, I, "ISREF(value)", "Returns TRUE if the value is a reference."),
    ("ISFORMULA", 1, 1, I, "ISFORMULA(reference)", "Returns TRUE if there is a reference to a cell that contains a formula."),
    ("FORMULATEXT", 1, 1, "Lookup & Reference", "FORMULATEXT(reference)", "Returns the formula at the given reference as text."),
    ("SHEET", 0, 1, I, "SHEET([value])", "Returns the sheet number of the referenced sheet."),
    ("SHEETS", 0, 1, I, "SHEETS([reference])", "Returns the number of sheets in a reference."),
    ("OFFSET", 3, 5, "Lookup & Reference", "OFFSET(reference, rows, cols, [height], [width])", "Returns a reference offset from a given reference."),
    ("INDIRECT", 1, 2, "Lookup & Reference", "INDIRECT(ref_text, [a1])", "Returns a reference indicated by a text value."),
    ("INDEX", 2, 4, "Lookup & Reference", "INDEX(array, row_num, [column_num], [area_num])", "Uses an index to choose a value from a reference or array."),
    ("CELL", 1, 2, I, "CELL(info_type, [reference])", "Returns information about the formatting, location, or contents of a cell."),
    ("TRIMRANGE", 1, 3, "Lookup & Reference", "TRIMRANGE(range, [trim_rows], [trim_cols])", "Scans in from the edges of a range or array until it finds a non-blank cell or value, it then excludes those blank rows or columns."),
    ("SUBTOTAL", 2, VAR, "Math & Trig", "SUBTOTAL(function_num, ref1, [ref2], …)", "Returns a subtotal in a list or database."),
    ("AGGREGATE", 3, VAR, "Math & Trig", "AGGREGATE(function_num, options, ref1, [ref2], …)", "Returns an aggregate in a list or database."),
    ("TEXT", 2, 2, "Text", "TEXT(value, format_text)", "Formats a number and converts it to text."),
]


def _special_placeholder(args, ctx):  # pragma: no cover - evaluated by calc.eval
    return CellError.VALUE


for _name, _lo, _hi, _cat, _sig, _desc in _SPECIAL_DOCS:
    fn(_name, _lo, _hi, _cat, A, _sig, _desc)(_special_placeholder)
