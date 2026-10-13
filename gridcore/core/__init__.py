"""Core types: addresses, values and Excel's coercion rules, error values,
date serials and typed-input parsing. No dependencies on the rest of the package."""

from .addr import (
    MAX_COLS,
    MAX_ROWS,
    RangeRef,
    a1,
    clamp_cell,
    col_to_letters,
    letters_to_col,
    offset_cell,
    parse_a1_prefix,
    parse_cell,
    parse_ranges,
)
from .date import DateSystem
from .value import (
    Array,
    CellError,
    XlError,
    compare,
    compare_numbers,
    compare_text,
    display,
    is_error,
    is_number,
    normalize,
    number,
    number_to_text,
    scalar,
    sort_compare,
    to_bool,
    to_number,
    to_text,
)
