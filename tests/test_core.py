"""Core tests: addresses, values, coercion, dates and typed input."""

from gridcore.core.addr import MAX_COLS, MAX_ROWS, RangeRef, clamp_cell, col_to_letters, letters_to_col, offset_cell, parse_cell
from gridcore.core.date import DateSystem, datetime_from_serial, days_in_month_in, serial_from_ymd
from gridcore.core.parse import parse_input
from gridcore.core.value import Array, CellError, XlError, compare, number, number_to_text, sort_compare, to_bool, to_number, to_text

S = DateSystem.D1900


def test_letters_roundtrip():
    for c, s in [(0, "A"), (25, "Z"), (26, "AA"), (51, "AZ"), (52, "BA"), (701, "ZZ"), (702, "AAA"), (16383, "XFD")]:
        assert col_to_letters(c) == s
        assert letters_to_col(s) == c
    assert letters_to_col("XFE") is None
    assert letters_to_col("") is None
    assert letters_to_col("A1") is None


def test_parse_cells():
    assert parse_cell("B7") == (6, 1)
    assert parse_cell("$b$7") == (6, 1)
    assert parse_cell("XFD1048576") == (MAX_ROWS - 1, MAX_COLS - 1)
    assert parse_cell("A0") is None
    assert parse_cell("A1048577") is None
    assert parse_cell("A1x") is None
    assert parse_cell("") is None
    assert parse_cell("$") is None


def test_ranges():
    r = RangeRef.parse("C3:A1")
    assert r.a1() == "A1:C3"
    assert r.count() == 9
    assert RangeRef.parse("A:C").a1() == "A:C"
    assert RangeRef.parse("3:5").a1() == "3:5"
    assert RangeRef.parse("0:5") is None
    a = RangeRef.parse("A1:C3")
    b = RangeRef.parse("B2:D4")
    assert a.intersection(b).a1() == "B2:C3"
    assert a.union(b).a1() == "A1:D4"
    assert len(list(a.cells())) == 9
    assert offset_cell((0, 0), -1, 0) is None
    assert clamp_cell(-5, 3) == (0, 3)


def test_number_text():
    assert number_to_text(1.0) == "1"
    assert number_to_text(-2.5) == "-2.5"
    assert number_to_text(0.1 + 0.2) == "0.3"
    assert number_to_text(123456789012345.0) == "123456789012345"
    assert number_to_text(1e15) == "1E+15"
    assert number_to_text(1234567890123456.0) == "1.23456789012346E+15"
    assert number_to_text(0.0001) == "0.0001"
    assert number_to_text(1e-10) == "1E-10"
    assert number_to_text(1.0 / 3.0) == "0.333333333333333"


def test_coercion():
    assert to_number("12") == 12.0
    assert to_number(" 1,234.5 ") == 1234.5
    assert to_number("50%") == 0.5
    try:
        to_number("abc")
        raise AssertionError
    except XlError as e:
        assert e.error is CellError.VALUE
    assert to_number(True) == 1.0
    assert to_text(None) == ""
    assert to_bool("true") is True
    assert number(float("nan")) is CellError.NUM


def test_comparisons():
    assert compare("abc", "ABC") == 0
    assert compare(5.0, "1") < 0
    assert compare("z", False) < 0
    assert compare(None, 0.0) == 0
    assert sort_compare(None, 1.0) > 0
    assert compare(0.1 + 0.2, 0.3) == 0
    assert compare(1.1 * 3.0, 3.3) == 0
    assert compare(1.0, 1.00000000000001) < 0
    for e in CellError:
        assert CellError.parse(e.value) is e
    assert CellError.parse("#div/0!") is CellError.DIV0


def test_arrays():
    a = Array.new(2, 2, [1.0, 2.0, 3.0, 4.0])
    assert a.transpose().get(0, 1) == 3.0
    assert Array.new(2, 2, []) is None
    row = Array.row([1.0, 2.0])
    assert row.get_broadcast(5, 1) == 2.0


def test_known_serials():
    assert serial_from_ymd(S, 1900, 1, 1) == 1.0
    assert serial_from_ymd(S, 1900, 2, 28) == 59.0
    assert serial_from_ymd(S, 1900, 3, 1) == 61.0
    assert serial_from_ymd(S, 2000, 1, 1) == 36526.0
    assert serial_from_ymd(S, 2026, 10, 7) == 46302.0
    assert serial_from_ymd(S, 9999, 12, 31) == 2958465.0
    assert serial_from_ymd(S, 2020, 13, 1) == serial_from_ymd(S, 2021, 1, 1)
    assert serial_from_ymd(S, 2020, 3, 0) == serial_from_ymd(S, 2020, 2, 29)
    assert serial_from_ymd(S, 120, 1, 1) == serial_from_ymd(S, 2020, 1, 1)
    assert serial_from_ymd(DateSystem.D1904, 1904, 1, 1) == 0.0
    assert serial_from_ymd(DateSystem.D1904, 2000, 1, 1) == 35064.0


def test_fictitious_1900_02_29():
    assert serial_from_ymd(S, 1900, 2, 29) == 60.0
    assert serial_from_ymd(S, 1900, 3, 0) == 60.0
    assert serial_from_ymd(S, 1900, 2, 30) == 61.0
    assert serial_from_ymd(S, 1900, 1, 60) == 60.0
    assert serial_from_ymd(S, 1900, 4, -30) == 61.0
    assert serial_from_ymd(S, 1900, 4, -31) == 60.0
    assert serial_from_ymd(S, 1901, -10, 29) == 60.0
    assert serial_from_ymd(S, 1900, 1, 0) == 0.0
    assert serial_from_ymd(S, 1900, 1, -1) is None
    assert serial_from_ymd(S, 1900, 0, 1) is None
    assert serial_from_ymd(S, 1904, 2, 29) == 1521.0
    for serial in range(0, 367):
        d = datetime_from_serial(S, float(serial))
        assert serial_from_ymd(S, d.year, d.month, d.day) == float(serial)
    assert days_in_month_in(S, 1900, 2) == 29
    assert days_in_month_in(DateSystem.D1904, 1900, 2) == 28
    s = DateSystem.D1904
    assert serial_from_ymd(s, 1904, 2, 29) == 59.0
    assert serial_from_ymd(s, 1904, 3, 1) == 60.0
    assert serial_from_ymd(s, 1904, 1, 0) is None


def test_roundtrip_serials():
    for serial in [1.0, 59.0, 60.0, 61.0, 36526.0, 46302.0, 2958465.0]:
        d = datetime_from_serial(S, serial)
        assert serial_from_ymd(S, d.year, d.month, d.day) == serial
    d = datetime_from_serial(S, 60.0)
    assert (d.year, d.month, d.day) == (1900, 2, 29)
    d = datetime_from_serial(S, 46302.75)
    assert (d.hour, d.minute, d.weekday) == (18, 0, 3)
    assert datetime_from_serial(S, 1.0).weekday == 0


def p(s):
    return parse_input(s, S)


def test_typed_numbers():
    assert p("42").value == 42.0
    assert p("-3.5").value == -3.5
    assert p("1,234").format == "#,##0"
    assert p("1,234").value == 1234.0
    assert p("1,23").value == "1,23"
    assert p("12%").value == 0.12
    assert p("$1,200.50").value == 1200.5
    assert p("$1,200.50").format == '"$"#,##0.00'
    assert p("(100)").value == -100.0
    assert p("1e3").value == 1000.0
    assert p("1 1/2").value == 1.5
    assert p(".5").value == 0.5


def test_typed_texts_and_specials():
    assert p("hello").value == "hello"
    assert p("'123").value == "123"
    assert p("true").value is True
    assert p("#N/A").value is CellError.NA
    assert p("e").value == "e"
    assert p("-").value == "-"
    assert p("1.2.3.4").value == "1.2.3.4"
    assert p("1_000").value == "1_000"


def test_typed_dates_and_times():
    assert p("10/7/2026").value == 46302.0
    assert p("2026-10-07").value == 46302.0
    assert p("7-Oct-2026").value == 46302.0
    assert p("Oct 7, 2026").value == 46302.0
    assert p("2/30/2026").value == "2/30/2026"
    assert p("7/4/1776").value == "7/4/1776"
    assert p("12/31/1899").value == "12/31/1899"
    assert p("1/1/1900").value == 1.0
    assert p("12:00").value == 0.5
    assert p("6:00 PM").value == 0.75
    assert p("12:00 AM").value == 0.0
    assert p("10/7/2026 18:00").value == 46302.75
    assert p("25:00").value == 25.0 / 24.0
    assert p("25:00").format == "[h]:mm:ss"
    assert p("23:59").format == "h:mm"
    assert p("2/29/1900").value == 60.0
    assert parse_input("2/29/1904", DateSystem.D1904).value == 59.0
