"""Number format tests (Excel-observed results)."""

import math
import random

from gridcore.core.date import DateSystem
from gridcore.core.value import CellError, XlError
from gridcore.numfmt import builtin_format, builtin_id, format_general_fit, format_value, parse, text_function
from gridcore.numfmt.decimal import decimal_digits, fixed

S = DateSystem.D1900


def f(code, n):
    return format_value(float(n), parse(code), S).text


def full(code, v):
    return format_value(v, parse(code), S)


def t(code, s):
    return format_value(s, parse(code), S).text


def test_numbers():
    assert f("0.00", 1234.567) == "1234.57"
    assert f("#,##0", 1234567.0) == "1,234,567"
    assert f("#,##0", -1234567.89) == "-1,234,568"
    assert f("#,##0.00", 0.0) == "0.00"
    assert f("0", 0.5) == "1"
    assert f("0", -2.5) == "-3"
    assert f("0.00", 2.675) == "2.68"
    assert f("0.00", 1.005) == "1.01"
    assert f("0.0", -0.04) == "0.0"  # rounds to zero: no minus sign (Excel)
    assert f("00000", 123.0) == "00123"
    assert f("000-00-0000", 123456789.0) == "123-45-6789"
    assert f("#.##", 0.5) == ".5"
    assert f("0.##", 1.0) == "1."
    assert f("#", 0.0) == ""
    assert f("0.0?", 1.5) == "1.5 "
    assert f("??0.0", 1.5) == "  1.5"
    assert f(".00", 12.5) == "12.50"
    assert f("0", 1e20) == "100000000000000000000"
    assert f("0.0", 1e-20) == "0.0"
    assert f("0,000", 5.0) == "0,005"
    assert f("#,##0.000", 1234.5) == "1,234.500"


def test_scaling_and_percent():
    assert f("#,##0,", 1234567.0) == "1,235"
    assert f("0.0,,", 12345678.0) == "12.3"
    assert f('0,"K"', 12345.0) == "12K"
    assert f("0%", 0.256) == "26%"
    assert f("0.00%", 0.12345) == "12.35%"
    assert f("0%%", 0.5) == "5000%%"
    assert f("0%", -0.5) == "-50%"
    assert parse("0.0%").is_percent()
    assert not parse("0.0").is_percent()


def test_sections_and_colors():
    r = full("#,##0.00;[Red](#,##0.00)", -5.0)
    assert r.text == "(5.00)" and r.color == "red" and r.numeric
    r = full("#,##0.00;[Red](#,##0.00)", 5.0)
    assert (r.text, r.color) == ("5.00", None)
    assert f("0;-0;;@", 0.0) == ""
    assert t("0;-0;;@", "hi") == "hi"
    assert f('0;-0;"zero"', 0.0) == "zero"
    assert f('0;-0;"zero"', 0.001) == "0"
    assert f("0_);(0)", 5.0) == "5 "
    assert f("0_);(0)", -5.0) == "(5)"
    assert f(";;;", 5.0) == ""
    assert t(";;;", "x") == ""
    assert full("[Color10]0", 1.0).color == "color10"
    assert full("[blue]0", 1.0).color == "blue"
    assert full("[Color99]0", 1.0).color is None


def test_conditions():
    c = '[>=100]"big";[<0]"neg";"small"'
    assert f(c, 150.0) == "big"
    assert f(c, -5.0) == "neg"
    assert f(c, 50.0) == "small"
    assert f("[<1]0.00;0", 0.5) == "0.50"
    assert f("[<1]0.00;0", 5.0) == "5"
    assert f("[<1]0.00;0", -5.0) == "5.00"
    assert f('[>=1000]0,"K";0', -5.0) == "5"
    assert f('[>=1000]0,"K"', -5.0) == "-5"
    phone = "[<=9999999]###-####;(###) ###-####"
    assert f(phone, 5551234.0) == "555-1234"
    assert f(phone, 2125551234.0) == "(212) 555-1234"
    assert full("[Red][<=100]0;[Blue][>100]0", 150.0).color == "blue"


def test_scientific():
    assert f("0.00E+00", 12345.0) == "1.23E+04"
    assert f("0.00E+00", 0.000123) == "1.23E-04"
    assert f("0.00E+00", 0.0) == "0.00E+00"
    assert f("0.00E-00", 12345.0) == "1.23E04"
    assert f("0.00e+00", 12345.0) == "1.23e+04"
    assert f("##0.0E+0", 12345.0) == "12.3E+3"
    assert f("0.0E+0", 0.00001) == "1.0E-5"
    assert f("0.000E+00", 1e-300) == "1.000E-300"
    assert f("0.00E+00", 9.999) == "1.00E+01"
    assert f("0.00E+00", -12345.0) == "-1.23E+04"


def test_fractions():
    assert f("# ?/?", 1.5) == "1 1/2"
    assert f("# ?/?", 0.25) == " 1/4"
    assert f("# ?/?", 2.0) == "2    "
    assert f("# ?/?", 0.0) == "0    "
    assert f("# ?/?", -1.5) == "-1 1/2"
    assert f("# ??/??", math.pi) == "3 14/99"
    assert f("# ??/??", 1.0 / 3.0) == "  1/3 "
    assert f("?/8", 0.5) == "4/8"
    assert f("# ?/8", 2.125) == "2 1/8"
    assert f("?/?", 1.5) == "3/2"
    assert f("# ?/?", 0.99) == "1    "
    assert f("# ???/???", 0.123) == " 100/813"
    assert f("# ???/???", 0.5) == "   1/2  "


def test_fractions_with_literals():
    cases = [
        (r"|#\:#/#|", 23.75, "|23:3/4|"),
        (r"|#\:#=/=#|", 23.75, "|23:3=/=4|"),
        (r"|#\:#=/=#|", -23.75, "-|23:3=/=4|"),
        (r"|#\:? ?=/=#|", 23.75, "|2:3 3=/=4|"),
        (r"|#-#-#\:#/#|", 23.75, "|-2-3:3/4|"),
        (r"|#\:#/#|", 0.75, "|3/4|"),
        (r"|#\:#/#|", -0.75, "-|3/4|"),
        (r"|#-#-#\:#/#|", 0.75, "|--3/4|"),
        (r"|#\:#=/=#|", 0.75, "|3=/=4|"),
        (r"|#\:0=/=?|", 0.75, "|3=/=4|"),
        (r"|#\:?=/=#|", 0.75, "| 3=/=4|"),
        (r"|?\:#=/=#|", 0.75, "|  3=/=4|"),
        (r"|?\:0=/=0|", -0.75, "-|  3=/=4|"),
        (r"|?\:?0=/=#|", 0.75, "|   3=/=4|"),
        (r"|0\:#=/=#|", 0.75, "|0:3=/=4|"),
        (r"|#\:#/#|", 0.0, "|0|"),
        (r"|#\:#/#|", 1.0, "|1|"),
        (r"|#\:#/#|", -1.0, "-|1|"),
        (r"|#-#-#\:#/#|", 0.0, "|--0|"),
        (r"|#-#-#\:#/#|", 1.0, "|--1|"),
        (r"|#\:#=/=0|", 0.0, "|0|"),
        (r"|0\:#=/=#|", 1.0, "|1|"),
        (r"|#\:#=/=?|", 0.0, "|0      |"),
        (r"|#\:?=/=0|", 1.0, "|1      |"),
        (r"|?\:#=/=#|", 0.0, "|0      |"),
        (r"|?\:#=/=#|", -1.0, "-|1      |"),
        (r"|#\:? ?=/=#|", 0.0, "|:0      |"),
        (r"|#\:? ?=/=?|", 1.0, "|:1      |"),
        (r"|#\:0=/=#|", 0.0, "|0=/=1|"),
        (r"|?\:0=/=?|", 0.0, "|0:0=/=1|"),
        (r"|0\:0=/=0|", 0.0, "|0:0=/=1|"),
        (r"|#\:0=/=#|", 1.0, "|1:0=/=1|"),
        (r"|?\:0=/=#|", -1.0, "-|1:0=/=1|"),
        (r"|#\:? ?0#/000", 0.0, "|:0  00/001"),
        (r"|#\:? ?0#/000", 1.0, "|:1  00/001"),
        (r"|#\:? ?0#/000", 23.75, "|2:3  03/004"),
        (r"|#\:? ?#/#|", 23.75, "|2:3  3/4|"),
        (r"|#_#/#|", 3.75, "|15 /4|"),
        (r"|#_#/#|", -0.75, "-|3 /4|"),
        (r"|#_#/#|", 0.0, "|0 /1|"),
        (r"|#_#/#|", -1.0, "-|1 /1|"),
        (r"|#_#=/=?|", 3.75, "|15 =/=4|"),
        (r"|#_?=/=0|", 0.0, "|0 =/=1|"),
        (r"|#_?=/=?|", 1.0, "|1 =/=1|"),
    ]
    for code, v, want in cases:
        assert f(code, v) == want, (code, v)


def test_question_mark_pads_thousands_separator():
    cases = [
        ("?,???????", 1234567.0, " 1,234,567"),
        ("?,????????", 1234567.0, "  1,234,567"),
        ("?,?????????", 1234567.0, "    1,234,567"),
        ("?,?????????", -1234567.0, "-    1,234,567"),
        ("?,??????", 1234567.0, "1,234,567"),
        ("?,????", 12345.0, "12,345"),
        ("0,000000000", 1234567.0, "0,001,234,567"),
    ]
    for code, v, want in cases:
        assert f(code, v) == want, (code, v)


def test_literals_and_fill():
    assert f('"$"#,##0.00', 3.0) == "$3.00"
    assert f("$#,##0.00", -3.0) == "-$3.00"
    assert f("[$€-407]#,##0.00", 1234.5) == "€1,234.50"
    assert f('0" units"', 5.0) == "5 units"
    assert f("\\A0", 5.0) == "A5"
    assert f("(0)", 5.0) == "(5)"
    r = full("*-0", 5.0)
    assert (r.text, r.fill) == ("5", ("-", 0))
    acct = builtin_format(44)
    r = full(acct, 0.0)
    assert (r.text, r.fill) == (" $-   ", (" ", 2))
    r = full(acct, 1234.5)
    assert (r.text, r.fill) == (" $1,234.50 ", (" ", 2))
    r = full(acct, -1234.5)
    assert (r.text, r.fill) == (" $(1,234.50)", (" ", 2))
    r = full(acct, "abc")
    assert (r.text, r.numeric) == (" abc ", False)
    r = full("-* 0", -5.0)
    assert (r.text, r.fill) == ("--5", (" ", 2))


def test_dates_and_times():
    assert f("m/d/yyyy", 46302.0) == "10/7/2026"
    assert f("dddd, mmmm d, yyyy", 46302.0) == "Wednesday, October 7, 2026"
    assert f("ddd mmm yy", 46302.0) == "Wed Oct 26"
    assert f("mmmmm", 46302.0) == "O"
    assert f("dd/mm/yyyy", 46302.0) == "07/10/2026"
    assert f("d-mmm-yy", 46302.0) == "7-Oct-26"
    assert f("e", 46302.0) == "2026"
    assert f("YYYY-MM-DD", 46302.0) == "2026-10-07"
    assert f("yyyy-mm-dd hh:mm:ss", 46302.5) == "2026-10-07 12:00:00"
    assert f("h:mm AM/PM", 0.75) == "6:00 PM"
    assert f("h AM/PM", 0.0) == "12 AM"
    assert f("h a/p", 0.25) == "6 a"
    assert f("hh:mm", 0.75) == "18:00"
    assert f("[h]:mm", 1.5) == "36:00"
    assert f("[h]:mm:ss", 2.25) == "54:00:00"
    assert f("[mm]:ss", 1.0 / 24.0) == "60:00"
    assert f("[s]", 1.0 / 1440.0) == "60"
    assert f("mm:ss.0", 61.2 / 86400.0) == "01:01.2"
    assert f("mm:ss.00", 61.25 / 86400.0) == "01:01.25"
    assert f("h:mm:ss", 0.5 + 59.6 / 86400.0) == "12:01:00"
    assert f("h:mm", 0.5 + 40.0 / 86400.0) == "12:00"
    assert f("m/d/yyyy", 60.0) == "2/29/1900"
    assert f("m/d/yyyy", 0.0) == "1/0/1900"
    assert format_value(0.0, parse("m/d/yyyy"), DateSystem.D1904).text == "1/1/1904"
    assert f("m/d/yyyy", -1.0) == "########"
    assert f("m/d/yyyy", 3e6) == "########"
    assert f("[$-409]m/d/yyyy", 46302.0) == "10/7/2026"
    assert f("m/d/yyyy h:mm", 46302.75) == "10/7/2026 18:00"
    assert f("mmss.0", 0.0) == "0000.0"
    assert parse("h:mm").is_date()
    assert not parse("0.00").is_date()


def test_text_and_other_values():
    assert t("@", "hello") == "hello"
    assert t('"<"@">"', "x") == "<x>"
    assert t("0.00", "abc") == "abc"
    assert f("@", 5.0) == "5"
    assert parse("@").is_text()
    assert not parse("0;@").is_text()
    g = parse("General")
    assert format_value(True, g, S).text == "TRUE"
    assert format_value(False, parse("0.00"), S).text == "FALSE"
    assert format_value(CellError.DIV0, g, S).text == "#DIV/0!"
    assert format_value(None, g, S).text == ""
    assert not format_value("x", g, S).numeric


def test_general():
    assert f("General", 1e11) == "1E+11"
    assert f("General", 12345678901.0) == "12345678901"
    assert f("General", 0.1 + 0.2) == "0.3"
    assert f("General", -1234.5) == "-1234.5"
    assert f("", 1.0 / 3.0) == "0.333333333"
    assert f("general", 123456.789012) == "123456.789"
    assert f('General" units"', 5.0) == "5 units"
    assert f("General;[Red]-General", -5.0) == "-5"
    assert parse("General").is_general()
    assert parse("0.0").code == "0.0"


def test_kinds():
    assert parse("General").kind() == "General"
    assert parse("0.00").kind() == "Number"
    assert parse("#,##0.00;[Red](#,##0.00)").kind() == "Number"
    assert parse('"$"#,##0.00').kind() == "Currency"
    assert parse(builtin_format(44)).kind() == "Accounting"
    assert parse("m/d/yyyy").kind() == "Date"
    assert parse("h:mm AM/PM").kind() == "Time"
    assert parse("0%").kind() == "Percentage"
    assert parse("# ?/?").kind() == "Fraction"
    assert parse("0.00E+00").kind() == "Scientific"
    assert parse("@").kind() == "Text"
    assert parse("00000").kind() == "Special"
    assert parse('"hello"').kind() == "Custom"


def test_builtins():
    assert builtin_format(0) == "General"
    assert builtin_format(14) == "m/d/yyyy"
    assert builtin_format(49) == "@"
    assert builtin_format(30) is None
    assert builtin_format(50) is None
    assert builtin_id("m/d/yyyy") == 14
    assert builtin_id("0.00%") == 10
    assert builtin_id("GENERAL") == 0
    assert builtin_id("[h]:mm:ss") == 46
    assert builtin_id("#,##0.00_);(#,##0.00)") is None
    for i in range(61):
        code = builtin_format(i)
        if code is not None:
            assert builtin_id(code) == i
    assert f(builtin_format(5), -1234.0) == "($1,234)"
    assert f(builtin_format(12), 0.5) == " 1/2"
    assert f(builtin_format(47), 61.2 / 86400.0) == "0101.2"


def test_text_fn():
    assert text_function(1234.5, "$#,##0.00") == "$1,234.50"
    assert text_function("5", "0.00") == "5.00"
    assert text_function("abc", "0.00") == "abc"
    try:
        text_function(CellError.NA, "0")
        raise AssertionError("expected error")
    except XlError as e:
        assert e.error is CellError.NA
    assert text_function(46302.0, "yyyy") == "2026"
    assert text_function(0.5, "General") == "0.5"
    assert text_function(1.0 / 3.0, "0.000") == "0.333"
    assert text_function(None, "0.00") == "0.00"
    assert text_function(5.0, "") == ""
    try:
        text_function(-1.0, "m/d/yyyy")
        raise AssertionError("expected error")
    except XlError as e:
        assert e.error is CellError.VALUE
    assert text_function(True, "0") == "TRUE"
    cases = [
        (True, ";;;@", "TRUE"),
        (True, ';;;"hi"', "hi"),
        (True, ';;;"h""I"', "hI"),
        (False, ";;;\\h\\i", "hi"),
        (True, ";;;\\@@\\@", "@TRUE@"),
        (False, ";;;-@-", "-FALSE-"),
        (True, ";;;@-", "TRUE-"),
        (False, ";;;-@-@-", "-FALSE-FALSE-"),
        (True, "0.00", "TRUE"),
    ]
    for b, code, want in cases:
        assert text_function(b, code) == want, (code, b)


def test_rounding_and_general_fit():
    assert fixed(2.675, 2) == ("2", "68")
    assert fixed(1.005, 2) == ("1", "01")
    assert fixed(0.5, 0) == ("1", "")
    assert fixed(0.04, 1) == ("", "0")
    assert fixed(0.05, 1) == ("", "1")
    assert fixed(9.999, 2) == ("10", "00")
    assert fixed(0.0004, 3) == ("", "000")
    assert fixed(0.0005, 3) == ("", "001")
    assert fixed(1e20, 0)[0] == "100000000000000000000"
    assert fixed(0.0, 2) == ("", "00")
    assert format_general_fit(1e11, 11) == "1E+11"
    assert format_general_fit(12345678901.0, 11) == "12345678901"
    assert format_general_fit(123456789012.0, 11) == "1.23457E+11"
    assert format_general_fit(0.1 + 0.2, 11) == "0.3"
    assert format_general_fit(1.0 / 3.0, 11) == "0.333333333"
    assert format_general_fit(123456.789012, 11) == "123456.789"
    assert format_general_fit(1e-10, 11) == "1E-10"
    assert format_general_fit(0.0000123456789, 11) == "1.23457E-05"
    assert format_general_fit(0.0001, 11) == "0.0001"
    assert format_general_fit(-5.0, 11) == "-5"
    assert format_general_fit(0.0, 11) == "0"
    assert format_general_fit(1234.5678, 6) == "1234.6"
    assert format_general_fit(1234.5678, 4) == "1235"
    assert format_general_fit(123456.0, 5) == "1E+05"
    assert format_general_fit(123456.0, 3) is None
    assert format_general_fit(9.99999999999, 5) == "10"
    assert format_general_fit(float("nan"), 11) is None
    assert format_general_fit(1.7976931348623157e308, 11) == "1.7977E+308"


def test_exact_ties():
    for v, whole in [
        (1234567890123405.0, "1234567890123400"),
        (1234567890123415.0, "1234567890123410"),
        (1234567890123425.0, "1234567890123420"),
        (1234567890123455.0, "1234567890123450"),
        (6631679258820775.0, "6631679258820770"),
        (8052991635850695.0, "8052991635850690"),
    ]:
        assert f("0", v) == whole
        assert f("0.00", v) == whole + ".00"
        assert f("0", -v) == "-" + whole
    assert f("#,##0", 1234567890123455.0) == "1,234,567,890,123,450"
    assert f("0", 123456789012345.5) == "123456789012345"
    assert f("0", 999999999999999.5) == "999999999999999"
    assert f("0", 4999999999999995.0) == "4999999999999990"
    assert f("0", 1234567890123455.0 + 0.25) == "1234567890123460"
    assert f("0", 1234567890123456.0) == "1234567890123460"
    assert f("0.0", 0.285) == "0.3"
    assert f("0.0", 99999999999999.95) == "100000000000000.0"
    assert f("0", 2.5) == "3"
    assert decimal_digits(1234567890123455.0) == ([1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5], 15)
    assert decimal_digits(999999999999999.5) == ([9] * 15, 14)
    assert decimal_digits(999999999999999.6) == ([1] + [0] * 14, 15)


def test_fuzz_no_crash():
    pieces = [
        "0", "#", "?", ".", ",", "%", "E+", "E-", "e", "/", "8", "16", '"', "\\", "_", "*", "@", ";", "[", "]",
        "[Red]", "[>=100]", "[<0]", "[Color7]", "[$€-407]", "[$-409]", "[h]", "[mm]", "[s]", "d", "dd", "ddd",
        "dddd", "m", "mm", "mmm", "mmmmm", "yy", "yyyy", "h", "hh", "s", "ss", ".000", "AM/PM", "A/P", "General",
        " ", "-", "(", ")", "$", "é", "€", "💥", "b", "g", "[>", "[=", '"abc', "00000000000000000000",
        "##########", ",,,,", "%%%%%%", "E+0000", "?/?????????", "/0",
    ]
    values = [
        0.0, -0.0, 1.0, -1.0, 0.5, -0.5, 1e-300, -1e-300, 1e300, -1e300, 1.7976931348623157e308,
        -1.7976931348623157e308, 2.2250738585072014e-308, 2958465.999999, 2958466.0, 46302.75,
        123456789.12345678, -987654.321, 1.0 / 3.0, math.pi, 1e15, 1e16, 0.999999999,
    ]
    rng = random.Random(42)
    for _ in range(600):
        code = "".join(rng.choice(pieces) for _ in range(rng.randrange(8)))
        fmt = parse(code)
        fmt.kind()
        for v in values:
            format_value(v, fmt, S)
            format_value(v, fmt, DateSystem.D1904)
            try:
                text_function(v, code)
            except XlError:
                pass
        format_value("text", fmt, S)
    for w in range(20):
        for v in values:
            s = format_general_fit(v, w)
            if s is not None:
                assert len(s) <= w
