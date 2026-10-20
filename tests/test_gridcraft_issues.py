"""Bugs that GridCraft (https://github.com/storytold/gridcraft), whose design gridcore follows, fixed
in its issues and pull requests, checked here so they stay fixed in gridcore. Each test names the
GridCraft issue (#…) or pull request (PR …). The expected answers were checked in Microsoft Excel 16
for Mac (en-US); where Excel answered otherwise than the issue said, the test says so and follows
Excel."""

import base64
import io
import os
import tempfile
import time
import warnings
import zipfile

import openpyxl
import pytest

from gridcore.calc import evaluate
from gridcore.core.addr import RangeRef
from gridcore.core.date import DateSystem
from gridcore.core.value import XlError, display
from gridcore.engine import EngineError, Session, viewport
from gridcore.model.workbook import Workbook
from gridcore.numfmt import text_function


def new():
    s = Session()
    s.new_workbook()
    return s


def val(s, a):
    return s.run("cell.get", cell=a)["value"]


def inp(s, a):
    return s.run("cell.get", cell=a)["input"]


def ev(formula, wb=None):
    wb = wb or Workbook()
    return display(evaluate(wb, wb.active_sheet, (0, 0), formula))


def cell(s, formula, at="Z1"):
    s.run("cell.set", cell=at, input=formula)
    return val(s, at)


def roundtrip(s, name="rt.xlsx"):
    s2 = Session()
    s2.run("file.open", data=s.run("file.saveBytes")["data"], name=name)
    return s2


def part(s, name):
    raw = base64.b64decode(s.run("file.saveBytes")["data"])
    return zipfile.ZipFile(io.BytesIO(raw)).read(name).decode()


def wb1904():
    wb = Workbook()
    wb.date_system = DateSystem.D1904
    return wb


def minimal_xlsx(sheet_body, workbook_extra=""):
    """A one-sheet package around ``sheet_body`` (the children of ``<worksheet>``)."""
    buf = io.BytesIO()
    z = zipfile.ZipFile(buf, "w")
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    z.writestr(
        "[Content_Types].xml",
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/>'
        f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ct}.worksheet+xml"/>'
        f'<Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/></Types>',
    )
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    z.writestr(
        "_rels/.rels",
        f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    )
    z.writestr(
        "xl/workbook.xml",
        f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{rel}">'
        f'{workbook_extra}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    z.writestr(
        "xl/_rels/workbook.xml.rels",
        f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{rel}/styles" Target="styles.xml"/></Relationships>',
    )
    z.writestr(
        "xl/styles.xml",
        '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs>'
        '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>',
    )
    z.writestr(
        "xl/worksheets/sheet1.xml",
        f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{rel}">{sheet_body}</worksheet>',
    )
    z.close()
    return base64.b64encode(buf.getvalue()).decode()


def open_data(data, name):
    s = Session()
    s.run("file.open", data=data, name=name)
    return s


# ---------------------------------------------------------------- functions


def test_search_positions_with_a_letter_whose_lower_case_is_longer():
    # #571: İ lower-cases to two characters; positions must still count it as one.
    assert ev('=SEARCH("b","İb")') == "2"
    assert ev('=SEARCH("B","İb")') == "2"
    assert ev('=MID("İb",SEARCH("b","İb"),1)') == "b"
    assert ev('=SEARCH("b?","aİbc")') == "3"
    assert ev('=SEARCH("b","aİbc",4)') == "#VALUE!"
    assert ev('=TEXTBEFORE("aİbXc","x",1,1)') == "aİb"


def test_case_insensitive_text_folds_one_letter_to_one_letter():
    # #571 had İ match i; Excel 16 for Mac folds by Unicode simple case folding: İ and ı are letters
    # of their own, ß is not ss; ς, σ and Σ are one letter, µ and μ too.
    for f, want in [('=SEARCH("i","aİbc")', "#VALUE!"), ('=SEARCH("I","aİb")', "#VALUE!"), ('=SEARCH("ı","aIb")', "#VALUE!"),
                    ('=SEARCH("ς","aΣb")', "2"), ('=SEARCH("µ","aΜb")', "2"), ('=SEARCH("ss","aßb")', "#VALUE!"),
                    ('=SEARCH("ß","aẞb")', "2"), ('="ς"="Σ"', "TRUE"), ('="İ"="i"', "FALSE"), ('=MATCH("i",{"İ","i"},0)', "2")]:
        assert ev(f) == want, f
    s = new()
    s.run("cell.set", cell="A1", input="İstanbul")
    assert cell(s, '=COUNTIF(A1,"i*")') == 0 and cell(s, '=COUNTIF(A1,"İSTANBUL")') == 1


def test_upper_and_lower_change_each_letter_into_one_letter():
    # Excel 16 for Mac: UPPER keeps ß and the letters whose upper case is two letters; LOWER("İ") is i.
    for f, want in [('=UPPER("straße")', "STRAßE"), ('=LOWER("İSTANBUL")', "istanbul"), ('=PROPER("ßa")', "ßa"), ('=UPPER("ŉ")', "ŉ"),
                    ('=LEN(UPPER("ﬀ"))', "1"), ('=UPPER("ǆ")', "Ǆ"), ('=PROPER("ǆa")', "Ǆa")]:
        assert ev(f) == want, f


def test_find_and_search_start_past_the_end():
    # #541 made both #VALUE!; Excel 16 for Mac lets FIND start just past the end of the text, not
    # SEARCH, which finds nothing in empty text.
    for f, want in [('=FIND("","abc",3)', "3"), ('=FIND("","abc",4)', "4"), ('=FIND("","abc",5)', "#VALUE!"), ('=FIND("","")', "1"),
                    ('=FIND("","a",2)', "2"), ('=FIND("x","abc",4)', "#VALUE!"), ('=FINDB("","abc",4)', "4"), ('=SEARCH("","abc",3)', "3"),
                    ('=SEARCH("","abc",4)', "#VALUE!"), ('=SEARCH("","")', "#VALUE!")]:
        assert ev(f) == want, f


def test_aggregate_needs_its_data_argument():
    # #572
    assert ev("=AGGREGATE(9,0)") == "#VALUE!"
    assert ev("=AGGREGATE(4,0)") == "#VALUE!"
    # Functions 1-13 take references only (Excel 16 for Mac); the array form (14-19) takes arrays.
    s = new()
    s.run("range.setValues", range="A1", values=[[1], [2]])
    for f, want in [("=AGGREGATE(9,0,{1,2})", "#VALUE!"), ("=AGGREGATE(9,0,A1:A2*1)", "#VALUE!"), ("=AGGREGATE(9,0,1)", "#VALUE!"),
                    ("=AGGREGATE(9,0,A1:A2)", 3), ("=AGGREGATE(9,0,OFFSET(A1,0,0,2,1))", 3), ("=AGGREGATE(14,6,{1,2,3},1)", 3),
                    ("=SUBTOTAL(9,{1,2})", "#VALUE!"), ("=SUBTOTAL(9,A1:A2)", 3)]:
        assert cell(s, f) == want, f


def test_subtotal_leaves_out_nested_subtotals_whatever_its_number():
    # #556
    s = new()
    s.run("cell.set", cell="A1", input="2")
    s.run("cell.set", cell="A2", input="=SUBTOTAL(9,A1)")
    assert cell(s, "=SUBTOTAL(109,A1:A2)") == 2
    assert cell(s, "=SUBTOTAL(9,A1:A2)") == 2
    assert cell(s, "=AGGREGATE(9,4,A1:A2)") == 4  # AGGREGATE keeps its own rule (options 4-7)


def test_prob_checks_only_that_the_probabilities_add_up_to_1():
    # #539 asked for #NUM! on a 0 probability (as Microsoft's documentation says); Excel 16 for Mac
    # takes any probabilities, -0.1 too, as long as they add up to 1 to 15 digits.
    for f, want in [("=PROB({1,2,3},{0,0.5,0.5},1,3)", "1"), ("=PROB({1,2},{-0.1,1.1},1)", "-0.1"), ("=PROB({1,2},{2,-1},2)", "-1"),
                    ("=PROB({1,2,3},{0.2,0.3,0.5},2,3)", "0.8"), ("=PROB({1,2},{0.5,0.6},1)", "#NUM!"),
                    ("=PROB({1,2,3},{0.2,0.30000000000001,0.5},1,3)", "#NUM!"), ("=PROB({1,2,3},{0.2,0.300000000000001,0.5},1,3)", "1")]:
        assert ev(f) == want, f


def test_rounding_at_precisions_past_the_number_range():
    # #540 and 4941bf8d: toward zero it is 0, away from zero 10^309 can't be stored.
    for f in ("ROUNDDOWN", "TRUNC", "ROUND"):
        for x in ("1234", "-1234", "0.5"):
            for d in ("-308", "-309", "-400"):
                assert ev(f"={f}({x},{d})") == "0", (f, x, d)
    assert ev("=ROUNDUP(1234,-309)") == "#NUM!"
    assert ev("=ROUNDUP(0,-309)") == "0"
    assert ev("=ROUNDUP(5E+307,-308)") == "1E+308"


def test_roman_forms():
    # #733, and Microsoft's ROMAN examples for 499.
    expected = {
        (499, 0): "CDXCIX", (499, 1): "LDVLIV", (499, 2): "XDIX", (499, 3): "VDIV", (499, 4): "ID",
        (99, 2): "IC", (199, 2): "CIC", (990, 2): "XM", (1999, 2): "MXMIX", (995, 3): "VM",
        (495, 2): "XDV", (495, 3): "VD", (995, 2): "XMV", (3999, 0): "MMMCMXCIX", (0, 0): "",
    }
    for (n, form), want in expected.items():
        assert ev(f"=ROMAN({n},{form})") == want, (n, form)
    for f, want in [("=ROMAN(4)", "IV"), ("=ROMAN(1999)", "MCMXCIX"), ("=ROMAN(499,)", "CDXCIX"), ('=ROMAN(499,"2")', "XDIX"), ("=ROMAN(A1)", "")]:
        assert ev(f) == want, f
    assert ev("=ROMAN(499,TRUE)") == "CDXCIX"
    assert ev("=ROMAN(499,FALSE)") == "ID"
    assert ev("=ROMAN(499,5)") == "#VALUE!"


def test_edate_and_eomonth_in_early_1900():
    # #734: serial 0 is day 0 of January and February 1900 has 28 days there.
    for f, want in [("=EDATE(0,0)", "0"), ("=EDATE(0,1)", "31"), ("=EDATE(0,12)", "366"), ("=EOMONTH(0,1)", "59"),
                    ("=EDATE(31,1)", "59"), ("=EOMONTH(59,0)", "59"), ("=EOMONTH(DATE(1900,1,15),1)", "59"),
                    ("=EDATE(60,1)", "89"), ("=EOMONTH(0,0)", "31"), ("=DATE(1900,2,29)", "60")]:
        assert ev(f) == want, f


def test_mode_skips_typed_text_and_logical_values():
    # #736
    for f, want in [('=MODE(1,1,"x")', "1"), ('=MODE(1,2,"x")', "#N/A"), ('=MODE(1,"1",2,3)', "#N/A"),
                    ('=MODE("1",1,#DIV/0!,2,3)', "#DIV/0!"), ("=MODE(1,TRUE,TRUE,2)", "#N/A"), ("=MODE(1,1,TRUE)", "1"),
                    ('=MODE.SNGL(3,3,"a",TRUE)', "3")]:
        assert ev(f) == want, f


def test_mirr_countif_indirect_and_usdollar():
    # #738
    assert ev("=MIRR({-10000,-3400,-6500,-1000},0.07,0.01)") == "-1"
    assert ev("=MIRR({100,200,300},0.07,0.01)") == "#DIV/0!"
    s = new()
    s.run("cell.set", cell="H3", input='=""')
    assert cell(s, '=COUNTIF(H3,"<=")') == 0
    assert cell(s, '=COUNTIF(H3,">=")') == 0
    s.run("cell.set", cell="I3", input="30")
    assert cell(s, '=INDIRECT("I3",FALSE)') == "#REF!"
    assert cell(s, '=INDIRECT("R3C9",FALSE)') == 30
    assert ev("=USDOLLAR(1)") == "$1.00"


def test_median_of_huge_numbers_doesnt_overflow():
    # #1579
    assert ev("=MEDIAN(1E+308,1E+308)") == "1E+308"
    assert ev("=MEDIAN(-1E+308,-1E+308)") == "-1E+308"
    assert ev("=MEDIAN(9E+307,9E+307)") == "9E+307"  # Excel takes no number past 9.99999999999999E+307 typed in
    s = new()
    s.run("range.setValues", range="A1", values=[[9e307], [9e307]])
    assert cell(s, "=AGGREGATE(12,6,A1:A2)") == 9e307
    assert ev("=MEDIAN(1,2,3,4)") == "2.5"


def test_criteria_longer_than_255_characters():
    # PR 686: #VALUE!. With several criteria pairs too, in any position (Excel 16 for Mac; PR 686
    # matched the first 255 units there).
    s = new()
    s.run("cell.set", cell="A1", input="a" * 300)
    s.run("cell.set", cell="B1", input="a" * 300)
    s.run("cell.set", cell="A2", input="a" * 255)
    for f in ("=COUNTIF(A1,B1)", "=SUMIF(A1,B1,A1)", "=AVERAGEIF(A1,B1,A1)", "=COUNTIFS(A1,B1)", "=COUNTIFS(A1:A2,B1,A1:A2,B1)",
              '=COUNTIFS(A1:A2,"<>",A1:A2,B1)', '=SUMIFS(A1:A2,A1:A2,"<>",A1:A2,B1)'):
        assert cell(s, f) == "#VALUE!", f
    assert cell(s, "=COUNTIF(A1,LEFT(B1,255))") == 0
    assert cell(s, "=COUNTIF(A2,LEFT(B1,255))") == 1
    assert cell(s, "=COUNTIFS(A1:A2,LEFT(B1,255),A1:A2,LEFT(B1,255))") == 1


def test_ifs_functions_with_several_criteria_pairs():
    # Found while checking PR 686: past the first pairs, a criteria range was taken for a single
    # value, so COUNTIFS with two pairs (SUMIFS with three) gave #VALUE!.
    s = new()
    s.run("range.setValues", range="A1", values=[[1, 1, 1], [2, 0, 1], [3, 5, 0]])
    for f, want in [('=COUNTIFS(A1:A3,">0",B1:B3,">0")', 2), ('=COUNTIFS(A1:A3,">0",B1:B3,">=0",C1:C3,1)', 2),
                    ('=SUMIFS(A1:A3,A1:A3,">0",B1:B3,">0")', 4), ('=SUMIFS(A1:A3,A1:A3,">0",B1:B3,">0",C1:C3,">0")', 1),
                    ('=AVERAGEIFS(A1:A3,A1:A3,">0",B1:B3,">0")', 2), ('=MAXIFS(A1:A3,B1:B3,">0",C1:C3,">0")', 1),
                    ('=MINIFS(A1:A3,B1:B3,">=0",C1:C3,">0")', 1), ('=SUM(COUNTIFS(A1:A3,{1,2,3}))', 3)]:
        assert cell(s, f) == want, f


def test_textbefore_and_textafter_search_from_the_end_for_negative_instances():
    # PR 1941
    for f, want in [('=TEXTBEFORE("aaa","aa",-1)', "a"), ('=TEXTAFTER("aaa","aa",-1)', ""), ('=TEXTBEFORE("a-b-c","-",-1)', "a-b"),
                    ('=TEXTAFTER("a-b-c","-",-2)', "b-c"), ('=TEXTBEFORE("xAyaz","A",-1,1)', "xAy"), ('=TEXTBEFORE("aaa","aa",1)', "")]:
        assert ev(f) == want, f


def test_month_name_dates_as_excel_reads_them():
    # #566 took "Mar 4 26" for a date; Excel 16 for Mac (en-US) doesn't: "month day year" needs a
    # comma after the day (then a space), "day month year" takes none. Without a year the day is in
    # this year, and a number that can't be a day of the month is the year (Mar 32 is March 1932).
    from gridcore.core.date import serial_from_ymd

    s = new()

    def typed(t):
        s.run("cell.set", cell="A1", input=t)
        return val(s, "A1")

    for t, want in [("Mar 4, 26", 46085), ("March 4, 26", 46085), ("Mar, 4, 2026", 46085), ("Mar 4 , 26", 46085), ("March 4,  2026", 46085),
                    ("4 March 26", 46085), ("4 Mar 2026", 46085), ("4-Mar-26", 46085), ("4/Mar/26", 46085), ("Mar 4, 99", 36223),
                    ("Mar 4, 5", 38415), ("Mar 32", 11749), ("Mar 99", 36220), ("Mar 0", 36586), ("Mar-32", 11749), ("Mar 1900", 61),
                    ("Jan 2024", 45292), ("March-2026", 46082)]:
        assert typed(t) == want, t
    for t in ("Mar 4 26", "March 4 26", "Mar 4 2026", "Mar 26 2026", "Sep 30 99", "Mar 4,2026", "4 Mar, 26", "4 March, 2026", "Mar. 4", "Mar.26",
              "Mar/4/26", "32 Mar", "0 Mar", "Mar 100", "Mar 1899", "Mar 4, 1776", "Feb 30, 26", "Tue, Mar 4, 2026", "March 4th, 2026"):
        assert typed(t) == t, t
    y = time.localtime().tm_year
    for t, (m, d) in [("Mar 4", (3, 4)), ("4 Mar", (3, 4)), ("mar 4", (3, 4)), ("Sept 4", (9, 4)), ("Mar 26", (3, 26)), ("Mar-26", (3, 26)),
                      ("Mar/4", (3, 4)), ("Jan 24", (1, 24))]:
        assert typed(t) == serial_from_ymd(DateSystem.D1900, y, m, d), t
    s.run("cell.set", cell="B1", input="Jan 24")
    assert s.run("cell.get", cell="B1")["text"] == "24-Jan"
    assert typed("Mar 4 1:30 PM") == serial_from_ymd(DateSystem.D1900, y, 3, 4) + 0.5625


def test_lambda_takes_one_argument_position_per_parameter():
    # #557
    for f, want in [("=LAMBDA(a,b,a+b)(2)", "#VALUE!"), ("=LAMBDA(a,b,a+b)(2,3)", "5"), ("=LAMBDA(a,b,a+b)(2,3,4)", "#VALUE!"),
                    ('=LAMBDA(a,b,IF(ISOMITTED(b),"Missing",a+b))(1,)', "Missing"), ("=LAMBDA(7)()", "7"), ("=LAMBDA(x,ISOMITTED(x))()", "TRUE")]:
        assert ev(f) == want, f
    s = new()
    s.run("formulas.defineName", name="AddPair", refersTo="=LAMBDA(a,b,a+b)")
    assert cell(s, "=AddPair(2)") == "#VALUE!"
    assert cell(s, "=AddPair(2,3)") == 5


def test_functions_passed_by_name():
    # #902 and PR 486: GROUPBY(…, SUM) as Microsoft documents it, BYROW(…, SUM), MAP(…, ABS).
    s = new()
    s.run("range.setValues", range="A1", values=[["K", "V"], ["a", 1], ["b", 2], ["a", 3]])
    s.run("cell.set", cell="D1", input="=GROUPBY(A2:A4,B2:B4,SUM)")
    assert [val(s, a) for a in ("D1", "E1", "D2", "E2")] == ["a", 4, "b", 2]
    assert cell(s, "=BYROW(B2:B4,SUM)", "G1") == 1 and val(s, "G3") == 3
    assert cell(s, "=MAP(-B2:B4,ABS)", "H1") == 1
    assert cell(s, "=REDUCE(0,B2:B4,SUM)", "I1") == 6
    assert cell(s, "=SCAN(0,B2:B4,SUM)", "J1") == 1 and val(s, "J3") == 6
    assert cell(s, "=MAKEARRAY(2,2,PRODUCT)", "K1") == 1 and val(s, "L2") == 4
    # A LET variable or a defined name keeps its meaning.
    assert cell(s, "=LET(SUM,LAMBDA(x,x*10),BYROW(B2:B3,SUM))", "M1") == 10
    # Files spell such names with _xleta. (Excel's form) and read back unchanged.
    x = part(s, "xl/worksheets/sheet1.xml")
    assert "_xlfn.GROUPBY(A2:A4,B2:B4,_xleta.SUM)" in x and "_xlfn.BYROW(B2:B4,_xleta.SUM)" in x
    s2 = roundtrip(s)
    assert inp(s2, "D1") == "=GROUPBY(A2:A4,B2:B4,SUM)" and val(s2, "E1") == 4


def test_number_formats_drop_the_sign_of_a_negative_that_rounds_to_zero():
    # #729, Excel 365's table.
    for code, v, want in [("0.00", -0.0033, "0.00"), ("0.00", -2.89e-09, "0.00"), ("0.00", -0.005, "-0.01"), ("0", -0.4, "0"),
                          ("0", -0.5, "-1"), ("0.0", -0.004, "0.0"), ("#,##0.00", -0.004, "0.00"), ("0.00%", -2.89e-09, "0.00%"),
                          ("0.00%", -0.0033, "-0.33%"), ("0.00;(0.00)", -0.0033, "(0.00)"), ('0.00;-0.00;"zero"', -0.0033, "-0.00"),
                          ("0.00E+00", -2.89e-09, "-2.89E-09")]:
        assert text_function(v, code) == want, (code, v)


def test_scientific_formats():
    # #768, in the forms Excel 16 for Mac takes (its answers): % and commas don't scale an exponent,
    # zero fills # placeholders. A lowercase e is the era year to Excel, so those codes are refused.
    for code, v, want in [("#E+#%", 12345, "1E+4%"), ("#,E+#", 12345, "1E+4"), ("#%E-#", 12345, "1%E4"), ("####.####E+#", 12345, "1.2345E+4"),
                          ("#.#E+#", -1.23456789e-05, "-1.2E-5"), ("#.#E+#", 0, "0.E+0"), ("####.####E+#", 0, "0000.E+0"),
                          ("0.00E+00", 12345, "1.23E+04"), ("0.00E+00", 0, "0.00E+00"), ("##0.0E+0", 12345, "12.3E+3"),
                          ("0.0E+0%", 12345, "1.2E+4%"), ("0.0E+0%", 0.004, "4.0E-3%"), ("0.0E-0", 12345, "1.2E4")]:
        assert text_function(v, code) == want, (code, v)
    for code in ("0.00e+00", "#,e+#", "####.####e+#", "0E+", "E+0", "0.0E0", "0.0E+0E+0"):
        with pytest.raises(XlError):
            text_function(12345, code)


def test_format_codes_excel_refuses():
    # #807 (past 255 characters) and #810 (an open [), and the rest of what Excel 16 for Mac refuses
    # as a number format: TEXT() is #VALUE! and Format Cells says so.
    bad = ["0" * 256, "[0;0", "[0", "0[", '0"', "0\\", "0*", "0_", "0;0;0;@;0", "[xyz]0", "[Purple]0", "[Color57]0", "[ENG]0", "[d]", "[h]:[m]",
           "[Red][Blue]0", "[]0", "0 yyyy", "d 0", "yyyy%", "h:mm:ss.0#", "# ?/? d", "@ h", "0 @", "General 0"]
    good = ["0" * 255, "0 x", "[Red]0", "[color 10]0", "[DBNum1]0", "[$€-x-euro2]0", "[>=1][Red]0", "mm:ss.0", "hh.0", "[s].00", "d.m.yyyy",
            "h:mm,", "0 AM/PM", "General h", "0;0;0;@", ";;;", "@@", "0 _)", '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)']
    for code in bad:
        with pytest.raises(XlError):
            text_function(1, code)
    for code in good:
        text_function(1, code)
    s = new()
    s.run("cell.set", cell="A1", input="5")
    with pytest.raises(EngineError):
        s.run("home.numberFormat", range="A1", code="[0;0")
    with pytest.raises(EngineError):
        s.run("home.formatCells", range="A1", numFmt="0" * 256)
    s.run("home.numberFormat", range="A1", code="0.00")
    assert s.run("cell.get", cell="A1")["text"] == "5.00"


# ---------------------------------------------------------------- calculation


def test_cells_that_only_hold_a_format_dont_block_a_spill():
    # #1011 (and the bold empty cell before it): only content blocks a spill, as in Excel.
    s = new()
    s.run("home.bold", range="C1")
    s.run("cell.set", cell="A1", input="=SEQUENCE(1,3)")
    assert val(s, "A1") == 1 and val(s, "C1") == 3
    # A styled blocker far from the anchor, cleared: the spill comes back.
    s.run("cell.set", cell="A3", input="=SEQUENCE(1,80)")
    s.run("cell.set", cell="CB3", input="x")
    s.run("home.bold", range="CB3")
    assert val(s, "A3") == "#SPILL!"
    s.run("edit.clearContents", range="CB3")
    assert val(s, "A3") == 1 and val(s, "CB3") == 80


def test_a_formula_reading_a_cleared_spill_past_65536_cells_recalculates():
    # #558
    s = new()
    s.run("cell.set", cell="A1", input="=SEQUENCE(65537)")
    s.run("cell.set", cell="C2", input="=A65537")
    assert val(s, "C2") == 65537
    s.run("cell.set", cell="A1", input="5")
    assert val(s, "C2") == 0


def test_subtotal_follows_rows_hidden_shown_and_filtered():
    # PR 796, and filters: SUBTOTAL and AGGREGATE look again whenever rows are hidden or shown.
    s = new()
    s.run("range.setValues", range="A1", values=[["H"], [1], [2], [3]])
    s.run("cell.set", cell="C1", input="=SUBTOTAL(109,A2:A4)")
    s.run("cell.set", cell="C2", input="=SUBTOTAL(9,A2:A4)")
    s.run("cell.set", cell="C3", input="=AGGREGATE(9,5,A2:A4)")
    s.run("home.hideRows", rows="3:3")
    assert [val(s, a) for a in ("C1", "C2", "C3")] == [4, 6, 4]
    s.run("home.unhideRows", rows="2:4")
    assert val(s, "C1") == 6
    s.run("home.rowHeight", rows="3:3", height=0)
    assert val(s, "C1") == 4
    s.run("home.rowHeight", rows="3:3", height=20)
    assert val(s, "C1") == 6
    s.run("data.filter", range="A1:A4")
    s.run("data.filterColumn", col="A", criterion={"kind": "values", "values": ["1"]})
    assert [val(s, a) for a in ("C1", "C2", "C3")] == [1, 1, 1]
    s.run("data.clearFilter")
    assert [val(s, a) for a in ("C1", "C2", "C3")] == [6, 6, 6]


def test_a_table_total_row_follows_its_filter():
    s = new()
    s.run("range.setValues", range="A1", values=[["Item", "Qty"], ["a", 2], ["b", 3], ["c", 5]])
    s.run("insert.table", range="A1:B4", name="T")
    s.run("table.totalRow", on=True, table="T")
    assert inp(s, "B5") == "=SUBTOTAL(109,T[Qty])" and val(s, "B5") == 10
    s.run("selection.set", range="B2")
    s.run("data.filterColumn", col="A", criterion={"kind": "values", "values": ["a", "c"]})
    assert val(s, "B5") == 7


def test_long_operator_chains_calculate():
    # PR 484 and #527: Excel allows them up to its 8,192 characters.
    s = new()
    s.run("cell.set", cell="A1", input="1")
    chain = "=" + "+".join(["A1"] * 2000)
    assert cell(s, chain, "C1") == 2000
    assert cell(s, "=" + "&".join(['"a"'] * 1500), "C2") == "a" * 1500
    assert cell(s, "=" + "*".join(["1"] * 3000) + "=1", "C3") is True
    s.run("home.insertRows", rows="1:1")
    assert inp(s, "C2") == "=" + "+".join(["A2"] * 2000) and val(s, "C2") == 2000
    assert val(roundtrip(s), "C2") == 2000


def test_formula_text_is_at_most_8192_characters():
    # PR 1330: longer text, postfix % nesting past the limit: refused, never a crash.
    s = new()
    ok = "=" + "+".join(["1"] * 4096)  # 8,192 characters with the = (Excel 16 for Mac counts it)
    assert cell(s, ok, "A1") == 4096
    for bad in ("=" + "+".join(["1"] * 4095) + "+12", "=" + "+".join(["1"] * 4097), "=1" + "%" * 5000, "=" + "(" * 300 + "1" + ")" * 300):
        with pytest.raises(EngineError):
            s.run("cell.set", cell="A2", input=bad)
    assert cell(s, "=1" + "%" * 3, "A3") == pytest.approx(1e-6)


def test_today_and_now_in_a_1904_workbook():
    # #578
    wb = wb1904()
    assert ev("=YEAR(TODAY())", wb) == str(time.localtime().tm_year)
    assert ev("=TODAY()-DATE(YEAR(TODAY()),MONTH(TODAY()),DAY(TODAY()))", wb) == "0"
    assert ev("=INT(NOW())=TODAY()", wb) == "TRUE"


def test_date_text_in_a_1904_workbook():
    # #564 and #565
    wb = wb1904()
    assert ev('="1/1/1904"+0', wb) == "0"
    assert ev('=VALUE("1/2/1904")', wb) == "1"
    assert ev('=-"1/2/1904"', wb) == "-1"
    assert ev('="1/2/1904"+0') == "1463"  # 1900 workbooks unchanged
    assert ev('=TEXT(2957003,"m/d/yyyy")', wb) == "12/31/9999"
    assert ev('=TEXT(2957004,"m/d/yyyy")', wb) == "#VALUE!"
    assert ev('=TEXT(2958465,"m/d/yyyy")') == "12/31/9999"


def test_negative_dates_and_times_in_a_1904_workbook():
    # Excel 16 for Mac shows them with a minus sign in a 1904 workbook; a 1900 one can't (#VALUE!).
    wb = wb1904()
    for f, want in [('=TEXT(-1,"m/d/yyyy")', "-1/2/1904"), ('=TEXT(-1.5,"m/d/yyyy h:mm")', "-1/2/1904 12:00"), ('=TEXT(-0.25,"h:mm")', "-6:00"),
                    ('=TEXT(-1.5,"[h]:mm")', "-36:00")]:
        assert ev(f, wb) == want, f
    assert ev('=TEXT(-1,"m/d/yyyy")') == "#VALUE!"


def test_index_area_number_below_1():
    # PR 476 (an area number past the last area is #REF!), and one below 1 is #VALUE! (Excel 16 for
    # Mac; 0.5 counts as 0).
    s = new()
    s.run("range.setValues", range="A1", values=[[1, 2, 5, 6], [3, 4, 7, 8]])
    for f, want in [("=INDEX((A1:B2,C1:D2),1,1,2)", 5), ("=INDEX((A1:B2,C1:D2),1,1,3)", "#REF!"), ("=INDEX(A1:B2,1,1,1E+300)", "#REF!"),
                    ("=INDEX(A1:B2,1,1,-1E+300)", "#VALUE!"), ("=INDEX(A1:B2,1,1,-1)", "#VALUE!"), ("=INDEX(A1:B2,1,1,0)", "#VALUE!"),
                    ("=INDEX(A1:B2,1,1,0.5)", "#VALUE!"), ("=INDEX(A1:B2,1,1,1.9)", 1)]:
        assert cell(s, f) == want, f


def test_whole_rows_and_columns_take_one_dollar_sign():
    # PR 1922: $$1:$3 is no reference (Excel 16 for Mac: #REF! from INDIRECT).
    s = new()
    for f, want in [('=ROWS(INDIRECT("$$1:$3"))', "#REF!"), ('=ROWS(INDIRECT("$1:$$3"))', "#REF!"), ('=COLUMNS(INDIRECT("$$A:$B"))', "#REF!"),
                    ('=ROWS(INDIRECT("+1:+3"))', "#REF!"), ('=ROWS(INDIRECT("$1:$3"))', 3), ('=COLUMNS(INDIRECT("$A:B"))', 2)]:
        assert cell(s, f) == want, f
    assert RangeRef.parse("$$1:$3") is None and RangeRef.parse("$1:$3") is not None


# ---------------------------------------------------------------- formula text and files


def test_an_intersection_with_a_deleted_operand_reads_back():
    # #568
    s = new()
    for f in ("=A1:B2 #REF!", "=#REF! A1:B2", "=(A1:B2) #REF!"):
        assert cell(s, f, "H1") == "#REF!" and inp(s, "H1") == f
    s.run("cell.set", cell="H2", input="=SUM(A1:B2 D1:E2)")
    s.run("home.deleteColumns", cols="D:E")
    assert inp(s, "F2") == "=SUM(A1:B2 #REF!)"
    s2 = roundtrip(s)
    s2.run("home.insertColumns", cols="A:A")
    assert inp(s2, "G2") == "=SUM(B1:C2 #REF!)"


def test_sheet_scoped_names_keep_their_sheet_quoted():
    # #567, and renames to names that need quotes.
    s = new()
    s.run("home.insertSheet", name="My Sheet", index=1)
    s.run("sheet.activate", sheet="My Sheet")
    s.run("cell.set", cell="A1", input="42")
    s.run("formulas.defineName", name="Total", refersTo="='My Sheet'!$A$1", scope="My Sheet")
    s.run("sheet.activate", sheet="Sheet1")
    assert cell(s, "='My Sheet'!Total", "B1") == 42 and inp(s, "B1") == "='My Sheet'!Total"
    s2 = roundtrip(s)
    s2.run("sheet.activate", sheet="Sheet1")
    assert val(s2, "B1") == 42
    for name, shown in (("It's", "='It''s'!Total"), ("Plain", "=Plain!Total")):
        cur = s.doc().wb.sheets[1].name
        s.run("sheet.rename", sheet=cur, name=name)
        s.run("sheet.activate", sheet="Sheet1")
        assert inp(s, "B1") == shown and val(s, "B1") == 42


def test_user_defined_prefix_stays():
    # #726: _xludf. marks a user-defined function, not the built-in of the same name.
    s = new()
    s.run("cell.set", cell="A1", input='=_xludf.CONCAT("a","b")')
    assert val(s, "A1") == "#NAME?" and inp(s, "A1") == '=_xludf.CONCAT("a","b")'
    assert '_xludf.CONCAT("a","b")' in part(s, "xl/worksheets/sheet1.xml")
    assert cell(s, '=CONCAT("a","b")', "A2") == "ab"


def test_table_columns_keep_spaces_at_their_edges():
    # #732, Excel 365's table: the column is "Qty ", and Table1[Qty] names no column.
    s = new()
    s.run("range.setValues", range="A1", values=[["Item", "Qty "], ["a", 2], ["b", 3]])
    s.run("insert.table", range="A1:B3", name="Table1")
    s.run("table.totalRow", on=True, table="Table1")
    assert val(s, "B1") == "Qty "
    assert cell(s, "=SUM(Table1[Qty ])", "F1") == 5 and inp(s, "F1") == "=SUM(Table1[[Qty ]])"
    assert cell(s, "=SUM(Table1[[Qty ]])", "F2") == 5
    assert cell(s, "=Table1[[#Totals],[Qty ]]", "F3") == 5
    assert cell(s, "=SUBTOTAL(109,Table1[Qty ])", "F4") == 5
    assert cell(s, "=SUM(Table1[Qty])", "F5") == "#REF!"
    assert inp(s, "B4") == "=SUBTOTAL(109,Table1[[Qty ]])"
    t = part(s, "xl/tables/table1.xml")
    assert 'name="Qty "' in t
    s2 = roundtrip(s)
    assert val(s2, "F1") == 5 and val(s2, "B4") == 5


def test_a_workbook_opens_by_its_content_not_its_name():
    # #569
    s = new()
    s.run("cell.set", cell="A1", input="5")
    data = s.run("file.saveBytes")["data"]
    for name in ("book.csv", "book.ods", "book.dat"):
        assert val(open_data(data, name), "A1") == 5, name
    assert val(open_data(base64.b64encode(b"a,b\r\n1,2\r\n").decode(), "data.xlsx"), "B2") == 2


def test_document_properties_keep_their_spaces():
    # #554
    s = new()
    props = {"title": "  Title  ", "subject": "\n", "author": "\tA ", "company": " C ", "keywords": " k ", "description": "\n d \n"}
    s.run("file.properties", **props)
    p = roundtrip(s).run("file.properties")
    assert {k: p.get(k) for k in props} == props


def test_a_table_always_has_a_data_row():
    # #511 saved a table of a header and a total row only; Excel 16 for Mac finds that file damaged,
    # and one with a header row only too. As in Excel: a table made from a header row takes the row
    # below as its data row, and deleting all its data rows keeps the first, emptied.
    s = new()
    s.run("range.setValues", range="A1", values=[["Name", "Qty"]])
    s.run("insert.table", range="A1:B1", name="Tab")
    s.run("table.totalRow", on=True, table="Tab")
    t = roundtrip(s).doc().sheet.tables
    assert [(x.name, x.range.a1(), x.totals_row) for x in t] == [("Tab", "A1:B3", True)]
    with pytest.raises(EngineError):
        s.run("table.resize", range="A1:B2", table="Tab")
    s = new()
    s.run("range.setValues", range="A1", values=[["Name", "Qty"], ["a", 1], ["b", 2], ["c", 3]])
    s.run("range.setValues", range="D1", values=[["d1"], ["d2"], ["d3"], ["d4"], ["d5"]])
    s.run("insert.table", range="A1:B4", name="Tab")
    s.run("table.totalRow", on=True, table="Tab")
    s.run("home.deleteRows", rows="2:4")
    t = s.doc().sheet.tables[0]
    assert t.range.a1() == "A1:B3" and t.totals_row
    assert [val(s, a) for a in ("A2", "B2", "D2", "D3")] == [None, None, "d2", "d5"]
    assert inp(s, "B3") == "=SUBTOTAL(109,Tab[Qty])"
    # A table without a data row (from another program) isn't written.
    from gridcore.model.features import Table, TableColumn
    from gridcore.xlsx import write_xlsx

    wb = Workbook()
    wb.sheets[0].tables.append(Table(1, "Bare", RangeRef(0, 0, 1, 1), totals_row=True, columns=[TableColumn("A"), TableColumn("B")]))
    assert not [n for n in zipfile.ZipFile(io.BytesIO(write_xlsx(wb))).namelist() if n.startswith("xl/tables/")]


def test_every_range_of_a_conditional_format_and_validation_is_read():
    # #512
    sq = " ".join(f"A{2 * i + 1}" for i in range(12000))
    body = (
        "<sheetData/>"
        f'<conditionalFormatting sqref="{sq}"><cfRule type="expression" dxfId="0" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>'
        f'<dataValidations count="1"><dataValidation type="whole" sqref="{sq}"><formula1>1</formula1></dataValidation></dataValidations>'
    )
    sh = open_data(minimal_xlsx(body), "x.xlsx").doc().sheet
    assert len(sh.cond_formats[0].ranges) == 12000 and len(sh.validations[0].ranges) == 12000


def test_hyperlinks_past_the_cell_cap_are_reported():
    # #513
    s = open_data(minimal_xlsx('<sheetData/><hyperlinks><hyperlink ref="A1:A12000" location="Sheet1!B1"/></hyperlinks>'), "x.xlsx")
    assert len(s.doc().sheet.hyperlinks) == 10000
    assert any("hyperlinks past the first 10,000 cells were dropped" in w for w in s.doc().warnings)
    s = open_data(minimal_xlsx('<sheetData/><hyperlinks><hyperlink ref="A1:A3" location="Sheet1!B1"/></hyperlinks>'), "x.xlsx")
    assert len(s.doc().sheet.hyperlinks) == 3 and not s.doc().warnings


def test_a_link_to_another_file_stays_a_link_to_that_file():
    # #800: files are external relationships, places in the workbook are locations.
    s = new()
    s.run("home.insertSheet", name="Sheet2", index=1)
    s.run("sheet.activate", sheet="Sheet1")
    for c, target in (("A1", "Source.xlsx"), ("A2", "..\\Reports\\Q1.xlsx"), ("A3", "https://example.com"), ("A4", "Sheet2!B5"), ("A5", "B5")):
        s.run("insert.hyperlink", cell=c, target=target)
    x = part(s, "xl/worksheets/sheet1.xml")
    rels = part(s, "xl/worksheets/_rels/sheet1.xml.rels")
    assert 'location="Source.xlsx"' not in x and 'Target="Source.xlsx"' in rels and "Q1.xlsx" in rels
    assert 'location="Sheet2!B5"' in x and 'location="B5"' in x
    links = roundtrip(s).doc().sheet.hyperlinks
    assert links[(0, 0)].external and links[(0, 0)].target == "Source.xlsx"
    assert not links[(3, 0)].external


def test_links_open_only_web_and_mail_addresses():
    # PR 780 (#777): a workbook's link can't run script in the browser or reach other schemes.
    s = new()
    s.run("home.insertSheet", name="Sheet2", index=1)
    s.run("sheet.activate", sheet="Sheet1")
    for target in ("javascript:alert(1)//x.y", " javascript:alert(1)", "JAVASCRIPT:x", "data:text/html,x", "file:///etc/passwd",
                   "vbscript:x", "java\tscript:alert(1)", "Source.xlsx"):
        s.run("insert.hyperlink", cell="A1", target=target)
        with pytest.raises(EngineError):
            s.run("hyperlink.follow", cell="A1")
    for target, url in (("https://example.com/a", "https://example.com/a"), ("mailto:a@b.co", "mailto:a@b.co"),
                        ("www.example.com", "http://www.example.com")):
        s.run("insert.hyperlink", cell="A1", target=target)
        assert s.run("hyperlink.follow", cell="A1") == {"url": url}
    s.run("insert.hyperlink", cell="A1", target="Sheet2!B5")
    assert s.run("hyperlink.follow", cell="A1") == {"sheet": "Sheet2", "range": "B5"}
    # A range on the same sheet is a place, not a link with the scheme `B5`.
    s.run("sheet.activate", sheet="Sheet1")
    for target, rng in (("B5:C9", "B5:C9"), ("#D3:E4", "D3:E4"), ("AB10:AC12", "AB10:AC12")):
        s.run("insert.hyperlink", cell="A1", target=target)
        assert s.run("hyperlink.follow", cell="A1") == {"sheet": "Sheet1", "range": rng}
    # A link read from a file goes through the same check.
    body = '<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks>'
    data = base64.b64decode(minimal_xlsx(body))
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(buf, "w") as zout:
        for n in zin.namelist():
            zout.writestr(n, zin.read(n))
        zout.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="javascript:alert(document.cookie)//x.y" TargetMode="External"/></Relationships>',
        )
    s2 = open_data(base64.b64encode(buf.getvalue()).decode(), "x.xlsx")
    with pytest.raises(EngineError):
        s2.run("hyperlink.follow", cell="A1")


def test_a_workbook_with_a_password_to_open_is_not_called_an_xls_file():
    # #11: such a workbook is a compound file holding the encrypted package.
    from gridcore.xlsx.cfb import Storage, write_cfb

    root = Storage()
    root.set_stream("EncryptionInfo", b"\x04\x00\x04\x00" + b"x" * 100)
    root.set_stream("EncryptedPackage", b"\x00" * 600)
    with pytest.raises(EngineError, match="needs a password to open"):
        open_data(base64.b64encode(write_cfb(root)).decode(), "secret.xlsx")
    old = Storage()
    old.set_stream("Workbook", b"\x09\x08" + b"\x00" * 600)
    with pytest.raises(EngineError, match="Legacy .xls"):
        open_data(base64.b64encode(write_cfb(old)).decode(), "old.xls")


def test_csv_of_two_cells_far_apart_is_refused_at_once():
    # #528: A1 and XFD1048576 would be 16 GiB of commas (and kept the server busy for ever).
    from gridcore.xlsx.csv_io import CsvTooLarge, write_csv

    s = new()
    s.run("cell.set", cell="A1", input="near")
    s.run("cell.set", cell="XFD1048576", input="far")
    t = time.time()
    with pytest.raises(EngineError, match="over the limit of 512 MiB"):
        s.run("file.saveBytes", format="csv")
    assert time.time() - t < 1
    # A sparse sheet is written quickly and exactly: only the filled cells are formatted.
    s = new()
    s.run("range.setValues", range="A1", values=[["a", "b,c", 'x"y'], [1, None, 3]])
    s.run("cell.set", cell="D5", input="=SEQUENCE(2)")
    s.run("cell.set", cell="B8", input="z")
    data = base64.b64decode(s.run("file.saveBytes", format="csv")["data"])
    assert data == b'\xef\xbb\xbfa,"b,c","x""y",\r\n1,,3,\r\n,,,\r\n,,,\r\n,,,1\r\n,,,2\r\n,,,\r\n,z,,\r\n'
    # The running size is checked too (long text in few cells).
    s = new()
    s.run("cell.set", cell="A1", input="x" * 3000)
    s.run("cell.set", cell="A2", input="y" * 3000)
    sh = s.doc().sheet
    with pytest.raises(CsvTooLarge):
        write_csv(sh, lambda rc: sh.value(rc) or "", limit=5000)
    assert len(write_csv(sh, lambda rc: sh.value(rc) or "", limit=7000)) > 6000


def test_saving_to_an_unknown_extension_is_refused():
    # #878
    s = new()
    s.run("cell.set", cell="A1", input="1")
    d = tempfile.mkdtemp()
    with pytest.raises(EngineError):
        s.run("file.saveAs", path=os.path.join(d, "report.cvs"))
    assert not os.listdir(d)
    s.run("file.saveAs", path=os.path.join(d, "report.xlsx"))
    assert os.listdir(d) == ["report.xlsx"]


# ---------------------------------------------------------------- commands


def test_a_monthly_series_counts_from_the_first_date():
    # #132: Jan 31, Feb 28, Mar 31, Apr 30.
    s = new()
    s.run("cell.set", cell="A1", input="2026-01-31")
    s.run("edit.fillSeries", range="A1:A4", type="date", unit="month")
    assert [s.run("cell.get", cell=f"A{i}")["text"] for i in range(1, 5)] == ["1/31/2026", "2/28/2026", "3/31/2026", "4/30/2026"]
    s.run("cell.set", cell="B1", input="2024-02-29")
    s.run("edit.fillSeries", range="B1:B3", type="date", unit="year")
    assert [s.run("cell.get", cell=f"B{i}")["text"] for i in range(1, 4)] == ["2/29/2024", "2/28/2025", "2/28/2026"]


def test_a_falling_growth_series_stops_at_its_stop_value():
    # b5a82554
    s = new()
    s.run("cell.set", cell="A1", input="100")
    s.run("edit.fillSeries", range="A1:A10", type="growth", step=0.5, stop=10)
    assert [val(s, f"A{i}") for i in range(1, 7)] == [100, 50, 25, 12.5, None, None]
    s.run("cell.set", cell="B1", input="1")
    s.run("edit.fillSeries", range="B1:B10", type="growth", step=2, stop=10)
    assert [val(s, f"B{i}") for i in range(1, 6)] == [1, 2, 4, 8, None]


def test_a_weekday_series_with_a_huge_step_is_refused_at_once():
    # #535: it used to keep the server busy for ever.
    s = new()
    s.run("cell.set", cell="A1", input="45000")
    t = time.time()
    with pytest.raises(EngineError):
        s.run("edit.fillSeries", range="A1:A3", type="date", unit="weekday", step=1e12)
    assert time.time() - t < 2 and val(s, "A2") is None
    s.run("cell.set", cell="B1", input="2026-10-16")  # a Friday
    s.run("edit.fillSeries", range="B1:B4", type="date", unit="weekday", step=1)
    assert [s.run("cell.get", cell=f"B{i}")["text"] for i in range(1, 5)] == ["10/16/2026", "10/19/2026", "10/20/2026", "10/21/2026"]


def test_autofill_times_quarters_ordinals_and_mixed_cells():
    # PR 483
    def fill(values, n):
        s = new()
        for i, v in enumerate(values):
            s.run("cell.set", cell=f"A{i + 1}", input=v)
        s.run("edit.autoFill", source=f"A1:A{len(values)}", target=f"A1:A{len(values) + n}")
        return [s.run("cell.get", cell=f"A{i + 1}")["text"] for i in range(len(values), len(values) + n)]

    assert fill(["9:00"], 2) == ["10:00", "11:00"]
    assert fill(["Qtr3"], 3) == ["Qtr4", "Qtr1", "Qtr2"]
    assert fill(["Quarter 3"], 2) == ["Quarter 4", "Quarter 1"]
    assert fill(["Q3"], 2) == ["Q4", "Q1"]
    assert fill(["1st Period"], 3) == ["2nd Period", "3rd Period", "4th Period"]
    assert fill(["11th"], 3) == ["12th", "13th", "14th"]
    assert fill(["Product 1", "On backorder"], 4) == ["Product 2", "On backorder", "Product 3", "On backorder"]
    assert fill(["Item 5"], 2) == ["Item 6", "Item 7"]  # unchanged
    # What Excel 16 for Mac does besides: the number a text starts with, else the one it ends with
    # (a number inside is copied); in a source without one pattern, neighbours of a kind continue
    # together and a number alone counts on by 1.
    assert fill(["a1b"], 2) == ["a1b", "a1b"]
    assert fill(["Item 5th"], 1) == ["Item 5th"]
    assert fill(["1 to 5"], 2) == ["2 to 5", "3 to 5"]
    assert fill(["1", "a"], 4) == ["2", "a", "3", "a"]
    assert fill(["1", "2", "x", "5"], 8) == ["3", "4", "x", "6", "5", "6", "x", "7"]
    assert fill(["Item 1", "Item 3", "x"], 6) == ["Item 5", "Item 7", "x", "Item 9", "Item 11", "x"]
    assert fill(["Mon", "Tue", "x"], 3) == ["Wed", "Thu", "x"]
    assert fill(["1", "", "a"], 3) == ["2", "", "a"]


def test_autofit_row_height_shows_hidden_rows():
    # PR 796, and what Excel 16 for Mac does: AutoFit Row Height shows the hidden rows it fits, empty
    # ones too, and SUBTOTAL follows; AutoFit Column Width shows a hidden column only when it holds
    # something.
    s = new()
    s.run("range.setValues", range="A1", values=[["H"], [1], [2], [3]])
    s.run("cell.set", cell="C1", input="=SUBTOTAL(109,A2:A4)")
    s.run("home.hideRows", rows="3:3")
    s.run("home.hideRows", rows="6:6")
    assert val(s, "C1") == 4
    s.run("home.autofitRowHeight", rows="2:6")
    assert not any(s.doc().sheet.is_row_hidden(r) for r in range(8)) and val(s, "C1") == 6
    s.run("cell.set", cell="E1", input="some text")
    s.run("home.hideColumns", cols="E:E")
    s.run("home.hideColumns", cols="F:F")
    s.run("home.autofitColumnWidth", cols="D:F")
    sh = s.doc().sheet
    assert not sh.col_info(4).hidden and sh.col_info(5).hidden


def test_removing_a_filter_without_criteria_keeps_hidden_rows():
    # #508
    s = new()
    s.run("range.setValues", range="A1", values=[["H"], [1], [2], [3], [4]])
    s.run("home.hideRows", rows="3:3")
    s.run("home.hideRows", rows="10:10")
    s.run("data.filter", range="A1:A5")
    s.run("data.filter")
    sh = s.doc().sheet
    assert [r + 1 for r in range(12) if sh.is_row_hidden(r)] == [3, 10]
    # A filter with criteria still shows the rows it hid when it goes.
    s.run("home.unhideRows", rows="1:12")
    s.run("data.filter", range="A1:A5")
    s.run("data.filterColumn", col="A", criterion={"kind": "values", "values": ["1"]})
    assert s.doc().sheet.is_row_hidden(2)
    s.run("data.filter")
    assert not any(s.doc().sheet.is_row_hidden(r) for r in range(6))


def test_committing_nothing_into_an_empty_cell_is_no_step():
    # #1855: Backspace, then Enter, is one undo step.
    s = new()
    s.run("cell.set", cell="A1", input="1")
    s.run("cell.set", cell="A2", input="2")
    s.run("edit.clearContents", range="A2")
    s.run("cell.set", cell="A2", input="")
    s.run("edit.undo")
    assert val(s, "A2") == 2 and val(s, "A1") == 1


def test_redo_selects_the_cells_of_the_step():
    # #1855
    s = new()
    s.run("selection.set", range="A1")
    s.run("cell.set", cell="A1", input="1")
    s.run("selection.set", range="D5")
    s.run("edit.undo")
    assert s.doc().selection.current().a1() == "A1"
    s.run("selection.set", range="F9")
    s.run("edit.redo")
    assert s.doc().selection.current().a1() == "A1" and val(s, "A1") == 1


def test_double_clicking_the_fill_handle_fills_down_to_the_adjacent_data():
    # #507
    s = new()
    s.run("range.setValues", range="A1", values=[[1, 10], [2, 20], [3, 30], [4, None], [5, None], [6, None]])
    s.run("selection.set", range="B1:B3")
    assert s.run("edit.autoFillDown") == {"range": "B1:B6"}
    assert [val(s, f"B{i}") for i in range(1, 8)] == [10, 20, 30, 40, 50, 60, None]
    assert s.doc().selection.current().a1() == "B1:B6"
    s.run("edit.undo")
    assert val(s, "B4") is None
    # It stops above a cell that holds something, and uses the right column when the left is empty.
    s = new()
    s.run("range.setValues", range="B1", values=[["=C1*2", 1], [None, 2], [None, 3], ["keep", 4], [None, 5]])
    s.run("selection.set", range="B1")
    s.run("edit.autoFillDown")
    assert [inp(s, f"B{i}") for i in range(1, 6)] == ["=C1*2", "=C2*2", "=C3*2", "keep", ""]
    s.run("selection.set", range="H1")
    assert s.run("edit.autoFillDown") == {"range": "H1"}
    # Both sides have data: as far as the longer run under the selection goes (Excel 16 for Mac);
    # a gap ends a run.
    for left, right, want in [(3, 6, "B1:B6"), (6, 3, "B1:B6"), (3, 2, "B1:B3")]:
        s = new()
        s.run("range.setValues", range="A1", values=[[i] for i in range(1, left + 1)])
        s.run("range.setValues", range="C1", values=[[i] for i in range(1, right + 1)])
        s.run("cell.set", cell="B1", input="=ROW()")
        s.run("selection.set", range="B1")
        assert s.run("edit.autoFillDown") == {"range": want}, (left, right)
    s = new()
    s.run("range.setValues", range="A1", values=[[1], [2], [None], [4], [5]])
    s.run("cell.set", cell="B1", input="=ROW()")
    s.run("selection.set", range="B1")
    assert s.run("edit.autoFillDown") == {"range": "B1:B2"}


def test_files_with_these_fixes_read_back_in_openpyxl():
    # External file links (#800), a header-and-totals table (#511), a column name with an edge space
    # (#732), functions by name (#902) and _xludf. (#726), read by another library.
    s = new()
    s.run("insert.hyperlink", cell="A1", target="Source.xlsx")
    s.run("insert.hyperlink", cell="A2", target="Sheet1!C3")
    s.run("range.setValues", range="C1", values=[["Name", "Qty "], ["a", 2], ["b", 3]])
    s.run("insert.table", range="C1:D3", name="Sales")
    s.run("table.totalRow", on=True, table="Sales")
    s.run("range.setValues", range="G1", values=[["H", "V"]])
    s.run("insert.table", range="G1:H1", name="Empty")
    s.run("table.totalRow", on=True, table="Empty")
    s.run("cell.set", cell="J1", input="=BYROW(D2:D3,SUM)")
    s.run("cell.set", cell="K1", input='=_xludf.CONCAT("a","b")')
    raw = base64.b64decode(s.run("file.saveBytes")["data"])
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ws = openpyxl.load_workbook(io.BytesIO(raw)).active
    assert ws["A1"].hyperlink.target == "Source.xlsx" and not ws["A1"].hyperlink.location
    assert ws["A2"].hyperlink.location == "Sheet1!C3"
    assert dict(ws.tables.items()) == {"Sales": "C1:D4", "Empty": "G1:H3"}
    assert [c.name for c in ws.tables["Sales"].tableColumns] == ["Name", "Qty "] and ws["D1"].value == "Qty "
    assert ws["D4"].value == "=SUBTOTAL(109,Sales[[Qty ]])"
    assert getattr(ws["J1"].value, "text", ws["J1"].value) == "=_xlfn.BYROW(D2:D3,_xleta.SUM)"  # a spill: an array formula
    assert ws["K1"].value == '=_xludf.CONCAT("a","b")'
