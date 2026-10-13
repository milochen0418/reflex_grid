"""Formula language tests."""

import pytest

from gridcore.core.addr import MAX_ROWS, RangeRef
from gridcore.formula import Binary, Call, Num, ParseError, Ref, StructRef, Unary, parse, print_formula, print_r1c1, reference_a1, tokenize
from gridcore.formula.adjust import Edit, adjust, delete_sheet, rename_sheet, shift_relative


def rt(src):
    return print_formula(parse(src))


def test_roundtrips():
    for f in [
        "1+2*3", "SUM(A1:B2)", "SUM($A$1:B$2,C3)", 'IF(A1>0,"yes","no")', "-2^2", "A1%",
        "Sheet2!A1+'My Sheet'!$B$3", "SUM(Sheet1:Sheet3!A1)", "SUM(A:A)", "SUM(3:5)", "{1,2;3,4}",
        '{-1,"a",TRUE,#N/A}', "VLOOKUP(A1,Table1[#All],2,FALSE)", "Table1[Sales]", "Table1[@Sales]",
        "Table1[[#Headers],[A]:[C]]", '"a""b"&C1', "SUM((A1:A3,C1:C3))", "A1:C3 B2:B5", "LAMBDA(x,x+1)(2)",
        "F(1,,3)", "@A1:A10", "A1<>B1", "A1<=B1", "#REF!+1", "1E+20", "MyName*2", "Sheet1!MyName", "SUM(A1:B2) ",
    ]:
        assert rt(f) == f.strip(), f


def test_normalises():
    assert rt("=sum(a1:b2)") == "SUM(A1:B2)"
    assert rt("= 1 + 2") == "1+2"
    assert rt("_xlfn.XLOOKUP(1,A:A,B:B)") == "XLOOKUP(1,A:A,B:B)"
    assert rt("true") == "TRUE"
    assert rt("'Sheet1'!A1") == "Sheet1!A1"


def test_spill_operator():
    for f in ["A1#", "Sheet1!$D$1#", "SUM(B2#)", "'My Sheet'!A1#*2", "ROWS(A1#)+1", "IF(A1,#N/A,#REF!)"]:
        assert rt(f) == f
    assert isinstance(parse("A1#"), Unary) and parse("A1#").op == "#"
    assert rt("_xlfn.ANCHORARRAY(Sheet1!$D$1)") == "Sheet1!$D$1#"
    assert rt("SUM(ANCHORARRAY(B2))") == "SUM(B2#)"
    with pytest.raises(ParseError):
        parse("A1:B2#")
    with pytest.raises(ParseError):
        parse("Name#")
    assert print_formula(delete_sheet(parse("SUM(Data!A1#)"), "Data", [])) == "SUM(#REF!)"


def test_delete_sheet_shrinks_3d_span():
    order = ["Sheet1", "Sheet2", "Sheet3", "Results"]

    def d(f, n):
        return print_formula(delete_sheet(parse(f), n, order))

    assert d("SUM(Sheet1:Sheet3!A1)", "Sheet1") == "SUM(Sheet2:Sheet3!A1)"
    assert d("SUM(Sheet1:Sheet3!A1)", "Sheet3") == "SUM(Sheet1:Sheet2!A1)"
    assert d("SUM(Sheet3:Sheet1!A1)", "Sheet1") == "SUM(Sheet3:Sheet2!A1)"
    assert d("SUM(Sheet1:Sheet3!A1)", "Sheet2") == "SUM(Sheet1:Sheet3!A1)"
    assert d("SUM(Sheet1:Sheet2!A1)", "Sheet1") == "SUM(Sheet2!A1)"
    assert d("Sheet1!A1+Sheet3!A1", "Sheet1") == "#REF!+Sheet3!A1"


def test_precedence():
    e = parse("1+2*3")
    assert isinstance(e, Binary) and e.op == "+"
    e = parse("-2^2")
    assert isinstance(e, Binary) and e.op == "^"
    e = parse('1&2="12"')
    assert isinstance(e, Binary) and e.op == "="
    e = parse("2^3^2")
    assert isinstance(e, Binary) and isinstance(e.left, Binary) and e.left.op == "^"


def test_errors():
    for bad in ["1+", "SUM(1,2", "(1", "{1,2;3}", '"abc', "A1:", "1 2", "{A1}", ")", "Sheet1!", "#FOO!"]:
        with pytest.raises(ParseError):
            parse(bad)


def test_no_crash_on_garbage():
    for s in ["(((((", "[[[", "'''", "!!!", "$$$", "=@@@", "A1:B2:C3:D4", "{,}", "{;}", "F(,,,,)", ":::", "1e", "1e+", ".", "..", "$A$", "XFD1048577"]:
        try:
            parse(s)
        except ParseError:
            pass
    deep = "(" * 5000 + "1" + ")" * 5000
    with pytest.raises(ParseError):
        parse(deep)


def test_references():
    e = parse("SUM(A1:B2)+Sheet2!C3*D4")
    assert [reference_a1(r) for r in e.references()] == ["A1:B2", "Sheet2!C3", "D4"]


def test_copy_shift():
    assert print_formula(shift_relative(parse("A1+$A1+A$1+$A$1+SUM(B2:C3)"), 2, 1)) == "B3+$A3+B$1+$A$1+SUM(C4:D5)"
    assert print_formula(shift_relative(parse("A1"), -1, 0)) == "#REF!"
    assert print_formula(shift_relative(parse("SUM(A:A)"), 5, 1)) == "SUM(B:B)"


def test_insert_delete():
    ins = Edit("insert", axis="rows", at=2, count=2)
    assert print_formula(adjust(parse("A1+A3+SUM(A2:A5)+Other!A3"), "Sheet1", "Sheet1", ins)) == "A1+A5+SUM(A2:A7)+Other!A3"
    dele = Edit("delete", axis="rows", at=2, count=2)
    assert print_formula(adjust(parse("A1+A3+SUM(A2:A5)+A6"), "Sheet1", "Sheet1", dele)) == "A1+#REF!+SUM(A2:A3)+A4"
    dele = Edit("delete", axis="cols", at=0, count=1)
    assert print_formula(adjust(parse("SUM(A1:C1)+B2"), "S", "S", dele)) == "SUM(A1:B1)+A2"
    assert print_formula(adjust(parse("Sheet1!B2"), "Other", "Sheet1", dele)) == "Sheet1!A2"
    assert print_formula(adjust(parse("SUM(A:A)"), "S", "S", dele)) == "SUM(#REF!)"


def test_insert_delete_cells():
    def adj(f, e):
        return print_formula(adjust(parse(f), "S", "S", e))

    ins = Edit("insert_cells", axis="rows", range=RangeRef.parse("A3:A4"))
    assert adj("SUM(A1:A5)+A2+A3+$A$9+B3+SUM(A1:B5)+SUM(A:A)", ins) == "SUM(A1:A7)+A2+A5+$A$11+B3+SUM(A1:B5)+SUM(A:A)"
    ins = Edit("insert_cells", axis="cols", range=RangeRef.parse("B2"))
    assert adj("A2+B2+C2+B3+SUM(A2:C2)", ins) == "A2+C2+D2+B3+SUM(A2:D2)"
    dele = Edit("delete_cells", axis="rows", range=RangeRef.parse("A3:A4"))
    assert adj("A2+A3+A4+A5+B5+SUM(A1:A10)+SUM(A3:A4)+SUM(A1:B10)", dele) == "A2+#REF!+#REF!+A3+B5+SUM(A1:A8)+SUM(#REF!)+SUM(A1:B10)"
    dele = Edit("delete_cells", axis="cols", range=RangeRef.parse("B1:C2"))
    assert adj("A1+B1+D1+D2+D3+SUM(A1:E1)", dele) == "A1+#REF!+B1+B2+D3+SUM(A1:C1)"


def test_moves_and_renames():
    mv = Edit("move", range=RangeRef.parse("A1:B2"), to_row=9, to_col=3)
    assert print_formula(adjust(parse("A1+B2+C3+SUM(A1:B2)"), "S", "S", mv)) == "D10+E11+C3+SUM(D10:E11)"
    mv = Edit("move", range=RangeRef.parse("B:C"), to_row=0, to_col=4)
    assert print_formula(adjust(parse("SUM(B:B)+SUM($C:$C)+D1"), "S", "S", mv)) == "SUM(E:E)+SUM($F:$F)+D1"
    assert print_formula(rename_sheet(parse("Old!A1+'Old'!B2"), "Old", "New Name")) == "'New Name'!A1+'New Name'!B2"
    assert print_formula(rename_sheet(parse("Data!MyName+MyName+Other!MyName"), "data", "Numbers")) == "Numbers!MyName+MyName+Other!MyName"
    assert print_formula(rename_sheet(parse("Data!MyName"), "Data", "New Name")) == "'New Name'!MyName"


def test_moves_between_sheets():
    mv = Edit("move_to_sheet", range=RangeRef.parse("A1:B2"), to_sheet="New Name", to_row=9, to_col=3)
    for host, inp, exp in [
        ("Source", "A1+$B$2+C3+SUM(A1:B2)", "'New Name'!D10+'New Name'!$E$11+C3+SUM('New Name'!D10:E11)"),
        ("Other", "A1+sOuRcE!$B2+Other!B2", "A1+'New Name'!$E11+Other!B2"),
        ("Source", "SUM(A1:C3)+SUM(Source:Other!A1)", "SUM(A1:C3)+SUM(Source:Other!A1)"),
    ]:
        assert print_formula(adjust(parse(inp), host, "Source", mv)) == exp
    edge = Edit("move_to_sheet", range=RangeRef.parse("A1:B2"), to_sheet="New Name", to_row=MAX_ROWS - 1, to_col=3)
    assert print_formula(adjust(parse("B2"), "Source", "Source", edge)) == "#REF!"


def test_r1c1():
    assert print_r1c1(parse("A1+$B$2+SUM(C:C)"), (1, 1)) == "R[-1]C[-1]+R2C2+SUM(C[1])"


def test_spans():
    toks = tokenize("SUM(A1, B2)")
    a1 = next(t for t in toks if t[0] == "word" and t[1] == "A1")
    assert (a1[2], a1[3]) == (4, 6)


def test_unbracketed_this_row_column():
    e = parse("[[#This Row],Units]*[[#This Row],Price]")
    assert isinstance(e, Binary) and isinstance(e.left, StructRef)
    assert e.left.specifiers == ("#This Row",) and e.left.col_start == "Units"
    assert rt("[[#This Row],Units]*[[#This Row],Price]") == "[@Units]*[@Price]"
    assert rt("Sales[[#This Row],Units]") == "Sales[@Units]"
    assert rt("Sales[[#This Row], Units ]") == "Sales[@Units]"
    assert rt("[[#This Row],Units]+[@[Unit Price]]+Sales[[#This Row],[Qty]]") == "[@Units]+[@[Unit Price]]+Sales[@Qty]"
    for bad in ["[[#This Row]x]", "[[#This Row],[A]]x", "[[#Data],Col]", "[[#This Row],A:B]", "[x[#This Row]]",
                "[[#This Row],]x", "Sales[[#All],[A]]junk", "Sales[[A],[B],[C]]", "Sales[@[A]x]"]:
        with pytest.raises(ParseError):
            parse(bad)


def test_trim_references():
    for f in ["SUM(A1:.E10)", "SUM(A1.:E10)", "SUM(A1.:.E10)", "SUM(A:.A)", "SUM($A.:.$C)", "ROWS(1:.5)",
              "ROWS($1.:$5)", "ROWS(1.:.5)", "Sheet2!$A$1:.B9", "Q1.2026!A1.:.B2", "'Q1 2026'!A1:.B2",
              "SUM(A1:.A9,B1.:B9)", "A1:.A9*2"]:
        assert rt(f) == f, f
    e = parse("SUM(A1:.E10)")
    assert isinstance(e, Call) and isinstance(e.args[0], Unary) and e.args[0].op == ":." and isinstance(e.args[0].expr, Ref)
    assert parse("1.5+.5") == Binary("+", Num(1.5), Num(0.5))
    for f in ["SUM(1:5)", "Q1.2026!A1:B2", "my.name+1", "SUM(A1:B2)*1.5", "1.:5"]:
        assert rt(f) == f, f
    assert rt("'Q1.2026'!A1:B2") == "Q1.2026!A1:B2"
    assert isinstance(parse("Q1.2026!A1:B2"), Ref)
    assert rt("SUM(_xlfn._TRO_TRAILING(A1:E10))") == "SUM(A1:.E10)"
    assert rt("_xlfn._TRO_LEADING(A:A)") == "A.:A"
    assert rt("_xlfn._TRO_ALL(Sheet1!A1:B2)") == "Sheet1!A1.:.B2"
    assert rt("_xlfn._TRO_ALL(B2)") == "TRIMRANGE(B2,3,3)"
    for bad in ["A1.:", ".:A1", "A1:.:B2", "A1.:.:B2"]:
        with pytest.raises(ParseError):
            parse(bad)
    assert print_r1c1(parse("SUM(A1:.B2)+ROWS(2.:2)"), (0, 0)) == "SUM(RC:.R[1]C[1])+ROWS(R[1].:R[1])"


def test_trim_references_adjust():
    ins = Edit("insert", axis="rows", at=2, count=2)
    assert print_formula(adjust(parse("SUM(A1:.A5)+SUM(B:.B)+ROWS(2.:.4)"), "S", "S", ins)) == "SUM(A1:.A7)+SUM(B:.B)+ROWS(2.:.6)"
    dele = Edit("delete", axis="cols", at=0, count=1)
    assert print_formula(adjust(parse("SUM(A1.:C1)+SUM(A:.A)"), "S", "S", dele)) == "SUM(A1.:B1)+SUM(#REF!)"
    assert print_formula(shift_relative(parse("SUM(A1.:.B2)"), 1, 1)) == "SUM(B2.:.C3)"
    assert [reference_a1(r) for r in parse("SUM(Sheet2!A1:.B5)").references()] == ["Sheet2!A1:B5"]
