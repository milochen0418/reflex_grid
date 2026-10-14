"""Excel object model cases: macros that use Range, Worksheet, Application… on a sheet of their
own, run by Microsoft Excel (``scripts/excel_vba_oracle_mac.py --objects``) and by ReflexGrid
(``test_vba_object_cases.py``).

Each case is the body of a Sub that starts on a new, empty, active sheet with A1 selected. It
reports what it sees with ``Out x, y…`` (``TypeName|text`` of each value) and leaves cells behind;
both are compared. ``evlog`` collects the workbook's Change/SelectionChange events.
"""

from __future__ import annotations

from typing import List, Tuple

from tests.vba_excel_cases import repr_vba

CASES: List[Tuple[str, str]] = [
    # ------------------------------------------------------------ writing values
    ("write number", '''
Range("A1").Value = 1.5
Out Range("A1").Value, Range("A1").Text, Range("A1").Formula, Range("A1").NumberFormat'''),
    ("write numeric strings", '''
Range("A1").Value = "00123"
Range("A2").Value = " 12 "
Range("A3").Value = "1,234"
Range("A4").Value = "50%"
Range("A5").Value = "$5"
Range("A6").Value = "1e3"
Range("A7").Value = "(5)"
Range("A8").Value = "-"
Range("A9").Value = "1 1/2"
Range("A10").Value = "-0"
Range("A11").Value = "+7"
Range("A12").Value = ".5"'''),
    ("write date strings", '''
Range("A1").Value = "1/2/2024"
Range("A2").Value = "2024-01-02"
Range("A3").Value = "13:45"
Range("A4").Value = "1/2/2024 3:04 PM"
Range("A5").Value = "Jan 2024"
Range("A6").Value = "2-Jan-2024"'''),
    ("write logical strings", '''
Range("A1").Value = "TRUE"
Range("A2").Value = "false"
Range("A3").Value = "yes"
Range("A4").Value = "#N/A"'''),
    ("write formula text", '''
Range("A1").Value = "=1+2"
Range("A2").Value = "=1/0"
Out Range("A1").Value, Range("A1").HasFormula, Range("A2").Text'''),
    ("apostrophe", '''
Range("A1").Value = "'123"
Out Range("A1").Value, Range("A1").Formula, Range("A1").Text, Range("A1").PrefixCharacter, Range("B1").PrefixCharacter'''),
    ("empty string", '''
Range("A1").Value = ""
Out IsEmpty(Range("A1").Value), Range("A1").Formula, WorksheetFunction.CountA(Range("A1")), WorksheetFunction.CountBlank(Range("A1"))'''),
    ("assign Empty and Null", '''
Range("A1").Value = 5
Range("A1").Value = Empty
Range("A2").Value = 5
Range("A2").Value = Null
Out IsEmpty(Range("A1").Value), IsEmpty(Range("A2").Value)'''),
    ("date value", '''
Range("A1").Value = DateSerial(2024, 1, 15)
Out Range("A1").Value, Range("A1").Value2, Range("A1").Text, Range("A1").NumberFormat'''),
    ("date and time value", '''
Range("A1").Value = DateSerial(2024, 1, 15) + TimeSerial(13, 30, 0)
Out Range("A1").Text, Range("A1").NumberFormat'''),
    ("time value", '''
Range("A1").Value = TimeSerial(13, 30, 15)
Out Range("A1").Text, Range("A1").NumberFormat, Range("A1").Value'''),
    ("date before 1900", '''
Range("A1").Value = DateSerial(1850, 1, 1)
Out TypeName(Range("A1").Value), Range("A1").NumberFormat'''),
    ("currency value", '''
Range("A1").Value = CCur(1234.5)
Out Range("A1").Text, Range("A1").NumberFormat, TypeName(Range("A1").Value), TypeName(Range("A1").Value2)'''),
    ("logical value", '''
Range("A1").Value = True
Out Range("A1").Value, Range("A1").Text, Range("A1").Formula'''),
    ("error value", '''
Range("A1").Value = CVErr(2042)
Out IsError(Range("A1").Value), Range("A1").Text, CStr(Range("A1").Value), Range("A1").Formula'''),
    ("formula error", '''
Range("A1").Formula = "=1/0"
Out Range("A1").Value = CVErr(xlErrDiv0), Range("A1").Text'''),
    ("long digit string", '''
Range("A1").Value = "123456789012345678"
Out Range("A1").Value, Range("A1").Text'''),
    ("text format", '''
Range("A1").NumberFormat = "@"
Range("A1").Value = "0123"
Range("A2").NumberFormat = "@"
Range("A2").Value = 5
Out TypeName(Range("A1").Value), TypeName(Range("A2").Value), Range("A2").Text'''),
    ("format kept on write", '''
Range("A1").NumberFormat = "0.0%"
Range("A1").Value = 0.256
Range("A2").NumberFormat = "m/d/yyyy"
Range("A2").Value = 45000
Out Range("A1").Text, TypeName(Range("A2").Value), Range("A2").Text'''),
    ("general text of numbers", '''
Range("A1").Value = 1 / 3
Range("A2").Value = 123456789012#
Range("A3").Value = 0.000001234
Range("A4").Value = -1234.5
Out Range("A1").Text, Range("A2").Text, Range("A3").Text, Range("A4").Text'''),
    ("integer types", '''
Range("A1").Value = CInt(5)
Range("A2").Value = CLng(7)
Range("A3").Value = CByte(9)
Range("A4").Value = CSng(1.5)
Out TypeName(Range("A1").Value), TypeName(Range("A2").Value), TypeName(Range("A3").Value), TypeName(Range("A4").Value)'''),
    ("formula result", '''
Range("A1").Formula = "=0.1+0.2"
Out Range("A1").Value = 0.3, Range("A1").Value'''),
    ("text of formats", '''
Range("A1").Value = 0.5
Range("A1").NumberFormat = "0%"
Range("A2").Value = 1234.5
Range("A2").NumberFormat = "#,##0.00"
Range("A3").Value = 12345
Range("A3").NumberFormat = "0.00E+00"
Range("A4").Value = DateSerial(2024, 3, 9)
Range("A4").NumberFormat = "mmm d, yyyy"
Range("A5").Value = 1.5
Range("A5").NumberFormat = "[h]:mm"
Range("A6").Value = -5
Range("A6").NumberFormat = "0;(0)"
Out Range("A1").Text, Range("A2").Text, Range("A3").Text, Range("A4").Text, Range("A5").Text, Range("A6").Text'''),
    # ------------------------------------------------------------ formulas
    ("Formula to R1C1", '''
Range("C3").Formula = "=A1+$B$2+B$1"
Out Range("C3").FormulaR1C1'''),
    ("R1C1 to Formula", '''
Range("C3").FormulaR1C1 = "=R[-1]C[-1]*2+R1C1"
Out Range("C3").Formula'''),
    ("formula on a range", '''
Range("A1:A3").Value = 1
Range("B1:B3").Formula = "=A1*2"
Out Range("B3").Formula, Range("B3").Value'''),
    ("formulas from an array", '''
Range("A1:B1").Formula = Array("=1+1", "x")
Out Range("A1").Value, Range("B1").Value'''),
    ("HasFormula", '''
Range("A1").Formula = "=1"
Range("A2").Value = 1
Out Range("A1:A2").HasFormula, Range("A1").HasFormula, Range("A2").HasFormula, Range("A3").HasFormula'''),
    ("Formula of constants", '''
Range("A1").Value = DateSerial(2024, 3, 4)
Range("A2").Value = 1234567.891
Range("A3").Value = "text"
Out Range("A1").Formula, Range("A2").Formula, Range("A3").Formula, Range("A4").Formula'''),
    ("formula with strings", '''
Range("A1").Formula = "=""a""&""b"""
Out Range("A1").Value, Range("A1").Formula'''),
    ("formula normalised", '''
Range("A1").Formula = "=sum(b1:b2)"
Range("A2").Formula = "= 1 + 2"
Out Range("A1").Formula, Range("A2").Formula'''),
    ("invalid formula", '''
On Error Resume Next
Range("A1").Formula = "=SUM(1,"
Out Err.Number, IsEmpty(Range("A1").Value)'''),
    ("unknown function", '''
Range("A1").Formula = "=FOOBAR(1)"
Out Range("A1").Text'''),
    ("FormulaArray", '''
Range("A1:A3").FormulaArray = "={1;2;3}"
Out Range("A2").Value, Range("A1").HasArray, Range("A2").FormulaArray, Range("A1").CurrentArray.Address'''),
    ("formula with sheet name", '''
Range("B1").Value = 4
Range("A1").Formula = "='" & ActiveSheet.Name & "'!B1+1"
Out Range("A1").Value, Range("A1").Formula = "=" & ActiveSheet.Name & "!B1+1"'''),
    # ------------------------------------------------------------ reading and writing blocks
    ("value of several areas", '''
Range("A1").Value = 1
Range("C1").Value = 2
Out Range("A1,C1").Value'''),
    ("2-D value", '''
Range("A1:B2").Value = 3
Dim v
v = Range("A1:B2").Value
Out TypeName(v), LBound(v, 1), UBound(v, 1), LBound(v, 2), UBound(v, 2), v(2, 2)
Out IsArray(Range("A1").Value), IsArray(Range("A1:A2").Value)'''),
    ("1-D array across", '''
Range("A1:C1").Value = Array(1, 2, 3)'''),
    ("1-D array down", '''
Range("A1:A3").Value = Array(1, 2, 3)'''),
    ("short array", '''
Range("A1:D1").Value = Array(1, 2)
Out Range("D1").Text'''),
    ("small 2-D array into a bigger range", '''
Dim v(1 To 2, 1 To 1)
v(1, 1) = 1
v(2, 1) = 2
Range("A1:B3").Value = v'''),
    ("0-based 2-D array", '''
Dim w(0 To 1, 0 To 1)
w(0, 0) = "a"
w(1, 1) = "d"
Range("A1:B2").Value = w
Out IsEmpty(Range("B1").Value)'''),
    ("Transpose", '''
Range("A1:A3").Value = Application.Transpose(Array("x", "y", "z"))
Dim v
v = Application.Transpose(Range("A1:A3").Value)
Out TypeName(v), LBound(v), UBound(v), v(3)'''),
    ("range to range", '''
Range("A1:A2").Value = 7
Range("C1:C2").Value = Range("A1:A2").Value
Range("A4").Value = 4
Range("B4") = Range("A4")
Range("D1:D2").Value = Range("A1:A2")'''),
    ("range object as a value", '''
Range("A1:A2").Value = 7
Range("D1:D2").Value = 9
Range("D1:D2").Value = Range("A1:A2")
Range("E1:E2").Value = 9
Range("E1:E2").Value2 = Range("A1:A2")
Range("F1") = Range("A1:A2")'''),
    ("empty cells in arrays", '''
Dim v
v = Range("A1:B2").Value
Out IsEmpty(v(1, 1)), TypeName(v(1, 1))'''),
    ("Value and Value2 of currency", '''
Range("A1").Value = 5
Range("A1").NumberFormat = "$#,##0.00"
Out TypeName(Range("A1").Value), TypeName(Range("A1").Value2)'''),
    ("dates in arrays", '''
Range("A1:B1").Value = Array(DateSerial(2024, 1, 1), CCur(2.5))
Out TypeName(Range("A1").Value), TypeName(Range("B1").Value)'''),
    ("scalar to a block", '''
Range("A1:B2").Value = "x"
Range("C1:C2").Value = Range("A1").Value'''),
    ("Text of several cells", '''
Range("A1").Value = 1
Range("A2").Value = 2
Range("B1:B2").Value = 3
Out Range("A1:A2").Text, Range("B1:B2").Text'''),
    # ------------------------------------------------------------ navigation
    ("Cells and Item", '''
Out Cells(2, 3).Address, Range("B2:C3").Cells(2, 2).Address, Range("B2:C3")(3).Address, Range("B2:C3").Item(5).Address, Range("B2").Cells(0, 0).Address, Cells(1, "C").Address'''),
    ("Offset and Resize", '''
Out Range("B2").Offset(1, 2).Address, Range("B2:C3").Offset(-1).Address, Range("B2").Resize(2, 3).Address, Range("B2:C3").Resize(, 1).Address, Range("B2").Offset(, 1).Address'''),
    ("Offset off the sheet", '''
On Error Resume Next
Dim r As Range
Set r = Range("A1").Offset(-1, 0)
Out Err.Number, r Is Nothing'''),
    ("Rows and Columns", '''
Out Rows.Count, Columns.Count, Range("A1:C5").Rows.Count, Range("A1:C5").Columns(2).Address, Range("A1:C5").Rows(2).Address, Columns("B:C").Address, Rows("2:3").Address, Range("B:B").Column'''),
    ("EntireRow and EntireColumn", '''
Out Range("B2:C3").EntireRow.Address, Range("B2").EntireColumn.Address, Range("B2").EntireRow.Rows.Count, Range("B2").EntireRow.Columns.Count'''),
    ("End", '''
Range("A1:A5").Value = 1
Range("C1").Value = 1
Out Range("A1").End(xlDown).Address, Range("A5").End(xlUp).Address, Range("A1").End(xlToRight).Address, Range("A10").End(xlUp).Address
Out Range("A7").End(xlDown).Address, Range("D1").End(xlToLeft).Address, Cells(Rows.Count, 1).End(xlUp).Row, Range("E1").End(xlToRight).Address'''),
    ("End with gaps", '''
Range("A1,A2,A4,A5").Value = 1
Out Range("A1").End(xlDown).Address, Range("A2").End(xlDown).Address, Range("A3").End(xlDown).Address, Range("A5").End(xlUp).Address'''),
    ("CurrentRegion", '''
Range("A1:B3").Value = 1
Range("C4").Value = 1
Range("F6").Value = 1
Out Range("A1").CurrentRegion.Address, Range("B2").CurrentRegion.Address, Range("F6").CurrentRegion.Address, Range("H10").CurrentRegion.Address'''),
    ("UsedRange", '''
Range("B2").Value = 1
Range("D5").Value = 1
Out ActiveSheet.UsedRange.Address, ActiveSheet.UsedRange.Rows.Count'''),
    ("Address forms", '''
Out Range("B2:C3").Address(0, 0), Range("B2:C3").Address(True, False), Range("B2:C3").Address(ReferenceStyle:=xlR1C1), Range("B2:C3").Address(0, 0, xlR1C1, False, Range("A1")), Range("A1:B2,D4").Address, Range("B2").Address(RowAbsolute:=False)'''),
    ("Count", '''
Out Range("A1:C3").Count, Range("A:A").Count, Range("A:B").Count, Cells.CountLarge
On Error Resume Next
Dim n
n = Cells.Count
Out Err.Number'''),
    ("Areas", '''
Dim r As Range
Set r = Range("A1:B2,D4:D5")
Out r.Areas.Count, r.Areas(2).Address, r.Count, r.Rows.Count, r.Cells(3).Address, r.Row, r.Column'''),
    ("Union and Intersect", '''
Out Union(Range("A1:A2"), Range("A2:A3")).Address, Union(Range("A1"), Range("A2")).Address, Intersect(Range("A1:C3"), Range("B2:D4")).Address
Out Intersect(Range("A1"), Range("B2")) Is Nothing, Union(Range("A1:A2"), Range("A2:A3")).Count, Union(Range("A1:B1"), Range("A2:B2")).Address'''),
    ("Range of two corners", '''
Out Range("B2", "D4").Address, Range(Cells(1, 1), Cells(2, 3)).Address, Range(Range("C3"), Range("A1")).Address, Range("A1:B2", "D4").Address'''),
    ("Range relative to a range", '''
Out Range("B2:C3").Range("A1").Address, Range("B2").Range("B2").Address, Range("B2:D4").Range("B1:C2").Address, Range("C3").Range("A1:B2").Cells.Count'''),
    ("invalid references", '''
On Error Resume Next
Dim r As Range
Set r = Range("A0")
Out Err.Number
Err.Clear
Set r = Range("XFE1")
Out Err.Number
Err.Clear
Set r = Range("")
Out Err.Number
Err.Clear
Set r = Cells(0, 1)
Out Err.Number'''),
    ("For Each order", '''
Dim c As Range, s As String
For Each c In Range("A1:B2")
    s = s & c.Address(0, 0) & ","
Next
For Each c In Range("A1:B2").Columns
    s = s & c.Address(0, 0) & ";"
Next
For Each c In Range("A1:B2").Rows
    s = s & c.Address(0, 0) & "|"
Next
For Each c In Range("A1,C1:C2")
    s = s & c.Address(0, 0) & "."
Next
Out s'''),
    ("rows and columns of a block", '''
Out Range("A1:C3").Rows(2).Cells(1).Address, Range("A1:C3").Columns(2).Cells(2).Address, Range("A1:C3").Columns(1).Rows.Count, Range("A1:C3").Columns.Item(3).Address, Range("A1:C3").Rows(2).Columns(3).Address'''),
    ("positions", '''
Out Range("C5:D6").Row, Range("C5:D6").Column, Range("C5:D6").Rows(2).Row, Range("C5").Height, Range("C5").Width, Range("B1").Left, Range("A2").Top, Range("B2:C3").Width'''),
    ("last cell idioms", '''
Range("A1:C1").Value = 1
Range("A2:A10").Value = 2
Range("E3").Value = 3
Out Cells(Rows.Count, "A").End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column, ActiveSheet.UsedRange.Rows.Count
Out Cells.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Out Cells.SpecialCells(xlCellTypeLastCell).Address'''),
    # ------------------------------------------------------------ formats
    ("Font", '''
With Range("A1").Font
    .Bold = True
    .Italic = True
    .Size = 14
    .Name = "Arial"
    .Color = RGB(255, 0, 0)
    .Underline = xlUnderlineStyleSingle
End With
With Range("A1").Font
    Out .Bold, .Italic, .Size, .Name, .Color, .Underline, .ColorIndex, .Strikethrough
End With'''),
    ("Font of several cells", '''
Range("A1").Font.Bold = True
Range("B1:B2").Font.Size = 9
Out Range("A1:A2").Font.Bold, Range("B1:B2").Font.Size, Range("A1:B1").Font.Size'''),
    ("Interior", '''
Range("A1").Interior.Color = RGB(0, 255, 0)
Out Range("A1").Interior.Color, Range("A1").Interior.ColorIndex, Range("A1").Interior.Pattern, Range("B1").Interior.Color, Range("B1").Interior.ColorIndex, Range("B1").Interior.Pattern'''),
    ("Interior by index", '''
Range("A1").Interior.ColorIndex = 3
Range("A2").Interior.ColorIndex = 6
Range("A3").Interior.Color = 255
Range("A3").Interior.ColorIndex = xlNone
Out Range("A1").Interior.Color, Range("A2").Interior.Color, Range("A3").Interior.ColorIndex, Range("A3").Interior.Color'''),
    ("NumberFormat", '''
Out Range("A1").NumberFormat
Range("A1").NumberFormat = "0.00"
Out Range("A1").NumberFormat, Range("A1:A2").NumberFormat, Range("C1:C2").NumberFormat'''),
    ("alignment", '''
Range("A1").HorizontalAlignment = xlRight
Range("A2").VerticalAlignment = xlTop
Range("A3").WrapText = True
Range("A4").IndentLevel = 2
Out Range("A1").HorizontalAlignment, Range("B1").HorizontalAlignment, Range("A2").VerticalAlignment, Range("B2").VerticalAlignment
Out Range("A3").WrapText, Range("B3").WrapText, Range("A4").IndentLevel, Range("A4").HorizontalAlignment'''),
    ("borders", '''
Range("B2").Borders(xlEdgeTop).LineStyle = xlContinuous
Out Range("B2").Borders(xlEdgeTop).LineStyle, Range("B2").Borders(xlEdgeBottom).LineStyle, Range("B2").Borders(xlEdgeTop).Weight
Range("D4").BorderAround xlContinuous, xlMedium
Out Range("D4").Borders(xlEdgeLeft).Weight, Range("D4").Borders(xlEdgeRight).LineStyle'''),
    ("all borders", '''
Range("A1:B2").Borders.LineStyle = xlContinuous
Out Range("A1").Borders(xlEdgeRight).LineStyle, Range("B2").Borders(xlEdgeLeft).LineStyle, Range("A1:B2").Borders(xlInsideVertical).LineStyle, Range("B1").Borders(xlEdgeBottom).Weight'''),
    ("Merge", '''
Range("A1:B2").Merge
Out Range("A1").MergeCells, Range("B2").MergeCells, Range("B2").MergeArea.Address, Range("A1:C1").MergeCells, Range("A1:B2").MergeCells, Range("C3").MergeArea.Address
Range("A1").UnMerge
Out Range("A1").MergeArea.Address'''),
    ("Merge keeps the first value", '''
Range("A1").Value = "x"
Range("B1").Value = "y"
Range("A1:B1").Merge
Out Range("A1").Value, IsEmpty(Range("B1").Value)'''),
    ("Merge across", '''
Range("A1:B2").Merge True
Out Range("A1").MergeArea.Address, Range("A2").MergeArea.Address'''),
    ("value of a merged area", '''
Range("A1:B1").Merge
Range("A1:B1").Value = 5
Out Range("A1").Value, IsEmpty(Range("B1").Value)'''),
    ("ColumnWidth", '''
Out Columns(1).ColumnWidth, Range("A1").Width
Columns(1).ColumnWidth = 12
Columns("B:C").ColumnWidth = 5
Out Columns(1).ColumnWidth, Range("A1").Width, Columns("B").ColumnWidth, Range("B1:C1").ColumnWidth, Range("A1:B1").ColumnWidth'''),
    ("RowHeight", '''
Out Rows(1).RowHeight
Rows(1).RowHeight = 20.25
Out Rows(1).RowHeight, Range("A1:A2").RowHeight, Range("A1").Height'''),
    ("Hidden", '''
Rows(2).Hidden = True
Columns(2).Hidden = True
Out Rows(2).Hidden, Range("A2").EntireRow.Hidden, Rows(2).RowHeight, Columns(2).Hidden, Columns(2).ColumnWidth, Rows(3).Hidden
Rows(2).Hidden = False
Out Rows(2).RowHeight'''),
    ("AutoFit", '''
Range("A1").Value = "a fairly long text value in one cell"
Columns(1).AutoFit
Out Columns(1).ColumnWidth > 20, Columns(1).ColumnWidth < 60'''),
    ("Locked", '''
Out Range("A1").Locked, Range("A1").FormulaHidden
Range("A1").Locked = False
Out Range("A1").Locked, Range("A1:A2").Locked'''),
    # ------------------------------------------------------------ editing
    ("ClearContents keeps formats", '''
Range("A1").Value = 5
Range("A1").NumberFormat = "0.00"
Range("A1").Font.Bold = True
Range("A1").ClearContents
Out IsEmpty(Range("A1").Value), Range("A1").NumberFormat, Range("A1").Font.Bold'''),
    ("ClearFormats keeps values", '''
Range("A1").Value = 5
Range("A1").NumberFormat = "0.00"
Range("A1").Font.Bold = True
Range("A1").ClearFormats
Out Range("A1").Value, Range("A1").NumberFormat, Range("A1").Font.Bold'''),
    ("Clear", '''
Range("A1").Value = 5
Range("A1").Font.Bold = True
Range("A1").Clear
Out IsEmpty(Range("A1").Value), Range("A1").Font.Bold'''),
    ("Delete shift up", '''
Range("A1:A5").Value = Application.Transpose(Array(1, 2, 3, 4, 5))
Range("C1").Formula = "=SUM(A1:A5)"
Range("A2").Delete Shift:=xlUp
Out Range("A2").Value, Range("C1").Formula, Range("C1").Value'''),
    ("Delete shift left", '''
Range("A1:D1").Value = Array(1, 2, 3, 4)
Range("B1").Delete xlToLeft
Out Range("B1").Value, IsEmpty(Range("D1").Value)'''),
    ("deleting a referenced row", '''
Range("A1:A5").Value = 1
Range("C1").Formula = "=A5"
Rows(5).Delete
Out Range("C1").Formula'''),
    ("Insert shift down", '''
Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
Range("C1").Formula = "=A2"
Range("A2").Insert xlShiftDown
Out Range("A3").Value, Range("C1").Formula, IsEmpty(Range("A2").Value)'''),
    ("inserted rows take the format above", '''
Range("A1").Font.Bold = True
Range("A2").Value = 2
Rows(2).Insert
Out Range("A2").Font.Bold, Range("A3").Value'''),
    ("insert columns", '''
Range("A1:C1").Value = Array(1, 2, 3)
Range("A2").Formula = "=C1"
Columns("B").Insert
Out Range("C1").Value, Range("A2").Formula, IsEmpty(Range("B1").Value)'''),
    ("Copy to a destination", '''
Range("A1").Value = 1
Range("A2").Formula = "=A1+$A$1"
Range("A1:A2").Copy Range("C3")
Out Range("C4").Formula, Range("C4").Value'''),
    ("Copy to a bigger range", '''
Range("A1").Value = 1
Range("A1").Copy Range("B1:B3")
Range("A1:A2").Value = 2
Range("A1:A2").Copy Range("C1:C5")'''),
    ("Copy rows", '''
Range("A1:C1").Value = Array(1, 2, 3)
Range("A1").Font.Bold = True
Rows(1).Copy Rows(3)
Out Range("A3").Font.Bold'''),
    ("Cut moves references", '''
Range("A1").Value = 3
Range("B1").Formula = "=A1*2"
Range("A1").Cut Range("C1")
Out Range("B1").Formula, Range("B1").Value, IsEmpty(Range("A1").Value)'''),
    ("PasteSpecial values", '''
Range("A1").Value = 1
Range("A2").Formula = "=A1+1"
Range("A1:A2").Copy
Range("C1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Out Range("C2").Formula, Range("C2").HasFormula'''),
    ("PasteSpecial formats", '''
Range("A1").Font.Bold = True
Range("A1").NumberFormat = "0.0"
Range("C1").Value = 5
Range("A1").Copy
Range("C1").PasteSpecial xlPasteFormats
Application.CutCopyMode = False
Out Range("C1").Font.Bold, Range("C1").NumberFormat, Range("C1").Value'''),
    ("PasteSpecial transpose", '''
Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
Range("A1:A3").Copy
Range("C1").PasteSpecial Paste:=xlPasteAll, Transpose:=True
Application.CutCopyMode = False'''),
    ("PasteSpecial add", '''
Range("A1").Value = 5
Range("B1").Value = 1
Range("A1").Copy
Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd
Application.CutCopyMode = False
Out Range("B1").Value'''),
    ("FillDown and FillRight", '''
Range("A1").Value = 1
Range("B1").Formula = "=A1*2"
Range("A1:B3").FillDown
Range("D1").Formula = "=D2"
Range("D1:F1").FillRight
Out Range("B3").Formula, Range("F1").Formula'''),
    ("AutoFill series", '''
Range("A1").Value = 1
Range("A2").Value = 2
Range("A1:A2").AutoFill Range("A1:A6")
Range("B1").Value = 1
Range("B1").AutoFill Range("B1:B3")
Range("C1").Value = 1
Range("C1").AutoFill Range("C1:C3"), xlFillSeries'''),
    ("AutoFill text", '''
Range("A1").Value = "Item1"
Range("A1").AutoFill Range("A1:A3")
Range("B1").Value = "Mon"
Range("B1").AutoFill Range("B1:B3")
Range("C1").Value = "Jan"
Range("C1").AutoFill Range("C1:C3")
Range("D1").Value = "x"
Range("D1").AutoFill Range("D1:D3")'''),
    ("AutoFill dates and formulas", '''
Range("A1").Value = DateSerial(2024, 1, 31)
Range("A1").AutoFill Range("A1:A3")
Range("B1").Formula = "=ROW()*10"
Range("B1").AutoFill Range("B1:B3")
Out Range("A3").Text, Range("B3").Formula'''),
    # ------------------------------------------------------------ sorting, finding
    ("Sort with a header", '''
Range("A1:B1").Value = Array("Name", "Score")
Range("A2:A4").Value = Application.Transpose(Array("c", "a", "b"))
Range("B2:B4").Value = Application.Transpose(Array(3, 1, 2))
Range("A1:B4").Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes'''),
    ("Sort with two keys", '''
Range("A1:A5").Value = Application.Transpose(Array("x", "y", "x", "y", "x"))
Range("B1:B5").Value = Application.Transpose(Array(3, 1, 2, 5, 1))
Range("A1:B5").Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("B1"), Order2:=xlDescending, Header:=xlNo'''),
    ("Sort mixed types", '''
Range("A1:A6").Value = Application.Transpose(Array(3, "b", True, 1, "a", "10"))
Range("A1:A6").Sort Key1:=Range("A1"), Header:=xlNo'''),
    ("Sort keeps row formulas", '''
Range("A1:A3").Value = Application.Transpose(Array(3, 1, 2))
Range("B1:B3").Formula = "=A1*10"
Range("A1:B3").Sort Key1:=Range("A1"), Header:=xlNo
Out Range("B1").Formula, Range("B1").Value'''),
    ("Sort object", '''
Range("A1:B1").Value = Array("k", "v")
Range("A2:B4").Value = 0
Range("A2:A4").Value = Application.Transpose(Array("b", "c", "a"))
Range("B2:B4").Value = Application.Transpose(Array(1, 2, 3))
With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("A2:A4"), Order:=xlAscending
    .SetRange Range("A1:B4")
    .Header = xlYes
    .Apply
End With'''),
    ("Sort left to right", '''
Range("A1:C1").Value = Array(3, 1, 2)
Range("A2:C2").Value = Array("c", "a", "b")
Range("A1:C2").Sort Key1:=Range("A1"), Order1:=xlAscending, Orientation:=xlLeftToRight'''),
    ("RemoveDuplicates", '''
Range("A1:B1").Value = Array("k", "v")
Range("A2:B2").Value = Array("x", 1)
Range("A3:B3").Value = Array("x", 2)
Range("A4:B4").Value = Array("x", 1)
Range("A5:B5").Value = Array("y", 1)
Range("A1:B5").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes'''),
    ("Find starts after the first cell", '''
Range("A1:A3").Value = Application.Transpose(Array("abc", "b", "xbx"))
Dim c As Range, s As String, first As String
Set c = Range("A1:A3").Find("b", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
first = c.Address
Do
    s = s & c.Address(0, 0) & ","
    Set c = Range("A1:A3").FindNext(c)
Loop While c.Address <> first
Out s'''),
    ("Find options", '''
Range("A1:A4").Value = Application.Transpose(Array("Apple", "apple", "pineapple", "APPLE PIE"))
Out Range("A1:A4").Find("apple", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Address
Out Range("A1:A4").Find("apple", LookAt:=xlWhole, MatchCase:=False).Address
Out Range("A1:A4").Find("pie", After:=Range("A2"), LookAt:=xlPart, MatchCase:=False).Address
Out Range("A1:A4").Find("apple", SearchDirection:=xlPrevious, LookAt:=xlPart, MatchCase:=False).Address
Out Range("A1:A4").Find("a*e", LookAt:=xlWhole, MatchCase:=False).Address
Out Range("A1:A4").Find("zzz", LookAt:=xlPart, MatchCase:=False) Is Nothing
Out Range("A1:A4").Find("PIE").Address'''),
    ("Find in values or formulas", '''
Range("A1").Formula = "=1+1"
Range("A2").Value = 10
Out Range("A1:A2").Find("2", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, Range("A1:A2").Find("2", LookIn:=xlFormulas) Is Nothing, Range("A1:A2").Find("1", LookIn:=xlFormulas).Address
Out Range("A1:A2").Find(10, LookAt:=xlWhole).Address'''),
    ("Replace", '''
Range("A1:A3").Value = Application.Transpose(Array("abc", "b", "Bob"))
Out Range("A1:A3").Replace("b", "X", LookAt:=xlPart, MatchCase:=False)
Range("B1:B2").Value = Application.Transpose(Array("cat", "concat"))
Range("B1:B2").Replace What:="cat", Replacement:="dog", LookAt:=xlWhole, MatchCase:=False'''),
    ("SpecialCells", '''
Range("A1").Value = 1
Range("A2").Formula = "=A1"
Range("A4").Value = "x"
Range("C1").Value = 2
Out ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants).Address, ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas).Address
Out ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Address, ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers).Address, ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues).Address'''),
    ("SpecialCells finds nothing", '''
On Error Resume Next
Dim r As Range
Range("A1").Value = 1
Set r = Range("A1").SpecialCells(xlCellTypeFormulas)
Out Err.Number, r Is Nothing'''),
    ("TextToColumns", '''
Range("A1").Value = "a,b,c"
Range("A2").Value = "1,2,3"
Range("A1:A2").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, Comma:=True
Out TypeName(Range("C2").Value)'''),
    # ------------------------------------------------------------ notes, links, rules, names
    ("comments", '''
Range("A1").AddComment "hello"
Out Range("A1").Comment.Text, ActiveSheet.Comments.Count, Range("B1").Comment Is Nothing
On Error Resume Next
Range("A1").AddComment "again"
Out Err.Number
On Error GoTo 0
Range("A1").ClearComments
Out ActiveSheet.Comments.Count'''),
    ("hyperlinks", '''
ActiveSheet.Hyperlinks.Add Anchor:=Range("A1"), Address:="https://example.com/", TextToDisplay:="Example"
Out Range("A1").Value, Range("A1").Hyperlinks.Count, Range("A1").Hyperlinks(1).Address, ActiveSheet.Hyperlinks.Count
Range("A1").Hyperlinks.Delete
Out ActiveSheet.Hyperlinks.Count, Range("A1").Value'''),
    ("validation", '''
Range("A1").Validation.Add Type:=xlValidateList, Formula1:="a,b,c"
Out Range("A1").Validation.Formula1, Range("A1").Validation.Type'''),
    ("conditional formats", '''
Range("A1:A3").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1"
Out Range("A1:A3").FormatConditions.Count, Range("A1:A3").FormatConditions(1).Formula1, Range("A1:A3").FormatConditions(1).Type, Range("A1:A3").FormatConditions(1).Operator
Range("A1:A3").FormatConditions.Delete
Out Range("A1:A3").FormatConditions.Count'''),
    ("names", '''
Dim nm As String
nm = "Data" & ActiveSheet.Index
ActiveWorkbook.Names.Add Name:=nm, RefersTo:="='" & ActiveSheet.Name & "'!$A$1:$A$3"
Range(nm).Value = 2
Out Range(nm).Address, Names(nm).RefersTo = "=" & ActiveSheet.Name & "!$A$1:$A$3", Names(nm).RefersToRange.Count, Application.Sum(Range(nm))
Names(nm).Delete'''),
    ("constant names", '''
ActiveWorkbook.Names.Add Name:="TaxRate", RefersTo:="=0.05"
Range("A1").Formula = "=TaxRate*100"
Out Evaluate("TaxRate"), [TaxRate], Names("TaxRate").Value, Range("A1").Value
Names("TaxRate").Delete'''),
    ("naming a range", '''
Range("A1:A2").Name = "Pair" & ActiveSheet.Index
Out Range("Pair" & ActiveSheet.Index).Address, Range("A1:A2").Name.Name = "Pair" & ActiveSheet.Index
ActiveWorkbook.Names("Pair" & ActiveSheet.Index).Delete'''),
    # ------------------------------------------------------------ Evaluate and worksheet functions
    ("Evaluate", '''
Range("A1:A3").Value = 2
Out Evaluate("SUM(A1:A3)"), ActiveSheet.Evaluate("A1*3"), [A1:A3].Count, Evaluate("A1:A3").Address, Evaluate("=1/0"), Evaluate("""abc"""), [SUM(A1:A3)]'''),
    ("WorksheetFunction", '''
Range("A1:A5").Value = Application.Transpose(Array(1, 2, 3, 4, 5))
With WorksheetFunction
    Out .Sum(Range("A1:A5")), .Average(Range("A1:A5")), .CountIf(Range("A1:A5"), ">2"), .SumIf(Range("A1:A5"), "<3"), .Index(Range("A1:A5"), 3)
    Out .Match(4, Range("A1:A5"), 0), .VLookup(3, Range("A1:A5"), 1, False), .Max(Range("A1:A5"), 10), .Large(Range("A1:A5"), 2), .Round(2.5, 0)
    Out .Text(1234.5, "#,##0.00"), .Trim("  a   b  "), .Proper("hello world"), .Substitute("aXa", "X", "Y"), .CountA(Range("A:A")), .Ceiling(7, 5)
End With'''),
    ("WorksheetFunction errors", '''
Range("A1:A3").Value = 1
On Error Resume Next
Dim x
x = WorksheetFunction.VLookup(99, Range("A1:A3"), 1, False)
Out Err.Number
Err.Clear
x = Application.VLookup(99, Range("A1:A3"), 1, False)
Out Err.Number, IsError(x), CStr(x)
x = Application.Match("zz", Range("A1:A3"), 0)
Out CStr(x)
x = WorksheetFunction.Sum(Range("A1:A3"), "x")
Out Err.Number'''),
    ("worksheet functions on arrays", '''
Out Application.Sum(Array(1, 2, 3)), Application.Max(Array(1, 5, 2)), Application.Index(Array(10, 20, 30), 2), Application.Match("b", Array("a", "b"), 0)
Dim t
t = Application.Transpose(Array(1, 2))
Out LBound(t, 1), UBound(t, 1), t(2, 1)
Range("A1:B2").Value = 4
Out Application.Index(Range("A1:B2").Value, 2, 1), Application.CountIf(Range("A1:B2"), 4)'''),
    ("VBA and Excel rounding", '''
Out Round(2.5), Application.Round(2.5, 0), Round(3.5), WorksheetFunction.RoundDown(-2.5, 0), Application.WorksheetFunction.Floor(7, 2)'''),
    # ------------------------------------------------------------ sheets
    ("adding sheets", '''
Dim n As Long, ws As Worksheet, home As Worksheet
Set home = ActiveSheet
n = Worksheets.Count
Set ws = Worksheets.Add
Out Worksheets.Count - n, ws.Index = home.Index - 1, ActiveSheet.Name = ws.Name, ws.Next.Name = home.Name
ws.Delete
Out Worksheets.Count - n, ActiveSheet.Name = home.Name'''),
    ("invalid sheet names", '''
On Error Resume Next
ActiveSheet.Name = "a/b"
Out Err.Number
Err.Clear
ActiveSheet.Name = "Results"
Out Err.Number
Err.Clear
ActiveSheet.Name = String(32, "x")
Out Err.Number
Err.Clear
ActiveSheet.Name = "results"
Out Err.Number'''),
    ("copying a sheet", '''
Dim home As Worksheet
Set home = ActiveSheet
Range("A1").Value = 5
ActiveSheet.Copy After:=ActiveSheet
Out ActiveSheet.Name = home.Name & " (2)", ActiveSheet.Range("A1").Value, ActiveSheet.Index = home.Index + 1
ActiveSheet.Delete
home.Activate'''),
    ("sheet order", '''
Out ActiveSheet.Index = Worksheets.Count, ActiveSheet.Next Is Nothing, ActiveSheet.Previous.Index = ActiveSheet.Index - 1, Worksheets(1).Name, Sheets(ActiveSheet.Index).Name = ActiveSheet.Name'''),
    ("hiding sheets", '''
Dim ws As Worksheet, home As Worksheet
Set home = ActiveSheet
Set ws = Worksheets.Add(After:=home)
ws.Visible = xlSheetHidden
Out ws.Visible, ActiveSheet.Name = home.Name
ws.Visible = xlSheetVeryHidden
Out ws.Visible
ws.Visible = xlSheetVisible
ws.Delete'''),
    ("type names", '''
Out TypeName(Sheets(1)), TypeName(ActiveSheet), TypeName(ActiveWorkbook), TypeName(Application), TypeName(Range("A1")), TypeName(Range("A1").Font), TypeName(ActiveSheet.Rows)
Out TypeName(Selection), TypeName(Worksheets), TypeName(Sheets), TypeName(Application.WorksheetFunction), TypeName(ActiveWorkbook.Names), TypeName(Range("A1").Interior), TypeName(Range("A1").Borders), TypeName(Columns(1))'''),
    ("protection", '''
ActiveSheet.Protect
On Error Resume Next
Range("A1").Value = 1
Out Err.Number, ActiveSheet.ProtectContents, IsEmpty(Range("A1").Value)
Err.Clear
On Error GoTo 0
ActiveSheet.Unprotect
Range("A1").Value = 1
Out ActiveSheet.ProtectContents, Range("A1").Value'''),
    ("unlocked cells on a protected sheet", '''
Range("A1").Locked = False
ActiveSheet.Protect Password:="pw"
On Error Resume Next
Range("A1").Value = 1
Range("A2").Value = 2
Out Err.Number, Range("A1").Value, IsEmpty(Range("A2").Value)
On Error GoTo 0
ActiveSheet.Unprotect "pw"'''),
    ("protection for the user only", '''
ActiveSheet.Protect UserInterfaceOnly:=True
On Error Resume Next
Range("A1").Value = 1
Out Err.Number, Range("A1").Value
On Error GoTo 0
ActiveSheet.Unprotect'''),
    # ------------------------------------------------------------ selection
    ("Select and Activate", '''
Range("B2:C3").Select
Out Selection.Address, ActiveCell.Address
Range("C3").Activate
Out ActiveCell.Address, Selection.Address
Range("E5").Activate
Out ActiveCell.Address, Selection.Address'''),
    ("moving the selection", '''
Range("A1").Select
ActiveCell.Offset(2, 1).Select
Out Selection.Address
Range("A1,C3").Select
Out Selection.Areas.Count, ActiveCell.Address
Application.Goto Range("D4")
Out ActiveCell.Address, Selection.Address'''),
    ("writing through the selection", '''
Range("B2:B3").Select
Selection.Value = 9
ActiveCell.Offset(0, 1).Value = "x"
Selection.Font.Bold = True
Out Range("B3").Font.Bold'''),
    # ------------------------------------------------------------ events and calculation
    ("events of macro changes", '''
evlog = ""
Range("A1").Value = 1
Range("B1:C2").Value = 2
Range("D1").ClearContents
Range("E1").Formula = "=1"
Range("A1").Font.Bold = True
Range("A1").Select
Out evlog'''),
    ("EnableEvents", '''
evlog = ""
Application.EnableEvents = False
Range("A1").Value = 1
Application.EnableEvents = True
Range("A2").Value = 2
Out evlog'''),
    ("manual calculation", '''
Application.Calculation = xlCalculationManual
Range("A1").Value = 1
Range("A2").Formula = "=A1*2"
Range("A1").Value = 5
Out Range("A2").Value
Application.Calculate
Out Range("A2").Value
Application.Calculation = xlCalculationAutomatic
Out Range("A2").Value'''),
    # ------------------------------------------------------------ tables and filters
    ("tables", '''
Range("A1:B1").Value = Array("Name", "Qty")
Range("A2:B2").Value = Array("x", 1)
Range("A3:B3").Value = Array("y", 2)
Dim lo As ListObject
Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:B3"), , xlYes)
lo.Name = "tblCase" & ActiveSheet.Index
Out lo.ListColumns.Count, lo.ListRows.Count, lo.HeaderRowRange.Address, lo.DataBodyRange.Address, lo.ListColumns(2).Name, lo.ListColumns("Qty").Index, lo.Range.Address, lo.ShowTotals
Range("D1").Formula = "=SUM(" & lo.Name & "[Qty])"
Out Range("D1").Value'''),
    ("table rows and columns", '''
Range("A1:B1").Value = Array("Name", "Qty")
Range("A2:B2").Value = Array("x", 1)
Dim lo As ListObject
Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:B2"), , xlYes)
lo.ListRows.Add
lo.ListRows(2).Range.Value = Array("y", 2)
lo.ListRows.Add 1
lo.ListColumns.Add
Out lo.ListRows.Count, lo.Range.Address, lo.ListColumns(3).Name, Range("A2").Value, IsEmpty(Range("A2").Value)'''),
    ("writing next to a table", '''
Range("A1:B1").Value = Array("Name", "Qty")
Range("A2:B2").Value = Array("x", 1)
Dim lo As ListObject
Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:B2"), , xlYes)
Range("A3").Value = "z"
Range("C1").Value = "New"
Out lo.Range.Address'''),
    ("AutoFilter", '''
Range("A1:B1").Value = Array("Name", "Qty")
Range("A2:B2").Value = Array("x", 1)
Range("A3:B3").Value = Array("y", 2)
Range("A4:B4").Value = Array("x", 3)
Range("A1:B4").AutoFilter Field:=1, Criteria1:="x"
Out ActiveSheet.AutoFilterMode, ActiveSheet.FilterMode, Rows(3).Hidden, Rows(4).Hidden, Range("A1:B4").SpecialCells(xlCellTypeVisible).Address, ActiveSheet.AutoFilter.Range.Address
ActiveSheet.ShowAllData
Out Rows(3).Hidden, ActiveSheet.FilterMode
ActiveSheet.AutoFilterMode = False
Out ActiveSheet.AutoFilterMode'''),
    ("AutoFilter conditions", '''
Range("A1").Value = "n"
Range("A2:A6").Value = Application.Transpose(Array(1, 2, 3, 4, 5))
Range("A1:A6").AutoFilter Field:=1, Criteria1:=">1", Operator:=xlAnd, Criteria2:="<4"
Out Range("A1:A6").SpecialCells(xlCellTypeVisible).Address
Range("A1:A6").AutoFilter Field:=1, Criteria1:=Array("1", "5"), Operator:=xlFilterValues
Out Range("A1:A6").SpecialCells(xlCellTypeVisible).Address
ActiveSheet.AutoFilterMode = False'''),
    # ------------------------------------------------------------ typical macros
    ("delete rows in a loop", '''
Dim i As Long
For i = 1 To 6
    Cells(i, 1).Value = IIf(i Mod 2 = 0, "del", "keep" & i)
Next
For i = 6 To 1 Step -1
    If Cells(i, 1).Value = "del" Then Rows(i).Delete
Next'''),
    ("totals by key", '''
Dim i As Long, d As Object, k
Range("A1:A6").Value = Application.Transpose(Array("x", "y", "x", "z", "y", "x"))
Range("B1:B6").Value = Application.Transpose(Array(1, 2, 3, 4, 5, 6))
Set d = New Collection
Dim keys As String
For i = 1 To 6
    If InStr(keys, "|" & Cells(i, 1).Value & "|") = 0 Then keys = keys & "|" & Cells(i, 1).Value & "|"
Next
i = 1
For Each k In Split(Replace(Mid(keys, 2, Len(keys) - 2), "||", "|"), "|")
    Cells(i, 4).Value = k
    Cells(i, 5).Formula = "=SUMIF(A:A,D" & i & ",B:B)"
    i = i + 1
Next
Range("D1:E3").Sort Key1:=Range("E1"), Order1:=xlDescending, Header:=xlNo
Range("E1:E3").NumberFormat = "0.00"'''),
    ("report with formats", '''
Dim r As Long
Range("A1:C1").Value = Array("Item", "Qty", "Price")
For r = 2 To 4
    Cells(r, 1).Value = "Item " & (r - 1)
    Cells(r, 2).Value = r * 2
    Cells(r, 3).Value = r * 1.25
Next
Cells(5, 1).Value = "Total"
Cells(5, 2).Formula = "=SUM(B2:B4)"
Cells(5, 3).FormulaR1C1 = "=SUMPRODUCT(R2C2:R4C2,R2C3:R4C3)"
With Range("A1:C1")
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
    .HorizontalAlignment = xlCenter
End With
Range("C2:C5").NumberFormat = "#,##0.00"
Range("A5:C5").Borders(xlEdgeTop).LineStyle = xlContinuous
Out Range("C5").Text, Range("C5").Value'''),
    ("arrays in and out", '''
Dim v, i As Long
Range("A1:A5").Value = Application.Transpose(Array(5, 3, 8, 1, 9))
v = Range("A1:A5").Value
For i = LBound(v) To UBound(v)
    v(i, 1) = v(i, 1) * 10
Next
Range("B1").Resize(UBound(v), 1).Value = v
ReDim w(1 To 1, 1 To 3)
w(1, 1) = "a": w(1, 2) = 2: w(1, 3) = True
Range("C1:E1").Value = w'''),
]


def build_module() -> Tuple[str, List[str], str]:
    """(Module1's code, the case names, ThisWorkbook's code)."""
    names = [n for n, _ in CASES]
    lines = [
        "Option Explicit",
        "Public outs As String",
        "Public evlog As String",
        repr_vba(),
        "Public Sub Out(ParamArray xs() As Variant)",
        "    Dim i As Long",
        "    For i = LBound(xs) To UBound(xs)",
        "        outs = outs & Show(xs(i)) & \";\"",
        "    Next",
        "End Sub",
        "",
        "Private Function Dump(ws As Worksheet) As String",
        "    On Error GoTo bad",
        "    Dim c As Range, s As String, ur As Range",
        "    Set ur = ws.UsedRange",
        "    If ur.CountLarge > 2000 Then Dump = \"big \" & ur.Address(0, 0): Exit Function",
        "    For Each c In ur.Cells",
        "        If Not IsEmpty(c.Value) Or c.HasFormula Then",
        "            s = s & c.Address(0, 0) & \"=\" & c.Formula & \"{\" & Show(c.Value) & \"}\"",
        "            If c.NumberFormat <> \"General\" Then s = s & \"[\" & c.NumberFormat & \"]\"",
        "            s = s & \" \"",
        "        End If",
        "    Next",
        "    Dump = s",
        "    Exit Function",
        "bad:",
        "    Dump = \"?\" & Err.Number",
        "End Function",
        "",
    ]
    for i, (_, body) in enumerate(CASES, 1):
        lines.append(f"Private Sub Case{i}()")
        lines += ["    " + ln for ln in body.strip("\n").split("\n")]
        lines += ["End Sub", ""]
    lines += ["Private Sub Try(ByVal i As Long)", "    On Error GoTo bad", "    Select Case i"]
    for i in range(1, len(CASES) + 1):
        lines.append(f"    Case {i}: Case{i}")
    lines += [
        "    End Select",
        "    Exit Sub",
        "bad:",
        "    outs = outs & \"E\" & Err.Number & \";\"",
        "End Sub",
        "",
        "Public Sub RunAll()",
        "    Dim i As Long, ws As Worksheet, res As Worksheet",
        "    Application.DisplayAlerts = False",
        "    Set res = Worksheets(1)",
        "    res.Name = \"Results\"",
        f"    For i = 1 To {len(CASES)}",
        "        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))",
        "        ws.Name = \"Case\" & i",
        "        outs = \"\"",
        "        evlog = \"\"",
        "        Application.EnableEvents = True",
        "        Try i",
        "        Application.EnableEvents = True",
        "        Application.Calculation = xlCalculationAutomatic",
        "        Application.CutCopyMode = False",
        "        res.Cells(i, 1).Value = \"'\" & outs",
        "        res.Cells(i, 2).Value = \"'\" & Dump(ws)",
        "    Next",
        "End Sub",
    ]
    this_workbook = "\n".join([
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)",
        "    If Sh.Name <> \"Results\" Then Module1.evlog = Module1.evlog & \"chg:\" & Target.Address(0, 0) & \",\"",
        "End Sub",
        "",
        "Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)",
        "    If Sh.Name <> \"Results\" Then Module1.evlog = Module1.evlog & \"sel:\" & Target.Address(0, 0) & \",\"",
        "End Sub",
    ])
    return "\n".join(lines), names, this_workbook
