"""VBA cases compared with Excel's own VBA (``scripts/excel_vba_oracle_mac.py`` runs them in
Microsoft Excel; ``tests/test_vba_excel_cases.py`` checks ReflexGrid still gives the answers
Excel gave). Each expression is evaluated as ``v = <expr>``; each function body is a Function
returning its result. Results are ``TypeName|CStr`` text, or ``E<number>`` for a run-time error."""

from typing import List, Tuple

# (expression) cases: evaluated as `v = <expr>`.
EXPRESSIONS: List[str] = [
    # VBA's enumeration constants are Longs
    'TypeName(vbYes)', 'vbYes * 10000', 'TypeName(vbOKOnly)', 'TypeName(vbTrue)', 'TypeName(vbString)', 'TypeName(vbTextCompare)', 'TypeName(vbObjectError)', 'TypeName(vbCrLf)',
    # conversions and number text
    'CStr(1/3)', 'CStr(2/3)', 'CStr(1E15)', 'CStr(1E14)', 'CStr(123456789012345#)', 'CStr(1234567890123456#)',
    'CStr(0.0001)', 'CStr(0.00001)', 'CStr(0.000001)', 'CStr(1.5E-10)', 'CStr(-0.5)', 'CStr(100#)', 'CStr(1E+20)',
    'CStr(CSng(1.1))', 'CStr(CSng(1/3))', 'CStr(CSng(123456789))', 'CStr(CCur(1.5))', 'CStr(CCur(1/3))',
    'CStr(CDec(1)/3)', 'CStr(True)', 'CStr(#1/2/2020#)', 'CStr(#1/2/2020 10:30:00 AM#)', 'CStr(#10:30:00 PM#)',
    'CStr(#12/30/1899#)', 'CStr(CDate(0.5))', 'CStr(CDate(-1.25))', 'CStr(Empty)',
    'CInt(2.5)', 'CInt(3.5)', 'CInt(-2.5)', 'CInt("12")', 'CInt(" 12 ")', 'CInt("1,234")', 'CInt("&HFF")',
    'CInt("1e3")', 'CInt("$12")', 'CInt("(12)")', 'CInt(True)', 'CInt("True")', 'CInt("abc")', 'CInt(40000)',
    'CLng(2147483647.4)', 'CLng("2147483648")', 'CDbl("1,234.5")', 'CDbl("1/1/2000")', 'CDbl(#1/1/2000#)',
    'CBool("True")', 'CBool("yes")', 'CBool(2)', 'CBool("0")', 'CBool(0.0)', 'CDate("2020-01-15")',
    'CDate("1/15/2020")', 'CDate("15/1/2020")', 'CDate("January 15, 2020")', 'CDate("15 Jan 2020")',
    'CDate("10:30 PM")', 'CDate("1/15/20")', 'CDate("1/15/30")', 'CDate(43831)', 'CCur("0.1") * 3',
    'CCur(1) + 1.5', 'CCur(1) * 1.5#', 'CCur(10) / 4', 'CByte(255)', 'CByte(256)', 'CByte(-1)', 'CSng("1.5")',
    'CDec("0.1") + CDec("0.2")', 'CVErr(2042)', 'CStr(CVErr(2042))', 'Val("1 2 3")', 'Val("&HFF")',
    'Val("1e3x")', 'Val("  -12.5abc")', 'Val("1,000")', 'Val(".5")', 'Str(5)', 'Str(-5)', 'Str(0.5)',
    'Hex(255)', 'Hex(-1)', 'Hex(-1&)', 'Hex(65536)', 'Oct(8)', 'Int(-2.5)', 'Fix(-2.5)', 'Int(2.5)',
    'Round(2.5)', 'Round(3.5)', 'Round(2.675, 2)', 'Round(0.125, 2)', 'Round(-2.5)', 'Round(1.005, 2)',
    # arithmetic and types
    '200 * 200', '32767 + 1', 'CInt(32767) + CLng(1)', '1 + 1', '1 + 1.5', 'CByte(1) + CByte(2)', 'CByte(200) + CByte(100)',
    'True + True', 'True * 2', 'Empty + Empty', 'Empty + 1', 'Empty & "x"', '"1" + 2', '"1" + "2"', '"a" + 1',
    '1 / 0', '0 / 0', '5 \\ 2', '-5 \\ 2', '5.5 \\ 2', '5 Mod 3', '-5 Mod 3', '5.5 Mod 2', '6.5 Mod 2', '2 ^ 0.5', '(-8) ^ (1/3)',
    '0 ^ -1', '2 ^ 1024', '#1/1/2020# + 1', '#1/2/2020# - #1/1/2020#', '#1/1/2020# * 2', '-#1/1/2020#', 'Null + 1',
    'Null & "x"', 'Null & Null', 'Null = Null', 'TypeName(Null And False)', 'Null Or True', 'Not Null', 'Not 5',
    '5 And 3', 'True And 5', 'True Or False', '&HFFFF', '&HFFFF&', '&H8000', '&H10000', '-32768', '1E300 * 1E10',
    'CSng(1) / CSng(3)', 'CSng(1) + 1', 'CSng(1) + 100000', 'CLng(1) / 2', '10 \\ 4.6', 'CCur(5) \\ 2',
    # comparisons
    '1 = 1#', '"1" = 1', 'CVar("1") = CVar(1)', 'CVar("a") > CVar(1)', '"abc" < "abd"', '"a" = "A"',
    'Empty = 0', 'Empty = ""', '#1/1/2020# > 43000', 'Null = 1', '1 Like "1"', '"abc" Like "a*"', '"aBc" Like "a?c"',
    '"a1" Like "a#"', '"b" Like "[a-c]"', '"d" Like "[!a-c]"', '"*" Like "[*]"',
    # strings
    'Left("hello", 2)', 'Right("hello", 10)', 'Mid("hello", 2)', 'Mid("hello", 10, 2)', 'Mid("hello", 0, 1)',
    'Len(Null)', 'Trim("  a b  ")', 'LTrim(vbTab & "x")', 'InStr("hello", "l")', 'InStr(3, "hello", "l")',
    'InStr("hello", "")', 'InStr(10, "hello", "l")', 'InStr(1, "Hello", "h", vbTextCompare)', 'InStr(Null, "x")',
    'InStrRev("hello", "l")', 'InStrRev("hello", "l", 3)', 'Replace("abcabc", "a", "x", 3)',
    'Replace("abcabc", "A", "x", , , vbTextCompare)', 'Replace("aaa", "a", "b", 1, 2)', 'UBound(Split("a,b,,c", ","))',
    'UBound(Split(""))', 'UBound(Split("a b  c"))', 'Join(Split("a-b-c", "-", 2), "|")', 'StrComp("a", "B")',
    'StrComp("a", "B", vbTextCompare)', 'StrConv("hello world", vbProperCase)', 'StrReverse("abc")', 'Space(3) & "|"',
    'String(3, "ab")', 'String(3, 65)', 'Chr(65)', 'Chr(128)', 'ChrW(8364)', 'Asc("A")', 'Asc("€")', 'AscW("€")',
    'UCase("straße")', 'LCase("ÀB")', 'Format(1234.5, "#,##0.00")', 'Format(0.5)', 'Format(1234.5)', 'Format(-1234.5, "Currency")',
    'Format(0.256, "Percent")', 'Format(1234.5, "Fixed")', 'Format(1234.5, "Standard")', 'Format(12345.678, "Scientific")',
    'Format(0, "Yes/No")', 'Format(#1/15/2020 3:05:09 PM#, "yyyy-mm-dd hh:nn:ss")', 'Format(#1/15/2020#, "dddd, mmmm d, yyyy")',
    'Format(#1/15/2020#, "Long Date")', 'Format(#1/15/2020#, "Medium Date")', 'Format(#1/15/2020#, "Short Date")',
    'Format(#3:05:09 PM#, "Long Time")', 'Format(#3:05:09 PM#, "Medium Time")', 'Format(#3:05:09 PM#, "Short Time")',
    'Format(#1/15/2020#, "q")', 'Format(#1/15/2020#, "ww")', 'Format(#1/15/2020#, "y")', 'Format(#1/15/2020#, "w")',
    'Format(#3:05:09 PM#, "h:m:s AM/PM")', 'Format(#3:05:09 PM#, "hh:mm")', 'Format("abc", "@@@@@")', 'Format("abc", "&&&&&")',
    'Format("abc", ">")', 'Format("5551234", "@@@-@@@@")', 'Format("abc", "!@@@@@")', 'Format(5, "000")',
    'Format(-5, "0;(0)")', 'Format(0, "0;(0);\\z\\e\\r\\o")', 'Format(Null, "0;0;0;\\n\\u\\l\\l")', 'Format("12", "0.00")',
    'Format("x", "0.00")', 'Format(43831, "yyyy-mm-dd")', 'Format(#1/15/2020#, "0.00")', 'Format(1.5, "#.##")',
    'Format(True)', 'Format(Empty)', 'FormatNumber(1234.567)', 'FormatNumber(-0.5, 1)', 'FormatCurrency(-1234.5)',
    'FormatPercent(0.1234)', 'FormatDateTime(#1/15/2020 3:05:09 PM#)', 'FormatDateTime(#1/15/2020#, vbLongDate)',
    'MonthName(2)', 'MonthName(2, True)', 'WeekdayName(1)', 'WeekdayName(1, True, vbMonday)',
    # dates
    'DateSerial(2020, 13, 1)', 'DateSerial(2020, 2, 30)', 'DateSerial(20, 1, 1)', 'DateSerial(2020, 1, 0)', 'TimeSerial(25, 0, 0)',
    'TimeSerial(0, -1, 0)', 'DateAdd("m", 1, #1/31/2020#)', 'DateAdd("yyyy", 1, #2/29/2020#)', 'DateAdd("d", -1, #3/1/2020#)',
    'DateAdd("h", 25, #1/1/2020#)', 'DateAdd("n", 90, #1/1/2020#)', 'DateAdd("ww", 2, #1/1/2020#)', 'DateAdd("q", 1, #11/30/2020#)',
    'DateDiff("d", #1/1/2020#, #3/1/2020#)', 'DateDiff("m", #1/31/2020#, #2/1/2020#)', 'DateDiff("yyyy", #12/31/2019#, #1/1/2020#)',
    'DateDiff("ww", #1/1/2020#, #1/31/2020#)', 'DateDiff("w", #1/1/2020#, #1/31/2020#)', 'DateDiff("h", #1/1/2020 10:59#, #1/1/2020 11:00#)',
    'DateDiff("s", #1/1/2020#, #1/2/2020#)', 'DateDiff("d", #1/2/2020#, #1/1/2020#)', 'DatePart("ww", #1/1/2021#)',
    'DatePart("ww", #1/1/2021#, vbMonday, vbFirstFourDays)', 'DatePart("q", #8/15/2020#)', 'DatePart("y", #12/31/2020#)',
    'Weekday(#1/15/2020#)', 'Weekday(#1/15/2020#, vbMonday)', 'Year(#1/15/2020#)', 'Month("2020-03-04")', 'Day(43831)',
    'Hour(#3:05:09 PM#)', 'Minute(0.75)', 'Second(#3:05:09 PM#)', 'DateValue("1/15/2020 3:00 PM")', 'TimeValue("3:00 PM")',
    'IsDate("1/15/2020")', 'IsDate("13/13/2020")', 'IsDate(43831)', 'IsDate("2020-01-15")', 'IsDate("hello")',
    # information
    'IsNumeric("12")', 'IsNumeric("1e5")', 'IsNumeric("$1,000")', 'IsNumeric("&HFF")', 'IsNumeric("1,2")',
    'IsNumeric("")', 'IsNumeric(Empty)', 'IsNumeric(Null)', 'IsNumeric(#1/1/2020#)', 'IsNumeric(True)', 'IsNumeric(" 12 ")',
    'IsEmpty(Empty)', 'IsEmpty("")', 'IsNull(Null)', 'IsArray(Array())', 'IsObject(Nothing)', 'VarType(1)', 'VarType(1&)',
    'VarType(1.5)', 'VarType("x")', 'VarType(Array(1))', 'VarType(Empty)', 'VarType(Null)', 'VarType(CCur(1))',
    'VarType(CDec(1))', 'VarType(#1/1/2020#)', 'VarType(True)', 'TypeName(Array(1, 2))', 'TypeName(CVErr(5))',
    'TypeName(Nothing)', 'TypeName(New Collection)', 'LBound(Array(1, 2))', 'UBound(Array())', 'IIf(1 > 2, "a", "b")',
    'Choose(2, "a", "b", "c")', 'Choose(5, "a")', 'Switch(False, 1, True, 2)', 'RGB(255, 128, 0)', 'QBColor(12)',
    'Error(11)', 'Sgn(-3.5)', 'Abs(-32768)', 'Abs(CInt(-5))', 'Sqr(16)', 'Log(Exp(1))', 'Atn(1) * 4', 'Fix(-0.5)',
    'Pmt(0.1 / 12, 12, 1000)', 'FV(0.05, 10, -100)', 'PV(0.05, 10, -100)', 'NPer(0.01, -100, 1000)', 'SLN(1000, 100, 5)',
    'SYD(1000, 100, 5, 1)', 'DDB(1000, 100, 5, 1)', 'Rnd(-1)', 'Rnd(-2)',
    # Like / Option Compare binary default and misc
    'TypeName(1 + 1)', 'TypeName(1 + 100000)', 'TypeName(30000 + 30000)', 'TypeName(CInt(1) * 1.5)', 'TypeName(1 / 1)',
    'TypeName(CSng(1) / 2)', 'TypeName(10 \\ 3)', 'TypeName(10& \\ 3)', 'TypeName("1" + 1)', 'TypeName(#1/1/2020# + 1)',
    'TypeName(True And True)', 'TypeName(5 And 3)', 'TypeName(CByte(5) And CByte(3))', 'TypeName(-CByte(1))',
    'TypeName(Len("a"))', 'TypeName(InStr("a", "a"))', 'TypeName(Left("a", 1))', 'TypeName(Split("a"))',
    'TypeName(Format(1))', 'TypeName(Now)', 'TypeName(Timer)', 'TypeName(Rnd)', 'TypeName(Int(1.5))', 'TypeName(Int("1.5"))',
    'TypeName(CCur(1) + 1.5)', 'TypeName(CCur(1) * 2.5)', 'TypeName(CCur(1) / 2)', 'TypeName(Abs(-1))', 'TypeName(Round(1.5))',
]

# Cases that need statements: Functions returning a value.
FUNCTIONS: List[Tuple[str, str]] = [
    ('ForVariantType', 'Dim i: For i = 1 To 3: Next: ForVariantType = TypeName(i) & i'),
    ('ForLongLimit', 'Dim i As Long: For i = 1 To 10 Step 3: Next: ForLongLimit = i'),
    ('ForDouble', 'Dim i, s: For i = 0 To 1 Step 0.25: s = s & i & ",": Next: ForDouble = s'),
    ('ForDown', 'Dim i As Integer, s As String: For i = 5 To 1 Step -2: s = s & i: Next: ForDown = s & "|" & i'),
    ('ForEachArr', 'Dim a(1 To 2, 1 To 2), x, s: a(1, 1) = 1: a(2, 1) = 2: a(1, 2) = 3: a(2, 2) = 4: For Each x In a: s = s & x: Next: ForEachArr = s'),
    ('ResumeNextIf', 'On Error Resume Next: Dim s: s = "start": If 1 / 0 = 1 Then s = s & " then" Else s = s & " else"\nResumeNextIf = s'),
    ('ErrInHandler', 'On Error GoTo h: Err.Raise 1000, "src", "custom": Exit Function\nh: ErrInHandler = Err.Number & " " & Err.Source & " " & Err.Description'),
    ('ErrClearedAfterResume', 'On Error GoTo h: Dim x: x = 1 / 0: ErrClearedAfterResume = "after " & Err.Number: Exit Function\nh: Resume Next'),
    ('RaiseDefaultDescription', 'On Error Resume Next: Err.Raise 11: RaiseDefaultDescription = Err.Description'),
    ('ErrObjectError', 'On Error Resume Next: Err.Raise vbObjectError + 513: ErrObjectError = Err.Number & "|" & Err.Description'),
    ('ReDimPreserve', 'Dim a() As Long: ReDim a(1 To 2): a(2) = 7: ReDim Preserve a(1 To 4): ReDimPreserve = UBound(a) & a(2) & a(4)'),
    ('ReDimPreserveBad', 'On Error Resume Next: Dim a(): ReDim a(1 To 2, 1 To 2): ReDim Preserve a(1 To 3, 1 To 2): ReDimPreserveBad = Err.Number'),
    ('ArrayCopy', 'Dim a, b: a = Array(1, 2): b = a: b(0) = 9: ArrayCopy = a(0) & b(0)'),
    ('ByRefDefault', 'Dim x As Long: x = 1: Bump x: ByRefDefault = x'),
    ('ByValParen', 'Dim x As Long: x = 1: Bump (x): ByValParen = x'),
    ('CallByRef', 'Dim x As Long: x = 1: Call Bump(x): CallByRef = x'),
    ('StaticCounter', 'Counter: Counter: StaticCounter = Counter()'),
    ('OptionalMissing', 'OptionalMissing = HasArg() & HasArg(1)'),
    ('ParamArrayCount', 'ParamArrayCount = CountArgs(1, "a", 3) & CountArgs()'),
    ('CollectionKeys', 'Dim c As New Collection: c.Add 1, "A": On Error Resume Next: c.Add 2, "a": CollectionKeys = Err.Number & "|" & c("a")'),
    ('CollectionIndexErr', 'Dim c As New Collection: On Error Resume Next: Dim x: x = c(1): CollectionIndexErr = Err.Number'),
    ('DictionaryAutoAdd', 'Dim d: Set d = CreateObject("Scripting.Dictionary"): Dim x: x = d("k"): DictionaryAutoAdd = d.Count & TypeName(x)'),
    ('DictionaryNumKey', 'Dim d: Set d = CreateObject("Scripting.Dictionary"): d(1) = "a": d("1") = "b": DictionaryNumKey = d.Count & d(1)'),
    ('DictionaryIntDouble', 'Dim d: Set d = CreateObject("Scripting.Dictionary"): d(1) = "a": d(1#) = "b": DictionaryIntDouble = d.Count & d(1)'),
    ('WithBlock', 'Dim c As New Collection: With c: .Add 1: .Add 2: WithBlock = .Count: End With'),
    ('SelectCaseString', 'Dim s: s = "b": Select Case s: Case "a", "b": SelectCaseString = "ab": Case Else: SelectCaseString = "no": End Select'),
    ('SelectCaseIs', 'Select Case 5: Case Is < 3: SelectCaseIs = 1: Case 4 To 6: SelectCaseIs = 2: End Select'),
    ('DoUntil', 'Dim i: Do Until i >= 3: i = i + 1: Loop: DoUntil = i'),
    ('WhileWend', 'Dim i: While i < 4: i = i + 2: Wend: WhileWend = i'),
    ('GoSubReturn', 'Dim s: GoSub sub1: GoSubReturn = s & "end": Exit Function\nsub1: s = s & "in ": Return'),
    ('ExitForNested', 'Dim i, j, s\nFor i = 1 To 3\nFor j = 1 To 3\nIf j = 2 Then Exit For\ns = s & i & j & " "\nNext\nNext\nExitForNested = s'),
    ('MidStatement', 'Dim s As String: s = "abcdef": Mid(s, 2, 3) = "XYZW": MidStatement = s'),
    ('LSetRSet', 'Dim s As String: s = "12345": LSet s = "ab": Dim t As String: t = "12345": RSet t = "ab": LSetRSet = "[" & s & "][" & t & "]"'),
    ('FixedString', 'Dim s As String * 5: s = "ab": FixedString = "[" & s & "]" & Len(s)'),
    ('IntegerOverflowVar', 'Dim a: a = CInt(30000): a = a + a: IntegerOverflowVar = TypeName(a) & a'),
    ('IntegerOverflowTyped', 'On Error Resume Next: Dim a As Integer: a = 30000: a = a + a: IntegerOverflowTyped = Err.Number'),
    ('StringToNumberCompare', 'On Error Resume Next: Dim v: v = "abc": Dim b: b = (v = 5): StringToNumberCompare = Err.Number & "|" & b'),
    ('NullCondition', 'If Null Then NullCondition = "t" Else NullCondition = "f"'),
    ('UDTCopy', 'Dim a As Pt, b As Pt: a.X = 1: b = a: b.X = 2: UDTCopy = a.X & b.X'),
    ('ClassBasics', 'Dim o As New Class1: o.Name = "x": UDTCopy2 o: ClassBasics = o.Name & o.Twice(4)'),
    ('ClassDefaultMember', 'Dim o As New Class1: o.Name = "v": ClassDefaultMember = o'),
    ('ObjectIs', 'Dim a As Collection, b As Collection: Set a = New Collection: Set b = a: ObjectIs = (a Is b) & (a Is Nothing)'),
    ('ArrayFunctionBase', 'ArrayFunctionBase = LBound(Array(1)) & LBound(VBA.Array(1))'),
    ('EraseFixed', 'Dim a(2) As Long: a(1) = 5: Erase a: EraseFixed = UBound(a) & a(1)'),
    ('EraseDynamic', 'On Error Resume Next: Dim a() As Long: ReDim a(2): Erase a: Dim u: u = UBound(a): EraseDynamic = Err.Number'),
    ('ConstExpr', 'Const A As Long = 2 * 3, B = A + 1: ConstExpr = B & TypeName(B)'),
    ('EnumValues', 'EnumValues = eB & eC'),
    ('NestedFunctionReturnArray', 'Dim a: a = MakeArr(): NestedFunctionReturnArray = UBound(a) & a(1)'),
    ('StrPrintTabs', 'StrPrintTabs = "a" & vbTab & "b"'),
    ('RndSequence', 'Dim s, i: Rnd -5: For i = 1 To 3: s = s & Format(Rnd, "0.000000") & " ": Next: RndSequence = s'),
    ('RandomizeSeeded', 'Dim s, i: Rnd -1: Randomize 42: For i = 1 To 3: s = s & Format(Rnd, "0.000000") & " ": Next: RandomizeSeeded = s'),
    ('ForEachCollectionModify', 'Dim c As New Collection, x, s: c.Add 1: c.Add 2: For Each x In c: s = s & x: Next: ForEachCollectionModify = s'),
    ('SplitLimit', 'SplitLimit = Join(Split("a,b,c,d", ",", 3), "|")'),
    ('FilterArray', 'FilterArray = Join(Filter(Array("apple", "banana", "cherry"), "an"), "|")'),
    ('LikeTextCompareDefault', 'LikeTextCompareDefault = ("ABC" Like "abc")'),
    ('OnErrorGoToMinus1', 'On Error GoTo h1: Err.Raise 5\nh1: On Error GoTo -1: On Error GoTo h2: Err.Raise 6: Exit Function\nh2: OnErrorGoToMinus1 = "second " & Err.Number'),
    ('ResumeLabel', 'On Error GoTo h: Err.Raise 5: Exit Function\nh: Resume there\nthere: ResumeLabel = "there " & Err.Number'),
    ('LenOfLong', 'Dim n As Long: n = 12345: Dim v: v = 12345: LenOfLong = Len(n) & "|" & Len(v)'),
    ('CmpVarStrLit', 'Dim v: v = "10": CmpVarStrLit = (v > 5) & "|" & (v = 10) & "|" & (CVar("3") > 5) & "|" & (CVar("abc") = 5)'),
    ('CmpVarStrTyped', 'Dim v, i As Integer: v = "10": i = 5: CmpVarStrTyped = (v > i) & "|" & (v = 10#)'),
    ('CmpVarStrBad', 'On Error Resume Next: Dim v, i As Integer: v = "abc": i = 5: Dim r: r = (v > i): CmpVarStrBad = Err.Number & "|" & r'),
    ('CmpStrTypedNum', 'On Error Resume Next: Dim s As String: s = "10": Dim r: r = (s > 5): CmpStrTypedNum = Err.Number & "|" & r'),
    ('CmpStrTypedBad', 'On Error Resume Next: Dim s As String: s = "abc": Dim r: r = (s > 5): CmpStrTypedBad = Err.Number & "|" & r'),
    ('CmpVarVar', 'Dim a, b: a = "10": b = 5: CmpVarVar = (a > b) & "|" & (a = 10)'),
    ('CmpCellLike', 'Dim v: v = "5": CmpCellLike = (v = 5) & "|" & (v < 6) & "|" & (v = "5")'),
    ('CmpVarNumStrLit', 'Dim v: v = 5: CmpVarNumStrLit = (v = "5") & "|" & (v < "abc")'),
    ('RandomizePoint1', 'Dim s, i: Rnd -1: Randomize 0.1: For i = 1 To 2: s = s & Format(Rnd, "0.000000") & " ": Next: RandomizePoint1 = s'),
    ('RandomizeBig', 'Dim s, i: Rnd -1: Randomize 123456789: For i = 1 To 2: s = s & Format(Rnd, "0.000000") & " ": Next: RandomizeBig = s'),
    ('RandomizeNeg', 'Dim s, i: Rnd -1: Randomize -3.7: For i = 1 To 2: s = s & Format(Rnd, "0.000000") & " ": Next: RandomizeNeg = s'),
    ('DateDiffMinutes', 'DateDiffMinutes = DateDiff("n", #10:00:59 AM#, #10:01:00 AM#) & "|" & DateDiff("h", #10:30:00 AM#, #11:29:00 AM#) & "|" & DateDiff("h", #10:30:00 AM#, #11:30:00 AM#) & "|" & DateDiff("s", #10:00:00 AM#, #10:00:01 AM#)'),
    ('ErlLine', 'On Error GoTo h\n10 Dim x\n20 x = 1 / 0\nh: ErlLine = Erl'),
]

HELPERS = """
Public Type Pt
    X As Long
    Y As Long
End Type

Public Enum Letters
    eA = 1
    eB
    eC = 10
End Enum

Sub Bump(n As Long)
    n = n + 1
End Sub

Function Counter() As Long
    Static n As Long
    n = n + 1
    Counter = n
End Function

Function HasArg(Optional x As Variant) As String
    HasArg = IIf(IsMissing(x), "m", "p")
End Function

Function CountArgs(ParamArray a() As Variant) As String
    CountArgs = (UBound(a) - LBound(a) + 1) & "/" & LBound(a)
End Function

Sub UDTCopy2(o As Class1)
    o.Name = o.Name & "y"
End Sub

Function MakeArr() As Variant
    MakeArr = Array(10, 20, 30)
End Function
"""

CLASS1 = """Private mName As String

Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mName
End Property

Public Property Let Name(v As String)
    mName = v
End Property

Public Function Twice(n As Long) As Long
    Twice = n * 2
End Function
"""


def repr_vba() -> str:
    return '''
Private Function Show(v As Variant) As String
    On Error GoTo bad
    If IsObject(v) Then
        If v Is Nothing Then Show = "Nothing|" Else Show = "Object|" & TypeName(v)
        Exit Function
    End If
    If IsArray(v) Then
        Show = TypeName(v) & "|" & LBound(v) & ".." & UBound(v)
        Exit Function
    End If
    If IsNull(v) Then Show = "Null|": Exit Function
    If IsEmpty(v) Then Show = "Empty|": Exit Function
    If IsError(v) Then Show = "Error|" & CStr(v): Exit Function
    Show = TypeName(v) & "|" & CStr(v)
    Exit Function
bad:
    Show = "?" & Err.Number
End Function
'''



# Functions Excel for Mac's VBA doesn't have (a compile error there): their cases run only here.
WINDOWS_ONLY = ("CDec(", "Scripting.Dictionary", "Chr(128)", "€", "ß", "À", "FormatNumber(1234.567)", "FormatPercent(0.1234)")


def mac_cases() -> Tuple[List[str], List[Tuple[str, str]]]:
    exprs = [e for e in EXPRESSIONS if not any(w in e for w in WINDOWS_ONLY)]
    funcs = [(n, b) for n, b in FUNCTIONS if not any(w in b for w in WINDOWS_ONLY)]
    return exprs, funcs


def build_module(windows: bool = False) -> Tuple[str, List[str]]:
    """The test module. VBA limits a procedure to 64 KB of compiled code, so the cases are
    spread over several Run<n> procedures filling one array."""
    names: List[str] = []
    calls: List[str] = []
    exprs, funcs = (EXPRESSIONS, FUNCTIONS) if windows else mac_cases()
    for e in exprs:
        names.append(e)
        calls.append(f"v = {e}")
    for fname, _ in funcs:
        names.append(fname + "()")
        calls.append(f"v = {fname}()")
    total = len(calls)
    lines = ["Option Explicit", f"Private r({total - 1}) As String", HELPERS, repr_vba()]
    chunk = 40
    parts = []
    for k in range(0, total, chunk):
        pname = f"Run{k // chunk}"
        parts.append(pname)
        lines += [f"Private Sub {pname}()", "    Dim v As Variant"]
        for i in range(k, min(k + chunk, total)):
            lines += [
                "    On Error Resume Next",
                "    Err.Clear",
                "    v = Empty",
                f"    {calls[i]}",
                f"    If Err.Number <> 0 Then r({i}) = \"E\" & Err.Number Else r({i}) = Show(v)",
                "    On Error GoTo 0",
            ]
        lines += ["End Sub", ""]
    lines += ["Public Function RunAll() As Variant"] + [f"    {p}" for p in parts] + ["    RunAll = r", "End Function", ""]
    for fname, body in funcs:
        lines.append(f"Function {fname}() As Variant")
        for part in body.split("\n"):
            lines.append("    " + part)
        lines.append("End Function")
        lines.append("")
    lines += ["Sub WriteAll()", "    Dim i As Long"]
    for k, pn in enumerate(parts):
        lines += [f"    Worksheets(1).Range(\"B1\").Value = \"{pn}\"", f"    {pn}"]
    lines += [
        "    Worksheets(1).Range(\"B1\").Value = \"done\"",
        f"    For i = 0 To {total - 1}",
        "        Worksheets(1).Cells(i + 1, 1).Value = \"'\" & r(i)",
        "    Next",
        "End Sub",
    ]
    return "\n".join(lines), names
