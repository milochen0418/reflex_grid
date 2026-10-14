# Excel macros (VBA) in ReflexGrid

ReflexGrid opens, keeps, edits and runs the VBA macros of Excel workbooks (`.xlsm`). The macros run on the
server, inside a sandboxed VBA interpreter that is part of gridcore, against ReflexGrid's own workbook. This
note covers why it is built this way, how it works, how it was checked against Excel, and what doesn't work.

## What you get

| | |
|---|---|
| Files | `.xlsm` opens with its VBA project; the project is written back unchanged (with Excel's compiled code and any digital signature) until someone changes the code, then rebuilt from source so Excel recompiles it. Workbooks with macros are saved on the server as `<id>.xlsm` and downloaded as `.xlsm`; *Save without Macros* writes a macro-free `.xlsx` |
| Editing | Developer tab › **Visual Basic** (Alt+F11): module tree (sheet and workbook modules, modules, class modules; forms are listed and kept but not run), a code editor with VBA highlighting, compile errors per module, Insert Module / Class, Import (`.bas`, `.cls`) and Export, Rename, Remove, and an **Immediate** window (`? expr` or statements) |
| Running | Developer tab › **Macros** (Alt+F8) or F5 in the editor (runs the Sub the cursor is in). `MsgBox` and `InputBox` show as dialogs and the macro waits for the answer; long macros run in time slices with a **Stop** button; a run-time error says where it happened and the editor jumps to the line |
| Worksheet functions | A public `Function` in a standard module can be used in formulas (`=Twice(A1)`, `Application.Caller`, arrays spill), recalculated like built-in functions |
| Events | `Worksheet_Change`, `Worksheet_SelectionChange`, `Worksheet_Activate`/`Deactivate`, and the `Workbook_Sheet…` events run when people edit — after someone turns macros on (the security bar, or Developer › Macros On), as Excel's "Enable Content". Turning them on also runs `Workbook_Open` / `Auto_Open`. Edits made by macros fire `Worksheet_Change` too, as in Excel |
| Undo | A whole macro run is one undo step (also across its dialogs), and so is an edit together with what its `Worksheet_Change` did. On a shared workbook it is the person's own step, like any other edit |
| Together | Everyone on a shared workbook sees a macro's changes as they happen (including while it waits on a dialog); dialogs show only to the person who ran the macro (or whose edit fired the event) |

Commands (usable from the Python API like every other command): `macro.list`, `macro.getModule`,
`macro.setModule`, `macro.addModule`, `macro.deleteModule`, `macro.renameModule`, `macro.importModule`,
`macro.exportModule`, `macro.removeAll`, `macro.check`, `macro.enableEvents`, `macro.run`, `macro.open`,
`macro.answer`, `macro.continue`, `macro.stop`, `macro.immediate`.

```python
from gridcore.engine import Session

s = Session()
s.new_workbook()
s.run("macro.setModule", name="Module1", code='Sub Hello()\n    Range("A1").Value = 42\nEnd Sub\n')
s.run("macro.run", name="Hello")                       # {'status': 'done', 'output': [], …}
s.run("macro.run", name="Ask", answers=[6, "Ann"])     # scripted answers: MsgBox vbYes, then an InputBox
```

## Why an interpreter of our own

Before writing one, the existing open-source options were checked:

- **LibreOffice's VBA compatibility layer** (Basic with `Option VBASupport`, the `ooo.vba.excel` API in
  `sc/source/ui/vba`). It is the most complete open implementation, but it can't be reused here: it is part of
  the LibreOffice process (UNO, `soffice`), works on LibreOffice's own document model, and would need a
  headless office per workbook. Running it means converting the workbook to LibreOffice and back for every
  macro (losing what ReflexGrid keeps and LibreOffice doesn't, and every cached value), it can't be called
  from gridcore's recalculation (VBA worksheet functions), can't fire events on ReflexGrid's edits, can't show
  a `MsgBox` in the browser, and its macros can reach the server (files, `Shell`, UNO) — exactly what a
  shared web app must not allow. It is also several hundred MB of dependency and not portable to every host
  ReflexGrid runs on.
- **ViperMonkey** (Python) emulates VBA to analyse malicious documents: semantics are best-effort, no Excel
  object model beyond stubs, and it is no longer actively developed.
- **vb2py** translates VB to Python ahead of time: no runtime semantics of VBA's own (Variant coercions,
  `On Error`, ByRef…), and the project has wound down.
- **Rubberduck**'s ANTLR grammar parses VBA very well but is a parser only (and GPL-3).
- VBScript engines (Wine's `vbscript`, …) are a different language (no types, no `Declare`, different
  library) and native code.

None gives Excel's VBA semantics over ReflexGrid's workbook, safely, inside the engine. So gridcore has its
own VBA implementation, written from the language specification (MS-VBAL) and checked against the real
Excel; the file format is written from the specifications (MS-CFB, MS-OVBA) and checked by Excel opening
the files.

## How it works

```
gridcore/xlsx/cfb.py        compound files (vbaProject.bin is one): read and write, v3
gridcore/xlsx/ovba.py       the VBA project format: dir stream, compression, PROJECT stream, protection
gridcore/model/vba.py       VbaProject / VbaModule (immutable; part of the workbook snapshot, so undo
                            and collaboration treat code like cells)
gridcore/vba/               the language: lexer, parser (AST), compiler to closures, runtime (Variant
                            values, ByRef slots, arrays, UDTs, classes, On Error, Err), the VBA library
                            (~180 functions, Format, Collection, Scripting.Dictionary, RegExp…)
gridcore/engine/macro/      Excel's object model over the engine (Application, Workbook, Worksheet,
                            Range…), the host that batches writes into undoable edits, runs (dialogs,
                            time slices), events, worksheet functions; Excel's constants
gridcore/engine/cmd_macro.py   the macro.* commands
reflex_grid/components/vba.py  Developer tab, VBA editor, Macros dialog, MsgBox/InputBox, security bar
assets/grid/vba_editor.jsx     the code editor (a local React component)
```

`gridcore/vba` doesn't know about Excel: a *host* gives it global names (`Range`, `Cells`, `ActiveSheet`…),
constants and objects. `gridcore/engine/macro` is that host. The layering stays
`core → numfmt → formula → functions → model → calc → xlsx → vba → engine`.

**Compilation.** Each module is parsed once and compiled to Python closures (statements become a flat list of
instructions so `GoTo`, `On Error GoTo`, `Resume` and `Resume Next` work as in VBA). Static types are tracked
so typed variables coerce and overflow as VBA's do (`Integer` overflows at 32,767; `Dim x As Long` keeps a
Long). A compile error stops only its own module, and is reported with its line; other modules still run.
The compiled project lives as long as the code doesn't change (module-level variables keep their values
between runs; `End` resets them; editing the code starts fresh, as in Excel).

**The object model.** `Range.Value = …` doesn't commit a new workbook per cell: writes go into one open edit
(the engine's `Ctx`) that is committed — recalculated, protection-checked — when the macro reads something
that depends on it, runs a command (sort, delete rows…), shows a dialog, or ends. Everything else is an
engine command run on the macro's behalf (`Range.Sort` is `data.sort`, `Rows(3).Delete` is
`home.deleteRows`…), so macros behave like the UI, formulas adjust, tables and names follow.

**Runs.** A `MacroRun` executes on a worker thread that only works while the person's event handler holds
the workbook (`book.act`). When the macro shows a dialog, or after 0.25 s of work, the worker parks and the
handler returns: the workbook is free for everyone, and the browser shows the dialog (or asks for the next
slice). Answering resumes the worker where it stopped. Each step joins the run's undo group, so the run is
one undo step unless someone else edited in between (then it is the person's own steps, as usual).

**Events.** After each command `Session.execute` asks the macro host whether the edit changed cells, the
selection or the sheet; if the workbook has handlers and macros are on, it runs them as part of the same edit.
Macro edits fire `Worksheet_Change` from inside the run, and `Application.EnableEvents = False` stops that, as
in Excel.

**Worksheet functions.** The formula evaluator asks the macro package for names it doesn't know; a public
`Function` of a standard module is called with the evaluator's values (ranges stay references, so `For Each
c In r` works) through a read-only host: a function used in a cell can't change the workbook (as in Excel),
and has a 5 s limit.

## Security

Macros run on the server, so they get a sandbox, not the machine:

- **No way out of the workbook.** No file system (`Open`, `Kill`, `Dir`, `FileCopy`, `MkDir`…), no `Shell`,
  `SendKeys`, `AppActivate`, no `Declare`d DLL functions (except `Sleep`/`GetTickCount`-style harmless ones,
  emulated), no COM/ActiveX other than the built-in `Collection`, `Scripting.Dictionary`, `VBScript.RegExp`
  and `System.Collections.ArrayList` (`CreateObject("WScript.Shell")`, `Scripting.FileSystemObject`,
  `MSXML2.XMLHTTP`… raise error 429), no other workbooks (`Workbooks.Open`, `SaveAs`), no changing code from
  code (`VBProject`), no registry (`GetSetting` returns the default), and `Environ` returns nothing. Each of
  these raises a normal VBA run-time error (445/429) that `On Error` can handle.
- **Limits.** A run stops after 60 s of work (dialog time doesn't count), a worksheet function after 5 s; the
  call depth is bounded ("Out of stack space", error 28); a paused run that nobody answers is abandoned after
  30 minutes (or when its tab goes away) without touching the workbook; at most 32 runs work at once.
- **Off until enabled.** A workbook's event macros (and `Workbook_Open`) don't run until someone turns macros
  on for it — the same decision as Excel's *Enable Content*. Running a macro on purpose (Macros dialog, F5) is
  always possible. On a shared workbook the switch is per workbook while it is open on the server, and starts
  off again when the server opens it next time. Since anyone who can edit a shared workbook can edit its code,
  turning macros on means trusting the people who edit it; the sandbox bounds what that code can do to the
  workbook itself.

## Checked against Excel

`scripts/excel_vba_oracle_mac.py` (opt-in; it drives the Microsoft Excel installed on a Mac) writes a
macro-enabled workbook with ReflexGrid, opens it in Excel, runs the cases there and compares the results
with ReflexGrid's:

- **Language and library**: 457 cases — expressions, conversions, coercion and overflow rules, string, date,
  math and financial functions, `Format`, arrays, `On Error`/`Resume`, ByRef, classes, UDTs, `Collection`,
  `Like`, `Rnd`/`Randomize`, the types of VBA's constants… All 444 cases Excel for Mac can run match; the
  answers are locked in `tests/vba_excel_expected.json` (`tests/test_vba_excel_cases.py`). The other 13 use
  what only Windows has (`CDec`, `Scripting.Dictionary`, non-ASCII source) and follow Windows.
- **Object model**: 158 scenarios (`tests/vba_object_cases.py`): what writing values, dates, currency, text and
  formulas to cells stores and what reading gives back, formats, navigation (`End`, `CurrentRegion`,
  `Offset`, `Areas`, `Union`…), editing (insert, delete, copy, paste special, fill, merge), sorting,
  finding, tables, filters, names, comments, links, rules, protection, selection, events of macro edits,
  manual calculation… 150 give Excel's answers, locked in `tests/vba_object_expected.json`
  (`tests/test_vba_object_cases.py`); the other 8 are the differences listed below. Running them also
  found engine bugs that the UI had too (F9 in manual calculation, `$` and ISO dates typed into cells,
  numbers typed with more than 15 digits, formats of inserted rows), now fixed.
- **Files**: Excel opens the `.xlsm` ReflexGrid writes (macros prompt, compiles, runs, `=Twice(21)` in a cell
  gives 42); ReflexGrid reads the project Excel saves; code edited in ReflexGrid runs in Excel.
- Comparing found one more thing worth knowing: Excel compiles a procedure only when it first runs, so a
  compile error in a macro nobody runs goes unnoticed there; ReflexGrid reports it for the module when the
  code is applied (Developer › Compile shows them all).

## Compatibility

### Language

Supported: `Sub`, `Function`, `Property Get/Let/Set` (also in standard modules), `Optional` with defaults,
`ParamArray`, `ByRef`/`ByVal` (ByRef to variables, array elements and UDT fields), named arguments, `Dim`/
`Private`/`Public`/`Static`/`Const`, all data types (`Byte`, `Integer`, `Long`, `LongLong`, `LongPtr`,
`Single`, `Double`, `Currency`, `Decimal`, `Date`, `String`, `String * n`, `Boolean`, `Variant`, `Object`),
type suffixes, `Def…`, `Option Explicit`/`Base`/`Compare`/`Private Module`, fixed and dynamic arrays
(`ReDim [Preserve]`, `Erase`, `LBound`/`UBound`, multi-dimensional, column-major like VBA), user-defined types,
`Enum`, class modules (`New`, `Set`, `Nothing`, `Is`, `TypeOf`, `Me`, `Class_Initialize`, `Class_Terminate`
when the last reference goes (at the next statement), default members via `Attribute …VB_UserMemId = 0`,
`Implements`, `Event`/`RaiseEvent`/`WithEvents` between your own classes), `If`/`ElseIf` (block and single-line), `Select Case` (ranges, `Is`, lists), `For`/`For Each`/`Do`/
`While`/`Exit`, `GoTo`/`GoSub`/`Return`/labels/line numbers, `On Error GoTo/Resume Next/GoTo 0`, `Resume`,
`Err` (`Raise`, `Number`, `Description`, `Source`, `Clear`), `Erl`, `With` (nested), `Mid` statement,
`LSet`/`RSet`, `Debug.Print`/`Debug.Assert`, `Stop`, `End`, line continuations, `[A1]` evaluation,
`VBA.`/`Excel.` qualified names, and Excel's enum types (`Dim d As XlDirection`, `XlDirection.xlUp`).

### Library

About 180 VBA functions (strings, conversions, dates, math, financial, `Format`/`FormatNumber`/…,
`Split`/`Join`/`Filter`, `InStr`/`Replace`/`StrConv`, `Rnd` with VBA's generator, `IIf`/`Choose`/`Switch`,
`TypeName`/`VarType`, `CallByName`…), all VBA constants, and the 4,692 enumeration constants of Excel's and
Office's object libraries (`xlUp`, `xlPasteValues`, `msoTrue`, `rgbRed`…), generated from Excel's type
libraries by `scripts/gen_vba_constants_mac.py`.

### Excel object model

| Object | Members |
|---|---|
| Application | `ActiveWorkbook/Sheet/Cell`, `Selection`, `ThisWorkbook`, `Workbooks`, `Worksheets`, `Sheets`, `Range`, `Cells`, `Rows`, `Columns`, `Names`, `WorksheetFunction`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation` (manual mode honoured during the run), `Calculate`, `StatusBar`, `CutCopyMode`, `Evaluate`, `Intersect`, `Union`, `Run`, `InputBox` (with `Type`), `Goto`, `Wait`, `ConvertFormula`, `Caller`, `Volatile`, `UserName`, `Version`, `International`…, and worksheet functions called directly (`Application.Sum`, `Application.VLookup` returning an error value) |
| Workbook | `Name`, `Worksheets`, `Sheets`, `ActiveSheet`, `Names`, `Windows`, `Save` (no-op: shared workbooks save themselves), `Saved`, `Protect`/`Unprotect`, `BuiltinDocumentProperties`, `Date1904`, `CodeName`, its own module's members (`ThisWorkbook.MyProc`) |
| Worksheets / Sheets | `Count`, `Item` by index or name, `Add` (`Before`/`After`/`Count`), `For Each`, `Visible`, `Select` |
| Worksheet | `Name`, `Index`, `CodeName`, `Range`, `Cells`, `Rows`, `Columns`, `UsedRange`, `Activate`, `Select`, `Delete`, `Copy`, `Move`, `Visible`, `Next`/`Previous`, `Protect` (incl. `UserInterfaceOnly`) / `Unprotect`, `ProtectContents`, `Paste`, `Evaluate`, `Calculate`, `ListObjects`, `Names`, `Hyperlinks`, `Comments`, `AutoFilter`/`AutoFilterMode`/`ShowAllData`/`FilterMode`, `Sort`, `PageSetup`, `Tab.Color`, `Shapes`/`ChartObjects` (basic), its own module's members |
| Range | `Value`/`Value2`/`Text`/`Formula`/`FormulaR1C1`/`Formula2`/`FormulaArray`/`HasFormula`/`HasArray`/`CurrentArray`/`HasSpill`/`SpillingToRange`, 2-D arrays in and out (Excel's broadcasting and `#N/A` fill), `Cells`, `Item`/default member, `Range` (relative), `Offset`, `Resize`, `Rows`, `Columns`, `EntireRow`/`EntireColumn`, `End`, `CurrentRegion`, `Areas`, `Count`/`CountLarge`, `Row`/`Column`, `Address` (all arguments, R1C1), `Find`/`FindNext`/`FindPrevious` (wildcards), `Replace`, `SpecialCells`, `Sort`, `RemoveDuplicates`, `AutoFilter`, `TextToColumns`, `Copy`/`Cut`/`PasteSpecial`, `Insert`/`Delete`, `Clear`/`ClearContents`/`ClearFormats`/`ClearComments`, `FillDown`/`Right`/`Up`/`Left`, `AutoFill`, `Merge`/`UnMerge`/`MergeCells`/`MergeArea`, `NumberFormat`, `Font`, `Interior`, `Borders`/`BorderAround`, alignment, `WrapText`, `IndentLevel`, `ColumnWidth`/`RowHeight`/`Width`/`Height`/`Left`/`Top`, `Hidden`, `AutoFit`, `Locked`, `Select`/`Activate`, `Name`, `Comment`/`AddComment`/`NoteText`, `Hyperlinks`, `Validation`, `FormatConditions`, `GoalSeek`, `Calculate`, `PrefixCharacter`, `Characters` |
| Others | `Names`/`Name` (`Add`, `RefersTo`, `RefersToRange`, `Value`, `Delete`), `ListObjects`/`ListObject`/`ListRows`/`ListColumns` (tables), `Sort`/`SortFields` (the recorder's style), `AutoFilter`, `Validation`, `FormatConditions`/`FormatCondition` (`Add`, `Interior`, `Font`, `Formula1`, `Modify`, `Delete`…), `Hyperlinks`, `Comments`, `Windows`/`Window` (`FreezePanes`, `Zoom`, `DisplayGridlines`, `ScrollRow`…), `PageSetup`, `WorksheetFunction` (every worksheet function ReflexGrid has, raising error 1004 on error results) |

## Not supported (and what happens instead)

- **UserForms** and ActiveX controls on sheets: forms are kept in the file (and in Excel) but not shown; using
  one raises an error. Form controls (buttons on a sheet) aren't drawn, so assign-to-button macros are run
  from the Macros dialog.
- **Anything outside the workbook**: files, other workbooks, the clipboard of the computer, the network,
  `Shell`, the registry, Windows API calls, other Office applications, `Application.OnTime`/`OnKey`,
  `FileDialog`/`GetOpenFilename` (see *Security*).
- **Macro recording** and stepping through code (breakpoints, Watch, Locals): there is no debugger; errors
  report their module and line, and `Debug.Print` goes to the Immediate window.
- **Charts and shapes** from VBA are basic (`ChartObjects.Add`, `SetSourceData`, `ChartType`, titles);
  pivot tables, slicers and Power Query objects are absent.
- **Password-protected (locked) VBA projects** open and are written back unchanged, but their code can't be
  viewed, edited or run here.
- **`WithEvents` on Excel's own objects** (`Private WithEvents App As Application` for application-level
  events, or on a `Worksheet` variable) gets no events; use the sheet and workbook modules' event procedures.
- **Signed projects**: the signature is kept while the code is unchanged; editing the code drops it (Excel
  does the same).
- **Locale**: VBA in ReflexGrid behaves as Excel does with US English settings (`m/d/yyyy` dates, `.` decimal,
  `$` currency, English names in `Format`). Excel itself follows the computer's region.
- **Windows vs Mac**: ReflexGrid follows Excel for Windows where the two differ (`CDec`,
  `Scripting.Dictionary` and `TypeName` of arrays exist and behave as on Windows; source code is read in the
  project's code page).
- **Calculation in manual mode** is honoured inside a run (`Application.Calculation = xlCalculationManual`,
  `Calculate`); the setting ends with the run, as the workbook's own setting comes back.
- **Performance**: VBA here is an interpreter in Python — roughly a million simple loop iterations a second,
  so fast enough for typical macros (reading and writing tens of thousands of cells, array processing), not
  for heavy numerical code. Writing through arrays (`Range.Value = v`) is much faster than cell by cell, as
  in Excel.
- **Known differences from Excel** (the 8 object-model scenarios above): `.Text` of a number in General
  format shows up to 11 characters whatever the column's width (Excel shows fewer digits, or `####`, in a
  narrow column); spaces typed inside a formula aren't kept (`= 1 + 2` becomes `=1+2`); `SpecialCells`
  returns the same cells split into differently shaped areas; `ColumnWidth`/`RowHeight` of columns or rows
  of different sizes are Null, as Microsoft documents (Excel for Mac gives the first one's), and sizes are
  those of Excel for Windows (Excel for Mac measures in its own points).
- `Find` and `Replace` remember `LookIn`, `LookAt`, `SearchOrder` (and `Replace` `MatchCase`) between calls,
  as Excel does — per person here, across that person's macro runs. Pass them explicitly for predictable
  results, in Excel too.

## Tests

```bash
poetry run pytest -q tests/test_vba_file.py tests/test_vba_excel_cases.py tests/test_vba_macros.py
poetry run ./run_test_suite.sh vba_macros            # the editor, dialogs, events, .xlsm in the browser
poetry run python scripts/excel_vba_oracle_mac.py              # language cases against Excel (Mac, opt-in)
poetry run python scripts/excel_vba_oracle_mac.py --objects    # object model cases against Excel
poetry run python scripts/gen_vba_constants_mac.py             # regenerate Excel's constants
```
