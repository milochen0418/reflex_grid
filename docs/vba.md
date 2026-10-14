# Excel macros (VBA) in ReflexGrid

ReflexGrid opens, keeps, edits and runs the VBA macros of Excel workbooks (`.xlsm`), UserForms included. The
macros run on the server, inside a sandboxed VBA interpreter that is part of gridcore, against ReflexGrid's own
workbook; their UserForms show in the browser. This note covers why it is built this way, how it works, how it
was checked against Excel, and what doesn't work.

## What you get

| | |
|---|---|
| Files | `.xlsm` opens with its VBA project; the project is written back unchanged (with Excel's compiled code and any digital signature) until someone changes the code, then rebuilt from source so Excel recompiles it. Workbooks with macros are saved on the server as `<id>.xlsm` and downloaded as `.xlsm`; *Save without Macros* writes a macro-free `.xlsx` |
| Editing | Developer tab › **Visual Basic** (Alt+F11): the Project (sheet and workbook modules, modules, class modules, forms), a code editor with VBA highlighting, compile errors per module, Insert Module / Class Module / UserForm, Import (`.bas`, `.cls`, and a form's `.frm` with its `.frx`) and Export, Rename, Remove, and an **Immediate** window (`? expr` or statements) |
| Designing forms | A form opens in the **form designer**, as in the VBA editor: a Toolbox (every Microsoft Forms control), a Properties window (name, caption, position and size, colours, font, behaviour, list and number settings… of the control selected, or of the form), and the form to work on — click to select, drag to move, the handles to resize, Delete to remove, the arrow keys to nudge, controls dropped into Frames and onto MultiPage pages, a 6-point grid. Double-clicking a control opens its event procedure, written when it doesn't exist yet. **Run Form** (or F5) shows the form. Every change is an undoable edit of the workbook |
| UserForms | The forms in a workbook's project show in the browser, drawn from their design, and work: typing, lists, check boxes, option groups, toggle buttons, frames, MultiPage tabs, tab strips, scroll bars, spin buttons, pictures; the Default and Cancel buttons answer Enter and Esc. Their code runs as in Excel: `UserForm_Initialize`, `Show` (modal: the macro waits, like on a `MsgBox`; modeless: the form stays open beside the grid and its events run as macros), `Hide`, `Load`/`Unload`, `QueryClose`, `Terminate`, the controls' `Click`/`Change`/`DblClick`/`Enter`/`Exit`/`BeforeUpdate`/`AfterUpdate`/`KeyPress`/`KeyDown`/`SpinUp`… events, `ControlSource`/`RowSource` cells, `Controls.Add` with `WithEvents` handlers |
| Running | Developer tab › **Macros** (Alt+F8) or F5 in the editor (runs the Sub the cursor is in). `MsgBox` and `InputBox` show as dialogs and the macro waits for the answer; long macros run in time slices with a **Stop** button; a run-time error says where it happened and the editor jumps to the line |
| Worksheet functions | A public `Function` in a standard module can be used in formulas (`=Twice(A1)`, `Application.Caller`, arrays spill), recalculated like built-in functions |
| Events | `Worksheet_Change`, `Worksheet_SelectionChange`, `Worksheet_Activate`/`Deactivate`, and the `Workbook_Sheet…` events run when people edit — after someone turns macros on (the security bar, or Developer › Macros On), as Excel's "Enable Content". Turning them on also runs `Workbook_Open` / `Auto_Open`. Edits made by macros fire `Worksheet_Change` too, as in Excel |
| Undo | A whole macro run is one undo step (also across its dialogs), and so is an edit together with what its `Worksheet_Change` did. On a shared workbook it is the person's own step, like any other edit |
| Together | Everyone on a shared workbook sees a macro's changes as they happen (including while it waits on a dialog); dialogs and forms show only to the person who ran the macro (or whose edit fired the event), and each person has their own forms |

Commands (usable from the Python API like every other command): `macro.list`, `macro.getModule`,
`macro.setModule`, `macro.addModule`, `macro.deleteModule`, `macro.renameModule`, `macro.importModule`,
`macro.exportModule`, `macro.removeAll`, `macro.check`, `macro.enableEvents`, `macro.run`, `macro.open`,
`macro.answer`, `macro.continue`, `macro.stop`, `macro.immediate`, and for forms `macro.showForm`,
`macro.formEvent`, `macro.forms`, `macro.formDesign`.

```python
from gridcore.engine import Session

s = Session()
s.new_workbook()
s.run("macro.setModule", name="Module1", code='Sub Hello()\n    Range("A1").Value = 42\nEnd Sub\n')
s.run("macro.run", name="Hello")                       # {'status': 'done', 'output': [], …}
s.run("macro.run", name="Ask", answers=[6, "Ann"])     # scripted answers: MsgBox vbYes, then an InputBox
# A form's events as answers: type into TextBox1, click OK (whose Click unloads the form).
s.run("macro.run", name="ShowOrder", answers=[{"event": "change", "control": "TextBox1", "value": "Ann"},
                                              {"event": "click", "control": "CommandButton1"}])
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
gridcore/model/forms.py     FormDesign / ControlDesign: a UserForm as designed
gridcore/xlsx/oforms.py     the forms' format (MS-OFORMS): reading designs, writing them, .frx files
gridcore/xlsx/oforms_edit.py   changing a form's design in place (the form designer)
gridcore/vba/msforms.py     the Forms object model at run time (UserForm, controls, Controls, Pages…),
                            form events, modal forms, each person's loaded forms
reflex_grid/components/vba.py  Developer tab, VBA editor, Macros dialog, MsgBox/InputBox, security bar
reflex_grid/components/userform.py, assets/grid/userform.jsx   the forms in the browser
assets/grid/form_designer.jsx  the form designer (Toolbox and Properties window are in vba.py)
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

## UserForms

A form lives in two places in `vbaProject.bin`: its code is a module like any other (in `VBA/`), and its
design is a storage named after it beside `VBA/` — the form's properties and one record per control, in the
binary format Microsoft publishes as [MS-OFORMS] (version 9.1 was used). Before writing anything, three
projects that read that format were studied: **pyOpenVBA** (Python, MIT; its reader and writer were tested
against live Excel, and its notes say what the specification leaves open: the caption Excel shows comes from
the `VBFrame` header, a MultiPage keeps its tab captions in an unnamed TabStrip, only the page on top has its
site visible), **LibreOffice** (`oox/source/ole` for the records, and its VBA compatibility layer
`vbahelper/source/msforms` and `scripting/source/vbaevents` for what a form does at run time: `Show` runs the
form modally, `Hide` ends that without disposing of it, `Unload` disposes of it, assigning to a control sets
its default property, controls are found by name through Frames and pages, and which events exist), and
**FrxEdit** (C#; the `.frx` container). None of them could be used as is: pyOpenVBA's runtime doesn't run
forms, LibreOffice's form layer is its own dialog toolkit inside the office process (the reasons in *Why an
interpreter of our own* apply), and FrxEdit is a file editor. ReflexGrid reads and writes the format itself
(`gridcore/xlsx/oforms.py`, table-driven, field by field from the specification) and runs forms on its own
interpreter.

**Keeping forms.** The design storage is kept byte for byte, its CLSIDs included, whatever happens to the code:
editing a form's code rewrites only its module, so Excel finds its form exactly as it left it. The designer
changes a form in place (`gridcore/xlsx/oforms_edit.py`): a change rewrites only the records and sites it
touches, and merges into them — the writer turns the control as it was and as it is now into records, and
only the fields where those differ go into the original record (bit by bit for the property bits) — so a form
made in Excel keeps whatever ReflexGrid doesn't model (mouse icons, IME modes, column widths, ActiveX
classes…). New controls are written as Excel writes them. Reading the
design (`read_design`) turns it into a `FormDesign`; the reader cross-checks every length in the format (each
record's size, the sites' byte count, each control's share of the `o` stream, every child storage claimed by
a container) and reports what it can't account for as warnings instead of guessing. A form whose design
can't be read still has its code. The writer (`write_design`) is the reader's inverse and lays records out as
Excel does; it serves `.frm`/`.frx` export and tests.

**The object model.** A form module compiles like a class with a default instance. Using its name
(`UserForm1.Show`, `UserForm1.TextBox1`) loads that instance when needed (`UserForm_Initialize` runs) and
`Unload` drops it, so the next use loads a fresh one, as in VBA; `New UserForm1` makes another. Inside the
form's module the controls and the form's own members (`Caption`, `Hide`, `Controls`…) need no `Me.`.
`gridcore/vba/msforms.py` implements the form (`Show`/`Hide`/`Move`/`Controls`/`ActiveControl`…, and
`Me("Name")`/`Me!Name` through `Controls`), every MSForms control with its properties and methods (lists:
`AddItem`, `RemoveItem`, `Clear`, `List`, `Column`, `ListIndex`, `Selected`, `BoundColumn`, `TextColumn`,
`ColumnCount`…; text: `Text`, `Value`, `MaxLength`, `SelStart`/`SelLength`/`SelText`…; numbers: `Min`,
`Max`, `SmallChange`…; containers: `Controls`, `Pages`, `Tabs`, `SelectedItem`), `Controls.Add`, the
`MSForms.ReturnBoolean`/`ReturnInteger` arguments of events, `UserForms`, and `MSForms.DataObject`. Setting a
property fires what Excel fires (`TextBox1.Text = "x"` runs `TextBox1_Change`; an OptionButton turning on
turns its group off), and `WithEvents` variables (class modules handling the controls of a form) get events
too.

**Showing a form.** A modal `Show` doesn't return until the form is hidden or unloaded. It works like
`MsgBox`: the run pauses with a `form` prompt (`macro.run` returns `status: "prompt"`, `prompt.kind:
"form"`, and every step returns the person's forms as snapshots); `macro.formEvent` (or `macro.answer`)
brings what the person did — `{"event": "click", "control": "OK", "values": {"TextBox1": "Ann"}, "focus":
"OK"}` — and the run resumes inside `Show`, applies the typed values (Change), moves the focus (Exit,
BeforeUpdate, the `ControlSource` cell, AfterUpdate, Enter), runs the click, and pauses again or goes on
after `Show`. A `MsgBox` in a form's event code shows over the form. A modeless form (`Show vbModeless`, or
`ShowModal = False` in its design) stays open after its macro: each event on it is a run of its own, through
the same commands. Without a person (scripts, tests) a modal form gets scripted events (`answers`) and is then
closed as with its close button.

**In the browser.** `assets/grid/userform.jsx` draws each form from its snapshot at its designed size (points
to pixels), with the designed fonts, colours, borders and pictures. It sends a click at once; it keeps typing
until the next event carries it, unless the code handles that control's `Change` or `KeyPress`, in which
case each change (or key) goes at once, as VBA runs those events key by key. A modal form has a backdrop that
keeps the workbook out of reach; a modeless one leaves the grid usable. The VBA editor makes way while forms
are up (it would hold the keyboard) and comes back when they close.

**Each person's forms.** Forms belong to the person (their session), not to the workbook: on a shared
workbook two people can use the same form at once, each their own instance, and a form's code changes the
workbook for everyone, as their edits do. The compiled project (with its module-level variables) is shared,
as before: editing the code, or `End`, resets it, which unloads everyone's forms, as resetting the project
does in Excel.

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
- **UserForms** (`--forms`, `--form-events`): Excel loads forms ReflexGrid wrote — every kind of control, a
  Frame, a MultiPage with pages, a TabStrip, a picture — and describes each control (type, position and size,
  tab order, caption, value, list sizes from `RowSource`, `ControlSource`, MultiPage and TabStrip values,
  fonts, `Controls.Count`…); ReflexGrid running the same module gives the same 30 answers. The events a
  form's code fires — setting a value to what it already is (none), a CheckBox (`Change`, then `Click`), an
  option group, list and combo indexes, a TabStrip (no `Click` from code), `Load`, `Show vbModeless`,
  `Hide`, `Unload` (`QueryClose` with `vbFormCode`, then `Terminate`) — come in Excel's order in all 20
  steps. Both sets of answers are locked in `tests/vba_forms_expected.json` (`tests/test_vba_forms.py`). The
  same description check passed for a form Excel made and ReflexGrid imported as `.frm`/`.frx` (21 of 21),
  and for a workbook whose form code was edited in ReflexGrid (13 of 13, the edited `UserForm_Initialize`
  running in Excel). It also settled that a form's `Controls` includes the controls on a MultiPage's pages,
  not the pages. Forms changed in the designer (`--form-designer`): a form made from nothing (every kind of
  control, some in a Frame and on a page, properties changed, one moved, one removed) is loaded and described
  by Excel as by ReflexGrid (18 of 18, locked in the same file); so are two Excel-made forms after moving,
  retitling, adding and removing controls (49 rows, three forms).
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

- **Controls on sheets** (Form controls such as an assigned button, and ActiveX controls placed on a
  worksheet) aren't drawn, and aren't kept when ReflexGrid saves the workbook; run assign-to-button macros
  from the Macros dialog. (This is about sheets: UserForms are kept and run.)
- **The form designer** doesn't add or remove a MultiPage's pages or a TabStrip's tabs (it renames them),
  set pictures, or align, size and order several controls at once; its undo is the workbook's (each change is
  one step). A control's list contents and column widths come from code or `RowSource`, as in Excel.
- **UserForms, what differs**: controls that aren't part of Microsoft Forms (ActiveX controls inside a form:
  calendars, `RefEdit`, list or tree views…) are shown as labelled placeholders; pictures in WMF/EMF aren't
  shown (BMP, GIF, JPEG, PNG and icons are); fonts are the browser's (a font the computer doesn't have falls
  back to a similar one); `MouseDown`/`MouseUp`/`MouseMove`, drag and drop, `Layout`, `Zoom`, scrolling a form
  or frame and `KeyDown` for character keys (other keys work) don't happen; `Picture` can't be changed from
  code (`LoadPicture` reads files); `MousePointer`, `PictureTiling` and `ScrollBars` on forms are kept but have
  no effect. A form shows only while its person's tab is open.
- **`MSForms.DataObject`** works with a clipboard of the person's own: a macro can't read or change the
  computer's clipboard (see *Security*).
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
poetry run pytest -q tests/test_vba_file.py tests/test_vba_excel_cases.py tests/test_vba_macros.py tests/test_vba_forms.py
poetry run ./run_test_suite.sh vba_macros            # the editor, dialogs, events, .xlsm in the browser
poetry run ./run_test_suite.sh vba_userforms         # UserForms in the browser, modal and modeless
poetry run python scripts/excel_vba_oracle_mac.py              # language cases against Excel (Mac, opt-in)
poetry run python scripts/excel_vba_oracle_mac.py --objects    # object model cases against Excel
poetry run python scripts/excel_vba_oracle_mac.py --forms        # forms ReflexGrid wrote, loaded by Excel
poetry run python scripts/excel_vba_oracle_mac.py --form-events  # the events a form's code fires, in order
poetry run python scripts/excel_vba_oracle_mac.py --form-designer  # a form made in the designer, loaded by Excel
poetry run python scripts/gen_vba_constants_mac.py             # regenerate Excel's constants
```
