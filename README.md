# ReflexGrid

ReflexGrid is an Excel-compatible spreadsheet that runs in the browser, built with [Reflex](https://reflex.dev/).
It opens and saves real `.xlsx` workbooks (formulas, cached values, styles, tables, charts…), so files move
between ReflexGrid and Microsoft Excel without conversion. Workbooks are shared: everyone who opens a
workbook's link edits it together, live — they see each other's selections, and can call everyone on it.

Its engine, **gridcore**, is pure Python with a layered design (core → number formats → formulas → functions →
model → calc → xlsx → vba → engine), copy-on-write workbook snapshots for undo, dependency-graph recalculation with
dynamic arrays, an "every action is a command" API, and a sandboxed VBA interpreter that runs Excel macros.

▶ [Demo video](https://www.youtube.com/watch?v=I5rBceuf1tM) (the first version)

## Why ReflexGrid

- **Excel files in, Excel files out**: reads and writes Office Open XML (`.xlsx`) — formulas with Excel's
  file spellings (`_xlfn.` prefixes, dynamic-array metadata, structured references), cached values, styles,
  and the features listed below. CSV/TSV import and export too.
- **A real calculation engine**: 478 worksheet functions, dynamic arrays that spill (`FILTER`, `SORT`,
  `UNIQUE`, `SEQUENCE`…), `LET`/`LAMBDA`/`MAP`/`REDUCE`, structured table references, 3-D references,
  legacy array formulas, iterative calculation, incremental recalculation.
- **The Excel way of working**: ribbon, name box and formula bar, in-cell editing (IME friendly), fill handle,
  drag-to-move, copy/paste with relative references, Paste Special, AutoSum, Format Cells, conditional
  formatting, data validation dropdowns, AutoFilter, sorting, tables, charts, notes, freeze panes, zoom,
  undo/redo, and the familiar keyboard shortcuts.
- **Excel macros**: `.xlsm` workbooks keep their VBA project; edit it in the VBA editor, run macros (with their
  `MsgBox` and `InputBox`), use VBA functions in formulas, let `Worksheet_Change`-style event macros run, and use
  the workbook's **UserForms** in the browser, modal or modeless, with their event code — on the server, in a
  sandbox, checked against Excel's own answers. See [docs/vba.md](docs/vba.md).
- **Scriptable**: every user-visible action is a command with a stable id (`cell.set`, `home.bold`,
  `data.sortAscending`…); the UI, the tests and your own scripts all use the same `Session.execute` API.
- **Edit together, live**: a workbook list like [CoDoc in MD](https://github.com/milochen0418/codoc_in_md)'s
  document manager, shared links, who is here and where (everyone's selection in their colour, as in
  Google Sheets), changes reaching everyone at once, undo that takes back only your own changes, and
  "Call everyone" on a workbook.
- **Sign-in and calls on re_ddns**: install it from the [re_ddns](https://github.com/milochen0418/re_ddns)
  Agentic App Store — see [With re_ddns](#with-re_ddns-sign-in-and-calls).

## Features

| Area | What you get |
|---|---|
| Editing | Enter/Tab/Shift navigation, F2, Alt+Enter line breaks, Ctrl+Enter (fill selection), Ctrl+Shift+Enter (array formula), formula point mode (click or arrow keys to insert references), F4 to toggle `$`, function autocomplete and argument hints, coloured reference highlights |
| Selection | Ranges, multi-ranges (Ctrl+click), whole rows/columns, Ctrl+Arrow jumps, Ctrl+A, Go To / Go To Special |
| Clipboard | Copy/Cut/Paste between cells (formulas adjust), from/to other apps (TSV), Paste Special (values, formulas, formats, transpose, skip blanks, add/subtract/multiply/divide, column widths) |
| Fill | Fill handle series (numbers, dates, months, weekdays, `Item 1`…), Ctrl+drag to copy, Fill Down/Right/Up/Left, Series |
| Structure | Insert/delete rows, columns and cells; move blocks; hide/unhide; row heights and column widths (AutoFit); merge (& center / across) — formulas, names, merges, rules, tables, filters, notes, links and charts follow |
| Formatting | Fonts, sizes, bold/italic/underline/strike, super/subscript, font & fill colours (theme colours with tints), borders (presets and styles), alignment, wrap, shrink, indent, rotation, number formats (General, number, currency, accounting, dates, times, percent, fraction, scientific, text, custom codes with sections/colours/conditions), cell styles gallery, Format Painter |
| Conditional formatting | Highlight rules, top/bottom, above/below average, duplicates, text, dates occurring, formulas, data bars, colour scales, icon sets; rules manager |
| Data | Sort (multi-key, by colour, case-sensitive), AutoFilter (values, conditions, top 10, average, colours), tables (styles, banding, total row, calculated columns, auto-expansion), remove duplicates, data validation (lists, numbers, dates, text length, custom; input and error messages), text to columns |
| Formulas | Defined names (workbook/sheet scope), Name Manager, create from selection, AutoSum, Insert Function, trace precedents/dependents, error checking, show formulas, manual/automatic/iterative calculation |
| Insert | Charts (column, bar, line, area, pie, doughnut, scatter, radar; stacked/100%), pictures, links, notes and threaded comments, sparklines |
| Sheets | Add, delete, rename, move/copy, duplicate, tab colour, hide/unhide, protect (with password) and workbook structure protection |
| View | Freeze panes, gridlines, headings, zoom 10–400 %, right-to-left |
| What-if | Goal Seek |
| Macros (VBA) | Developer tab, VBA editor (modules, class modules, UserForms with a form designer — Toolbox and Properties window —, import/export `.bas`/`.cls`/`.frm`+`.frx`, compile errors, Immediate window), Macros dialog (Alt+F8), `MsgBox`/`InputBox` dialogs, UserForms (shown from their design; `Initialize`/`Show`/`Hide`/`Unload`/`QueryClose`, control events, `ControlSource`/`RowSource`, modal and modeless), Stop for long macros, VBA functions in formulas, event macros (off until enabled, like Excel's *Enable Content*), a macro run as one undo step; the Excel object model macros use (`Range`, `Cells`, `Worksheets`, `Application`, tables, filters, names, formats…); no file, shell, network or COM access ([details and limits](docs/vba.md)) |
| Files | Open `.xlsx`/`.xlsm`/`.csv`/`.tsv` (upload), download `.xlsx` (`.xlsm` with macros), export CSV, print (with the sheet's page setup), save as web page, document properties; workbooks are saved on the server as `.xlsx` (`.xlsm` with macros) while you edit |
| Together | Workbook list (search, filters, who is editing now, new, upload, download, delete), shared links, presence (avatars, "N here"), everyone's selection and typing cell in their colour, changes highlighted for a moment, per-person sheets, zoom and undo, a shared title, calls |

### Excel compatibility

Read and written in `.xlsx`: cell values and formulas (shared, array, dynamic-array with spill metadata;
data-table formulas are read), cached results, shared strings and rich text, styles (fonts, fills, borders,
alignment, protection, number formats, themes, differential formats), column widths and row heights,
hidden and outline state, merged cells, frozen panes and view settings, defined names (incl. print areas and
titles), tables (styles, totals, calculated columns), AutoFilter and filter criteria, conditional formatting
(incl. data bars, colour scales and icon sets), data validation, hyperlinks, notes and threaded comments,
charts (DrawingML with cached data), pictures, sparklines, page setup and margins, sheet and workbook
protection, the 1904 date system, document properties, and VBA projects (`vbaProject.bin`: kept byte for byte
until the code changes, then rebuilt so Excel recompiles it; Excel opens and runs what ReflexGrid writes). Unknown parts are skipped with a warning, so a
file always opens.

Files written by ReflexGrid are checked in the test suite by reading them back with
[openpyxl](https://openpyxl.readthedocs.io/). On a Mac with Microsoft Excel installed,
`poetry run python scripts/excel_check_mac.py` goes further (it opens Excel windows): Excel opens each
generated workbook, its values, formulas and displayed text are compared with what ReflexGrid computed,
Excel saves the file, and ReflexGrid reads Excel's own file back and recalculates it to match Excel's
results. Functions the local Excel licence lacks (older versions have no IFS, XLOOKUP, SEQUENCE…)
are detected first and reported as skipped.

## Getting Started

This project uses [Poetry](https://python-poetry.org/) for dependency management.
Make sure you have **Python 3.11** and Poetry installed.

### 1) Install

```bash
cd reflex_grid
poetry env use python3.11
poetry install
```

### 2) Run

```bash
poetry run ./reflex_rerun.sh
```

Then open http://localhost:3000 (the backend listens on port 8000): the workbook list. **New workbook**
offers a blank one and three samples (Sales, Budget, Grades) that show tables, charts, conditional formats
and validation; **Upload .xlsx** opens a file as a new workbook. Open the same workbook's address in a
second browser (or a private window) to see two people editing it.

To rebuild the environment from scratch: `./proj_reinstall.sh` (see `--help`).

### 3) Use the engine from Python

```python
from gridcore.engine import Session

s = Session()
s.new_workbook()
s.run("range.setValues", range="A1", values=[["Item", "Qty"], ["Pen", 3], ["Ink", 5]])
s.run("cell.set", cell="B4", input="=SUM(B2:B3)")
print(s.run("cell.get", cell="B4")["value"])   # 8.0
s.run("home.bold", range="A1:B1")
s.run("file.saveAs", path="/tmp/demo.xlsx")      # opens in Excel
```

`gridcore.engine.COMMANDS` lists every command with its label, ribbon location, shortcut and parameters.

### 4) Command line

```bash
poetry run python -m gridcore.cli eval '=TEXTJOIN("-",TRUE,SEQUENCE(1,4)*2)'   # 2-4-6-8
poetry run python -m gridcore.cli cat book.xlsx --sheet Sales --formulas
poetry run python -m gridcore.cli run book.xlsx script.json -o out.xlsx         # JSON list of commands
poetry run python -m gridcore.cli convert book.xlsx book.csv
poetry run python -m gridcore.cli commands                                      # every command id
poetry run python -m gridcore.cli functions lookup
```

## Working together

Every workbook has an address, `/book/<id>`; everyone who opens it edits the same workbook. The first visit
to a new address creates the workbook, so a link is all it takes to share one (**Share** in the title bar
copies it). The workbook list at `/` shows every workbook, who is editing which right now, and lets you
search, filter (by creator, date, "being edited now"), create, upload, download and delete them.

What people on a workbook share, and what is their own:

| Shared, live | Each person's own |
|---|---|
| Cells, formulas and formats (every change reaches everyone at once and lights up briefly in its author's colour), rows, columns, merges, rules, tables, charts, sheets and their names, the workbook's title, the call | The sheet on screen, the selection, zoom, *Show Formulas*, the clipboard, Find and the Format Painter, undo and redo |

- **Who is where**: avatars and "N here" in the title bar (click an avatar to jump to that person); on the
  grid, each person's selection in their colour with their name, a dashed box with ✎ while they type.
- **Conflicts**: two people may change the same cell; the last change wins, as in Google Sheets.
- **Undo is per person**: it takes back your own changes, not the ones others made since. When nobody has
  changed the workbook after you, undo restores the workbook exactly as it was; otherwise it puts back what
  you changed and nobody touched since, and keeps (and tells you about) what others changed meanwhile.
  Matching goes by address, so after someone moves cells (inserts or deletes rows or columns, sorts, moves
  cells, or adds, removes or reorders sheets — adding one at the end is fine), your earlier changes can't be
  undone any more, and Undo says why.
- **Saving**: workbooks are saved on the server as `.xlsx` a moment after each change ("All changes saved").
  They live in `REFLEX_GRID_HOME` (default `~/Library/Application Support/reflex_grid` on macOS,
  `~/.local/share/reflex_grid` on Linux), as `books/<id>.xlsx` plus `books/<id>.json` (title, creator,
  times); a workbook with macros is `books/<id>.xlsm`.

## With re_ddns: sign-in and calls

ReflexGrid runs fine on its own (everyone is a guest). Sign-in and **Call everyone** come from
[re_ddns](https://github.com/milochen0418/re_ddns) and its **Agentic App Store**: once the App Store is up,
install a few apps by hand and they work.

1. Start re_ddns (see its README) and open the App Store at `https://aapps.reflex-ddns.com`.
2. Install **ReflexGrid**, then open it at `https://grid.reflex-ddns.com`.
3. Install **Relack** for sign-in: people then appear with their name and picture, and **Mine** filters
   the workbook list.
4. Install a call app, e.g. **LiveKit Audio Chat (self-hosted)**, **LiveKit Video Chat** or **LiveKit Avatar
   Chat**: **Call everyone** rings the people on the workbook and opens the call in a dialog. With several
   call apps installed, whoever starts the call picks one.

Apps on re_ddns open each other in dialogs through **DDNS Intents** (a call is the intent `call.join`); this
video explains the idea: [DDNS Intent](https://www.youtube.com/watch?v=Y5YKXsWkrxg).

## Tests

```bash
poetry run pytest -q                              # engine, collaboration, formulas, functions, number formats, xlsx
poetry run ./run_test_suite.sh smoke_home         # Playwright end-to-end suites (see testcases/README.md)
poetry run ./run_test_suite.sh collab_live        # three people editing one workbook
poetry run ./run_test_suite.sh vba_macros         # the VBA editor, macro dialogs, event macros, .xlsm
poetry run ./run_test_suite.sh vba_userforms      # UserForms: modal and modeless, the editor's form view
poetry run python scripts/excel_vba_oracle_mac.py --objects   # macros compared with Microsoft Excel (Mac, opt-in)
```

## Project structure

```
gridcore/                 the spreadsheet engine (no UI dependencies)
  core/                   addresses, values and errors, dates, typed-input parsing
  numfmt/                 number-format codes: parsing and rendering
  formula/                lexer, Pratt parser, AST, printer, reference adjustment
  functions/              the worksheet function library (478 functions)
  model/                  workbook, sheets, copy-on-write cell store, styles, features
  calc/                   evaluator, dependency graph, recalculation, dynamic arrays
  xlsx/                   XLSX and CSV readers/writers; vbaProject.bin (cfb.py, ovba.py), UserForm designs
                          (oforms.py)
  vba/                    the VBA language: lexer, parser, compiler, runtime, VBA's library, the Forms
                          object model (msforms.py)
  engine/                 sessions, commands (cmd_*.py), undo/redo, clipboard, fill, viewport payloads,
                          several people on one workbook (collab.py)
    macro/                Excel's object model for VBA, macro runs, event macros, VBA worksheet functions
reflex_grid/              the Reflex app
  state.py                bridges a tab's UI events to the shared workbook it shows
  books.py                the shared workbooks: stored as .xlsx, opened once, saved as they change
  room.py                 BookRoom, the SharedState that tells everyone on a workbook about changes and people
  sessions.py             per-tab engine sessions (kept outside Reflex state)
  workbooks.py            the workbook list
  call.py                 calls on a workbook (DDNS Intent call.join)
  api.py                  the backend's own route (a closing tab's goodbye)
  components/             workbook list, ribbon, dialogs, chrome, calls, macros (vba.py), UserForms
                          (userform.py), and the grid's Reflex wrapper
assets/grid/              the grid itself, a React component (sheet_view.jsx, sheet_charts.jsx), the
                          VBA code editor (vba_editor.jsx) and UserForms (userform.jsx)
tests/                    unit tests
testcases/                Playwright E2E suites
docs/architecture.md      design notes
docs/vba.md               Excel macros: design, security, compatibility and limits
```

## Credits

- [Reflex](https://reflex.dev/) for the web framework; [Radix Themes](https://www.radix-ui.com/themes) and
  [Lucide](https://lucide.dev/) icons as bundled by Reflex.
- The workbook list, presence and calls follow [CoDoc in MD](https://github.com/milochen0418/codoc_in_md).

ReflexGrid is not affiliated with Microsoft. "Excel" is a trademark of Microsoft Corporation, used here only to
describe file compatibility.

## License

MIT — see [LICENSE](LICENSE).
