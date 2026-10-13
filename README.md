# ReflexGrid

ReflexGrid is an Excel-compatible spreadsheet that runs in the browser, built with [Reflex](https://reflex.dev/).
It opens and saves real `.xlsx` workbooks (formulas, cached values, styles, tables, charts…), so files move
between ReflexGrid and Microsoft Excel without conversion.

Its engine, **gridcore**, is pure Python with a layered design (core → number formats → formulas → functions →
model → calc → xlsx → engine), copy-on-write workbook snapshots for undo, dependency-graph recalculation with
dynamic arrays, and an "every action is a command" API.

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
- **Scriptable**: every user-visible action is a command with a stable id (`cell.set`, `home.bold`,
  `data.sortAscending`…); the UI, the tests and your own scripts all use the same `Session.execute` API.

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
| Files | Open `.xlsx`/`.xlsm`/`.csv`/`.tsv` (upload), save `.xlsx`, export CSV, print (with the sheet's page setup), save as web page, document properties, several open workbooks |

### Excel compatibility

Read and written in `.xlsx`: cell values and formulas (shared, array, dynamic-array with spill metadata;
data-table formulas are read), cached results, shared strings and rich text, styles (fonts, fills, borders,
alignment, protection, number formats, themes, differential formats), column widths and row heights,
hidden and outline state, merged cells, frozen panes and view settings, defined names (incl. print areas and
titles), tables (styles, totals, calculated columns), AutoFilter and filter criteria, conditional formatting
(incl. data bars, colour scales and icon sets), data validation, hyperlinks, notes and threaded comments,
charts (DrawingML with cached data), pictures, sparklines, page setup and margins, sheet and workbook
protection, the 1904 date system and document properties. Unknown parts are skipped with a warning, so a
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

Then open http://localhost:3000 (the backend listens on port 8000). **File › New** offers three sample
workbooks (Sales, Budget, Grades) that show tables, charts, conditional formats and validation.

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

## Tests

```bash
poetry run pytest -q                              # engine, formulas, functions, number formats, xlsx
poetry run ./run_test_suite.sh smoke_home         # Playwright end-to-end suites (see testcases/README.md)
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
  xlsx/                   XLSX and CSV readers/writers
  engine/                 sessions, commands (cmd_*.py), undo/redo, clipboard, fill, viewport payloads
reflex_grid/              the Reflex app
  state.py                bridges UI events to the engine session of each browser tab
  sessions.py             per-tab engine sessions (kept outside Reflex state)
  components/             ribbon, dialogs, chrome, and the grid (sheet_view.jsx, sheet_charts.jsx)
tests/                    unit tests
testcases/                Playwright E2E suites
docs/architecture.md      design notes
```

## Credits

- [Reflex](https://reflex.dev/) for the web framework; [Radix Themes](https://www.radix-ui.com/themes) and
  [Lucide](https://lucide.dev/) icons as bundled by Reflex.

ReflexGrid is not affiliated with Microsoft. "Excel" is a trademark of Microsoft Corporation, used here only to
describe file compatibility.

## License

MIT — see [LICENSE](LICENSE).
