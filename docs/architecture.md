# ReflexGrid architecture

ReflexGrid has two halves: **gridcore**, a spreadsheet engine with no UI dependencies, and **reflex_grid**,
the Reflex web app. This note explains the parts you
need to know before changing either half.

## gridcore layers

Each layer only imports the ones before it.

| Layer | Module | Responsibility |
|---|---|---|
| core | `gridcore/core` | `RangeRef` and A1 addresses (1,048,576 × 16,384), values (`None`, `float`, `str`, `bool`, `CellError`, `Array`), Excel comparison and coercion rules, the 1900/1904 date systems (including the Lotus 1900 leap-year bug), parsing what people type (`50%`, `$1,200`, `3/14/2026`, `1 1/2`) |
| numfmt | `gridcore/numfmt` | Number-format codes: sections, conditions, colours, dates/times, fractions, scientific, `*` fill and `_` padding; General formatting that fits a column width |
| formula | `gridcore/formula` | Lexer, Pratt parser into a frozen-dataclass AST (refs, ranges, structured refs, arrays, `@`, `#` spill, `.:` trim), printer (A1 and R1C1), reference adjustment for copy, insert/delete, move, sheet rename/delete |
| functions | `gridcore/functions` | The registry (arity, argument kinds, signature, description) and 478 functions with Excel's array lifting |
| model | `gridcore/model` | `Workbook` → `Sheet` → copy-on-write `CellStore`; interned immutable `Style`s; tables, conditional formats, validation, notes, links, filters, charts, pictures, sparklines, print settings, protection, defined names |
| calc | `gridcore/calc` | The evaluator (references stay lazy, special forms such as `IF`, `LET`, `LAMBDA`, `OFFSET`, `INDIRECT`) and `Calc`: a dependency graph indexed by cell, row block, tall and wide ranges; dirty closure + topological order; spills and `#SPILL!`; cycles (`#CIRC!` or iterative calculation) |
| xlsx | `gridcore/xlsx` | Streaming XLSX reader and writer (ElementTree), styles/theme/drawing/chart/comment parts, CSV/TSV |
| engine | `gridcore/engine` | `Session` (open documents, clipboard, journal), the command registry, undo/redo, structural edits, fill, conditional-format evaluation and the viewport payloads the UI draws |

### Snapshots, undo and recalculation

A `Workbook` is never edited in place by a command. `commit()` clones it (sheets are shared until
`sheet_mut()` copies one; inside a sheet the cell store shares 64-row bands until one is written), runs the
edit on the clone, recalculates, and swaps it in. The previous workbook object *is* the undo snapshot, so undo
is O(1) and safe: nothing else holds a mutable reference to it.

Recalculation is incremental: changed cells seed a dirty closure through the dependency graph, which is
evaluated in topological order. Structural edits (rows, columns, sheets, names) rebuild the graph and
recalculate everything. Dynamic arrays spill into `Sheet.spill`; a blocked spill turns its anchor into
`#SPILL!` and is re-checked when the blocking cells change.

### Commands

Every user-visible action is a command (`@command("home.bold", "Bold", menu=…, shortcut=…, params=…)`).
`Session.execute(id, params)` runs it, records it in the journal, pushes an undo entry when the workbook
changed, and turns any unexpected exception into an `EngineError` after restoring the workbook — the UI
never sees a half-applied edit.

## reflex_grid

- **Sessions** (`reflex_grid/sessions.py`): one engine `Session` per browser tab, keyed by Reflex's client
  token and kept in the backend process. Reflex state carries only what the browser needs to draw.
- **State** (`reflex_grid/state.py`): `GridState.on_event` receives every grid event
  (`viewport`, `select`, `commit`, `command`, `copy`, `paste`, `resize`, `fill`, `move`, `filter`, `link`,
  `object`, `dialog`, `namebox`), runs engine commands and refreshes four props:
  `geometry` (sizes, freeze, view flags), `data` (formatted cells, styles, merges, CF looks, table looks,
  filter buttons, notes, sparklines for the requested window), `objects` (charts with resolved data,
  pictures) and `ui` (selection, formula bar, name box, status statistics, ribbon state, dropdown lists).
  Each event carries a sequence number; the browser only adopts the server's selection once the server has
  processed its latest event.
- **Grid** (`reflex_grid/components/sheet_view.jsx`): a local React component. Gridlines, fills, data bars
  and borders are drawn on a canvas; text, selection, editor, notes and objects are DOM. The view is
  virtualised: only visible cells render, and it asks the server for the cells around the visible window.
  Typing goes into a hidden textarea placed over the active cell, which becomes the in-cell editor — so IME
  input (e.g. Chinese) composes in place.
- **Charts** (`reflex_grid/components/sheet_charts.jsx`): an SVG renderer for the chart kinds the engine
  models, plus sparklines and conditional-format icons.

## Excel compatibility rules worth knowing

- Newer functions are written with their `_xlfn.` / `_xlws.` prefixes, LET/LAMBDA parameters with `_xlpm.`.
- A formula whose result spills is written as an array formula with `cm="1"` and the dynamic-array metadata
  part, which is how Excel marks dynamic arrays.
- Structured references are written the way Excel writes them in files — always with the table name and
  never with `@` (`Table1[[#This Row],[Units]]`) — and shown as `[@Units]` inside their own table.
- `#SPILL!` and `#CALC!` are cached as `#VALUE!`, as they are not error codes of the file format.
- Differential formats (conditional formatting) put a solid fill's colour in `bgColor`, as Excel does.
