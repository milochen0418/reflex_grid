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
| xlsx | `gridcore/xlsx` | Streaming XLSX reader and writer (ElementTree), styles/theme/drawing/chart/comment parts, CSV/TSV, VBA projects (`cfb.py` compound files, `ovba.py` the `vbaProject.bin` format) |
| vba | `gridcore/vba` | The VBA language, independent of Excel: lexer, parser, compiler to closures, runtime (Variants, ByRef, arrays, classes, `On Error`), VBA's function library; a *host* supplies the application's objects (see [vba.md](vba.md)) |
| engine | `gridcore/engine` | `Session` (open documents, clipboard, journal), the command registry, undo/redo, structural edits, fill, conditional-format evaluation, the viewport payloads the UI draws, several people on one workbook (`collab.py`), and macros (`macro/`: Excel's object model as the VBA host, runs, event macros, VBA worksheet functions; `cmd_macro.py`) |

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

### Several people on one workbook (`collab.py`)

A `SharedBook` holds the one workbook everybody edits (a `DocState`: snapshots, calculation, names) and a
`View` per person — per browser tab: the sheet they look at (by name, so it survives others adding or
deleting sheets), their selection, zoom and *Show Formulas*, and their own undo and redo stacks. Each tab
has a `CollabSession` (its clipboard, Find, Format Painter and preferences) whose document is the shared
book.

Everything a person does runs inside `book.act(view)`, under the book's lock: their sheet and selection are
put in place on the shared `DocState`, the commands run as in a single-user session, and on the way out the
view takes its new selection back. If the workbook changed, the act records it: `rev` goes up, `changes`
gets an entry (who, and the bounding box of the cells whose content or format changed, for highlighting),
listeners hear about it (the app saves), and the person's undo stack gets a `Transition` from the workbook
before to the workbook after. The engine's own undo stacks stay empty; `edit.undo`, `edit.redo`, `view.zoom`
and `formulas.showFormulas` are answered by the session for the person, and the file commands
(`file.open`, `file.saveAs`…) are refused: the app opens, saves and closes shared workbooks.

Undo is per person. A transition is undone exactly — the workbook goes back to its snapshot — when nobody
changed the workbook after it (`book.doc.wb is t.dst`). Otherwise it is undone selectively: for each sheet,
the cells it changed (`CellStore.diff_keys`, ignoring formulas whose cached result merely changed) go back
where the current workbook still holds what the transition left; row and column sizes, notes and links
likewise entry by entry; any other sheet or workbook attribute (merges, rules, tables, names…) as a whole
when it still is as the transition left it. What others changed since is kept and counted, and the session
tells the person. Redo is the same thing in the other direction (undoing an undo).

Selective undo matches by address, which is only right while cells stay where they were. Edits that move
cells mark `Ctx.moved` (rows or columns inserted or deleted, cells shifted or moved, sorts, removing
duplicates, sheets inserted before others, deleted, moved or copied), and the commit sets
`DocState.cells_moved`; the book then starts a new *epoch*. A transition recorded in another epoch, or one
that moved cells itself, can only be undone exactly; otherwise undo refuses and says why. Undoing a move
exactly puts its epoch back, so earlier transitions become undoable again.

## reflex_grid

- **Books** (`reflex_grid/books.py`): the shared workbooks. Each is `books/<id>.xlsx` (the workbook, written
  by the xlsx writer, so Excel opens it; `<id>.xlsm` when it has macros) plus `books/<id>.json` (title, creator, times, sheet names) under
  `REFLEX_GRID_HOME`. `STORE.open(id)` reads one into a `Book` (a `SharedBook`) shared by every tab on it; a
  new address creates a blank book. A change schedules a save `SAVE_DELAY` seconds later on a timer thread:
  the workbook snapshot is taken under the book's lock and encoded outside it. Books nobody has open are
  dropped from memory after `IDLE_CLOSE`.
- **Room** (`reflex_grid/room.py`): `BookRoom`, a Reflex `SharedState`. Every tab on a book links its
  `BookRoom` to the book (`book-<id>`), so what one tab assigns there reaches all of them: the book's `rev`,
  `title`, `peers` (each tab's name, colour, sheet, selection and the cell it is typing in), the latest
  `changes`, whether it is saved, whether it was deleted, and the call. Cells are not in the room: a sheet can
  hold millions and each person scrolls elsewhere, so each tab keeps its own window in `GridState`, and a tab
  that drew an older revision than the room's asks for its window again (a `sync` event).
- **Sessions** (`reflex_grid/sessions.py`): one engine `CollabSession` per browser tab, keyed by Reflex's
  client token and kept in the backend process, with the tab's book and view. Reflex state carries only what
  the browser needs to draw.
- **State** (`reflex_grid/state.py`): `GridState.on_load` opens the book in the URL, joins it and links the
  room. `GridState.on_event` receives every grid event
  (`viewport`, `select`, `commit`, `command`, `copy`, `paste`, `resize`, `fill`, `move`, `filter`, `link`,
  `object`, `dialog`, `namebox`, and `sync`, `ping`, `presence`), runs engine commands inside
  `book.act(view)` and refreshes four props:
  `geometry` (sizes, freeze, view flags), `data` (formatted cells, styles, merges, CF looks, table looks,
  filter buttons, notes, sparklines for the requested window), `objects` (charts with resolved data,
  pictures) and `ui` (selection, formula bar, name box, status statistics, ribbon state, dropdown lists).
  Each event carries a sequence number; the browser only adopts the server's selection once the server has
  processed its latest event, and never from a *passive* refresh (someone else's change, a `ping`, a scroll),
  which would undo a move it hasn't sent yet. After every event the handler publishes the book to the room
  (only what changed is assigned, so only that is sent).
- **Presence**: each tab pings every 5 s; a tab not heard from for `PEER_IDLE` seconds is dropped. A closing
  tab says goodbye with `navigator.sendBeacon` to `POST /_reflex_grid/bye` (`api.py`, named by its client
  token), and `push()` shows the book to the others at once by publishing it from one of them: the room, a
  SharedState, carries it to the rest. The same push tells everyone when a save finishes and when the book
  is deleted from the list.
- **Calls** (`reflex_grid/call.py`, from CoDoc in MD): the call's members and its private id stay in the
  process (`CallStore`); the room carries what everyone must see (how many are in it, who is ringing, whom it
  rings and who answered). The call runs in the app chosen for the DDNS Intent `call.join`.
- **Cost of a shared state**: Reflex re-sends its router data with each update a linked tab gets (about 2 KB
  next to the room's own data), so every move of a selection costs each other person a small message —
  fine for a team, not meant for hundreds of people on one workbook.
- **Grid** (`assets/grid/sheet_view.jsx`, wrapped in `reflex_grid/components/sheet_view.py`): a local React component. Gridlines, fills, data bars
  and borders are drawn on a canvas; text, selection, the other people's selections (with name tags),
  change highlights, editor, notes and objects are DOM. The view is
  virtualised: only visible cells render, and it asks the server for the cells around the visible window.
  Typing goes into a hidden textarea placed over the active cell, which becomes the in-cell editor — so IME
  input (e.g. Chinese) composes in place.
- **Charts** (`assets/grid/sheet_charts.jsx`): an SVG renderer for the chart kinds the engine
  models, plus sparklines and conditional-format icons.

## Excel compatibility rules worth knowing

- Newer functions are written with their `_xlfn.` / `_xlws.` prefixes, LET/LAMBDA parameters with `_xlpm.`.
- A formula whose result spills is written as an array formula with `cm="1"` and the dynamic-array metadata
  part, which is how Excel marks dynamic arrays.
- Structured references are written the way Excel writes them in files — always with the table name and
  never with `@` (`Table1[[#This Row],[Units]]`) — and shown as `[@Units]` inside their own table.
- `#SPILL!` and `#CALC!` are cached as `#VALUE!`, as they are not error codes of the file format.
- Differential formats (conditional formatting) put a solid fill's colour in `bgColor`, as Excel does.

## Macros

A workbook's VBA project is part of its snapshot (`Workbook.vba`, an immutable `VbaProject`), so editing code
is an ordinary undoable, shared edit. A macro runs through the engine like a person: `Range.Value` writes are
batched into one open edit, everything else is an engine command, and the whole run is one undo step (a
`group` on the history entry, which `collab.py` merges into one transition while nobody else edits). A run
that shows a `MsgBox` pauses on a worker thread outside `book.act`, so nobody waits on a dialog. Details,
security and compatibility: [vba.md](vba.md).
