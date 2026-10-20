# E2E testcases

This folder contains Playwright end-to-end suites for ReflexGrid.

## Structure

Each suite lives under:

- `testcases/<suite_name>/run_test.py`
- `testcases/<suite_name>/output/` (artifacts, ignored by git)

Shared helpers (opening a workbook, clicking cells, reading cell text, several people on one workbook) are
in `testcases/_lib/grid.py`. Each suite works on new workbooks of its own (a fresh `/book/<id>` address),
and the runner keeps them in the suite's `output/data` (`REFLEX_GRID_HOME`), not with your workbooks.

The `run_test.py` script should:
- Exit `0` on success
- Exit non-zero on failure
- Write screenshots/logs into `OUTPUT_DIR` (provided by the runner)

## Suites

| Suite | What it checks |
|---|---|
| `smoke_home` | A workbook loads: ribbon, formula bar, grid headers, sheet tabs, status bar, title bar (title, save status, who is here); no console errors |
| `editing_formulas` | Typing values and formulas, recalculation, formula bar, autocomplete, AutoSum, fill handle, undo/redo, copy/paste (a block of more than 1 MB from another app, then copied and pasted again inside the sheet, and pasting after Esc) |
| `formatting_ribbon` | Bold, fill colour, number formats, merge, borders, Format Cells dialog, column resize |
| `sheets_data` | Sheet tabs (add/rename/switch), sort, filter dialog, tables, conditional formatting, charts from samples, dragging a chart and a picture (they move without resizing; one undo per drag) |
| `insert_tab` | The Insert tab button by button: every chart kind of the chart menus and All Charts, sparklines of each type, links (ScreenTip on hover, Ctrl/Cmd+click), comment threads (post, reply, resolve) and notes shown on hover, a cell holding a note or a comment but not both, Insert Function, a table and a picture |
| `xlsx_roundtrip` | Upload an .xlsx built with openpyxl, check what renders, download it again and re-read it with openpyxl |
| `collab_live` | Three people on one workbook: presence, each other's selections and typing, edits and formulas reaching everyone, change highlights, per-person sheets, per-person undo, the shared title, jumping to someone, leaving and closing a tab |
| `workbook_manager` | The workbook list: who is editing, a new workbook and its title, search and filters, uploading, downloading, deleting a workbook someone has open |
| `vba_macros` | Excel macros: the Developer tab and VBA editor, F5 with `MsgBox`/`InputBox` answers, a VBA function in a formula, one undo for a run, the Macros dialog, event macros behind the security bar, stopping a long macro, downloading `.xlsm` and opening it again |
| `vba_userforms` | UserForms: an `.xlsm` with forms (built by the suite) opened in the browser; a modal form filled in (typing, a drop-down list, a check box, a `KeyPress` handler, a `MsgBox` over the form, OK writing the cells, Esc as Cancel), a modeless form beside the grid, the VBA editor's form view (double-click for an event procedure, Run Form), designing a new form (Toolbox, Properties window, moving a control, its code, running it), the forms kept in the downloaded `.xlsm`, and the editor staying light in a dark-mode browser |
| `collab_call` | Calling everyone on a workbook: who is rung, declining, accepting, joining later; with the call app running, the call itself (needs a `call.join` provider, see below) |

## Running

Use the repo-level runner (it manages starting/stopping the Reflex server and collecting artifacts):

```bash
poetry run ./run_test_suite.sh smoke_home
```

Optional env vars:
- `BASE_URL` (default `http://127.0.0.1:3000`)
- `HEADLESS=0` to watch the browser
- `PW_SLOWMO_MS=100` to slow actions down so you can see interactions
- `PW_TIMEOUT_MS=60000` to increase Playwright timeouts

`collab_call` needs a `call.join` provider named for the server; the call app itself only has to run for the
check of the call itself (e.g. reflex_ddns_livekit_audio_chat on 3200/8200):

```bash
DDNS_INTENT_PROVIDER_CALL_JOIN=livekit DDNS_INTENT_URL_LIVEKIT=http://localhost:3200 \
poetry run ./run_test_suite.sh collab_call
```

Example (headed + slow motion):

```bash
HEADLESS=0 PW_SLOWMO_MS=100 poetry run ./run_test_suite.sh editing_formulas
```
