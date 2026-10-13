# E2E testcases

This folder contains Playwright end-to-end suites for ReflexGrid.

## Structure

Each suite lives under:

- `testcases/<suite_name>/run_test.py`
- `testcases/<suite_name>/output/` (artifacts, ignored by git)

Shared helpers (opening the app, clicking cells, reading cell text) are in `testcases/_lib/grid.py`.

The `run_test.py` script should:
- Exit `0` on success
- Exit non-zero on failure
- Write screenshots/logs into `OUTPUT_DIR` (provided by the runner)

## Suites

| Suite | What it checks |
|---|---|
| `smoke_home` | The page loads: ribbon, formula bar, grid headers, sheet tabs, status bar; no console errors |
| `editing_formulas` | Typing values and formulas, recalculation, formula bar, autocomplete, AutoSum, fill handle, undo/redo, copy/paste |
| `formatting_ribbon` | Bold, fill colour, number formats, merge, borders, Format Cells dialog, column resize |
| `sheets_data` | Sheet tabs (add/rename/switch), sort, filter dialog, tables, conditional formatting, charts from samples |
| `xlsx_roundtrip` | Upload an .xlsx built with openpyxl, check what renders, download it again and re-read it with openpyxl |

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

Example (headed + slow motion):

```bash
HEADLESS=0 PW_SLOWMO_MS=100 poetry run ./run_test_suite.sh editing_formulas
```
