# AGENTS.md

This file describes how an AI coding agent (and humans) should work in this repository.

## 0) Always read the docs first

Before making changes or running commands, **read all README.md files in this repo** and follow them.

- Repo overview, architecture & dev workflow: `README.md`
- Engine design notes: `docs/architecture.md`
- Excel macros (VBA): `docs/vba.md`
- E2E test runner conventions: `testcases/README.md`

If you already read them earlier in the session, **re-check** when unsure.

## 1) Python / Poetry requirements (strict)

This project targets **Python 3.15** (always set the Poetry env to 3.15 before installing deps or running commands).
`pyproject.toml` still allows 3.11, the Python of the re_ddns App Store image: keep the range covering both.

- Prefer Poetry-managed environments.
- If Poetry selected another Python (e.g. 3.11), switch it:

```bash
poetry env use python3.15
poetry env info
poetry install
```

If `python3.15` is not available on PATH, use an absolute path to the 3.15 executable.

## 2) Default rule: use `poetry run` for scripts

When running any repo scripts, prefer:

- `poetry run <command>`

This ensures the correct virtualenv + dependencies are used.

## 3) Running the app: prefer `reflex_rerun.sh`

When you need to start/restart the Reflex app during development, **prefer the repo helper script**:

```bash
poetry run ./reflex_rerun.sh
```

Notes:
- Prefer `poetry run ./reflex_rerun.sh` over `poetry run reflex run`.
- It frees and uses the expected ports: frontend `3000`, backend `8000`.
- If you must run Reflex directly, still do it via Poetry: `poetry run reflex run`.

## 3.5) Recreate the env: use `proj_reinstall.sh`

To fully rebuild the local Poetry environment and clean Reflex artifacts, run:

```bash
./proj_reinstall.sh
```

Flags:

- `--remove-only`: only clean artifacts and remove Poetry envs; no new env/install.
- `--with-rerun`: after reinstall, run `poetry run ./reflex_rerun.sh` automatically.
- `--with-rerun-debug`: same, with `--loglevel debug`.

The script will auto-run `poetry lock` if `poetry install` fails due to a stale lockfile.

## 4) Tests

Unit tests (engine, formulas, functions, number formats, XLSX):

```bash
poetry run pytest -q
```

E2E suites live under `testcases/<suite_name>/run_test.py`. Use the repo runner, which starts/stops the
Reflex server, sets `OUTPUT_DIR` and collects artifacts:

```bash
poetry run ./run_test_suite.sh smoke_home
```

Optional environment variables (see `testcases/README.md`): `BASE_URL` (default `http://127.0.0.1:3000`),
`HEADLESS=0` to watch the browser. The runner keeps the workbooks a suite creates in its `output/data`;
when you run the app yourself, point `REFLEX_GRID_HOME` at a scratch folder to keep test workbooks out of
the user's own.

GitHub Actions (`.github/workflows/tests.yml`) runs the unit tests on Python 3.11 and 3.15 and each browser
suite on every push, and on `main` publishes the docs site with the Excel compatibility report
(`scripts/compat_report.py`). Keep it green: a new suite goes in the workflow's `suite` matrix, a new unit
test file in `AREAS` in `scripts/compat_report.py`.

## 5) How the code is organised (keep it that way)

- `gridcore/` is the spreadsheet engine and must not import Reflex. Layers depend only downward:
  `core → numfmt → formula → functions → model → calc → xlsx → vba → engine`. `gridcore/vba` (the VBA
  language) knows nothing of Excel; `gridcore/engine/macro` is its Excel host (see `docs/vba.md`).
- Every user-visible action is an engine **command** (`gridcore/engine/cmd_*.py`, registered with `@command`).
  The UI, tests and scripts all go through `Session.execute(id, params)`. Add a command (with a unit test)
  before wiring a button to it.
- The engine never crashes the app: unexpected errors become `EngineError` and the workbook is restored.
- `reflex_grid/` is the UI: `state.py` bridges events to the engine; `assets/grid/sheet_view.jsx` is the grid
  (a local React component, wrapped in `components/sheet_view.py`); ribbon, dialogs and chrome are Reflex
  components. Keep JSX as plain assets: `rx.asset(..., shared=True)` writes machine-specific symlinks into
  `assets/external/` that break the App Store dev mode, which mounts this checkout into a container.
- Workbooks are shared (see `docs/architecture.md`): a tab's engine work on its workbook runs inside
  `book.act(view)` (`gridcore/engine/collab.py`), never outside it and never across an `await`; after it,
  publish the book to its room (`GridState._publish`). Keep cells out of Reflex state: the room
  (`reflex_grid/room.py`, a `SharedState`) carries only small, broadcast things.
- Engine edits that move cells to other addresses set `ctx.moved` (per-person undo depends on it).
- The UI is light only, like Excel's: the theme is `rx.theme(appearance="light")`, the page declares
  `color-scheme: light`, and each page sets Reflex's colour mode back to light when it mounts (a browser
  may remember "system" or "dark" from before, which Reflex would apply over the theme and turn dialogs
  dark), so an OS in dark mode doesn't turn its text light on our white surfaces. On surfaces we colour
  ourselves (the title bar, the backstage menu, the grid), give text an explicit colour.
- Excel compatibility is a feature: anything written to `.xlsx` must open in Excel. When touching
  `gridcore/xlsx`, round-trip with openpyxl in a test.
- No Microsoft assets (icons, fonts, templates).

## 6) Expectations when acting as an agent

- Prefer the smallest, focused change set.
- After code changes, run the closest relevant check (unit tests, and the related suite via `run_test_suite.sh`).
- Report outcomes: what you ran, pass/fail, and where artifacts/logs were written.

## 7) Smart way to help the Agent to debug

```
poetry run ./reflex_rerun.sh --loglevel debug
```
