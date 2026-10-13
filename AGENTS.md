# AGENTS.md

This file describes how an AI coding agent (and humans) should work in this repository.

## 0) Always read the docs first

Before making changes or running commands, **read all README.md files in this repo** and follow them.

- Repo overview, architecture & dev workflow: `README.md`
- Engine design notes: `docs/architecture.md`
- E2E test runner conventions: `testcases/README.md`

If you already read them earlier in the session, **re-check** when unsure.

## 1) Python / Poetry requirements (strict)

This project targets **Python 3.11** (always set the Poetry env to 3.11 before installing deps or running commands).

- Prefer Poetry-managed environments.
- If Poetry selected another Python (e.g. 3.12+), switch it:

```bash
poetry env use python3.11
poetry env info
poetry install
```

If `python3.11` is not available on PATH, use an absolute path to the 3.11 executable.

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
`HEADLESS=0` to watch the browser.

## 5) How the code is organised (keep it that way)

- `gridcore/` is the spreadsheet engine and must not import Reflex. Layers depend only downward:
  `core → numfmt → formula → functions → model → calc → xlsx → engine`.
- Every user-visible action is an engine **command** (`gridcore/engine/cmd_*.py`, registered with `@command`).
  The UI, tests and scripts all go through `Session.execute(id, params)`. Add a command (with a unit test)
  before wiring a button to it.
- The engine never crashes the app: unexpected errors become `EngineError` and the workbook is restored.
- `reflex_grid/` is the UI: `state.py` bridges events to the engine; `components/sheet_view.jsx` is the grid
  (a local React component); ribbon, dialogs and chrome are Reflex components.
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
