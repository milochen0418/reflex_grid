---
description: Open-source, self-hosted online Excel. ReflexGrid opens and saves real .xlsx and .xlsm workbooks, runs Excel VBA macros and UserForms in the browser, and lets people edit the same workbook together, live.
---

ReflexGrid is a spreadsheet that runs in the browser, made for workbooks that have to keep working as they do
in Microsoft Excel:

- **Excel files in, Excel files out**: real `.xlsx` and `.xlsm` (formulas, styles, tables, charts,
  conditional formats, validation, the VBA project), and CSV or TSV.
- **Excel VBA macros in the browser**: a VBA editor, `MsgBox` and `InputBox`, event macros such as
  `Worksheet_Change`, VBA functions in formulas, and UserForms with a form designer. Macros run on the server,
  in a sandbox. Excel for the web can't run VBA; ReflexGrid can.
- **Edit together, live**: share a workbook's link; everyone sees changes at once and each other's
  selections, and undo takes back only your own changes.
- **A Python spreadsheet engine**: gridcore calculates 478 Excel functions, with dynamic arrays and
  `LET`/`LAMBDA`, and runs VBA macros headless, with no Excel installed.
- **Self-hosted, MIT-licensed.**

<a href="https://youtu.be/Zl0tysSTTOI"><img src="https://i.ytimg.com/vi/Zl0tysSTTOI/maxresdefault.jpg" width="640" alt="ReflexGrid demo video: co-editing, VBA macros and UserForms"></a>

▶ [Demo video: co-editing, VBA macros and UserForms](https://youtu.be/Zl0tysSTTOI) (3 min, English) ·
[the first version](https://www.youtube.com/watch?v=I5rBceuf1tM)

## Try it

With Python 3.11 or newer and [Poetry](https://python-poetry.org/):

```bash
git clone https://github.com/milochen0418/reflex_grid.git
git clone https://github.com/milochen0418/reflex_ddns_auth.git
cd reflex_grid
poetry install
poetry run ./reflex_rerun.sh        # then open http://localhost:3000
```

## Read more

- [README](https://github.com/milochen0418/reflex_grid#readme): every feature,
  [how it compares](https://github.com/milochen0418/reflex_grid#how-it-compares) with ONLYOFFICE, Collabora
  Online, Univer, Grist, Excel for the web and Google Sheets, and the
  [FAQ](https://github.com/milochen0418/reflex_grid#faq)
- [Excel macros (VBA) in ReflexGrid](vba.md): how macros and UserForms run, the sandbox, checks against
  Excel, limits
- [Architecture](architecture.md): the engine's layers, commands, undo, recalculation and collaboration
