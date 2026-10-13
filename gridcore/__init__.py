"""gridcore: an Excel-compatible spreadsheet engine in pure Python.

Layered: ``core`` (addresses, values, dates, typed input) → ``numfmt`` (number formats) →
``formula`` (parser, printer, reference adjustment) → ``functions`` (worksheet functions) →
``model`` (workbook, sheets, styles) → ``calc`` (evaluator, dependency graph, dynamic arrays) →
``xlsx`` (XLSX/CSV files) → ``engine`` (session, commands, undo, clipboard). Nothing in here
depends on the UI.
"""

__version__ = "0.1.0"
