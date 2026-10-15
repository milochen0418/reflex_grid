# VBA-JSON's spec workbook (third-party test case)

`VBA-JSON-Specs.xlsm` is the spec workbook of [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) v2.3.1
by Tim Hall, MIT-licensed (see `LICENSE`), unchanged except for its file name:

| | |
|---|---|
| Source | `specs/VBA-JSON - Specs.xlsm` at commit [`06b6c45`](https://github.com/VBA-tools/VBA-JSON/tree/06b6c45a2e8086a7a761f00bead7460852d6b602) |
| SHA-256 | `c5a4db4d19b2992b2aeb4f9b398ce82aea1c77058b1a369893c8d497857e7cee` |

It was written for Microsoft Excel by VBA-JSON's authors, with their own spec suite (VBA-TDD), so it
checks ReflexGrid with code and expectations that come from outside ReflexGrid.
`tests/test_third_party_vba.py` opens it in ReflexGrid, adds a small module that runs the suite and
writes each spec's result to a sheet, and turns each spec into a test. Its macros only ever run in
ReflexGrid's sandbox, never in Excel.
