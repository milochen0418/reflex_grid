"""VBA: the macro language of Excel workbooks, run in a sandbox.

``lexer`` → ``parser`` → ``ast`` → ``compiler`` (closures) → ``runtime``; ``values`` holds the
Variant types and VBA's conversion and operator rules, ``stdlib`` the VBA library (strings,
math, dates, Format, Collection, Dictionary…). Nothing here knows Excel: the workbook objects
(Application, Range…) come from a host (:mod:`gridcore.engine.macro`)."""
