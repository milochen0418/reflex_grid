"""Sample workbooks (File › New from sample), built entirely with engine commands."""

from __future__ import annotations

from typing import Callable, Dict

from .session import DocState, Session


def _sales(s: Session) -> None:
    run = s.run
    run("sheet.rename", name="Sales")
    rows = [["Region", "Product", "Month", "Units", "Price", "Revenue"]]
    regions = ["North", "South", "East", "West"]
    products = ["Widget", "Gadget", "Gizmo"]
    prices = {"Widget": 12.5, "Gadget": 24.0, "Gizmo": 7.25}
    n = 0
    for m in range(1, 7):
        for i, reg in enumerate(regions):
            prod = products[(m + i) % 3]
            units = 20 + ((m * 37 + i * 53) % 90)
            rows.append([reg, prod, f"2026-{m:02d}-01", units, prices[prod], None])
            n += 1
    run("range.setValues", range="A1", values=rows)
    run("selection.set", range="A1")
    run("insert.table", style="TableStyleMedium2")
    run("cell.set", cell="F2", input="=[@Units]*[@Price]")
    run("home.numberFormat", range=f"C2:C{n + 1}", code="mmm yyyy")
    run("home.numberFormat", range=f"E2:F{n + 1}", format="currency")
    run("home.conditionalFormat", range=f"F2:F{n + 1}", rule={"type": "dataBar", "color": "#4EA72E"})
    run("range.setValues", range="H1", values=[["Region", "Revenue", "Share"]] + [[r, f"=SUMIFS(Table1[Revenue],Table1[Region],H{k + 2})", f"=I{k + 2}/SUM($I$2:$I$5)"] for k, r in enumerate(regions)])
    run("home.bold", range="H1:J1")
    run("home.numberFormat", range="I2:I5", format="currency")
    run("home.numberFormat", range="J2:J5", code="0.0%")
    run("home.columnWidth", cols="A:C", width=80)
    run("home.columnWidth", cols="D:E", width=64)
    run("home.columnWidth", cols="F", width=96)
    run("home.columnWidth", cols="H:J", width=90)
    run("selection.set", range="H1:I5")
    run("insert.chart", kind="columnClustered", title="Revenue by Region")
    run("selection.set", range="A1")
    run("view.freezePanes", cell="A2")


def _budget(s: Session) -> None:
    run = s.run
    run("sheet.rename", name="Budget")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    cats = [("Rent", 1800), ("Utilities", 240), ("Groceries", 620), ("Transport", 180), ("Insurance", 150), ("Leisure", 300), ("Savings", 900)]
    run("cell.set", cell="A1", input="Household Budget 2026")
    run("home.cellStyle", range="A1", name="Title")
    run("range.setValues", range="A3", values=[["Category"] + months + ["Total", "Average"]])
    run("home.cellStyle", range="A3:O3", name="Heading 3")
    for i, (cat, base) in enumerate(cats):
        r = 4 + i
        vals = [cat] + [round(base * (1 + ((m * 7 + i * 3) % 11 - 5) / 100.0), 2) for m in range(12)]
        run("range.setValues", range=f"A{r}", values=[vals])
        run("cell.set", cell=f"N{r}", input=f"=SUM(B{r}:M{r})")
        run("cell.set", cell=f"O{r}", input=f"=AVERAGE(B{r}:M{r})")
    last = 3 + len(cats)
    tot = last + 1
    run("cell.set", cell=f"A{tot}", input="Total")
    run("selection.set", range=f"B{tot}:O{tot}")
    for c in "BCDEFGHIJKLMNO":
        run("cell.set", cell=f"{c}{tot}", input=f"=SUM({c}4:{c}{last})")
    run("home.cellStyle", range=f"A{tot}:O{tot}", name="Total")
    run("home.numberFormat", range=f"B4:O{tot}", code="#,##0.00")
    run("home.conditionalFormat", range=f"B4:M{last}", rule={"type": "colorScale", "scale": "greenYellowRed"})
    run("home.columnWidth", cols="A", width=110)
    run("home.columnWidth", cols="B:O", width=84)
    run("selection.set", range="A3")
    run("view.freezePanes", cell="A4")


def _grades(s: Session) -> None:
    run = s.run
    run("sheet.rename", name="Grades")
    names = ["Avery", "Blake", "Casey", "Devon", "Emery", "Finley", "Gray", "Harper", "Indy", "Jordan"]
    rows = [["Student", "Quiz 1", "Quiz 2", "Midterm", "Final", "Average", "Grade"]]
    for i, n in enumerate(names):
        scores = [55 + (i * 17 + k * 29) % 45 for k in range(4)]
        rows.append([n] + scores + [None, None])
    run("range.setValues", range="A1", values=rows)
    last = len(names) + 1
    for r in range(2, last + 1):
        run("cell.set", cell=f"F{r}", input=f"=AVERAGE(B{r}:E{r})")
        run("cell.set", cell=f"G{r}", input=f'=IFS(F{r}>=90,"A",F{r}>=80,"B",F{r}>=70,"C",F{r}>=60,"D",TRUE,"F")')
    run("home.bold", range="A1:G1")
    run("home.fillColor", range="A1:G1", color="#DCE6F2")
    run("home.borders", range=f"A1:G{last}", preset="all", color="#A6A6A6")
    run("home.numberFormat", range=f"F2:F{last}", code="0.0")
    run("home.alignCenter", range=f"G2:G{last}")
    run("data.validation", range=f"B2:E{last}", kind="decimal", operator="between", f1="0", f2="100",
        errorTitle="Invalid score", errorMessage="Scores are between 0 and 100.", inputTitle="Score", inputMessage="Enter a score from 0 to 100.")
    run("home.conditionalFormat", range=f"F2:F{last}", rule={"type": "iconSet", "iconSet": "3TrafficLights1"})
    run("home.conditionalFormat", range=f"G2:G{last}", rule={"type": "cellIs", "operator": "equal", "value": "F", "preset": "lightRed"})
    run("range.setValues", range="I1", values=[["Look up", "Average"], ["Casey", f"=XLOOKUP(I2,A2:A{last},F2:F{last})"], ["Class average", f"=AVERAGE(F2:F{last})"], ["Top score", f"=MAX(F2:F{last})"], ["Best student", f"=INDEX(A2:A{last},MATCH(J4,F2:F{last},0))"]])
    run("home.bold", range="I1:J1")
    run("home.numberFormat", range="J2:J4", code="0.0")
    run("data.validation", range="I2", kind="list", f1=f"=$A$2:$A${last}")
    run("home.columnWidth", cols="A", width=90)
    run("home.columnWidth", cols="I", width=110)
    run("review.newNote", cell="I2", text="Pick a student from the list.")
    run("selection.set", range="A1")


SAMPLES: Dict[str, Callable[[Session], None]] = {"Sales": _sales, "Budget": _budget, "Grades": _grades}


def new_sample(s: Session, name: str) -> DocState:
    build = SAMPLES[name]
    d = s.new_workbook()
    build(s)
    d.name = f"{name}.xlsx"
    d.undo.clear()
    d.redo.clear()
    d.saved = d.wb
    return d
