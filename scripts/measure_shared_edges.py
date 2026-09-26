"""Which border Excel reports for an edge two cells both draw: the upper cell's bottom against the lower one's top.

A macro setting an edge clears the neighbour's copy of it, but a cell
style's border does not: two stacked cells given styles can each draw the
edge between them. Custom styles here each give one border, and pairs of
cells in a row of columns take every pair of the thirteen line styles in
black, every pair of sixteen colours on thin lines, and a line style
against a colour. Both cells' Borders are read for the edge.

    python scripts/measure_shared_edges.py

writes tests/fixtures/cell_styles/shared_edges.json and the workbook,
shared_edges.xlsx, which tests/test_excel_cell_styles.py opens and reads.
"""

from __future__ import annotations

import itertools
import json
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "cell_styles"

#: (name, LineStyle, Weight) for the file's thirteen line styles.
LINES = [("hair", 1, 1), ("thin", 1, 2), ("medium", 1, -4138), ("thick", 1, 4), ("dashed", -4115, 2),
         ("mediumDashed", -4115, -4138), ("dashDot", 4, 2), ("mediumDashDot", 4, -4138), ("dashDotDot", 5, 2),
         ("mediumDashDotDot", 5, -4138), ("dotted", -4118, 2), ("double", -4119, 4), ("slantDashDot", 13, -4138)]
COLORS = [("black", (0, 0, 0)), ("white", (255, 255, 255)), ("red", (255, 0, 0)), ("green", (0, 255, 0)),
          ("blue", (0, 0, 255)), ("yellow", (255, 255, 0)), ("cyan", (0, 255, 255)), ("magenta", (255, 0, 255)),
          ("gray", (128, 128, 128)), ("maroon", (100, 0, 0)), ("forest", (0, 100, 0)), ("navy", (0, 0, 100)),
          ("rust", (200, 100, 50)), ("steel", (50, 100, 200)), ("light", (178, 178, 178)), ("dark", (63, 63, 63))]

#: (upper cell's bottom, lower cell's top) as (line, colour) pairs.
PAIRS = [((a, "black"), (b, "black")) for a, b in itertools.product([name for name, *_ in LINES], repeat=2)]
PAIRS += [(("thin", a), ("thin", b)) for a, b in itertools.product([name for name, _ in COLORS], repeat=2) if a != b]
PAIRS += [(("medium", "white"), ("thin", "black")), (("thin", "black"), ("medium", "white")),
          (("hair", "black"), ("thin", "white")), (("thin", "white"), ("hair", "black")),
          (("double", "white"), ("thick", "black")), (("thick", "black"), ("double", "white"))]
#: Pairs set by one procedure; a longer one grows past what VBA compiles.
BATCH = 40


def module() -> str:
    kinds = sorted({one for pair in PAIRS for one in pair})
    lines = {name: (style, weight) for name, style, weight in LINES}
    colors = dict(COLORS)
    code = ["Private Sub Made(wb As Object)", "Dim s As Object"]
    for line, color in kinds:
        style, weight = lines[line]
        for edge, index in (("B", -4107), ("T", -4160)):
            code += [f'Set s = wb.Styles.Add("{edge}_{line}_{color}")', "s.IncludeBorder = True",
                     f"s.Borders({index}).LineStyle = {style}", f"s.Borders({index}).Weight = {weight}",
                     f"s.Borders({index}).Color = RGB{colors[color]}"]
    code.append("End Sub")
    batches = [list(enumerate(PAIRS, start=1))[start:start + BATCH] for start in range(0, len(PAIRS), BATCH)]
    for number, batch in enumerate(batches):
        code += [f"Private Sub Pairs{number}(ws As Object)"]
        for column, ((upper, upper_color), (lower, lower_color)) in batch:
            code += [f'ws.Cells(1, {column * 2}).Style = "B_{upper}_{upper_color}"',
                     f'ws.Cells(2, {column * 2}).Style = "T_{lower}_{lower_color}"']
        code.append("End Sub")
    code += ["Public Function Probe(target As String) As String", "Dim wb As Object, ws As Object",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "Made wb", *(f"Pairs{number} ws" for number in range(len(batches))),
             "Probe = Edges(ws)", 'wb.SaveAs Filename:=target & "shared_edges.xlsx"', "wb.Close False", "End Function"]
    return "\n".join(code) + "\n" + READER


#: Both cells' view of each pair's edge; the replay reads the saved workbook with it.
READER = f'''Public Function Edges(ws As Object) As String
    Dim out As String, column As Long, top As Object, low As Object
    For column = 2 To {len(PAIRS) * 2} Step 2
        Set top = ws.Cells(1, column)
        Set low = ws.Cells(2, column)
        out = out & low.Borders(8).LineStyle & "," & low.Borders(8).Weight & "," & low.Borders(8).Color & "," _
            & top.Borders(9).LineStyle & "," & top.Borders(9).Weight & "," & top.Borders(9).Color & "|"
    Next
    Edges = out
End Function
'''


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", args=(str(folder) + "\\",), timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(PAIRS)]
    OUT.mkdir(parents=True, exist_ok=True)
    copy_saved(folder / "shared_edges.xlsx", OUT / "shared_edges.xlsx")
    record = {"reader": READER, "pairs": PAIRS, "answers": answers}
    (OUT / "shared_edges.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(len(answers), "pairs")


if __name__ == "__main__":
    main()
