"""Chart.Location: a chart moved to a chart sheet of its own, to another worksheet, and back.

Each case makes a workbook of three worksheets, One, Two and Three, with One active and ch a column chart on it, runs its statements,
and records r, a digest of the tabs and charts, and the parts a chart's place touches.

    python scripts/measure_chart_location.py

writes tests/fixtures/chart_location.json, which tests/test_excel_chart_location.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_location.json'
NEW = 'Set t = ch.Location(xlLocationAsNewSheet, "Mine"): '
CASES: list[str] = ['r = 0']
# The new sheet: its name, place and what the call answers.
CASES += [NEW + 'r = TypeName(t) & " " & t.Name', 'Set t = ch.Location(xlLocationAsNewSheet): r = t.Name', 'Set t = ch.Location(1): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, ""): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, "One"): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, "one"): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, "a/b"): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, "' + 'x' * 32 + '"): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, "' + 'x' * 31 + '"): r = t.Name', 'Set t = ch.Location(xlLocationAsNewSheet, 12): r = t.Name',
          'Set t = ch.Location(xlLocationAutomatic): r = TypeName(t) & " " & t.Name', 'Set t = ch.Location(3, "Mine"): r = t.Name', 'Set t = ch.Location(0): r = 1', 'Set t = ch.Location(4): r = 1', 'Set t = ch.Location("x"): r = 1', 'ch.Location xlLocationAsNewSheet, "Mine": r = ws.Parent.Sheets.Count',
          'ws.Parent.Worksheets("Two").Activate: ' + NEW + 'r = ws.Parent.Sheets(1).Name & " " & ws.Parent.Sheets(2).Name', 'ws.Parent.Worksheets("Three").Activate: ' + NEW + 'r = t.Index', 'ws.Parent.Worksheets("Two").Visible = False: ' + NEW + 'r = t.Index', NEW + 'r = t.Index & " " & t.Type & " " & t.Visible', NEW + 'r = TypeName(t.Parent) & " " & t.Parent.Name',
          NEW + 'r = ws.Parent.Charts.Count & " " & ws.Parent.Charts(1).Name & " " & ws.Parent.Sheets("Mine").Name & " " & TypeName(ws.Parent.Sheets("Mine"))', NEW + 'r = ws.Parent.ActiveSheet.Name & " " & ActiveSheet.Name & " " & TypeName(ActiveSheet) & " " & TypeName(ActiveChart)', NEW + 'r = ws.ChartObjects.Count & " " & ws.Shapes.Count', NEW + 'r = ws.Parent.Worksheets.Count', NEW + 'r = t.HasTitle & " " & t.ChartType & " " & t.SeriesCollection.Count & " " & t.SeriesCollection(1).Formula', NEW + 'r = TypeName(t.ChartArea) & " " & t.HasLegend',
          NEW + 'r = ch.Name', NEW + 'r = TypeName(ch)', NEW + 'r = ch.ChartType', NEW + 'r = co.Name', NEW + 'r = TypeName(co)', NEW + 'r = ch.SeriesCollection.Count']
# A second, and names that follow.
SECOND = 'Set c2 = ws.ChartObjects.Add(48, 300, 240, 144).Chart: c2.SetSourceData ws.Range("A1:B3"): '
CASES += [f'{SECOND}Set t = ch.Location(1): Set u = c2.Location(1): r = t.Name & " " & u.Name & " " & u.Index & " " & t.Index', f'Set t = ch.Location(1): r = ws.Parent.Charts.Count & " " & t.Name', f'{SECOND}Set t = ch.Location(1, "Chart2"): Set u = c2.Location(1): r = u.Name', 'ws.Parent.Worksheets.Add.Name = "Chart1": Set t = ch.Location(1): r = t.Name', f'{SECOND}ch.Location 1, "Mine": Set u = c2.Location(1, "Mine"): r = u.Name']
# After it.
AFTER = NEW
CASES += [f'{AFTER}t.Name = "Other": r = t.Name', f'{AFTER}t.Move After:=ws.Parent.Sheets(ws.Parent.Sheets.Count): r = t.Index', f'{AFTER}Application.DisplayAlerts = False: t.Delete: r = ws.Parent.Sheets.Count & " " & ws.Parent.Charts.Count & " " & ActiveSheet.Name', f'{AFTER}t.HasTitle = True: t.ChartTitle.Text = "Over": r = t.ChartTitle.Text', f'{AFTER}t.ChartType = xlLine: r = t.ChartType', f'{AFTER}ws.Activate: r = ActiveSheet.Name', f'{AFTER}ws.Parent.Worksheets("Two").Activate: t.Activate: r = TypeName(ActiveSheet) & " " & ActiveSheet.Name', f'{AFTER}ws.Range("B2").Value = 99: r = t.SeriesCollection(1).Values(1)', f'{AFTER}ws.Rows(1).Insert: r = t.SeriesCollection(1).Formula', f'{AFTER}ws.Name = "Renamed": r = t.SeriesCollection(1).Formula',
          f'{AFTER}r = t.ChartObjects.Count', f'{AFTER}r = TypeName(t.Shapes)', f'{AFTER}r = ws.Parent.Sheets.Count: ws.Copy After:=ws: r = r & " " & ws.Parent.Sheets.Count', f'{AFTER}Set u = ws.Parent.Worksheets.Add: r = u.Index & " " & u.Name', f'{AFTER}Set u = ws.Parent.Sheets.Add(After:=t): r = u.Index']
# To another worksheet, and back again.
TO = 'Set t = ch.Location(xlLocationAsObject, "Two"): '
CASES += [TO + 'r = TypeName(t) & " " & t.Name', TO + 'r = t.Parent.Name & " " & t.Parent.Parent.Name', TO + 'r = ws.ChartObjects.Count & " " & ws.Parent.Worksheets("Two").ChartObjects.Count', TO + 'r = ws.Parent.Worksheets("Two").ChartObjects(1).Name', TO + 'r = t.Parent.Left & " " & t.Parent.Top & " " & t.Parent.Width & " " & t.Parent.Height', TO + 'r = ActiveSheet.Name & " " & TypeName(ActiveChart)', TO + 'r = ch.ChartType', TO + 'r = co.Name', TO + 'r = t.SeriesCollection(1).Formula', 'Set t = ch.Location(xlLocationAsObject, "One"): r = t.Parent.Name & " " & ws.ChartObjects.Count', 'Set t = ch.Location(xlLocationAsObject, "two"): r = t.Parent.Parent.Name', 'Set t = ch.Location(xlLocationAsObject, "Nope"): r = 1', 'Set t = ch.Location(xlLocationAsObject): r = 1', 'ws.Parent.Worksheets("Two").ChartObjects.Add(10, 10, 100, 100).Name = "Taken": ' + TO + 'r = t.Parent.Name & " " & ws.Parent.Worksheets("Two").ChartObjects.Count',
          NEW + 'Set u = t.Location(xlLocationAsObject, "Two"): r = TypeName(u) & " " & u.Parent.Name & " " & ws.Parent.Sheets.Count & " " & ws.Parent.Charts.Count', NEW + 'Set u = t.Location(xlLocationAsObject, "One"): r = u.Parent.Left & " " & u.Parent.Top & " " & u.Parent.Width', NEW + 'Set u = t.Location(xlLocationAsNewSheet, "Next"): r = u.Name & " " & ws.Parent.Sheets.Count & " " & ws.Parent.Charts.Count', NEW + 'Set u = t.Location(xlLocationAutomatic): r = u.Name']
# With what the chart has.
CASES += ['ch.HasTitle = True: ch.ChartTitle.Text = "T": ' + NEW + 'r = t.ChartTitle.Text', 'ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255: ' + NEW + 'r = t.SeriesCollection(1).Format.Fill.ForeColor.RGB', 'ch.ChartArea.Font.Size = 14: ' + NEW + 'r = t.ChartArea.Font.Size', 'co.Name = "Named": ' + NEW + 'r = t.Name', 'Set p = ws.Shapes.AddChart2(201, xlColumnClustered).Chart: p.SetSourceData ws.Range("A1:C3"): Set t = p.Location(xlLocationAsNewSheet, "Mine"): r = t.ChartStyle', 'Set p = ws.Shapes.AddChart2(201, xlColumnClustered).Chart: p.SetSourceData ws.Range("A1:C3"): Set t = p.Location(xlLocationAsObject, "Two"): r = t.ChartStyle', 'ws.Protect: Set t = ch.Location(1): r = 1', 'ws.Parent.Protect Structure:=True: Set t = ch.Location(1): r = 1', 'ws.Parent.Worksheets("Two").Protect: Set t = ch.Location(xlLocationAsObject, "Two"): r = 1']

# Where a new sheet lands. The rivals: just before the sheet the chart was on, first of all, or before the active sheet.
ON_TWO = 'Set w2 = ws.Parent.Worksheets("Two"): Set c3 = w2.ChartObjects.Add(48, 96, 240, 144).Chart: c3.SetSourceData ws.Range("A1:C3"): '
ON_THREE = 'Set w3 = ws.Parent.Worksheets("Three"): Set c4 = w3.ChartObjects.Add(48, 96, 240, 144).Chart: c4.SetSourceData ws.Range("A1:C3"): '
CASES += [f'{ON_TWO}Set t = c3.Location(1, "Mine"): r = t.Index', f'{ON_THREE}Set t = c4.Location(1, "Mine"): r = t.Index', f'{ON_THREE}ws.Parent.Worksheets("Two").Activate: Set t = c4.Location(1, "Mine"): r = t.Index', f'{ON_TWO}{ON_THREE}Set t = c4.Location(1, "A"): Set u = c3.Location(1, "B"): r = t.Index & " " & u.Index', f'{ON_TWO}{ON_THREE}Set t = c3.Location(1, "A"): Set u = c4.Location(1, "B"): r = t.Index & " " & u.Index', f'{ON_TWO}Set t = ch.Location(1, "A"): Set u = c3.Location(1, "B"): r = t.Index & " " & u.Index', f'{ON_TWO}Set t = c3.Location(1, "A"): Set u = ch.Location(1, "B"): r = t.Index & " " & u.Index', 'Set t = ch.Location(1, "A"): r = ws.Previous.Name & " " & t.Next.Name']
# The chart object on the other sheet, a shape name taken there, and the shape's place in the list.
CASES += ['ws.Parent.Worksheets("Two").ChartObjects.Add(10, 10, 100, 100).Name = "Chart 1": Set t = ch.Location(xlLocationAsObject, "Two"): r = t.Parent.Name & " " & ws.Parent.Worksheets("Two").ChartObjects(1).Name & " " & ws.Parent.Worksheets("Two").ChartObjects(2).Name', 'ws.Parent.Worksheets("Two").Shapes.AddShape(1, 10, 10, 10, 10).Name = "Box": Set t = ch.Location(xlLocationAsObject, "Two"): r = ws.Parent.Worksheets("Two").Shapes(1).Name & " " & ws.Parent.Worksheets("Two").Shapes(2).Name', 'co.Name = "Named": Set t = ch.Location(xlLocationAsObject, "Two"): r = t.Parent.Name', 'Set t = ch.Location(xlLocationAsObject, "Two"): Set u = t.Location(xlLocationAsObject, "One"): r = u.Parent.Parent.Name & " " & ws.ChartObjects.Count', 'Set t = ch.Location(xlLocationAsObject, "Two"): t.Parent.Left = 96: r = t.Parent.Left', 'Set t = ch.Location(xlLocationAsObject, "Two"): r = ws.Parent.Worksheets("Two").ChartObjects(1).Chart.ChartType', 'Set t = ch.Location(xlLocationAsObject, "Two"): r = t.Parent.TopLeftCell.Address & " " & t.Parent.BottomRightCell.Address']
# A chart with one named series holds its automatic title once moved. The rivals: any move, a move to another sheet only,
# or a chart sheet only; a title taken away staying away; no name or two series giving none.
CASES += [f'{SECOND}Set t = c2.Location(1): r = t.HasTitle', f'{SECOND}Set t = c2.Location(2, "Two"): r = t.HasTitle', f'{SECOND}Set t = c2.Location(2, "One"): r = t.HasTitle',
          f'{SECOND}c2.HasTitle = False: Set t = c2.Location(1): r = t.HasTitle', f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: Set t = c2.Location(1): r = t.HasTitle',
          f'{SECOND}c2.HasTitle = True: c2.ChartTitle.Delete: Set t = c2.Location(1): r = t.HasTitle', f'{SECOND}c2.HasTitle = False: Set t = c2.Location(2, "Two"): r = t.HasTitle',
          'Set c2 = ws.ChartObjects.Add(48, 300, 240, 144).Chart: c2.SetSourceData ws.Range("B2:B3"): Set t = c2.Location(1): r = t.HasTitle',
          'Set t = ch.Location(1): r = t.HasTitle', 'ch.SeriesCollection(2).Delete: Set t = ch.Location(1): r = t.HasTitle', 'ch.ChartType = xlPie: Set t = ch.Location(1): r = t.HasTitle',
          'ch.ChartType = xlPie: Set t = ch.Location(2, "Two"): r = t.HasTitle']
# What earned it was the chart becoming active, on the way to the move: a move that fails earns it too, and Where refused does not.
CASES += [f'{SECOND}r = c2.HasTitle', f'{SECOND}c2.Parent.Activate: r = c2.HasTitle', f'{SECOND}c2.Parent.Select: r = c2.HasTitle', f'{SECOND}c2.ChartArea.Select: r = c2.HasTitle',
          f'{SECOND}On Error Resume Next: c2.Location 2: r = Err.Number: On Error GoTo 0: r = r & " " & c2.HasTitle', f'{SECOND}On Error Resume Next: c2.Location 2, "Nope": r = Err.Number: On Error GoTo 0: r = r & " " & c2.HasTitle',
          f'{SECOND}On Error Resume Next: c2.Location 1, "One": r = Err.Number: On Error GoTo 0: r = r & " " & c2.HasTitle', f'{SECOND}On Error Resume Next: c2.Location 9: r = Err.Number: On Error GoTo 0: r = r & " " & c2.HasTitle',
          f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: r = c2.HasTitle', f'{SECOND}c2.HasTitle = True: c2.ChartTitle.Text = "T": c2.HasTitle = False: r = c2.HasTitle', f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: c2.HasTitle = True: r = c2.HasTitle',
          f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: c2.Parent.Activate: r = c2.HasTitle', f'{SECOND}c2.HasTitle = False: c2.Parent.Activate: r = c2.HasTitle',
          f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: On Error Resume Next: r = TypeName(c2.ChartTitle): r = r & " " & Err.Number: On Error GoTo 0',
          'ch.HasTitle = True: ch.HasTitle = False: r = ch.HasTitle', 'ch.Parent.Activate: r = ch.HasTitle', 'ch.SeriesCollection(2).Delete: ch.Parent.Activate: r = ch.HasTitle',
          f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: c2.SeriesCollection(1).Name = "New": r = c2.HasTitle', f'{SECOND}c2.HasTitle = True: c2.HasTitle = False: ws.Copy After:=ws: r = ws.Parent.Sheets(2).ChartObjects(2).Chart.HasTitle',
          f'{SECOND}ws.Copy After:=ws: r = c2.HasTitle & " " & ws.Parent.Sheets(2).ChartObjects(2).Chart.HasTitle',
          f'{SECOND}ws.Protect: On Error Resume Next: c2.Parent.Activate: r = Err.Number: On Error GoTo 0: ws.Unprotect: r = r & " " & c2.HasTitle',
          f'{SECOND}ws.Protect: On Error Resume Next: c2.Location 1: r = Err.Number: On Error GoTo 0: ws.Unprotect: r = r & " " & ws.ChartObjects.Count & " " & c2.HasTitle']
# A protected sheet: its drawing objects, or only its cells, and a chart object unlocked.
TRY = 'On Error Resume Next: {}: r = Err.Number: On Error GoTo 0: ws.Unprotect: r = r & " " & c2.HasTitle & " " & TypeName(ActiveChart)'
CASES += [*[f'{SECOND}{guard}: ' + TRY.format(act) for guard in ('ws.Protect', 'ws.Protect DrawingObjects:=False', 'ws.Protect Contents:=False', 'ws.Protect UserInterfaceOnly:=True', 'c2.Parent.Locked = False: ws.Protect')
            for act in ('c2.Parent.Activate', 'c2.Parent.Select', 'c2.ChartArea.Select', 'c2.Location 2, "One"', 'c2.Location 1')]]
# A sheet's copy. The rivals for what earns its title: the data on another sheet, the copied sheet not being the active one, or reading the copy.
ELSEWHERE = 'Set w2 = ws.Parent.Worksheets("Two"): Set c3 = w2.ChartObjects.Add(1, 1, 100, 100).Chart: c3.SetSourceData ws.Range("A1:B3"): '
COPY_READ = 'r = c3.HasTitle & " " & ws.Parent.Worksheets("Two (2)").ChartObjects(1).Chart.HasTitle'
CASES += [f'{ELSEWHERE}w2.Copy After:=w2: {COPY_READ}', f'{ELSEWHERE}w2.Activate: w2.Copy After:=w2: {COPY_READ}', f'{ELSEWHERE}w2.Copy After:=w2: r = ws.Parent.Worksheets("Two (2)").ChartObjects(1).Chart.SeriesCollection(1).Formula',
          f'{ELSEWHERE}r = c3.SeriesCollection(1).Formula: r = r & " " & c3.HasTitle', f'{SECOND}r = c2.SeriesCollection(1).Formula: r = r & " " & c2.HasTitle',
          'Set w2 = ws.Parent.Worksheets("Two"): w2.Range("A1:B1").Value = Array("", "S"): w2.Range("A2:B2").Value = Array("a", 1): Set c3 = w2.ChartObjects.Add(1, 1, 100, 100).Chart: c3.SetSourceData w2.Range("A1:B2"): w2.Copy After:=w2: ' + COPY_READ,
          'Set w2 = ws.Parent.Worksheets("Two"): w2.Range("A1:B1").Value = Array("", "S"): w2.Range("A2:B2").Value = Array("a", 1): Set c3 = w2.ChartObjects.Add(1, 1, 100, 100).Chart: c3.SetSourceData w2.Range("A1:B2"): w2.Activate: w2.Copy After:=w2: ' + COPY_READ]
# Which of a copied sheet's charts: the first, the last, or any with one series when it is the only chart.
ON = 'Set w2 = ws.Parent.Worksheets("Two"): '
ONE_SERIES = 'w2.ChartObjects.Add(1, 1, 100, 100).Chart.SetSourceData ws.Range("A1:B3"): '
TWO_SERIES = 'w2.ChartObjects.Add(1, 1, 100, 100).Chart.SetSourceData ws.Range("A1:C3"): '
THE_COPY = ('w2.Copy After:=w2: Set w3 = ws.Parent.Worksheets("Two (2)"): For Each p In w3.ChartObjects: r = r & p.Chart.HasTitle & " ": Next: '
            'For Each p In w2.ChartObjects: r = r & p.Chart.HasTitle & " ": Next')
CASES += [f'{ON}{mix}{THE_COPY}' for mix in (ONE_SERIES * 2, ONE_SERIES * 3, TWO_SERIES + ONE_SERIES, ONE_SERIES + TWO_SERIES, TWO_SERIES + ONE_SERIES * 2)]
# All a sheet's chart objects selected together, there being one.
CASES += [f'{guard}: On Error Resume Next: ws.ChartObjects.Select: r = Err.Number: On Error GoTo 0: ws.Unprotect: r = r & " " & TypeName(ActiveChart)' for guard in ('ws.Protect', 'ws.Protect DrawingObjects:=False', 'ws.Protect UserInterfaceOnly:=True')]
# The only chart, with another shape beside it.
CASES += [f'{ON}{ONE_SERIES}w2.Shapes.AddShape(1, 1, 200, 50, 50).Name = "Box": {THE_COPY}']

MAKE = '''    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    wb.Worksheets.Add(After:=ws).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets("Two")).Name = "Three"
    ws.Activate
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set co = ws.ChartObjects.Add(48, 96, 240, 144)
    Set ch = co.Chart
    ch.SetSourceData ws.Range("A1:C3")
    ch.ChartType = xlColumnClustered
'''
DIGEST = ('wb.Sheets.Count', 'wb.Worksheets.Count', 'wb.Charts.Count', 'wb.ActiveSheet.Name', 'TypeName(wb.ActiveSheet)', 'TypeName(ActiveChart)', 'ws.ChartObjects.Count')


def probe() -> str:
    lines = ['Private Function Read1(wb As Workbook, ws As Worksheet) As String', '    Dim s As String, v As Variant, o As Object', '    On Error Resume Next']
    for expression in DIGEST:
        lines += ['    v = "E": Err.Clear', f'    v = {expression}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & v & " "']
    lines += ['    Err.Clear', '    For Each o In wb.Sheets', '        s = s & "[" & o.Name & ":" & TypeName(o) & ":" & o.Index & "]"', '    Next', '    Err.Clear', '    Read1 = s', 'End Function', '']
    for first in range(0, len(CASES), 60):
        lines += [f'Private Sub Act{first // 60}(ByVal which As Long, ws As Worksheet, co As Object, ch As Object, r As Variant)', '    Dim t As Object, u As Object, c2 As Object, p As Object, c3 As Object, c4 As Object, w2 As Object, w3 As Object', '    Select Case which']
        for index, statement in enumerate(CASES[first:first + 60], first):
            lines += [f'    Case {index}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, co As Object, ch As Object, r As Variant, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', MAKE.rstrip('\n'), '    On Error GoTo Failed', '    Select Case which \\ 60']
    lines += [f'    Case {block}: Act{block} which, ws, co, ch, r' for block in range((len(CASES) + 59) // 60)]
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v & "|" & Read1(wb, ws)', '    Err.Clear', '    wb.Unprotect: ws.Unprotect: wb.Worksheets("Two").Unprotect', '    Err.Clear', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    If Err.Number Then s = s & "|saveE" & Err.Number', '    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def parts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    keep = re.compile(r'\[Content_Types\]\.xml|xl/workbook\.xml|xl/_rels/workbook\.xml\.rels|xl/chartsheets/.*|xl/drawings/.*|xl/charts/chart\d+\.xml|xl/charts/_rels/.*|xl/worksheets/_rels/.*|docProps/app\.xml')
    with zipfile.ZipFile(path) as package:
        return {name: re.sub(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', '',
                             package.read(name).decode('utf-8'), flags=re.S)
                for name in package.namelist() if keep.fullmatch(name)}


def batch(source: str, folder: str, first: int, last: int) -> list[str]:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        done = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0, line_numbers=False)
        if done.ok and not done.dialogs:
            return str(done.value).split('@@@')
    if first == last:
        return [f'STOPPED {done.outcome}']
    return [answer for index in range(first, last + 1) for answer in batch(source, folder, index, index)]


def main() -> None:
    source = probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for first in range(0, len(CASES), 25):
            answers += batch(source, folder, first, min(first + 24, len(CASES) - 1))
        saved = [parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer, _ in record['cases']:
        print(statement[:110], '=>', answer)


if __name__ == '__main__':
    main()
