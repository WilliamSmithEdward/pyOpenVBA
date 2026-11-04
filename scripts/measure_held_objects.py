"""What an object answers when a macro still holds it after it has gone, across the families the sheet's objects fall into: ListObject and
its rows and columns, Comment, Hyperlink, Worksheet and what hangs off it, Name, PivotTable, ChartObject and its chart's parts, FormatCondition,
and a shape after Cut and Copy.

    python scripts/measure_held_objects.py

writes tests/fixtures/held_objects.json, which tests/test_excel_held_objects.py replays. Each case is [family, held, gone, tries, answers]: the
answers are made, gone, TypeName(h)/error, (h Is Nothing)/error, then one error number for each try, or the error of the Set if h was never set.
A run carries on from the cases an earlier one left beside the fixture. Cut and Copy take live.lock, since other sessions' probes share the clipboard.
"""
from __future__ import annotations

import json
import os
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/held_objects.json'
COMMON = ['h.Parent.Name', 'h.Application.Name', 'h.Creator', 'h.NoSuchMember']
DATA = 'ws.Range("A1").Value = "a": ws.Range("B1").Value = "b": ws.Range("A2").Value = 1: ws.Range("B2").Value = 2: ws.Range("A3").Value = 3: ws.Range("B3").Value = 4: '
TABLE = DATA + 'Set t = ws.ListObjects.Add(1, ws.Range("A1:B3"), , 1): '
NOTE = 'Set t = ws.Range("D5").AddComment("note"): '
LINK = 'Set t = ws.Hyperlinks.Add(ws.Range("D5"), "http://example.com/"): '
SHEET = 'Set t = ws.Parent.Worksheets("Sheet2"): t.Range("A1").Value = 1: '
GLOBAL = 'Set t = ws.Parent.Names.Add("Held", "=" & ws.Name & "!$A$1"): '
LOCAL = 'Set t = ws.Names.Add("Held", "=" & ws.Name & "!$A$1"): '
PIVOT = DATA + 'Set t = ws.Parent.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("F5")): t.PivotFields("a").Orientation = xlRowField: t.AddDataField t.PivotFields("b"): '
PIVOT_ELSEWHERE = DATA + 'Set t = ws.Parent.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Parent.Worksheets("Sheet2").Range("A3")): t.PivotFields("a").Orientation = xlRowField: t.AddDataField t.PivotFields("b"): '
CHART = DATA + 'Set t = ws.ChartObjects.Add(200, 10, 200, 100): t.Chart.SetSourceData ws.Range("A1:B3"): '
TITLED = CHART + 't.Chart.HasTitle = True: t.Chart.HasLegend = True: '
CONDITION = 'ws.Range("D5:D9").Value = 1: Set t = ws.Range("D5:D9").FormatConditions.Add(xlCellValue, xlGreater, "=0"): t.Interior.Color = 255: '
SCALE = 'ws.Range("D5:D9").Value = 1: Set t = ws.Range("D5:D9").FormatConditions.AddColorScale(2): '
BAR = 'ws.Range("D5:D9").Value = 1: Set t = ws.Range("D5:D9").FormatConditions.AddDatabar: '
ICONS = 'ws.Range("D5:D9").Value = 1: Set t = ws.Range("D5:D9").FormatConditions.AddIconSetCondition: '
RECT = 'Set t = ws.Shapes.AddShape(1, 10, 10, 100, 40): '
#: family: what is made, then (what is held, what removes it, what to try on it).
FAMILIES: dict[str, list[tuple[str, list[str], list[str], list[str]]]] = {
    'ListObject': [
        (TABLE, ['t'], ['t.Delete', 't.Unlist', 'ws.Range("A1:B3").Delete', 'ws.Rows("1:3").Delete', 'ws.Columns("A:B").Delete', 'ws.Range("A1:B3").Clear', 'ws.Cells.Clear', 'ws.Delete'],
         ['h.Name', 'h.DisplayName', 'h.Range.Address', 'h.ShowTotals', 'h.TableStyle', 'h.ListColumns.Count', 'h.ListRows.Count', 'h.Name = "Z"', 'h.ShowTotals = True', 'h.Resize ws.Range("A1:B4")', 'h.Unlist']),
        (TABLE, ['t'], ['t.Delete', 'ws.Delete'], ['h.Delete']),
        (TABLE, ['t.ListRows(1)', 't.ListRows', 't.ListColumns(1)', 't.ListColumns', 't.Range', 't.DataBodyRange', 't.HeaderRowRange', 't.AutoFilter', 't.Sort'], ['t.Delete', 't.Unlist', 'ws.Delete'],
         ['h.Count', 'h.Index', 'h.Name', 'h.Range.Address', 'h.Address', 'h.Delete', 'h.Add']),
        (TABLE, ['t.ListRows(1)', 't.ListRows(2)', 't.ListColumns(1)', 't.ListColumns(2)'], ['t.ListRows(1).Delete', 't.ListColumns(1).Delete', 'ws.Rows(2).Delete', 'ws.Columns(1).Delete', 't.ListRows.Add 1', 't.ListColumns.Add 1'],
         ['h.Index', 'h.Name', 'h.Range.Address', 'h.Range.Value', 'h.Delete', 'h.Name = "Z"']),
    ],
    'Comment': [
        (NOTE, ['t', 't.Shape', 't.Shape.TextFrame', 'ws.Comments'], ['t.Delete', 'ws.Range("D5").ClearComments', 'ws.Range("D5").ClearNotes', 'ws.Range("D5").Clear', 'ws.Rows(5).Delete', 'ws.Columns(4).Delete', 'ws.Range("D5").Delete', 'ws.Cells.Clear', 'ws.Shapes(1).Delete', 'ws.Delete'],
         ['h.Text', 'h.Author', 'h.Visible', 'h.Name', 'h.Count', 'h.Parent.Address', 'h.Visible = True', 'h.Text "x"', 'h.Shape.Name', 'h.Next', 'h.Previous', 'h.Delete']),
    ],
    'Hyperlink': [
        (LINK, ['t', 't.Range', 'ws.Hyperlinks', 'ws.Range("D5").Hyperlinks', 'ws.Hyperlinks(1)'], ['t.Delete', 'ws.Hyperlinks.Delete', 'ws.Range("D5").Hyperlinks.Delete', 'ws.Range("D5").ClearHyperlinks', 'ws.Range("D5").Clear', 'ws.Range("D5").Delete', 'ws.Rows(5).Delete', 'ws.Cells.Clear', 'ws.Delete'],
         ['h.Address', 'h.SubAddress', 'h.Name', 'h.Type', 'h.TextToDisplay', 'h.Range.Address', 'h.Count', 'h.Address = "http://example.org/"', 'h.ScreenTip = "x"', 'h.Delete']),
    ],
    'Worksheet': [
        (SHEET, ['t', 't.Range("A1")', 't.Cells', 't.Rows(1)', 't.Shapes', 't.ListObjects', 't.Names', 't.Comments', 't.Hyperlinks', 't.ChartObjects', 't.PivotTables', 't.PageSetup', 't.Tab', 't.Outline', 't.Protection', 't.AutoFilter', 't.UsedRange'],
         ['t.Delete', 'ws.Parent.Worksheets("Sheet2").Delete', 'ws.Parent.Sheets(2).Delete'],
         ['h.Name', 'h.Index', 'h.Count', 'h.Value', 'h.Address', 'h.Name = "Z"', 'h.Value = 1', 'h.Activate', 'h.Select', 'h.Delete', 'h.Visible', 'h.Visible = xlSheetHidden', 'h.Color']),
        ('t2 = 0: Set t = ws.Parent.Worksheets("Sheet2"): ws.Parent.Worksheets("Sheet2").Shapes.AddShape 1, 1, 1, 9, 9: ', ['t.Shapes(1)', 't.Shapes(1).Fill'], ['t.Delete'], ['h.Name', 'h.Visible', 'h.Name = "Z"']),
        (SHEET + 't.ListObjects.Add 1, t.Range("A1:B3"), , 1: ', ['t.ListObjects(1)', 't.ListObjects(1).ListRows(1)'], ['t.Delete'], ['h.Name', 'h.Index', 'h.Name = "Z"']),
        (SHEET + 't.Range("A1").AddComment "x": ', ['t.Range("A1").Comment'], ['t.Delete'], ['h.Text', 'h.Visible = True']),
        (SHEET + 't.Hyperlinks.Add t.Range("A1"), "http://example.com/": ', ['t.Hyperlinks(1)'], ['t.Delete'], ['h.Address', 'h.Address = "http://example.org/"']),
        (SHEET + 't.ChartObjects.Add 1, 1, 9, 9: ', ['t.ChartObjects(1)'], ['t.Delete'], ['h.Name', 'h.Left = 5']),
    ],
    'Name': [
        (GLOBAL, ['t'], ['t.Delete', 'ws.Parent.Names("Held").Delete', 'ws.Parent.Names.Item(1).Delete', 'ws.Delete'], ['h.Name', 'h.RefersTo', 'h.Value', 'h.Visible', 'h.Index', 'h.RefersToRange.Address', 'h.Comment', 'h.Name = "Z"', 'h.RefersTo = "=1"', 'h.Visible = False', 'h.Delete']),
        (LOCAL, ['t'], ['t.Delete', 'ws.Delete', 'ws.Names("Held").Delete'], ['h.Name', 'h.RefersTo', 'h.Value', 'h.Visible', 'h.Index', 'h.Name = "Z"', 'h.RefersTo = "=1"', 'h.Delete']),
        (GLOBAL, ['ws.Parent.Names'], ['t.Delete'], ['h.Count', 'h.Item(1).Name']),
        (LOCAL, ['ws.Names'], ['t.Delete', 'ws.Delete'], ['h.Count', 'h.Item(1).Name']),
    ],
    'PivotTable': [
        (PIVOT, ['t', 't.PivotFields("a")', 't.PivotFields("b")', 't.PivotFields', 't.DataFields(1)', 't.PivotCache', 't.TableRange1', 't.TableRange2', 't.DataBodyRange', 't.PivotFields("a").PivotItems(1)', 't.PivotFields("a").PivotItems'],
         ['t.TableRange2.Clear', 't.TableRange2.Delete', 'ws.Columns("F:H").Delete', 'ws.Rows("5:9").Delete', 'ws.Cells.Clear', 'ws.Delete', 't.PivotFields("a").Orientation = xlHidden', 't.ClearTable'],
         ['h.Name', 'h.Value', 'h.Caption', 'h.Orientation', 'h.Count', 'h.Address', 'h.SourceName', 'h.TableRange1.Address', 'h.Name = "Z"', 'h.Orientation = xlColumnField', 'h.RefreshTable', 'h.Refresh']),
        (PIVOT_ELSEWHERE, ['t', 't.PivotCache', 't.PivotFields("a")'], ['ws.Parent.Worksheets("Sheet2").Delete'], ['h.Name', 'h.TableRange1.Address', 'h.Name = "Z"', 'h.RefreshTable', 'h.Refresh']),
    ],
    # One member to a case here: some members of a chart's parts hang or kill Excel once the chart has gone.
    'ChartObject': [
        (CHART, ['t'], ['t.Delete', 'ws.Shapes(1).Delete', 'ws.ChartObjects.Delete', 'ws.Delete'],
         ['h.Name', 'h.Left', 'h.Index', 'h.Visible', 'h.Placement', 'h.ZOrder', 'h.TopLeftCell.Address', 'TypeName(h.Chart)', 'h.Name = "Z"', 'h.Left = 5', 'h.Select', 'h.Activate', 'h.BringToFront', 'h.Delete', *COMMON]),
        (CHART, ['t.Chart'], ['t.Delete', 'ws.Delete'], ['h.Name', 'h.ChartType', 'h.HasTitle', 'h.ChartType = xlLine', 'h.HasTitle = True', 'h.Activate', *COMMON]),
        (CHART, ['t.ShapeRange'], ['t.Delete', 'ws.Delete'], ['h.Name', 'h.Count', 'h.Left = 5', *COMMON]),
        (CHART, ['t.Chart.ChartArea', 't.Chart.PlotArea', 't.Chart.Axes(1)', 't.Chart.ChartArea.Format.Fill'], ['t.Delete', 'ws.Delete'], ['h.Left', 'h.Name', 'h.Visible', 'h.Left = 5', 'h.Parent.Name']),
        (CHART, ['t.Chart.SeriesCollection', 't.Chart.SeriesCollection(1)'], ['t.Delete', 'ws.Delete', 't.Chart.SeriesCollection(1).Delete', 'ws.Range("A1:B3").Delete'], ['h.Count', 'h.Name', 'h.Formula', 'h.Name = "Z"', 'h.Parent.Name']),
        (CHART, ['t.Chart.SeriesCollection(1).Points(1)', 't.Chart.SeriesCollection(1).Format'], ['t.Chart.SeriesCollection(1).Delete'], ['h.Name', 'h.Parent.Name']),
        (TITLED, ['t.Chart.ChartTitle'], ['t.Chart.HasTitle = False', 't.Chart.ChartTitle.Delete', 't.Delete'], ['h.Text', 'h.Left', 'h.Text = "x"', 'h.Delete', 'h.Parent.Name']),
        (TITLED, ['t.Chart.Legend'], ['t.Chart.HasLegend = False', 't.Chart.Legend.Delete', 't.Delete'], ['h.Position', 'h.Left', 'h.Left = 5', 'h.Delete', 'h.Parent.Name']),
    ],
    # One member to a case here too: some members of a format condition that has gone kill Excel.
    'FormatCondition': [
        (CONDITION, ['t'], ['t.Delete', 'ws.Range("D5:D9").FormatConditions.Delete', 'ws.Range("D5:D9").Clear', 'ws.Rows("5:9").Delete', 'ws.Delete'],
         ['h.Type', 'h.Operator', 'h.Formula1', 'h.Priority', 'h.StopIfTrue', 'h.AppliesTo.Address', 'h.Priority = 1', 'h.StopIfTrue = False', 'h.Modify xlCellValue, xlLess, "=5"', 'h.SetFirstPriority', 'h.Delete', *COMMON]),
        (CONDITION, ['t.Interior', 't.Font', 't.Borders'], ['t.Delete', 'ws.Range("D5:D9").Clear', 'ws.Delete'], ['h.Color', 'h.Color = 255', 'h.Parent.Name']),
        (CONDITION, ['ws.Range("D5:D9").FormatConditions', 't.AppliesTo'], ['t.Delete', 'ws.Range("D5:D9").Clear', 'ws.Rows("5:9").Delete', 'ws.Delete'], ['h.Count', 'h.Address', 'h.Parent.Name']),
        (CONDITION, ['t'], ['ws.Range("D5").Delete', 'ws.Range("D5").FormatConditions.Delete', 'ws.Rows(5).Delete'], ['h.AppliesTo.Address', 'h.Formula1', 'h.Priority']),
        *[(one, ['t'], ['t.Delete', 'ws.Range("D5:D9").Clear'], ['h.Type', 'h.Priority', 'h.Priority = 1', 'h.Delete']) for one in (SCALE, BAR, ICONS)],
    ],
}
CLIP = [(RECT, ['t', 't.Fill', 't.TextFrame2', 'ws.Shapes.Range(Array(t.Name))'], ['t.Cut', 't.Copy', 'ws.Shapes.Range(Array(t.Name)).Cut', 'ws.Shapes.Range(Array(t.Name)).Copy', 't.Cut: ws.Paste', 't.Copy: ws.Paste', 't.Cut: Application.CutCopyMode = False', 't.Select: Selection.Cut', 't.Select: Selection.Copy'], ['h.Name', 'h.Visible', 'h.Name = "Z"', 's2 = ws.Shapes.Count'])]


def _cases(family: str, rows: list[tuple[str, list[str], list[str], list[str]]]) -> list[list[object]]:
    if family in ('ChartObject', 'FormatCondition'):
        return [[family, made, held, gone, [one]] for made, helds, gones, tries in rows for held in helds for gone in gones for one in tries]
    return [[family, made, held, gone, [*tries, *COMMON]] for made, helds, gones, tries in rows for held in helds for gone in gones]


CASES = [one for family, rows in FAMILIES.items() for one in _cases(family, rows)]
CLIP_CASES = _cases('Clipboard', CLIP)


def _probe(cases: list[list[object]]) -> str:
    lines: list[str] = []
    size = 8
    chunks = [range(at, min(at + size, len(cases))) for at in range(0, len(cases), size)]
    for number, chunk in enumerate(chunks):
        lines += [f'Private Function Act{number}(ByVal which As Long, ByVal ws As Worksheet) As String', '    Dim t As Object, h As Object, s2 As Variant, t2 As Variant, v As Variant, s As String', '    On Error Resume Next', '    Select Case which']
        for index in chunk:
            _, made, held, gone, tries = cases[index]
            assert isinstance(tries, list)
            lines += [f'    Case {index}', f'        {made}'.rstrip(': '), f'        Set h = {held}', '        s = "made " & Err.Number: Err.Clear', f'        {gone}', '        s = s & "~gone " & Err.Number: Err.Clear',
                      '        v = "": v = TypeName(h): s = s & "~" & v & "/" & Err.Number: Err.Clear', '        v = "": v = (h Is Nothing): s = s & "~" & v & "/" & Err.Number: Err.Clear']
            for one in tries:
                assigning = ' = ' in one or one == 'h.Add' or one.startswith(('h.Delete', 'h.Unlist', 'h.Resize', 'h.Text "', 'h.Activate', 'h.Select', 'h.Modify', 'h.SetFirst', 'h.Refresh', 'h.BringTo'))
                lines += [f'        {one}: s = s & "~" & Err.Number: Err.Clear' if assigning else f'        v = Empty: v = {one}: s = s & "~" & Err.Number: Err.Clear']
        lines += ['    End Select', f'    Act{number} = s', 'End Function', '']
    lines += ['Private Function Tried(ByVal which As Long) As String', '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Sheet1"',
              '    wb.Worksheets.Add(After:=ws).Name = "Sheet2"', '    ws.Activate', f'    Select Case which \\ {size}']
    lines += [f'    Case {number}: Tried = Act{number}(which, ws)' for number in range(len(chunks))]
    lines += ['    End Select', '    Application.CutCopyMode = False', '    wb.Close False', 'End Function', '', 'Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              '    For i = first To last', '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _key(case: list[object]) -> str:
    return json.dumps(case[:5])


def _run(excel: ExcelSession, cases: list[list[object]], done: dict[str, str], partial: Path, started: float) -> bool:
    """One case at a time, since a member of a held object can hang or crash Excel: such a case is recorded so, and False says start a new
    session. Delete on a table that was unlisted or cleared away kills Excel outright, which is why that try is kept to the removals it survives."""
    source = _probe(cases)
    for first, case in enumerate(cases):
        if _key(case) in done:
            continue
        if time.monotonic() - started > 400:
            return True
        result = excel.run_vba(source, 'Probe', (first, first), timeout=30.0, line_numbers=False)
        assert not any('Compile error' in one.message for one in result.dialogs), result
        answer = str(result.value) if result.ok and not result.dialogs else result.outcome.upper()
        done[_key(case)] = answer
        partial.write_text(json.dumps(done), encoding='utf-8')
        if answer == result.outcome.upper():
            print(answer, case[2], '|', case[3], '|', result.message[:60] if result.message else '', [one.message[:60] for one in result.dialogs])
            return False
    return True


def main() -> None:
    partial = OUT.with_suffix('.part')
    done: dict[str, str] = json.loads(partial.read_text(encoding='utf-8')) if partial.exists() else {}
    started = time.monotonic()
    every = [*CASES, *CLIP_CASES]
    while any(_key(case) not in done for case in CASES) and time.monotonic() - started < 400:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            _run(excel, CASES, done, partial, started)
    if not any(_key(case) not in done for case in CASES):
        common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=OUT.parent).stdout.strip()
        lock = (OUT.parent / common).resolve() / 'excel-coordination' / 'live.lock'
        while any(_key(case) not in done for case in CLIP_CASES) and time.monotonic() - started < 520:
            try:
                os.mkdir(lock)
            except FileExistsError:
                time.sleep(10)
                continue
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                    excel.new_document()
                    _run(excel, CLIP_CASES, done, partial, started - 300)
            finally:
                os.rmdir(lock)
    print(sum(_key(case) in done for case in every), 'of', len(every))
    if all(_key(case) in done for case in every):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': [[*case, done[_key(case)]] for case in every]}, indent=1) + '\n', encoding='utf-8')
        partial.unlink()


if __name__ == '__main__':
    main()
