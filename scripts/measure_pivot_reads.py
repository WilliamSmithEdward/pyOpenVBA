"""Members a macro reads off a pivot table and its fields, filters, cells and cache that had no measurement: ActiveFilters,
AllItemsVisible, ChildItems, ShowingInAxis, AddPageItem, PageRangeCells, RefreshDate, RefreshName, and a PivotFilter's
and PivotCell's plainer members.

Each case makes a workbook with a four-column source and a pivot table (Region on rows, Kind on columns, Year as a page,
Sum of Amount), runs its statements, reads READS, each as its value or "E" and the error number. RefreshName is read as
whether it is Application.UserName, and RefreshDate as whether it is today, so the fixture holds neither.
PivotRowLine and PivotColumnLine are left alone: reading a PivotColumnLine's LineType can stop Excel.

    python scripts/measure_pivot_reads.py

writes tests/fixtures/pivot_reads.json, which tests/test_excel_pivot_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_reads.json'
CASES = [
    '',
    'pt.PivotFields("Region").PivotItems("East").Visible = False',
    'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionEquals, , "East"',
    'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionBeginsWith, , "E"',
    'pt.PivotFields("Region").PivotFilters.Add2 xlTopCount, pt.DataFields(1), 1',
    'pt.PivotFields("Region").PivotFilters.Add2 xlValueIsGreaterThan, pt.DataFields(1), 5',
    'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionEquals, , "East": '
    'pt.PivotFields("Kind").PivotFilters.Add2 xlCaptionEquals, , "A"',
    'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionEquals, , "East": pt.ClearAllFilters',
    'pt.PivotFields("Year").ClearAllFilters: pt.PivotFields("Year").CurrentPage = "2020"',
    'pt.PivotFields("Year").EnableMultiplePageItems = True: pt.PivotFields("Year").PivotItems("2021").Visible = False',
    'pt.PivotFields("Year").ClearAllFilters: pt.PivotFields("Year").AddPageItem "2020"',
    'pt.PivotFields("Year").ClearAllFilters: pt.PivotFields("Year").AddPageItem "2020": '
    'pt.PivotFields("Year").AddPageItem "2021"',
    'pt.PivotFields("Year").ClearAllFilters: pt.PivotFields("Year").AddPageItem "2020", True',
    'pt.PivotFields("Year").AddPageItem "1999"',
    'pt.PivotFields("Region").AddPageItem "East"',
    'pt.PivotFields("Region").AllItemsVisible = True',
    'pt.PivotFields("Region").ShowingInAxis = False',
    'pt.PivotFields("Region").ShowingInAxis = True',
    'pt.PivotFields("Note").Orientation = xlRowField',
    'pt.PivotFields("Region").Orientation = xlHidden',
    'pt.PivotFields("Kind").PivotFilters.Add2 xlCaptionEquals, , "A": '
    'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionEquals, , "East"',
    'pt.RefreshTable',
    'pt.PivotCache.Refresh',
    'pt.PivotFields("Kind").PivotFilters.Add2 xlBottomCount, pt.DataFields(1), 1',
    'pt.PivotFields("Kind").PivotFilters.Add2 xlValueIsLessThan, pt.DataFields(1), 15',
    'pt.PivotFields("Region").PivotFilters.Add2 xlValueIsGreaterThan, pt.DataFields(1), 15',
]
READS = [
    'pt.ActiveFilters.Count', 'TypeName(pt.ActiveFilters)', 'pt.ActiveFilters(1).Name',
    'pt.ActiveFilters(1).FilterType', 'pt.ActiveFilters(1).Description', 'TypeName(pt.ActiveFilters(1).Application)',
    'pt.ActiveFilters(1).Active', 'pt.ActiveFilters(1).PivotField.Name', 'pt.ActiveFilters(1).IsMemberPropertyFilter',
    'pt.ActiveFilters(1).Value1',
    'pt.PivotFields("Region").AllItemsVisible', 'pt.PivotFields("Kind").AllItemsVisible',
    'pt.PivotFields("Year").AllItemsVisible', 'pt.PivotFields("Amount").AllItemsVisible',
    'pt.PivotFields("Note").AllItemsVisible',
    'pt.PivotFields("Region").ShowingInAxis', 'pt.PivotFields("Kind").ShowingInAxis',
    'pt.PivotFields("Year").ShowingInAxis', 'pt.PivotFields("Amount").ShowingInAxis',
    'pt.PivotFields("Note").ShowingInAxis', 'pt.DataFields(1).ShowingInAxis',
    'TypeName(pt.PivotFields("Region").ChildItems)', 'pt.PivotFields("Region").ChildItems.Count',
    'pt.PivotFields("Year").PageRangeCells', 'pt.PageRangeCells.Address', 'pt.PageRange.Address',
    'pt.RefreshName = Application.UserName', 'Int(pt.RefreshDate) = Date', 'pt.PivotCache.RefreshName = Application.UserName',
    'Int(pt.PivotCache.RefreshDate) = Date', 'pt.PivotCache.MemoryUsed > 0', 'pt.PivotFields("Region").MemoryUsed > 0',
    'TypeName(pt.PivotRowAxis)', 'TypeName(pt.PivotColumnAxis)', 'TypeName(pt.PivotFormulas)', 'pt.PivotFormulas.Count',
    'TypeName(ws.Range("B7").PivotCell.Parent)', 'ws.Range("B7").PivotCell.CellChanged',
    'ws.Range("B7").PivotCell.DataSourceValue', 'ws.Range("B7").PivotCell.ServerActions.Count',
    'ws.Range("A7").PivotCell.CellChanged', 'ws.Range("A7").PivotCell.DataSourceValue',
    'pt.TableRange1.Address', 'pt.TableRange2.Address', 'ws.Range("A1").Value', 'ws.Range("B1").Value',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, data As Worksheet, ws As Worksheet, pt As Object, s As String, v As String',
             '    Dim i As Long, row As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set data = wb.Worksheets(1)',
             '    data.Range("A1:E1").Value = Array("Region", "Kind", "Year", "Amount", "Note")',
             '    row = Array(Array("East", "A", 2020, 1, "n"), Array("East", "B", 2021, 2, "n"), '
             'Array("West", "A", 2020, 4, "m"), Array("West", "B", 2021, 8, "m"), Array("North", "A", 2021, 16, "m"))',
             '    For i = 0 To 4: data.Range("A" & i + 2).Resize(1, 5).Value = row(i): Next',
             '    Set ws = wb.Worksheets.Add',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, data.Range("A1:E6")).CreatePivotTable(ws.Range("A5"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.PivotFields("Kind").Orientation = xlColumnField',
             '    pt.PivotFields("Year").Orientation = xlPageField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        if statements:
            lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & Replace(v, ",", ";")']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer] for statements, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, answer in record['cases']:
        print(statements, '->', answer[:1600])


if __name__ == '__main__':
    main()
