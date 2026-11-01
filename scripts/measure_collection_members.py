"""Application and Parent of more collections, WorkbookQuery's default member, DisplayFormat's MergeCells and Style, a
colour scale FormatColor's ColorIndex, and PivotTables.Add.

    python scripts/measure_collection_members.py

writes tests/fixtures/collection_members.json, which tests/test_excel_collection_members.py replays.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/collection_members.json'
QUERIES = 'power_query/loaded_to_sheet.xlsx'
#: Reads on a sheet with a pivot table, conditional formats on D1:D5, a merged F1:G1 and a chart.
READS = [
    'TypeName(ws.Range("D1:D5").FormatConditions.Application) & " " & TypeName(ws.Range("D1:D5").FormatConditions.Parent)',
    'TypeName(wb.IconSets.Application) & " " & TypeName(wb.IconSets.Parent)',
    'TypeName(wb.PivotCaches.Application) & " " & TypeName(wb.PivotCaches.Parent)',
    'TypeName(ws.PivotTables.Application) & " " & TypeName(ws.PivotTables.Parent)',
    'TypeName(pt.PivotFields("k").PivotFilters.Application) & " " & TypeName(pt.PivotFields("k").PivotFilters.Parent)',
    'TypeName(ws.Range("J5").PivotCell.RowItems.Application) & " " & TypeName(ws.Range("J5").PivotCell.RowItems.Parent)',
    'TypeName(ch.Axes.Application) & " " & TypeName(ch.Axes.Parent)',
    'TypeName(ch.SeriesCollection.Application) & " " & TypeName(ch.SeriesCollection.Parent)',
    'ws.Range("F1").DisplayFormat.MergeCells & " " & ws.Range("D1").DisplayFormat.MergeCells & " " & '
    'IsNull(ws.Range("E1:F1").DisplayFormat.MergeCells)',
    'TypeName(ws.Range("D1").DisplayFormat.Style) & " " & ws.Range("D1").DisplayFormat.Style.Name & " " & '
    'ws.Range("F1").DisplayFormat.Style.Name',
    'cs.ColorScaleCriteria(1).FormatColor.ColorIndex & " " & cs.ColorScaleCriteria(2).FormatColor.ColorIndex & " " & '
    'cs.ColorScaleCriteria(3).FormatColor.ColorIndex',
    'TypeName(ws.PivotTables.Add(pt.PivotCache, ws.Range("M3"), "PT2")) & " " & ws.PivotTables.Count & " " & '
    'ws.PivotTables("PT2").TableRange1.Address',
    'TypeName(ws.PivotTables.Add(pt.PivotCache, ws.Range("M3"))) & " " & ws.PivotTables(2).Name',
    'TypeName(ws.PivotTables.Add(PivotCache:=pt.PivotCache, TableDestination:="R3C13", TableName:="PT3"))',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, cs As Object, ch As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:B5").Value = Application.Transpose(Application.Transpose(Array(Array("k", "v"), '
             'Array("a", 1), Array("b", 2), Array("a", 3), Array("c", 4))))',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B5")).CreatePivotTable(ws.Range("I3"), "PT")',
             '    pt.PivotFields("k").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum',
             '    ws.Range("D1:D5").Value = Application.Transpose(Array(1, 2, 3, 4, 5))',
             '    Set cs = ws.Range("D1:D5").FormatConditions.AddColorScale(3)', '    ws.Range("F1:G1").Merge',
             '    Set ch = ws.ChartObjects.Add(300, 10, 200, 120).Chart', '    ch.SetSourceData ws.Range("A1:B5")',
             '    On Error GoTo Failed', '    Select Case which']
    for index, read in enumerate(READS):
        lines += [f'    Case {index}', f'        Tried = {read}']
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Queries(ByVal path As String) As String', '    Dim wb As Workbook, q As Object',
              '    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)', '    Set q = wb.Queries(1)',
              '    Queries = TypeName(wb.Queries.Application) & " " & TypeName(wb.Queries.Parent) & " " & '
              'TypeName(q.Application) & " " & TypeName(q.Parent) & " " & q & " " & q.Name',
              '    wb.Close False', 'End Function', '',
              'Public Function Probe(ByVal path As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(READS) - 1}',
              '        s = s & "@@@" & Tried(i)', '    Next', '    s = s & "@@@" & Queries(path)',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        # Excel reads a copy, so that nothing it does touches the fixture.
        copy = Path(folder) / 'queries.xlsx'
        shutil.copy(ROOT / 'tests/fixtures' / QUERIES, copy)
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (str(copy),), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'queries_file': QUERIES, 'source': source,
              'reads': [[read, answer] for read, answer in zip([*READS, 'Queries'], answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['reads']:
        print(read[:70], '->', answer)


if __name__ == '__main__':
    main()
