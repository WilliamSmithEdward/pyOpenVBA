"""Parent and Application of Areas, Filters, AutoFilter, WorksheetFunction and the table collections, and
ListColumn.Total with and without a totals row.

    python scripts/measure_object_parents.py

writes tests/fixtures/object_parents.json, which tests/test_excel_object_parents.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/object_parents.json'
#: Each read, on a sheet with a filtered range A1:B4, a table lo with a totals row at E1:F4, and one without, lo2.
READS = [
    'TypeName(ws.Range("A1:B2,D4").Areas.Parent)',
    'ws.Range("A1:B2,D4").Areas.Parent.Address',
    'TypeName(ws.Range("A1:B2,D4").Areas.Application)',
    'TypeName(ws.AutoFilter.Filters.Parent)',
    'TypeName(ws.AutoFilter.Filters.Application)',
    'TypeName(ws.AutoFilter.Application)',
    'TypeName(Application.WorksheetFunction.Parent)',
    'TypeName(Application.WorksheetFunction.Application)',
    'TypeName(ws.ListObjects.Application)',
    'TypeName(lo.ListColumns.Application)',
    'TypeName(lo.ListRows.Application)',
    'TypeName(lo.ListColumns(1).Application)',
    'TypeName(lo.ListColumns(1).Total)',
    'lo.ListColumns(2).Total.Address',
    'lo.ListColumns(2).Total.Formula',
    'TypeName(lo2.ListColumns(1).Total)',
]


def _probe() -> str:
    lines = ['Private Function T(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, lo As Object, lo2 As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:B4").Value = Application.Transpose(Application.Transpose(Array(Array("h1", "h2"), '
             'Array(3, 4), Array(1, 2), Array(5, 6))))',
             '    ws.Range("A1:B4").AutoFilter 1, ">1"',
             '    ws.Range("E1:F3").Value = Application.Transpose(Application.Transpose(Array(Array("x", "y"), '
             'Array(1, 2), Array(3, 4))))',
             '    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("E1:F3"), , xlYes)', '    lo.ShowTotals = True',
             '    ws.Range("H1:I3").Value = ws.Range("E1:F3").Value',
             '    Set lo2 = ws.ListObjects.Add(xlSrcRange, ws.Range("H1:I3"), , xlYes)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, read in enumerate(READS):
        lines += [f'    Case {index}', f'        T = {read}']
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    T = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(READS) - 1}',
              '        s = s & "@@@" & T(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'reads': [[read, answer] for read, answer in zip(READS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['reads']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
