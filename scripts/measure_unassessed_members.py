"""Members the model registers that no measurement had covered: FormatCondition.SetLastPriority, a table's Slicers,
PivotItemList.Parent, Queries.Add and a WorkbookQuery's members, and Workbooks.Count.

Each read is its value (TypeName for an object) or "E" and the error number. The workbooks are new and closed
unsaved.

    python scripts/measure_unassessed_members.py

writes tests/fixtures/unassessed_members.json, which tests/test_excel_unassessed_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/unassessed_members.json'
SETUP = [
    'n0 = Workbooks.Count',
    'Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
    'ws.Range("A1:B4").Value = Application.Transpose(Application.Transpose(Array(Array("k", "v"), Array("a", 1), '
    'Array("b", 2), Array("a", 3))))',
    'Set fc1 = ws.Range("B2:B4").FormatConditions.Add(xlCellValue, xlGreater, "1")',
    'Set fc2 = ws.Range("B2:B4").FormatConditions.Add(xlCellValue, xlLess, "3")',
    'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes)',
    'Set pc = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B4"))',
    'Set pt = pc.CreatePivotTable(ws.Range("E1"), "P1")', 'pt.PivotFields("k").Orientation = xlRowField',
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum',
]
READS = [
    'fc1.Priority & "/" & fc2.Priority', 'do: fc1.SetLastPriority', 'fc1.Priority & "/" & fc2.Priority',
    'TypeName(lo.Slicers)', 'lo.Slicers.Count', 'TypeName(lo.Slicers.Parent)', 'TypeName(lo.Slicers.Application)',
    'lo.Slicers.Creator', 'TypeName(lo.Slicers(1))',
    'TypeName(ws.Range("F3").PivotCell.RowItems)', 'TypeName(ws.Range("F3").PivotCell.RowItems.Parent)',
    'ws.Range("F3").PivotCell.RowItems.Count',
    'Workbooks.Count - n0', 'wb.Queries.Count', 'TypeName(wb.Queries.Add("Q1", "let Source = 1 in Source", "d"))',
    'wb.Queries.Count', 'wb.Queries(1).Name', 'wb.Queries("Q1").Formula', 'wb.Queries(1).Description',
    'TypeName(wb.Queries(1).Parent)', 'wb.Queries.Item(1).Name', 'wb.Queries(2).Name',
    'TypeName(wb.Queries.Add("Q1", "let Source = 2 in Source"))', 'do: wb.Queries(1).Delete', 'wb.Queries.Count',
]


def _probe() -> str:
    lines = ['Public Function Probe() As String',
             '    Dim wb As Workbook, ws As Worksheet, fc1 As Object, fc2 As Object, lo As Object, pc As Object',
             '    Dim pt As Object, n0 As Long, s As String, v As String']
    lines += [f'    {line}' for line in SETUP]
    lines += ['    On Error Resume Next']
    for read in READS:
        # A "do:" entry is a method run as a statement, answering "done" or its error.
        run = f'    {read[4:]}: v = "done"' if read.startswith('do: ') else f'    v = {read}'
        lines += ['    Err.Clear', '    v = ""', run, '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "@@@" & v']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[read, answer] for read, answer in zip(READS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['cases']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
