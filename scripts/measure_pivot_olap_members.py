"""The pivot members a table that is not OLAP answers only with an error, or with nothing: PivotCache's ResetTimer and
MakeConnection, PivotField's and PivotItem's DrillTo, PivotTable's AllocateChanges, CommitChanges, DiscardChanges and
RefreshDataSourceValues, and PivotCell's AllocateChange and DiscardChange. What each answers, or the error it
raises, and whether the table's cells are left as they were.

Each case makes a workbook with a list of regions, products and amounts on its first sheet, builds a pivot table on a
second sheet with Region and Product as row fields and the sum of Amount as data, runs the case's statement, and reads
the table's cells.

    python scripts/measure_pivot_olap_members.py

writes tests/fixtures/pivot_olap_members.json, which tests/test_excel_pivot_olap_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_olap_members.json'
CASES = [
    'pt.PivotCache.ResetTimer',
    'pt.PivotCache.MakeConnection',
    'pt.PivotFields("Region").DrillTo "Product"',
    'pt.PivotFields("Region").PivotItems("East").DrillTo "Product"',
    'pt.AllocateChanges',
    'pt.CommitChanges',
    'pt.DiscardChanges',
    'pt.RefreshDataSourceValues',
    'ws.Range("B4").PivotCell.AllocateChange',
    'ws.Range("B4").PivotCell.DiscardChange',
    's = TypeName(pt.PivotCache.ResetTimer)',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, c As Range, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("West", "Ink", 20)',
             '    Set ws = wb.Worksheets.Add(After:=ws)', '    ws.Name = "report"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R3C3").CreatePivotTable(ws.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Read:',
              '    For Each c In pt.TableRange1.Cells: Tried = Tried & "|" & c.Text: Next',
              '    wb.Close False', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Read',
              'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
