"""PivotTable's ConvertToFormulas, ChangeConnection and ChangePivotCache: what each answers, or the error it
raises, the table's cells after it, and how many caches the workbook keeps.

Each case makes a workbook with a list of regions, products and amounts on its first sheet, its fourth row below the
range the table reads, builds a pivot table on a second sheet with Region as a row field and the sum of Amount as data,
runs the case's statement, and reads the cache count and the table's cells.

    python scripts/measure_pivot_table_more.py

writes tests/fixtures/pivot_table_more.json, which tests/test_excel_pivot_table_more.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_table_more.json'
CASES = [
    'pt.ConvertToFormulas False',
    'pt.ConvertToFormulas True',
    'pt.ChangeConnection ws',
    'pt.ChangePivotCache wb.PivotCaches.Create(xlDatabase, "data!R1C1:R4C3")',
    's = TypeName(pt.ChangePivotCache(wb.PivotCaches.Create(xlDatabase, "data!R1C1:R4C3")))',
    'pt.ChangePivotCache "data!R1C1:R4C3"',
    'pt.ChangePivotCache wb.PivotCaches(1)',
    'pt.ChangePivotCache wb.PivotCaches.Create(xlDatabase, "data!R1C1:R4C3"): s = pt.PivotCache.SourceData',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, c As Range, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("West", "Ink", 20)',
             '    ws.Range("A4:C4").Value = Array("North", "Pad", 5)',
             '    Set ws = wb.Worksheets.Add(After:=ws)', '    ws.Name = "report"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R3C3").CreatePivotTable(ws.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Read:', '    Tried = Tried & "|caches=" & wb.PivotCaches.Count',
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
