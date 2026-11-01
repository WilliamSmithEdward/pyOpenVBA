"""PivotTable.SubtotalLocation: what each field's LayoutSubtotalLocation reads after it, the errors, the table's cells,
and each field's subtotalTop as the saved pivot table definition holds it.

Each case makes a workbook with a list of regions, products and amounts on its first sheet, builds a pivot table on a
second sheet with Region and Product as row fields and the sum of Amount as data, sets the case's layout and runs the
case's statements with pt as the table, then reads each row field's LayoutSubtotalLocation and every cell of the
table, and saves it.

    python scripts/measure_pivot_subtotal_location.py

writes tests/fixtures/pivot_subtotal_location.json, which tests/test_excel_pivot_subtotal_location.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_subtotal_location.json'
CASES = [
    ('xlCompactRow', ''),
    ('xlCompactRow', 'pt.SubtotalLocation xlAtBottom'),
    ('xlCompactRow', 'pt.SubtotalLocation xlAtBottom: pt.SubtotalLocation xlAtTop'),
    ('xlOutlineRow', 'pt.SubtotalLocation xlAtBottom'),
    ('xlTabularRow', 'pt.SubtotalLocation xlAtBottom'),
    ('xlCompactRow', 'pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom: pt.SubtotalLocation xlAtTop'),
    ('xlCompactRow', 'pt.SubtotalLocation 3'),
    ('xlCompactRow', 's = TypeName(pt.SubtotalLocation(xlAtBottom))'),
    ('xlCompactRow', 'pt.PivotFields("Product").Orientation = xlColumnField: pt.SubtotalLocation xlAtBottom'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, c As Range, s As String, f As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("East", "Ink", 20)',
             '    ws.Range("A4:C4").Value = Array("West", "Pen", 30)',
             '    ws.Range("A5:C5").Value = Array("West", "Ink", 40)',
             '    ws.Range("A6:C6").Value = Array("West", "Pad", 50)',
             '    Set ws = wb.Worksheets.Add(After:=ws)', '    ws.Name = "report"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R6C3").CreatePivotTable(ws.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.PivotFields("Product").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (layout, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        pt.RowAxisLayout {layout}'] + (
            [f'        {statements}'] if statements else [])
    lines += ['    End Select',
              '    For Each f In pt.RowFields: s = s & "|" & f.Name & "=" & f.LayoutSubtotalLocation: Next',
              '    s = s & "|"',
              '    For Each c In pt.TableRange1.Cells: s = s & c.Address(False, False) & ":" & c.Text & ";": Next',
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def subtotal_tops(path: Path) -> list[str]:
    """Each pivot field's name-free settings that place its subtotals: its axis and subtotalTop, as saved."""
    with zipfile.ZipFile(path) as package:
        name = next(one for one in package.namelist() if re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml', one))
        xml = package.read(name).decode('utf-8')
    return [' '.join(re.findall(r'\b(?:axis|subtotalTop|outline|compact)="[^"]*"', field))
            for field in re.findall(r'<pivotField\b[^>]*>', xml)]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [subtotal_tops(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{layout}: {statements}', answer, found]
                        for (layout, statements), answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case, '->', answer[:200], found)


if __name__ == '__main__':
    main()
