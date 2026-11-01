"""SparklineGroups.ClearGroups, Group and Ungroup on a sheet with a line group at E1:E2 and a column group at E3:
which groups are left, newest first, and the groups a save writes.

    python scripts/measure_sparkline_grouping.py

writes tests/fixtures/sparkline_grouping.json, which tests/test_excel_sparkline_grouping.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sparkline_grouping.json'
SETUP = ['ws.Range("E1:E2").SparklineGroups.Add xlSparkLine, "A1:D2"',
         'ws.Range("E3").SparklineGroups.Add xlSparkColumn, "A3:D3"']
#: Each case: statements after SETUP, what is read, and the file saved, if any.
CASES = [
    ([], 'Groups(ws)', ''),
    (['ws.Range("E2").SparklineGroups.ClearGroups'], 'Groups(ws)', 'cleargroups.xlsx'),
    (['ws.Range("E1:E3").SparklineGroups.Group ws.Range("E3")'], 'Groups(ws)', 'grouped.xlsx'),
    (['ws.Range("E1:E3").SparklineGroups.Group ws.Range("E1")'], 'Groups(ws)', 'grouped1.xlsx'),
    (['ws.Range("E2").SparklineGroups.Ungroup'], 'Groups(ws)', 'ungrouped.xlsx'),
    (['ws.Range("E1:E2").SparklineGroups.Ungroup'], 'Groups(ws)', 'ungrouped2.xlsx'),
    (['ws.Range("E1").SparklineGroups.Group ws.Range("E3")'], 'Groups(ws)', ''),
    (['ws.Range("E3").SparklineGroups.Group ws.Range("E1")'], 'Groups(ws)', ''),
    (['ws.Range("E1:E3").SparklineGroups.Group ws.Range("H9")'], 'Groups(ws)', ''),
    (['ws.Range("E1:E3").SparklineGroups.ClearGroups'], 'Groups(ws)', ''),
    (['ws.Range("G1").SparklineGroups.ClearGroups'], 'Groups(ws)', ''),
    (['ws.Range("G1").SparklineGroups.Ungroup'], 'Groups(ws)', ''),
    (['ws.Range("E1:E3").SparklineGroups(1).Points.Markers.Visible = True',
      'ws.Range("E1:E3").SparklineGroups.Group ws.Range("E3")'],
     'Groups(ws) & " " & ws.Range("E1").SparklineGroups(1).Points.Markers.Visible', 'kept.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Groups(ByVal ws As Worksheet) As String', '    Dim g As Object, s As String',
             '    For Each g In ws.Cells.SparklineGroups',
             '        s = s & "|" & g.Type & ":" & g.Location.Address(False, False) & ":" & g.SourceData', '    Next',
             '    Groups = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1:D4").Value = Application.Transpose(Application.Transpose('
             'Array(Array(1, 3, -2, 4), Array(-1, 2, -3, 4), Array(5, 4, 3, 2), Array(2, 2, 2, 2))))',
             '    On Error GoTo Failed', *(f'    {statement}' for statement in SETUP), '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {read}')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<extLst>.*</extLst>', sheet, re.DOTALL)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'setup': SETUP, 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, _, answer in record['cases']:
        print((statements[-1] if statements else read)[:70], '->', answer)


if __name__ == '__main__':
    main()
