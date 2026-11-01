"""A sparkline group's DateRange and ModifyDateRange, and its horizontal axis's RightToLeftPlotOrder: what each reads,
what a set does, and the group a save writes.

    python scripts/measure_sparkline_dates.py

writes tests/fixtures/sparkline_dates.json, which tests/test_excel_sparkline_dates.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sparkline_dates.json'
#: Each case: statements on a line group over A1:D1 at E1, with dates in A6:D6, what is read, and the file saved.
CASES = [
    ([], 'g.DateRange', ''),
    (['g.ModifyDateRange "A6:D6"'], 'g.DateRange & " " & g.Axes.Horizontal.IsDateAxis', 'dated.xlsx'),
    (['g.ModifyDateRange "Sheet1!A6:D6"'], 'g.DateRange', ''),
    (['g.ModifyDateRange "$A$6:$D$6"'], 'g.DateRange', 'absolute.xlsx'),
    (['g.ModifyDateRange "a6:c6"'], 'g.DateRange', ''),
    (['g.ModifyDateRange "A6"'], 'g.DateRange', 'cell.xlsx'),
    (['g.ModifyDateRange "nonsense"'], 'g.DateRange', ''),
    (['g.ModifyDateRange ""'], 'g.DateRange', ''),
    (['g.ModifyDateRange "Sheet9!A6:D6"'], 'g.DateRange', ''),
    (['g.ModifyDateRange "A6:D6"', 'g.ModifyDateRange ""'], 'g.DateRange & " " & g.Axes.Horizontal.IsDateAxis', ''),
    (['g.DateRange = "A6:D6"'], 'g.DateRange', ''),
    (['g.ModifyDateRange "A6:A9"'], 'g.DateRange', ''),
    (['g.ModifyDateRange "A6:D6"', 'ws.Range("A6:D6").Clear'], 'g.DateRange', ''),
    ([], 'g.Axes.Horizontal.RightToLeftPlotOrder', ''),
    (['g.Axes.Horizontal.RightToLeftPlotOrder = True'], 'g.Axes.Horizontal.RightToLeftPlotOrder', 'rtl.xlsx'),
    (['g.Axes.Horizontal.RightToLeftPlotOrder = True', 'g.Axes.Horizontal.RightToLeftPlotOrder = False'],
     'g.Axes.Horizontal.RightToLeftPlotOrder', 'ltr.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, g As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1:D4").Value = Application.Transpose(Application.Transpose('
             'Array(Array(1, 3, -2, 4), Array(-1, 2, -3, 4), Array(5, 4, 3, 2), Array(2, 2, 2, 2))))',
             '    ws.Range("A6:D6").Value = Application.Transpose(Application.Transpose('
             'Array(Array(45000, 45001, 45005, 45010))))',
             '    On Error GoTo Failed', '    Set g = ws.Range("E1").SparklineGroups.Add(xlSparkLine, "A1:D1")',
             '    Select Case which']
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
        groups = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<x14:sparklineGroup\b.*?</x14:sparklineGroup>', sheet, re.DOTALL)
                groups[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'groups': groups}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, _, answer in record['cases']:
        print((statements[-1] if statements else read)[:60], '->', answer)


if __name__ == '__main__':
    main()
