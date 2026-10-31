"""Range.SparklineGroups: Add's groups of line, column and win/loss sparklines, what each group and sparkline reads,
Modify and its kin, Clear, Delete and Range.Clear, and the worksheet extension a save writes for them.

    python scripts/measure_sparklines.py

writes tests/fixtures/sparklines.json, which tests/test_excel_sparklines.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sparklines.json'
ADD = 'Set g = ws.Range("E1").SparklineGroups.Add(xlSparkLine, "A1:D1")'
#: Each case: statements on a sheet with numbers in A1:D4, a read, and the file saved, if any.
CASES = [
    ([ADD], 'TypeName(g) & " " & g.Type & " " & g.SourceData & " " & g.Location.Address & " " & g.Count & " " & '
            'ws.Cells.SparklineGroups.Count & " " & ws.Range("E1").SparklineGroups.Count & " " & '
            'ws.Range("A1").SparklineGroups.Count', 'one.xlsx'),
    (['Set g = ws.Range("E1:E3").SparklineGroups.Add(xlSparkColumn, "A1:D3")'],
     'g.Type & " " & g.SourceData & " " & g.Location.Address & " " & g.Count & " " & g.Item(2).Location.Address & '
     '" " & g.Item(2).SourceData & " " & TypeName(g.Item(1))', 'three.xlsx'),
    (['Set g = ws.Range("E1").SparklineGroups.Add(xlSparkColumnStacked100, "A1:D1")'], 'g.Type', 'winloss.xlsx'),
    ([ADD], 'g.LineWeight & " " & g.SeriesColor.Color & " " & g.SeriesColor.ThemeColor & " " & '
            'g.SeriesColor.TintAndShade & " " & g.DisplayBlanksAs & " " & g.DisplayHidden', ''),
    ([ADD, 'g.SeriesColor.Color = RGB(255, 0, 0)', 'g.LineWeight = 1.5'], 'g.LineWeight & " " & g.SeriesColor.Color',
     'styled.xlsx'),
    ([ADD, 'ws.Range("E3").SparklineGroups.Add xlSparkColumn, "A2:D2"'],
     'ws.Cells.SparklineGroups.Count & " " & ws.Cells.SparklineGroups(1).Location.Address & " " & '
     'ws.Cells.SparklineGroups(2).Location.Address & " " & ws.Range("E1:E3").SparklineGroups.Count & " " & '
     'ws.Range("E2").SparklineGroups.Count', 'two.xlsx'),
    ([ADD, 'g.Delete'], 'ws.Cells.SparklineGroups.Count', ''),
    ([ADD, 'ws.Range("E1").SparklineGroups.Clear'], 'ws.Cells.SparklineGroups.Count', 'cleared.xlsx'),
    (['Set g = ws.Range("E1:E2").SparklineGroups.Add(xlSparkLine, "A1:D2")', 'ws.Range("E2").SparklineGroups.Clear'],
     'ws.Cells.SparklineGroups.Count & " " & g.Count & " " & g.Location.Address & " " & g.SourceData', ''),
    ([ADD, 'ws.Range("E1").ClearContents'], 'ws.Cells.SparklineGroups.Count', ''),
    ([ADD, 'ws.Range("E1").Clear'], 'ws.Cells.SparklineGroups.Count', ''),
    ([ADD, 'ws.Range("A1:D1").Clear'], 'ws.Cells.SparklineGroups.Count', ''),
    ([ADD], 'g.Creator & " " & TypeName(g.Parent) & " " & TypeName(g.Application)', ''),
    ([ADD], 'TypeName(g.Item(1).Parent) & " " & g.Item(1).SourceData & " " & g.Item(1).Location.Address', ''),
    ([ADD], 'TypeName(ws.Range("E1").SparklineGroups.Parent) & " " & ws.Range("E1").SparklineGroups.Parent.Address',
     ''),
    ([ADD], 'g.Item(2).SourceData', ''),
    ([ADD], 'g.Item(0).SourceData', ''),
    ([ADD], 'ws.Range("E1").SparklineGroups(2).Type', ''),
    ([ADD, 'g.Type = xlSparkColumn'], 'g.Type', 'retyped.xlsx'),
    ([ADD, 'g.Type = 9'], 'g.Type', ''),
    ([], 'ws.Range("E1").SparklineGroups.Add(4, "A1:D1").Type', ''),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "A1").SourceData', ''),
    ([], 'ws.Range("E1:F1").SparklineGroups.Add(xlSparkLine, "A1:D2").Location.Address', ''),
    ([], 'ws.Range("E1:F2").SparklineGroups.Add(xlSparkLine, "A1:D4").Count', ''),
    ([], 'ws.Range("E1:E4").SparklineGroups.Add(xlSparkLine, "A1:D4").Item(3).SourceData', ''),
    ([], 'ws.Range("A10:D10").SparklineGroups.Add(xlSparkLine, "A1:D4").Item(2).SourceData', 'across.xlsx'),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "Sheet2!A1:D1").SourceData', ''),
    (['wb.Worksheets.Add After:=ws'], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "Sheet2!A1:D1").SourceData',
     'elsewhere.xlsx'),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "nonsense").SourceData', ''),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "").SourceData', ''),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "$A$1:$D$1").SourceData', 'absolute.xlsx'),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "a1:d1").SourceData', ''),
    ([], 'ws.Range("E1").SparklineGroups.Add(xlSparkLine, "A1:D2").SourceData', ''),
    ([ADD, 'ws.Range("E1").SparklineGroups.Add xlSparkLine, "A2:D2"'],
     'ws.Cells.SparklineGroups.Count & " " & ws.Range("E1").SparklineGroups(1).SourceData', ''),
    ([ADD], 'ws.Range("E1").Value & "|" & IsEmpty(ws.Range("E1").Value)', ''),
    ([ADD, 'ws.Range("E1").Value = 5'], 'ws.Range("E1").Value & " " & ws.Cells.SparklineGroups.Count', ''),
    ([ADD, 'g.Location = ws.Range("F1")'], 'g.Location.Address', ''),
    ([ADD, 'g.SourceData = "A2:D2"'], 'g.SourceData', ''),
    ([ADD, 'g.Modify ws.Range("F1"), "A2:D2"'], 'g.Location.Address & " " & g.SourceData', 'modified.xlsx'),
    ([ADD, 'g.ModifySourceData "A3:D3"'], 'g.SourceData', ''),
    ([ADD, 'g.ModifyLocation ws.Range("G1")'], 'g.Location.Address', ''),
    ([ADD, 'g.Item(1).SourceData = "A2:D2"'], 'g.SourceData & " " & g.Item(1).SourceData', ''),
    ([ADD, 'g.Item(1).ModifyLocation ws.Range("H1")'], 'g.Location.Address', ''),
    ([ADD, 'g.SeriesColor.Color = vbGreen'], 'g.SeriesColor.Color & " " & g.SeriesColor.ThemeColor', ''),
    ([ADD, 'g.LineWeight = 0'], 'g.LineWeight', ''),
    ([ADD, 'g.LineWeight = -1'], 'g.LineWeight', ''),
    ([ADD, 'g.LineWeight = 2000'], 'g.LineWeight', ''),
    ([ADD, 'g.LineWeight = 0.75'], 'g.LineWeight', 'weight.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, g As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1:D4").Value = Application.Transpose(Application.Transpose('
             'Array(Array(1, 3, 2, 4), Array(-1, 2, -3, 4), Array(5, 4, 3, 2), Array(2, 2, 2, 2))))',
             '    On Error GoTo Failed', '    Select Case which']
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
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, saved, answer in record['cases']:
        print((statements[-1] if statements else read)[:70], '->', answer)
    for saved, part in parts.items():
        print(saved, part[:600])


if __name__ == '__main__':
    main()
