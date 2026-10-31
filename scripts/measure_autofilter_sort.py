"""AutoFilter.Sort: the filter's own sort settings, apart from the sheet's; what its range, header and fields read
and do, what removing the filter leaves, and the autoFilter element a save writes before and after Apply.

    python scripts/measure_autofilter_sort.py

writes tests/fixtures/autofilter_sort.json, which tests/test_excel_autofilter_sort.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/autofilter_sort.json'
#: Each case: statements on a sheet whose A1:B6 has an AutoFilter, a read, and the file saved, if any.
CASES = [
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")'],
     'ws.AutoFilter.Sort.SortFields.Count & " " & ws.Sort.SortFields.Count', 'added.xlsx'),
    (['ws.Sort.SortFields.Add ws.Range("A2:A6")'],
     'ws.AutoFilter.Sort.SortFields.Count & " " & ws.Sort.SortFields.Count', 'sheet.xlsx'),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")', 'ws.AutoFilter.Sort.Apply'], 'Grid(ws)', 'applied.xlsx'),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("B2:B6"), , xlDescending', 'ws.AutoFilter.Sort.Apply'],
     'Grid(ws) & " " & ws.AutoFilter.Sort.Rng.Address & " " & ws.Sort.SortFields.Count', 'down.xlsx'),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("B2:B6")', 'ws.AutoFilter.Sort.MatchCase = True',
      'ws.AutoFilter.Sort.Apply'], 'Grid(ws)', 'case.xlsx'),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")', 'ws.AutoFilter.Sort.Apply',
      'ws.AutoFilter.Sort.SortFields.Add ws.Range("B2:B6")'], 'ws.AutoFilter.Sort.SortFields.Count', 'later.xlsx'),
    (['ws.AutoFilter.Sort.SetRange ws.Range("A1:B3")'], 'ws.AutoFilter.Sort.Rng.Address', ''),
    ([], 'ws.AutoFilter.Sort.Header', ''),
    (['ws.AutoFilter.Sort.Header = xlNo'], 'ws.AutoFilter.Sort.Header', ''),
    (['ws.AutoFilter.Sort.MatchCase = True'], 'ws.AutoFilter.Sort.MatchCase & " " & ws.Sort.MatchCase', ''),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")', 'ws.AutoFilter.ShowAllData'],
     'ws.AutoFilter.Sort.SortFields.Count', ''),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")', 'ws.AutoFilterMode = False',
      'ws.Range("A1:B6").AutoFilter'], 'ws.AutoFilter.Sort.SortFields.Count', ''),
    ([], 'TypeName(ws.AutoFilter.Sort) & " " & TypeName(ws.AutoFilter.Sort.Parent) & " " '
         '& TypeName(ws.AutoFilter.Sort.SortFields)', ''),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("A2:A6")'], 'ws.AutoFilter.Sort.SortFields(1).Key.Address', ''),
    (['ws.AutoFilter.Sort.SortFields.Add ws.Range("D2:D6")', 'ws.AutoFilter.Sort.Apply'], 'Grid(ws)', ''),
    (['ws.AutoFilter.Sort.Apply'], 'Grid(ws)', 'empty.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Grid(ByVal ws As Worksheet) As String', '    Dim c As Range, s As String',
             '    For Each c In ws.Range("A2:B6")', '        s = s & "|" & c.Formula', '    Next',
             '    Grid = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:B6").Value = Application.Transpose(Application.Transpose(Array(Array("h1", "h2"), '
             'Array(3, "c"), Array(1, "a"), Array(5, "E"), Array(2, "b"), Array(4, "e"))))',
             '    ws.Range("A1:B6").AutoFilter', '    On Error GoTo Failed', '    Select Case which']
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
                found = re.search(r'<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>', sheet, re.DOTALL)
                outside = re.search(r'<sortState\b[^>]*/>|<sortState\b.*?</sortState>',
                                    re.sub(r'<autoFilter\b.*?</autoFilter>', '', sheet, flags=re.DOTALL), re.DOTALL)
                parts[saved] = [found.group(0) if found else '', outside.group(0) if outside else '']
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, saved, answer in record['cases']:
        print(statements[-1:] or [read], '->', answer, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
