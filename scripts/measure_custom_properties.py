"""Measure Worksheet.CustomProperties: adding, reading, changing and deleting a sheet's custom properties, what the
saved file holds for them, and what a reopened workbook and a copied sheet read.

Each case runs on a new workbook and answers as text; some save it, and the sheet part's customProperties element,
its relationships, the property parts and their content types are recorded.

    python scripts/measure_custom_properties.py

writes tests/fixtures/custom_properties.json, which tests/test_excel_custom_properties.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/custom_properties.json'
LIST = ('For Each p In ws.CustomProperties: r = r & "|" & p.Name & "=" & TypeName(p.Value) & ":" & p.Value: Next: '
        'r = ws.CustomProperties.Count & r')
#: Each case: statements on a new workbook wb with sheet ws; the file saved, if any.
CASES: list[tuple[list[str], str]] = [
    (['r = ws.CustomProperties.Count & ":" & TypeName(ws.CustomProperties)'], ''),
    (['Set p = ws.CustomProperties.Add("Key", "Val")', 'r = TypeName(p) & ":" & p.Name & "=" & p.Value'], ''),
    (['ws.CustomProperties.Add "Key", "Val"', LIST], 'one.xlsx'),
    (['ws.CustomProperties.Add "Num", 5', 'ws.CustomProperties.Add "Dbl", 1.5', 'ws.CustomProperties.Add "Bool", True',
      'ws.CustomProperties.Add "Day", #1/2/2024#', LIST], 'types.xlsx'),
    (['ws.CustomProperties.Add "Key", "a"', 'ws.CustomProperties.Add "Key", "b"', LIST], 'duplicate.xlsx'),
    (['ws.CustomProperties.Add "Key", "a"', 'ws.CustomProperties.Add "KEY", "b"', LIST], ''),
    (['ws.CustomProperties.Add "", "a"', LIST], ''),
    (['ws.CustomProperties.Add "Key", ""', LIST], 'blank.xlsx'),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties.Add "A", "2"', LIST], 'order.xlsx'),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties.Add "A", "2"',
      'r = ws.CustomProperties(1).Name & ws.CustomProperties.Item(2).Name'], ''),
    (['ws.CustomProperties.Add "B", "1"', 'r = ws.CustomProperties("B").Value'], ''),
    (['ws.CustomProperties.Add "B", "1"', 'r = ws.CustomProperties(2).Value'], ''),
    (['ws.CustomProperties.Add "B", "1"', 'r = ws.CustomProperties(0).Value'], ''),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties.Add "A", "2"', 'ws.CustomProperties(1).Delete',
      LIST], 'deleted.xlsx'),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties(1).Value = 7', LIST], ''),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties(1).Name = "C"', LIST], ''),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties(1).Value = Empty', LIST], ''),
    (['ws.CustomProperties.Add "B", Null', LIST], ''),
    (['Set wso = ws', 'wso.CustomProperties.Add "B"', LIST], ''),
    (['Set wso = ws', 'wso.CustomProperties.Add Value:="x"', LIST], ''),
    (['ws.CustomProperties.Add "B", ws.Range("A1")', LIST], ''),
    (['ws.Range("A1").Value = "cell"', 'ws.CustomProperties.Add "B", ws.Range("A1")', LIST], ''),
    (['ws.CustomProperties.Add "B", Array(1, 2)', LIST], ''),
    (['ws.CustomProperties.Add String(300, "n"), "v"', 'r = Len(ws.CustomProperties(1).Name)'], ''),
    (['ws.CustomProperties.Add "n", String(70000, "v")', 'r = Len(ws.CustomProperties(1).Value)'], ''),
    (['ws.CustomProperties.Add "a b!", "x y"', LIST], 'odd.xlsx'),
    (['ws.CustomProperties.Add "B", "1"', 'ws.Copy After:=ws', 'Set ws = wb.Worksheets(2)', LIST], ''),
    (['ws.CustomProperties.Add "B", "1"',
      'r = TypeName(ws.CustomProperties.Parent) & TypeName(ws.CustomProperties(1).Parent) & '
      'ws.CustomProperties.Creator & ws.CustomProperties(1).Creator'], ''),
    (['ws.CustomProperties.Add "B", "1"', 'ws.CustomProperties.Add "A", "2"',
      'wb.SaveAs folder & "\\reopen.xlsx"', 'wb.Close', 'Set wb = Workbooks.Open(folder & "\\reopen.xlsx")',
      'Set ws = wb.Worksheets(1)', LIST], ''),
    (['ws.CustomProperties.Add "Num", 5', 'ws.CustomProperties.Add "Day", #1/2/2024#',
      'wb.SaveAs folder & "\\reopen_types.xlsx"', 'wb.Close', 'Set wb = Workbooks.Open(folder & "\\reopen_types.xlsx")',
      'Set ws = wb.Worksheets(1)', LIST], ''),
    (['Set ws2 = wb.Worksheets.Add(After:=ws)', 'ws2.CustomProperties.Add "S2", "two"',
      'ws.CustomProperties.Add "S1", "one"', LIST], 'two_sheets.xlsx'),
    (['ws.CustomProperties.Add "B", 1E+300', 'ws.CustomProperties.Add "C", 1.23456789012345', LIST], 'numbers.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, wso As Object, p As Object, r As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    Tried = CStr(r)', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _parts(path: Path) -> dict[str, str]:
    """The sheet parts' customProperties, their relationships, the property parts as hex, and content types."""
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        for name in names:
            if re.fullmatch(r'xl/worksheets/sheet[0-9]+\.xml', name):
                xml = package.read(name).decode('utf-8')
                found[name] = (re.search(r'<customProperties>.*?</customProperties>', xml, re.DOTALL) or [''])[0]
                found[name + ':root'] = (re.search(r'<worksheet\b[^>]*>', xml) or [''])[0]
            elif re.fullmatch(r'xl/worksheets/_rels/sheet[0-9]+\.xml\.rels', name):
                found[name] = package.read(name).decode('utf-8')
            elif name.startswith('xl/customProperty'):
                found[name] = package.read(name).hex()
        found['[Content_Types].xml'] = package.read('[Content_Types].xml').decode('utf-8')
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {saved: _parts(Path(folder) / saved) for _, saved in CASES if saved and (Path(folder) / saved).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statements, saved, answer) in enumerate(record['cases']):
        print(index, statements[-1][:60], '->', answer[:120], '|', saved)


if __name__ == '__main__':
    main()
