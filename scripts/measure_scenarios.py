"""Measure Worksheet.Scenarios: adding, showing, changing and deleting what-if scenarios, what they read, and what a
save writes for them.

Each case runs on a new workbook and answers as text; some save it, and the sheet part's scenarios element is
recorded. Application.UserName is set to "Probe" for the case and put back after it.

    python scripts/measure_scenarios.py

writes tests/fixtures/scenarios.json, which tests/test_excel_scenarios.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/scenarios.json'
SETUP = ['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 2', 'ws.Range("B1").Formula = "=A1+A2"']
LIST = ('r = ws.Scenarios.Count: For Each sc In ws.Scenarios: r = r & "|" & sc.Name & ":" & '
        'sc.ChangingCells.Address & ":" & Join(Vals(sc), ",") & ":" & sc.Comment & ":" & sc.Locked & ":" & '
        'sc.Hidden & ":" & sc.Index: Next')
#: Each case: statements on a new workbook wb with sheet ws, after the setup; the file saved, if any.
CASES: list[tuple[list[str], str]] = [
    (['r = ws.Scenarios.Count & TypeName(ws.Scenarios)'], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', LIST], 'one.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios("Low").Show',
      'r = ws.Range("A1").Value & "," & ws.Range("A2").Value & "," & ws.Range("B1").Value'], 'shown.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20), "Note", False, True', LIST], 'flags.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios.Add "High", ws.Range("A1"), Array(99)',
      LIST], 'two.xlsx'),
    (['ws.Scenarios.Add "B", ws.Range("A1"), Array(5)', 'ws.Scenarios.Add "A", ws.Range("A1"), Array(6)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios.Add "Low", ws.Range("A1"), Array(5)', LIST],
     ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20, 30)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2")', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("B1"), Array(5)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1,A3"), Array(1, 3)', LIST], 'union.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), "text"', LIST], 'text.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), 1.5', LIST], 'fraction.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), #1/2/2024#', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), True', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), "=1+1"', 'ws.Scenarios(1).Show', 'r = ws.Range("A1").Formula'], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios(1).Delete', LIST], 'deleted.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios(1).ChangeScenario ws.Range("A1"), Array(7)',
      LIST], 'changed.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios(1).Name = "Lower"',
      'ws.Scenarios(1).Comment = "c"', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'r = ws.Scenarios(1).Values(2) & TypeName('
      'ws.Scenarios(1).Values(2)) & TypeName(ws.Scenarios(1).Values)'], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'r = ws.Scenarios("low").Name'], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'r = ws.Scenarios(2).Name'], ''),
    (['ws.Scenarios.Add "", ws.Range("A1"), Array(1)', LIST], ''),
    (['ws.Scenarios.Add String(256, "n"), ws.Range("A1"), Array(1)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), Array(1)', 'ws.Scenarios(1).Show', 'r = ws.Range("A1").Value'],
     'shown_one.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), Array(1)', 'ws.Scenarios.Add "High", ws.Range("A1"), Array(2)',
      'ws.Scenarios(2).Show', 'ws.Scenarios(1).Delete', LIST], 'shown_then_deleted.xlsx'),
    (['ws.Range("A1").NumberFormat = "0.00"', 'ws.Scenarios.Add "Low", ws.Range("A1"), Array(3)', 'ws.Scenarios(1).Show',
      'r = ws.Range("A1").Text'], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), Array(3)', 'ws.Copy After:=ws', 'Set ws = wb.Worksheets(2)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), Array(3)', 'r = TypeName(ws.Scenarios(1).Parent) & ws.Scenarios.Creator'],
     ''),
    (['ws.Protect', 'ws.Scenarios.Add "Low", ws.Range("A1"), Array(3)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("C5:D6"), Array(1, 2, 3, 4)', LIST], 'block.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), 5', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), CLng(5)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), CDbl(5)', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), "7"', LIST], ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1"), Array("=A2*2")', 'ws.Scenarios(1).Show', 'r = ws.Range("A1").Formula'],
     ''),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios(1).Locked = False',
      'ws.Scenarios(1).Hidden = True', LIST], 'set_flags.xlsx'),
    (['ws.Scenarios.Add "Low", ws.Range("A1:A2"), Array(10, 20)', 'ws.Scenarios.Add "High", ws.Range("A1:A2"), '
      'Array(30, 40)', 'ws.Scenarios("High").Show', 'ws.Scenarios("Low").Show', LIST], 'shown_second.xlsx'),
    (['ws.Range("A1").Value = 1.123456789012345', 'ws.Scenarios.Add "Low", ws.Range("A1")', LIST], 'precise.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Vals(ByVal sc As Object) As String()',
             '    Dim v As Variant, out() As String, i As Long', '    v = sc.Values',
             '    ReDim out(LBound(v) To UBound(v))',
             '    For i = LBound(v) To UBound(v): out(i) = TypeName(v(i)) & "=" & CStr(v(i)): Next', '    Vals = out',
             'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sc As Object, r As Variant, user As String',
             '    user = Application.UserName', '    Application.UserName = "Probe"',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)']
    lines += [f'    {statement}' for statement in SETUP]
    lines += ['    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    Tried = CStr(r)', 'Done:', '    wb.Close False', '    Application.UserName = user',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description',
              '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _scenarios(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    return (re.search(r'<scenarios\b.*?</scenarios>', xml, re.DOTALL) or [''])[0]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            before = excel.run_vba('Public Function U() As String\n    U = Application.UserName\nEnd Function\n',
                                   'U', (), timeout=60.0)
            result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
            after = excel.run_vba('Public Function U() As String\n    U = Application.UserName\nEnd Function\n',
                                  'U', (), timeout=60.0)
        assert result.ok and not result.dialogs, result
        assert before.value == after.value, (before, after)
        answers = str(result.value).split('@@@')
        saved = {name: _scenarios(Path(folder) / name) for _, name in CASES if name and (Path(folder) / name).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, name, answer] for (statements, name), answer in zip(CASES, answers, strict=True)],
              'scenarios': saved}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statements, name, answer) in enumerate(record['cases']):
        print(index, statements[-1][:55], '->', answer[:150], '|', saved.get(name, ''))


if __name__ == '__main__':
    main()
