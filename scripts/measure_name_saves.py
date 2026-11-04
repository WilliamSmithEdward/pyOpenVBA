"""What Excel's saved file holds for a name that does not come back as it was: a macro name, and a name that refers to nothing.

Each case changes one name, or a second beside it, saves to xlsx and xlsm, reads workbook.xml's definedNames out of the file, and
opens it again to say what the Names collection holds.

    python scripts/measure_name_saves.py

writes tests/fixtures/name_saves.json, which tests/test_excel_name_saves.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/name_saves.json'
CASES = {
    'plain': '',
    'function': 'n.MacroType = 1',
    'command': 'n.MacroType = 2',
    'function category': 'n.MacroType = 1: n.Category = "Financial"',
    'command key': 'n.MacroType = 2: n.ShortcutKey = "a"',
    'function with other': 'wb.Names.Add "Zed", "=Sheet1!$B$1": n.MacroType = 1',
    'command with other': 'wb.Names.Add "Zed", "=Sheet1!$B$1": n.MacroType = 2',
    'empty text': 'n.RefersTo = ""',
    'empty value': 'n.Value = ""',
    'empty Empty value': 'n.Value = Empty',
    'equals': 'n.RefersTo = "="',
    'empty with other': 'wb.Names.Add "Zed", "=Sheet1!$B$1": n.RefersTo = ""',
    'empty hidden': 'n.RefersTo = "": n.Visible = False',
    'empty comment': 'n.RefersTo = "": n.Comment = "c"',
    'macro then back': 'n.MacroType = 1: n.MacroType = -4142',
    'macro then refers': 'n.MacroType = 1: n.RefersTo = "=Sheet1!$C$1"',
    'macro to formula': 'n.RefersTo = "=1+2": n.MacroType = 2',
}

SOURCE = '''Public Function Tried(ByVal which As Long, ByVal path As String, ByVal kind As Long) As String
    Dim wb As Workbook, n As Name, s As String, one As Name
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Names.Add "Nm", "=Sheet1!$A$1"
    Set n = wb.Names("Nm")
    On Error GoTo Failed
    Select Case which
@@CASES@@
    End Select
    s = "count " & wb.Names.Count
    For Each one In wb.Names
        s = s & "|" & one.Name & " " & one.RefersTo & " " & one.Visible & " " & one.MacroType
    Next
    wb.SaveAs path, kind
    wb.Close False
    Set wb = Workbooks.Open(path, 0)
    s = s & "@reopened count " & wb.Names.Count
    For Each one In wb.Names
        s = s & "|" & one.Name & " " & one.RefersTo & " " & one.Visible & " " & one.MacroType
    Next
    wb.Close False
    Tried = s
    Exit Function
Failed:
    Tried = s & " Err " & Err.Number & " " & Err.Description
    On Error Resume Next
    wb.Close False
End Function
'''


def _names_of(path: Path) -> str:
    with zipfile.ZipFile(path) as saved:
        xml = saved.read('xl/workbook.xml').decode('utf-8')
    found = re.search(r'<definedNames>.*?</definedNames>|<definedNames/>', xml, re.DOTALL)
    return found.group(0) if found else ''


def main() -> None:
    lines = ''.join(f'    Case {index}\n        {statement}\n' for index, statement in enumerate(CASES.values()))
    source = SOURCE.replace('@@CASES@@', lines.rstrip('\n'))
    answers: dict[str, dict[str, str]] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for index, name in enumerate(CASES):
            for extension, kind in (('xlsx', 51), ('xlsm', 52)):
                path = Path(folder) / f'c{index}.{extension}'
                result = excel.run_vba(source, 'Tried', (index, str(path), kind), timeout=40.0)
                answer = str(result.value) if result.ok else f'breaks: {result.outcome}'
                file_names = _names_of(path) if path.exists() else ''
                answers[f'{name} {extension}'] = {'read': answer, 'file': file_names}
                print(name, extension, '=>', answer, '##', file_names, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': CASES, 'source': source, 'answers': answers}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
