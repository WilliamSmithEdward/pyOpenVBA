"""What a form control's OnAction reads and saves: set from another workbook's macro, set with its own workbook's name, after a
save, and after the file is opened again.

    python scripts/measure_form_control_macros.py

writes tests/fixtures/form_control_macros.json, which tests/test_excel_old_controls.py checks the rule against.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/form_control_macros.json'
SOURCE = '''
Public Function Probe(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, o As Object, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set o = ws.Buttons.Add(48, 96, 72, 24)
    o.Name = "b1"
    o.OnAction = "Clicked"
    s = "set from another workbook=" & o.OnAction & ";shape=" & ws.Shapes("b1").OnAction
    Set o = ws.Buttons.Add(48, 144, 72, 24)
    o.Name = "b2"
    o.OnAction = wb.Name & "!Own"
    s = s & ";own name=" & Replace(o.OnAction, wb.Name, "<new>")
    Set o = ws.Shapes.AddFormControl(xlButtonControl, 48, 192, 72, 24)
    o.Name = "b3"
    o.OnAction = "ViaAdd"
    s = s & ";added as a shape=" & o.OnAction
    Set o = ws.Shapes.AddShape(1, 10, 10, 20, 20)
    o.OnAction = "Plain"
    s = s & ";autoshape=" & o.OnAction
    wb.SaveAs folder & "\\made.xlsx", xlOpenXMLWorkbook
    s = s & ";after save=" & ws.Shapes("b1").OnAction
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\made.xlsx")
    Set ws = wb.Worksheets(1)
    s = s & ";reopened foreign=" & ws.Shapes("b1").OnAction & ";reopened own=" & ws.Shapes("b2").OnAction & ";reopened through Buttons=" & ws.Buttons("b2").OnAction
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Probe', (folder,), timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = dict(pair.split('=', 1) for pair in str(result.value).split(';'))
        files: dict[str, list[str]] = {}
        with zipfile.ZipFile(Path(folder) / 'made.xlsx') as package:
            for name in package.namelist():
                if 'vmlDrawing' in name or name.endswith('sheet1.xml'):
                    files[name] = re.findall(r'macro="[^"]*"|<x:FmlaMacro>[^<]*', package.read(name).decode('utf-8', 'replace'))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers, 'files': files}, indent=1) + '\n', encoding='utf-8')
    print(answers, files)


if __name__ == '__main__':
    main()
