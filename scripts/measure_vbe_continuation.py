"""CodeModule.InsertLines and AddFromString with text that goes on with " _": CountOfLines after them and the module
the workbook then saves, read back with ExcelFile.get_module.

    python scripts/measure_vbe_continuation.py

writes tests/fixtures/vbe_continuation.json, which tests/test_excel_vbe_continuation.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.excel import ExcelFile

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbe_continuation.json'
#: Statements on m, a new standard module's CodeModule.
CASES = [
    'm.AddFromString "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function"',
    'm.AddFromString "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function" & vbCrLf',
    'm.AddFromString "Function G( _" & vbCrLf & "x, _" & vbCrLf & "y)" & vbCrLf & "End Function" & vbCrLf',
    'm.AddFromString "Sub A( _" & vbCrLf & ")" & vbCrLf & "End Sub" & vbCrLf & "Sub B( _" & vbCrLf & ")" & vbCrLf & '
    '"End Sub" & vbCrLf',
    'm.AddFromString "Sub A()" & vbCrLf & "x = 1 + _" & vbCrLf & "2" & vbCrLf & "End Sub" & vbCrLf',
    'm.InsertLines 1, "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function" & vbCrLf',
    'm.AddFromString "Dim a As _" & vbCrLf & "Long" & vbCrLf',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub": '
    'm.AddFromString "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function"',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub": '
    'm.InsertLines 3, "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function"',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub": '
    'm.InsertLines 3, "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function" & vbCrLf',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub": '
    'm.InsertLines 1, "Function G( _" & vbCrLf & "x)" & vbCrLf & "End Function" & vbCrLf',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub": m.InsertLines 2, "x = 1 + _" & vbCrLf & "2"',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub" & vbCrLf: m.InsertLines 2, "x = 1 + _" & vbCrLf & "2"',
    'm.InsertLines 1, "x = 1 + _"',
    'm.InsertLines 1, "Sub A( _" & vbCrLf & ")" & vbCrLf & "End Sub": m.DeleteLines 1, 1',
    'm.AddFromString "Sub A()" & vbCrLf & "End Sub" & vbCrLf',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, m As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set m = wb.VBProject.VBComponents.Add(1).CodeModule', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = m.CountOfLines',
              '    wb.SaveAs folder & "\\c" & which & ".xlsm", xlOpenXMLWorkbookMacroEnabled', '    wb.Close False',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        modules = [ExcelFile(Path(folder) / f'c{index}.xlsm').get_module('Module1') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, module] for case, answer, module in zip(CASES, answers, modules, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, module in record['cases']:
        print(case[:70], '->', answer, repr(module[31:]))


if __name__ == '__main__':
    main()
