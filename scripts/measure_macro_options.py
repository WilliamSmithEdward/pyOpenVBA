"""Application.MacroOptions: the attribute lines it writes into a procedure, read back from the module of the workbook
saved after each call, and the errors it raises.

Each case makes a workbook, puts a module in it through its VBProject (the harness trusts access to the VBA project),
runs MacroOptions on its macros and saves it as .xlsm.

    python scripts/measure_macro_options.py

writes tests/fixtures/macro_options.json, which tests/test_excel_macro_options.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.excel import ExcelFile

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/macro_options.json'
#: The module each case's workbook holds, its lines joined by \n.
#: No line in it goes on with " _": Excel's AddFromString then ends the module with one more blank line, which the
#: model does not yet add.
MODULE = 'Function F(a, b)\\n    F = a + b\\nEnd Function\\nSub S()\\nEnd Sub\\n'
CASES = [
    'Application.MacroOptions Macro:="F", Description:="Adds two"',
    'Application.MacroOptions Macro:="F", Category:=5',
    'Application.MacroOptions Macro:="S", HasShortcutKey:=True, ShortcutKey:="q"',
    'Application.MacroOptions Macro:="S", Description:="Does it", HasShortcutKey:=True, ShortcutKey:="Q"',
    'Application.MacroOptions Macro:="F", Description:="Adds", Category:=14',
    'Application.MacroOptions Macro:="F", HelpContextID:=7',
    'Application.MacroOptions Macro:="F", Description:="x": Application.MacroOptions Macro:="F", Description:=""',
    'Application.MacroOptions Macro:="F", Description:="d", HelpContextID:=3, Category:=2',
    'Application.MacroOptions Macro:="Module1.F", Description:="d"',
    'Application.MacroOptions Macro:="Zzz", Description:="d"',
    'Application.MacroOptions Macro:="S", HasShortcutKey:=True, ShortcutKey:="q": '
    'Application.MacroOptions Macro:="S", Description:="later"',
    'Application.MacroOptions Macro:="S", HasShortcutKey:=True, ShortcutKey:="q": '
    'Application.MacroOptions Macro:="S", HasShortcutKey:=False',
    'Application.MacroOptions Macro:="F", Category:=5: Application.MacroOptions Macro:="F", Description:="d"',
    'Application.MacroOptions Macro:="S", ShortcutKey:="q"',
    'Application.MacroOptions Macro:="S", HasShortcutKey:=True',
    'Application.MacroOptions Macro:="F", Category:=0',
    'Application.MacroOptions Macro:="F", Category:=40',
    'Application.MacroOptions Macro:="F"',
    'Application.MacroOptions Macro:="F", Description:="say ""hi"""',
    'Application.MacroOptions Macro:="S", HasShortcutKey:=True, ShortcutKey:="ab"',
    'Application.MacroOptions Macro:="F", Description:="d": '
    'wb.VBProject.VBComponents("Module1").CodeModule.InsertLines 1, "Dim top As Long"',
    'Application.MacroOptions Macro:="S", Description:="d": '
    'wb.VBProject.VBComponents("Module1").CodeModule.DeleteLines 1, 3',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             f'    wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString Replace("{MODULE}", "\\n", vbCrLf)',
             '    wb.Activate', '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & wb.VBProject.VBComponents("Module1").CodeModule.CountOfLines',
              'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsm", xlOpenXMLWorkbookMacroEnabled', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description',
              '    Resume Saved', 'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
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
        print(case[24:80], '->', answer)


if __name__ == '__main__':
    main()
