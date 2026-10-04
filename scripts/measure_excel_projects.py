"""Measure macro lookup and ThisWorkbook across separate native Excel projects."""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/excel_projects.json"
MODULE = '''Public Function Echo() As String
Echo = ThisWorkbook.Name & ":" & ActiveWorkbook.Name
End Function
Private Function PrivateAnswer() As Long
PrivateAnswer = 42
End Function
Public Function CallOther() As String
CallOther = Application.Run("'beta.xlsm'!Echo") & "|" & ThisWorkbook.Name
End Function
Public Function HandledError() As Long
On Error Resume Next
Err.Raise 5
HandledError = Err.Number
End Function
Public Function ErrorAtEntry() As Long
ErrorAtEntry = Err.Number
End Function
Public Function ClearError() As Long
Err.Clear
ClearError = Err.Number
End Function
'''
CLASS = '''Public Function StaticAnswer() As Long
StaticAnswer = 42
End Function
'''
DOCUMENT = '''Public Function DocumentAnswer() As String
Me.Worksheets(1).Range("A1").Value = 99
DocumentAnswer = Me.Name & ":" & ThisWorkbook.Name
End Function
Public Function DocumentArg(ByVal n As Long) As Long
Me.Worksheets(1).Range("A1").Value = n
DocumentArg = n
End Function
'''
ACTIONS = {
    "unqualified_active_alpha": 'alpha.Activate\nanswer = Application.Run("Echo")',
    "unqualified_active_beta": 'beta.Activate\nanswer = Application.Run("Echo")',
    "qualified_alpha": 'beta.Activate\nanswer = Application.Run("\'alpha.xlsm\'!Echo")',
    "qualified_beta": 'alpha.Activate\nanswer = Application.Run("\'beta.xlsm\'!Echo")',
    "qualified_module": 'beta.Activate\nanswer = Application.Run("\'alpha.xlsm\'!Library.Echo")',
    "private_function": 'beta.Activate\nanswer = Application.Run("\'alpha.xlsm\'!PrivateAnswer")',
    "missing_workbook": 'answer = Application.Run("\'absent.xlsm\'!Echo")',
    "missing_macro": 'answer = Application.Run("\'alpha.xlsm\'!Absent")',
    "nested_project_call": 'alpha.Activate\nanswer = Application.Run("\'alpha.xlsm\'!CallOther")',
    "unqualified_caller_active": 'ThisWorkbook.Activate\nanswer = Application.Run("Echo")',
    "unqualified_single_project": 'beta.Close False\nalpha.Activate\nanswer = Application.Run("Echo")',
    "unqualified_caller_project": 'beta.Activate\nanswer = Application.Run("Echo")',
    "qualified_case_insensitive": 'beta.Activate\nanswer = Application.Run("\'ALPHA.XLSM\'!library.eCHO")',
    "callee_clears_error": 'Err.Raise 5\nanswer = Application.Run("\'alpha.xlsm\'!Echo")',
    "callee_handled_error": 'answer = Application.Run("\'alpha.xlsm\'!HandledError")',
    "callee_entry_error": 'Err.Raise 5\nanswer = Application.Run("\'alpha.xlsm\'!ErrorAtEntry")',
    "callee_explicit_clear": 'Err.Raise 5\nanswer = Application.Run("\'alpha.xlsm\'!ClearError")',
    "ordinary_class_macro": 'answer = Application.Run("\'alpha.xlsm\'!Box.StaticAnswer")',
    "workbook_document_macro": 'answer = Application.Run("\'alpha.xlsm\'!ThisWorkbook.DocumentAnswer")',
    "document_macro_side_effect": 'answer = CStr(Application.Run("\'alpha.xlsm\'!ThisWorkbook.DocumentAnswer")) & "#cell:" & CStr(alpha.Worksheets(1).Range("A1").Value)',
    "document_macro_argument": 'answer = CStr(Application.Run("\'alpha.xlsm\'!ThisWorkbook.DocumentArg", 42)) & "#cell:" & CStr(alpha.Worksheets(1).Range("A1").Value)',
}


def main() -> None:
    rows: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="pyopenvba-projects-") as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
            for name, action in ACTIONS.items():
                excel.new_document()
                lines = ['Function Probe() As String', 'Dim alpha As Workbook, beta As Workbook, item As Object',
                         'Dim answer As Variant, number As Long', 'Application.EnableEvents = False',
                         'Application.DisplayAlerts = False']
                for variable in ("alpha", "beta"):
                    lines.extend([f'Set {variable} = Workbooks.Add',
                                  f'Set item = {variable}.VBProject.VBComponents.Add(1)',
                                  'item.Name = "Library"'])
                    source = ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in MODULE.splitlines())
                    lines.append('item.CodeModule.AddFromString ' + source)
                    lines += [f'Set item = {variable}.VBProject.VBComponents.Add(2)', 'item.Name = "Box"']
                    source = ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in CLASS.splitlines())
                    lines.append('item.CodeModule.AddFromString ' + source)
                    lines.append(f'Set item = {variable}.VBProject.VBComponents("ThisWorkbook")')
                    source = ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in DOCUMENT.splitlines())
                    lines.append('item.CodeModule.AddFromString ' + source)
                    lines.append(f'{variable}.SaveAs "{Path(folder) / (variable + ".xlsm")}", 52')
                lines.extend(['On Error Resume Next', action, 'number = Err.Number', 'On Error GoTo 0',
                              'Probe = CStr(answer) & "#error:" & CStr(number)',
                              'On Error Resume Next', 'alpha.Close False', 'beta.Close False', 'End Function'])
                caller_source = ''
                if name == "unqualified_caller_project":
                    caller_source = 'Public Function Echo() As String\nEcho = "caller"\nEnd Function'
                    lines.extend(caller_source.splitlines())
                result = excel.run_vba('\n'.join(lines), 'Probe', timeout=30.0)
                assert result.ok, (name, result)
                rows.append({"name": name, "action": action, "caller_source": caller_source, "result": result.value})
                print(name, result.value, flush=True)
            version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
            assert version.ok, version
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(),
                               "excel_version_build": version.value, "module": MODULE, "class": CLASS,
                               "document": DOCUMENT, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == "__main__":
    main()
