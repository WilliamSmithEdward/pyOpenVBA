"""Members of Application that macros use and that need no screen: Undo, Repeat, OnUndo, OnRepeat, OnKey, CheckAbort, and settings
read but never set. What each answers or raises.

It leaves out SendKeys, InputBox, the file and other dialogs, which would act on this machine or wait for someone, and sets
no setting that outlasts the session. A key OnKey assigns is given back by the same statement. A number is recorded in place
of anything that names this machine or its user.

    python scripts/measure_application_undo.py

writes tests/fixtures/application_undo.json, which tests/test_excel_application_undo.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/application_undo.json'
A = 'app'
CASES = [
    # Undoing.
    f'{A}.Undo', f'r = {A}.Undo', f'ws.Range("A1").Value = 1: {A}.Undo: r = ws.Range("A1").Value', f'{A}.Repeat', f'r = {A}.Repeat', f'ws.Range("A1").Value = 1: {A}.Repeat: r = ws.Range("A2").Value',
    f'{A}.OnUndo "Undo it", "Nothing"', f'r = {A}.OnUndo("Undo it", "Nothing")', f'{A}.OnUndo "Undo it", "Nothing": {A}.Undo', f'{A}.OnUndo "Undo it"', f'{A}.OnUndo', f'{A}.OnUndo "", ""', f'{A}.OnUndo 5, 6',
    f'{A}.OnRepeat "Again", "Nothing"', f'r = {A}.OnRepeat("Again", "Nothing")', f'{A}.OnRepeat "Again"', f'{A}.OnRepeat', f'{A}.OnRepeat "", ""',
    # Keys, each given back.
    f'{A}.OnKey "^+{{F12}}", "": {A}.OnKey "^+{{F12}}"', f'r = {A}.OnKey("^+{{F12}}", ""): {A}.OnKey "^+{{F12}}"', f'r = {A}.OnKey("^+{{F12}}")', f'{A}.OnKey "^+{{F12}}", "Nothing": {A}.OnKey "^+{{F12}}"', f'{A}.OnKey "^+z", "": {A}.OnKey "^+z"', f'{A}.OnKey "%+{{F11}}", "": {A}.OnKey "%+{{F11}}"', f'{A}.OnKey "{{NOSUCHKEY}}", ""', f'{A}.OnKey "", ""', f'{A}.OnKey "^+{{F12}}{{F11}}", ""', f'{A}.OnKey', f'{A}.OnKey 5', f'{A}.OnKey Key:="^+{{F12}}", Procedure:="": {A}.OnKey Key:="^+{{F12}}"',
    # Macros' own stopping.
    f'r = {A}.CheckAbort', f'{A}.CheckAbort', f'r = {A}.CheckAbort(Nothing)', f'r = {A}.CheckAbort(ws.Range("A1"))', f'r = {A}.CheckAbort(5)',
    # Read, never set.
    f'r = {A}.WindowState', f'r = {A}.IgnoreRemoteRequests', f'r = {A}.AutomationSecurity', f'r = {A}.MultiThreadedCalculation.Enabled', f'r = TypeName({A}.MultiThreadedCalculation)', f'r = {A}.MultiThreadedCalculation.ThreadMode', f'r = {A}.AlwaysUseClearType', f'r = {A}.WindowsForPens', f'r = TypeName({A}.AddIns)', f'r = {A}.AddIns.Count >= 0', f'r = TypeName({A}.AddIns2)', f'r = TypeName({A}.RecentFiles)', f'r = {A}.RecentFiles.Maximum >= 0', f'r = TypeName({A}.AutoCorrect)', f'r = {A}.MaxSupportedCompatibilityVersion', f'r = TypeName({A}.PreviousSelections)', f'r = IsEmpty({A}.PreviousSelections)', f'r = {A}.DisplayXMLSourcePane', f'r = {A}.MailSystem', f'r = {A}.TransitionMenuKey', f'r = {A}.TransitionMenuKeyAction', f'r = TypeName({A}.Watches)', f'r = {A}.Watches.Count', f'r = TypeName({A}.Speech)', f'r = TypeName({A}.SpellingOptions)', f'r = TypeName({A}.DefaultWebOptions)', f'r = TypeName({A}.DefaultPivotTableLayoutOptions)', f'r = TypeName({A}.LanguageSettings)', f'r = TypeName({A}.AutoRecover)', f'r = {A}.AutoRecover.Enabled', f'r = TypeName({A}.UsedObjects)', f'r = {A}.UsedObjects.Count > 0', f'r = TypeName({A}.ClipboardFormats)', f'r = TypeName({A}.ODBCErrors)', f'r = {A}.ODBCErrors.Count', f'r = TypeName({A}.OLEDBErrors)', f'r = {A}.OLEDBErrors.Count', f'r = TypeName({A}.Excel4MacroSheets)', f'r = {A}.Excel4MacroSheets.Count', f'r = {A}.Excel4IntlMacroSheets.Count', f'r = TypeName({A}.FileConverters)', f'r = TypeName({A}.RegisteredFunctions)', f'r = TypeName({A}.ProtectedViewWindows)', f'r = {A}.ProtectedViewWindows.Count', f'r = TypeName({A}.ActiveProtectedViewWindow)', f'r = TypeName({A}.COMAddIns)', f'r = TypeName({A}.CommandBars)', f'r = TypeName({A}.VBE)', f'r = {A}.Hwnd > 0', f'r = {A}.Hinstance <> 0', f'r = TypeName({A}.Hwnd)', f'r = TypeName({A}.Height)', f'r = TypeName({A}.Width)', f'r = TypeName({A}.UsableHeight)', f'r = {A}.Height > 0', f'r = {A}.UsableWidth > 0',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Object, app As Object, s As String, v As String, r As Variant', '    Set app = Application', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v',
              '    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    """Each case alone, since one that breaks the run takes its session with it: it is recorded as what the run said."""
    source = _probe()
    answers: list[str] = []
    index = 0
    while index < len(CASES):
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            while index < len(CASES):
                result = excel.run_vba(source, 'Probe', (index, index), timeout=40.0)
                index += 1
                if not result.ok or result.dialogs:
                    answers.append(f'breaks: {result.outcome}')
                    break
                answers.append(str(result.value))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
