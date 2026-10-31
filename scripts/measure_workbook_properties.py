"""Measure Workbook properties a macro reads about the file and its sharing: each read on a new workbook, on one
saved as .xlsx and reopened, and each that can be set, set and saved, with workbook.xml's workbookPr and
fileSharing recorded.

    python scripts/measure_workbook_properties.py

writes tests/fixtures/workbook_properties.json, which tests/test_excel_workbook_properties.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_properties.json'
READ = ['AutoSaveOn', 'AutoUpdateFrequency', 'AutoUpdateSaveChanges', 'ChangeHistoryDuration', 'CompatibilityVersion',
        'ConflictResolution', 'ConnectionsDisabled', 'DoNotPromptForConvert', 'EncryptionProvider', 'EnvelopeVisible',
        'Excel8CompatibilityMode', 'HighlightChangesOnScreen', 'IsAddin', 'IsInplace', 'KeepChangeHistory',
        'ListChangesOnNewSheet', 'MultiUserEditing', 'Password', 'PersonalViewListSettings',
        'PersonalViewPrintSettings', 'RevisionNumber', 'ShowConflictHistory', 'TemplateRemoveExtData', 'VBASigned',
        'WritePassword', 'WriteReserved', 'UserStatus']
#: Each case: statements on wb (a new workbook) and wbo (it, late bound); the file saved, if any.
CASES: list[tuple[list[str], str]] = [([], '')]
CASES.append((['wb.SaveAs folder & "\\plain.xlsx", xlOpenXMLWorkbook', 'wb.Close',
               'Set wb = Workbooks.Open(folder & "\\plain.xlsx")', 'Set wbo = wb'], ''))
for name, value in (('PersonalViewListSettings', 'False'), ('PersonalViewPrintSettings', 'False'),
                    ('TemplateRemoveExtData', 'True'), ('DoNotPromptForConvert', 'True'), ('EnvelopeVisible', 'True'),
                    ('ChangeHistoryDuration', '30'), ('KeepChangeHistory', 'False'), ('ConflictResolution', '2'),
                    ('HighlightChangesOnScreen', 'True'), ('ListChangesOnNewSheet', 'True'),
                    ('ShowConflictHistory', 'True'), ('AutoUpdateFrequency', '5'), ('AutoUpdateSaveChanges', 'False'),
                    ('IsAddin', 'True'), ('AutoSaveOn', 'True'), ('WritePassword', '"w"'), ('Password', '"p"')):
    CASES.append(([f'wbo.{name} = {value}'], f'set_{name}.xlsx'))


def _probe() -> str:
    lines = ['Private Function Prop(ByVal o As Object, ByVal what As String) As String',
             '    Dim v As Variant', '    On Error Resume Next', '    v = CallByName(o, what, VbGet)',
             '    If Err.Number <> 0 Then',
             '        Prop = "E" & Err.Number', '    ElseIf IsArray(v) Then',
             '        Prop = "Array" & LBound(v) & "-" & UBound(v)', '    Else',
             '        Prop = TypeName(v) & ":" & CStr(v)', '    End If', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, wbo As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set wbo = wb',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    Dim s As String']
    lines += [f'    s = s & "|{name}=" & Prop(wbo, "{name}")' for name in READ]
    lines += ['    Tried = Mid(s, 2)', 'Done:', '    wb.Close False', '    Exit Function',
              'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function',
              '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _parts(path: Path) -> dict[str, str]:
    if not zipfile.is_zipfile(path):
        # A password to open makes an encrypted compound file, not a package.
        return {'encrypted': path.read_bytes()[:8].hex()}
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    return {name: (re.search(rf'<{name}\b[^>]*/?>', workbook) or [''])[0]
            for name in ('fileSharing', 'workbookPr', 'bookViews')}


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        answers: list[str] = []
        # Case by case: one setting hangs Excel, and the others should still be measured.
        for index in range(len(CASES)):
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                result = excel.run_vba(source, 'Probe', (folder, index, index), timeout=60.0)
            answers.append(str(result.value) if result.ok and not result.dialogs else f'Excel {result.outcome}')
        parts = {saved: _parts(Path(folder) / saved) for _, saved in CASES if saved and (Path(folder) / saved).exists()}
        parts['plain.xlsx'] = _parts(Path(folder) / 'plain.xlsx')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    base = dict(item.split('=', 1) for item in record['cases'][0][2].split('|'))
    print('new:', record['cases'][0][2])
    for statements, saved, answer in record['cases'][1:]:
        if answer.startswith(('Err', 'Excel ')):
            print(statements[0], '->', answer)
            continue
        changed = {k: v for k, v in (item.split('=', 1) for item in answer.split('|')) if base.get(k) != v}
        print(statements[-1][:45], '->', changed, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
