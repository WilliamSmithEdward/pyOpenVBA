"""Workbook.RemoveDocumentInformation: what the workbook's document properties read after each kind of removal, the
errors it raises, and the properties and workbook settings the saved file holds.

Each case makes a workbook, gives it a title, subject, author, keywords, comments, manager and company and a custom
property, runs the case's statements and reads the properties back, then saves it. The saved file's core.xml and
app.xml are recorded with their times cut out, and workbook.xml's workbookPr element.

    python scripts/measure_remove_document_information.py

writes tests/fixtures/remove_document_information.json, which tests/test_excel_remove_document_information.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/remove_document_information.json'
NAMES = ('Title', 'Subject', 'Author', 'Keywords', 'Comments', 'Manager', 'Company', 'Last author')
READ = ' & "|" & '.join(f'P(wb, "{name}")' for name in NAMES) + ' & "|" & wb.CustomDocumentProperties.Count'
CASES = [
    '',
    'wb.RemoveDocumentInformation xlRDIDocumentProperties',
    'wb.RemoveDocumentInformation xlRDIAll',
    'wb.RemoveDocumentInformation xlRDIComments',
    'wb.RemoveDocumentInformation xlRDIRemovePersonalInformation',
    'wb.RemoveDocumentInformation xlRDIDefinedNameComments',
    'wb.RemoveDocumentInformation 0',
    'wb.RemoveDocumentInformation 1234',
    'wb.RemoveDocumentInformation xlRDIDocumentProperties: wb.BuiltinDocumentProperties("Title") = "again"',
    'wb.RemoveDocumentInformation xlRDIRemovePersonalInformation: s = s & "|" & wb.RemovePersonalInformation',
]


def _probe() -> str:
    lines = ['Private Function P(ByVal wb As Workbook, ByVal name As String) As String',
             '    On Error Resume Next', '    P = CStr(wb.BuiltinDocumentProperties(name).Value)',
             '    If Err.Number <> 0 Then P = "E" & Err.Number', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1").Value = 1']
    for name, value in (('Title', 'T'), ('Subject', 'S'), ('Author', 'Ann'), ('Keywords', 'K'),
                        ('Comments', 'C'), ('Manager', 'M'), ('Company', 'Co')):
        lines.append(f'    wb.BuiltinDocumentProperties("{name}") = "{value}"')
    lines += ['    wb.CustomDocumentProperties.Add "Kept", False, msoPropertyTypeString, "v"',
              '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {case}'] if case else [])
    lines += ['    End Select', f'    Tried = "ok" & s & "|" & {READ}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def saved(path: Path) -> list[str]:
    """core.xml and app.xml with their times cut out, and workbook.xml's workbookPr."""
    with zipfile.ZipFile(path) as package:
        core = package.read('docProps/core.xml').decode('utf-8')
        app = package.read('docProps/app.xml').decode('utf-8')
        custom = package.read('docProps/custom.xml').decode('utf-8') if 'docProps/custom.xml' in package.namelist() \
            else ''
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    core = re.sub(r'(<dcterms:(?:created|modified)\b[^>]*>)[^<]*', r'\1TIME', core)
    found = re.search(r'<workbookPr\b[^>]*/>', workbook)
    return [core, app, custom, found.group() if found else '']


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        files = [saved(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, *found] for case, answer, found in zip(CASES, answers, files, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, core, app, custom, workbook in record['cases']:
        print(case, '->', answer)
        print('    ', workbook, '| custom' if custom else '| no custom')
        print('    ', re.sub(r'<\?xml[^>]*>|xmlns:\w+="[^"]*"', '', core)[:400])


if __name__ == '__main__':
    main()
