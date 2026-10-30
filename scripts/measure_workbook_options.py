"""Workbook settings a file keeps in its workbook part: DisplayDrawingObjects, DisplayInkComments,
RemovePersonalInformation, EnableAutoRecover and ReadOnlyRecommended; what each reads, takes and refuses, whether
setting it leaves the workbook unsaved, and where a save writes it.

Each case runs on a workbook of its own; the saved workbooks' workbook parts are recorded, and each is opened
again, read-only recommendation ignored, to read the setting back.

    python scripts/measure_workbook_options.py

writes tests/fixtures/workbook_options.json, which tests/test_excel_workbook_options.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_options.json'
READ = ('wb.DisplayDrawingObjects & " " & wb.DisplayInkComments & " " & wb.RemovePersonalInformation & " " & '
        'wb.EnableAutoRecover & " " & wb.ReadOnlyRecommended & " " & wb.Saved')
#: Each case: statements on the new workbook wb, and the file it saves, if any.
CASES = [
    ([], ''),
    (['wb.DisplayDrawingObjects = xlPlaceholders'], 'placeholders.xlsx'),
    (['wb.DisplayDrawingObjects = xlHide'], 'hide.xlsx'),
    (['wb.DisplayDrawingObjects = xlHide', 'wb.DisplayDrawingObjects = xlDisplayShapes'], 'shapes.xlsx'),
    (['wb.DisplayDrawingObjects = 0'], ''),
    (['wb.DisplayDrawingObjects = "x"'], ''),
    (['wb.DisplayInkComments = False'], 'ink.xlsx'),
    (['wb.RemovePersonalInformation = True'], 'personal.xlsx'),
    (['wb.EnableAutoRecover = False'], 'recover.xlsx'),
    (['wb.ReadOnlyRecommended = True'], 'recommended.xlsx'),
    (['wb.ReadOnlyRecommended = "x"'], ''),
    (['wb.EnableAutoRecover = Null'], ''),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {READ}')
        if saved:
            lines += [f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook', '        wb.Close False',
                      f'        Set wb = Workbooks.Open(folder & "\\{saved}", IgnoreReadOnlyRecommended:=True)',
                      f'        Tried = Tried & " / " & {READ}']
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description & " / " & ' + READ, '    Resume Done',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "|" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
        parts = {}
        for _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    workbook = package.read('xl/workbook.xml').decode('utf-8')
                parts[saved] = re.findall(r'<(?:fileSharing|workbookPr|fileRecoveryPr|bookViews|workbookView|'
                                          r'calcPr)\b[^>]*>', workbook)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, saved, answer in record['cases']:
        print(statements, '->', answer)
    for name, found in parts.items():
        print(name, found)


if __name__ == '__main__':
    main()
