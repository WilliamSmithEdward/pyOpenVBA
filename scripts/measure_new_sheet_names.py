"""What Worksheets.Add calls a new sheet. Each rival rule was written down first, as what its cases should say:

- a count: Sheet followed by the number of sheets, plus one, or the next free number from there;
- the lowest free number from 1;
- a high-water mark: one past the highest number the workbook has used this session, gone or not;
- a mark that starts again when the file is opened, from the numbers its sheets' names then hold.

    python scripts/measure_new_sheet_names.py

writes tests/fixtures/new_sheet_names.json, which tests/test_excel_new_sheet_names.py replays. A case that opens a file has saved
what the statements before OPEN built, then closed it. Each case runs in a session of its own.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/new_sheet_names.json'
ADD = 'wb.Worksheets.Add'
DROP = 'Application.DisplayAlerts = False: wb.Worksheets("{0}").Delete'
#: Statements in order; "name X" renames the first sheet, "add" adds and records what it was called, "OPEN" saves, closes and reopens.
CASES = (
    # New workbook.
    ('add', 'add', 'add'), ('drop Sheet1', 'add'), ('add', 'add', 'drop Sheet3', 'add'), ('add', 'add', 'drop Sheet2', 'add'), ('add', 'add', 'drop Sheet1', 'add'), ('add', 'drop Sheet2', 'add', 'add'),
    ('name One', 'add'), ('name One', 'add', 'add'), ('name Sheet5', 'add'), ('name Sheet5', 'add', 'add', 'add', 'add', 'add', 'add'), ('name sheet2', 'add'), ('name Sheet2', 'add'), ('add', 'rename Sheet2 Other', 'add'), ('add', 'rename Sheet2 Sheet9', 'add'),
    # Opened from a file.
    ('OPEN', 'add'), ('OPEN', 'add', 'add'), ('add', 'add', 'OPEN', 'add'), ('add', 'add', 'drop Sheet2', 'OPEN', 'add'), ('add', 'add', 'drop Sheet1', 'OPEN', 'add'), ('add', 'add', 'drop Sheet3', 'OPEN', 'add'),
    ('name One', 'OPEN', 'add'), ('name One', 'OPEN', 'add', 'add'), ('name One', 'add', 'rename Sheet2 Two', 'OPEN', 'add'), ('name One', 'add', 'rename Sheet2 Two', 'add', 'rename Sheet2 Three', 'add', 'rename Sheet2 Forms', 'add', 'rename Sheet2 Blank', 'OPEN', 'add'),
    ('name One', 'add', 'rename Sheet2 Two', 'add', 'rename Sheet2 Three', 'add', 'rename Sheet2 Forms', 'add', 'rename Sheet2 Blank', 'OPEN', 'drop Blank', 'add'),
    ('name Sheet5', 'OPEN', 'add'), ('name Sheet5', 'OPEN', 'add', 'add'), ('OPEN', 'drop Sheet1x', 'add'), ('add', 'OPEN', 'drop Sheet1', 'add'), ('add', 'OPEN', 'drop Sheet2', 'add'), ('add', 'add', 'OPEN', 'drop Sheet3', 'add'), ('name Report', 'OPEN', 'add'),
    ('name Sheet1x', 'OPEN', 'add'), ('name sheet3', 'OPEN', 'add'), ('name Sheet03', 'OPEN', 'add'),
    # Other ways of adding.
    ('Sheets.Add', 'Sheets.Add'), ('Add Count 2',), ('OPEN', 'Add Count 2'), ('Add After',), ('copy', 'add'), ('OPEN', 'copy', 'add'), ('chart', 'add'), ('name One', 'OPEN', 'chart', 'add'), ('Add xlWorksheet', 'add'),
    # A workbook opened in a session that has used names already; and two workbooks side by side.
    ('other', 'add'), ('other', 'OPEN', 'add'), ('other add', 'add'),
)


def statement(step: str) -> str:
    words = step.split()
    if step == 'add':
        return f'{ADD}: r = r & " " & ActiveSheet.Name'
    if words[0] == 'name':
        return f'wb.Worksheets(1).Name = "{words[1]}"'
    if words[0] == 'rename':
        return f'wb.Worksheets("{words[1]}").Name = "{words[2]}"'
    if words[0] == 'drop':
        return DROP.format(words[1])
    if step == 'OPEN':
        return 'wb.SaveAs folder & "\\o.xlsx", xlOpenXMLWorkbook: wb.Close False: Set wb = Workbooks.Open(folder & "\\o.xlsx")'
    return {'Sheets.Add': 'wb.Sheets.Add: r = r & " " & ActiveSheet.Name',
            'Add Count 2': 'wb.Worksheets.Add Count:=2: r = r & " " & wb.Worksheets(1).Name & " " & wb.Worksheets(2).Name',
            'Add After': 'wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): r = r & " " & ActiveSheet.Name',
            'copy': 'wb.Worksheets(1).Copy After:=wb.Worksheets(1): r = r & " " & ActiveSheet.Name',
            'chart': 'wb.Charts.Add: r = r & " " & ActiveSheet.Name',
            'Add xlWorksheet': 'wb.Sheets.Add Type:=xlWorksheet: r = r & " " & ActiveSheet.Name',
            'other': 'Set other = Workbooks.Add(xlWBATWorksheet): other.Worksheets.Add: r = r & " " & ActiveSheet.Name: wb.Activate',
            'other add': 'Set other = Workbooks.Add(xlWBATWorksheet): other.Worksheets.Add: wb.Activate: r = r & " " & ActiveSheet.Name'}[step]


def source(steps: tuple[str, ...]) -> str:
    body = '\n'.join(f'    {statement(step)}' for step in steps)
    return f'''Public Function Tried(ByVal folder As String) As String
    Dim wb As Workbook, other As Workbook, r As Variant, s As Object
    Application.DisplayAlerts = False
    r = ""
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo Failed
{body}
    For Each s In wb.Sheets
        r = r & " [" & s.Name & "]"
    Next
    Tried = "ok" & r
    GoTo Finish
Failed:
    Tried = "Err " & Err.Number & r
    Resume Finish
Finish:
    On Error Resume Next
    If Not other Is Nothing Then other.Close False
    wb.Close False
    Application.DisplayAlerts = True
End Function
'''


def main() -> None:
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for steps in CASES:
            for name in Path(folder).iterdir():
                name.unlink()
            with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
                excel.new_document()
                done = excel.run_vba(source(steps), 'Tried', (folder,), timeout=120.0, line_numbers=False)
                answers.append(str(done.value) if done.ok else f'STOPPED {done.outcome}')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': [[list(steps), answer] for steps, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for steps, answer in record['cases']:
        print(' / '.join(steps), '=>', answer)


if __name__ == '__main__':
    main()
