"""Six small members: what each answers in live Excel.

    python scripts/measure_small_rows.py

writes tests/fixtures/small_rows.json, which tests/test_excel_small_rows.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/small_rows.json'
CASES = [
    # CellFormat.NumberFormatLocal, through Application.FindFormat.
    'Application.FindFormat.Clear: r = Application.FindFormat.NumberFormatLocal',
    'Application.FindFormat.Clear: Application.FindFormat.NumberFormatLocal = "0.00": r = Application.FindFormat.NumberFormatLocal & "|" & Application.FindFormat.NumberFormat',
    'Application.FindFormat.Clear: Application.FindFormat.NumberFormat = "0.00": r = Application.FindFormat.NumberFormatLocal',
    'Application.ReplaceFormat.Clear: Application.ReplaceFormat.NumberFormatLocal = "0%": r = Application.ReplaceFormat.NumberFormat',
    'Application.FindFormat.Clear: Application.FindFormat.NumberFormatLocal = "nonsense[": r = Application.FindFormat.NumberFormatLocal',
    'ws.Range("A1").Value = 1.5: ws.Range("A1").NumberFormat = "0.00": Application.FindFormat.Clear: Application.FindFormat.NumberFormatLocal = "0.00": r = ws.Cells.Find("*", SearchFormat:=True).Address',
    # Characters.PhoneticCharacters.
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").Characters.PhoneticCharacters',
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").Characters(1, 2).PhoneticCharacters',
    'ws.Range("A1").Value = "abc": ws.Range("A1").Characters.PhoneticCharacters = "x": r = ws.Range("A1").Characters.PhoneticCharacters & "|" & ws.Range("A1").Value',
    'ws.Range("A1").Value = 5: r = ws.Range("A1").Characters.PhoneticCharacters',
    'r = ws.Range("A1").Characters.PhoneticCharacters',
    # DisplayFormat.Characters.
    'ws.Range("A1").Value = "abc": r = TypeName(ws.Range("A1").DisplayFormat.Characters)',
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").DisplayFormat.Characters.Text',
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").DisplayFormat.Characters(2, 1).Text',
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").DisplayFormat.Characters.Count',
    'ws.Range("A1").Value = "abc": r = ws.Range("A1").DisplayFormat.Characters.Font.Bold',
    'ws.Range("A1").Value = "abc": ws.Range("A1").Font.Bold = True: r = ws.Range("A1").DisplayFormat.Characters.Font.Bold',
    'ws.Range("A1").Value = "abc": ws.Range("A1").Characters(1, 1).Font.Bold = True: r = ws.Range("A1").DisplayFormat.Characters(1, 1).Font.Bold & "|" & ws.Range("A1").DisplayFormat.Characters(2, 1).Font.Bold',
    'ws.Range("A1").Value = "abc": ws.Range("A1").DisplayFormat.Characters.Text = "x": r = ws.Range("A1").Value',
    'ws.Range("A1").Value = 5: r = ws.Range("A1").DisplayFormat.Characters.Text',
    'ws.Range("A1").Value = 5: ws.Range("A1").NumberFormat = "0.00": r = ws.Range("A1").DisplayFormat.Characters.Text',
    # HeaderFooter.Picture.
    'r = TypeName(ws.PageSetup.CenterHeaderPicture)',
    'r = TypeName(ws.PageSetup.Pages(1).CenterHeader.Picture)',
    'r = ws.PageSetup.Pages(1).CenterHeader.Picture.Filename',
    'r = TypeName(ws.PageSetup.EvenPage.LeftFooter.Picture)',
    'r = TypeName(ws.PageSetup.FirstPage.RightHeader.Picture)',
    'r = ws.PageSetup.EvenPage.CenterHeader.Picture.Height',
    # ListColumn.XPath.
    'ws.Range("A1:B3").Value = 1: Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes): r = TypeName(o.ListColumns(1).XPath)',
    'ws.Range("A1:B3").Value = 1: Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes): r = o.ListColumns(1).XPath.Value',
    'ws.Range("A1:B3").Value = 1: Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes): r = o.ListColumns(1).XPath.Repeating',
    'ws.Range("A1:B3").Value = 1: Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes): r = TypeName(o.ListColumns(1).XPath.Map)',
    'ws.Range("A1:B3").Value = 1: Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes): o.ListColumns(1).XPath.Clear: r = "cleared"',
    'r = TypeName(ws.Range("A1").XPath) & "|" & ws.Range("A1").XPath.Value',
    # WorksheetView.DisplayDataTypeIcons.
    'r = ActiveWindow.SheetViews(1).DisplayDataTypeIcons',
    'ActiveWindow.SheetViews(1).DisplayDataTypeIcons = False: r = ActiveWindow.SheetViews(1).DisplayDataTypeIcons',
    'ActiveWindow.SheetViews(1).DisplayDataTypeIcons = True: r = ActiveWindow.SheetViews(1).DisplayDataTypeIcons',
]


def _probe() -> str:
    lines = ['Private Sub Act(ByVal which As Long, ws As Worksheet, r As Variant)', '    Dim o As Object', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '',
              'Public Function Probe(ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, r As Variant, s As String, v As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Act which, ws, r',
              '    s = "ok "', 'Answered:', '    On Error Resume Next', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    Err.Clear',
              '    Application.FindFormat.Clear', '    Application.ReplaceFormat.Clear', '    wb.Close False', '    Probe = s & "|" & v', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index in range(len(CASES)):
            result = excel.run_vba(source, 'Probe', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (CASES[index], result)
            answers.append(str(result.value))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}, indent=1) + '\n', encoding='utf-8')
    for case, answer in zip(CASES, answers, strict=True):
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
