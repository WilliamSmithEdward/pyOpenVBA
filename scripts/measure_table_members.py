"""ListObject members the model lacked: SourceType, Comment, AlternativeText, Summary, ShowAutoFilterDropDown,
InsertRowRange, Active, DisplayRightToLeft, Slicers, Application, and the ones a table of the sheet's own cells has
no use for (Refresh, QueryTable, Unlink, SharePointURL, XmlMap, TableObject).

Each case makes a new workbook with a table over A1:C4, runs its statements and reads its expression; a case that
saves records the saved table part, and reads the expression again from the workbook opened anew.

    python scripts/measure_table_members.py

writes tests/fixtures/table_members.json, which tests/test_excel_table_members.py replays.
"""
from __future__ import annotations

import json
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/table_members.json'
PLAIN = ('t.SourceType & "/" & t.Comment & "/" & t.AlternativeText & "/" & t.Summary & "/" & '
         't.ShowAutoFilter & "/" & t.ShowAutoFilterDropDown & "/" & TypeName(t.InsertRowRange) & "/" & '
         't.DisplayRightToLeft & "/" & t.Slicers.Count & "/" & TypeName(t.Application) & "/" & TypeName(t.Slicers)')
#: Each case: statements on the new workbook wb, its sheet ws and its table t; the expression read; the file saved.
CASES = [
    ([], PLAIN, ''),
    (['ws.Range("E1").Select'], 't.Active', ''),
    (['ws.Range("B2").Select'], 't.Active', ''),
    (['ws.Range("A1:E1").Select'], 't.Active', ''),
    (['t.Comment = "Note here"'], PLAIN, 'comment.xlsx'),
    (['t.AlternativeText = "Alt words"', 't.Summary = "Longer summary"'], PLAIN, 'alt.xlsx'),
    (['t.AlternativeText = "Alt only"'], PLAIN, 'alt_only.xlsx'),
    (['t.ShowAutoFilterDropDown = False'], PLAIN, 'dropdown.xlsx'),
    (['t.ShowAutoFilterDropDown = False', 't.ShowAutoFilterDropDown = True'], PLAIN, 'dropdown_back.xlsx'),
    (['t.ShowAutoFilter = False'], PLAIN, ''),
    (['t.ShowAutoFilter = False', 't.ShowAutoFilterDropDown = True'], PLAIN, ''),
    (['t.ShowAutoFilterDropDown = False', 't.ShowAutoFilter = False', 't.ShowAutoFilter = True'], PLAIN, ''),
    (['t.ShowAutoFilterDropDown = False', 't.Range.AutoFilter 1, "a"'], PLAIN, 'dropdown_filtered.xlsx'),
    (['t.Comment = "x"', 't.Comment = ""'], PLAIN, 'comment_cleared.xlsx'),
    (['t.Comment = String(300, "c")'], 'Len(t.Comment)', ''),
    (['t.Comment = 12'], 'TypeName(t.Comment) & "/" & t.Comment', ''),
    (['t.Delete', 'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C1"), , xlYes)'],
     'TypeName(t.InsertRowRange) & "/" & t.InsertRowRange.Address & "/" & t.ListRows.Count & "/" & '
     't.Range.Address', ''),
    (['t.Delete', 'Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)'],
     'TypeName(t.InsertRowRange) & "/" & t.ListRows.Count', ''),
    (['t.Refresh'], '"done"', ''),
    ([], 'TypeName(t.QueryTable)', ''),
    (['t.Unlink'], '"done"', ''),
    ([], 't.SharePointURL', ''),
    ([], 'TypeName(t.XmlMap)', ''),
    ([], 'TypeName(t.TableObject)', ''),
    (['ws.DisplayRightToLeft = True'], 't.DisplayRightToLeft', ''),
    (['t.AlternativeText = 7'], 'TypeName(t.AlternativeText) & "/" & t.AlternativeText', ''),
    (['t.Comment = String(255, "c")'], 'Len(t.Comment)', ''),
    (['t.Comment = String(256, "c")'], 'Len(t.Comment)', ''),
    (['t.AlternativeText = String(300, "a")', 't.Summary = String(300, "s")'],
     'Len(t.AlternativeText) & "/" & Len(t.Summary)', ''),
    (['ws.Range("A1:E1").Select', 'ws.Range("E1").Activate'], 'ActiveCell.Address & "/" & t.Active', ''),
    (['ws.Range("E1:F2").Select', 'ws.Range("E1").Activate'], 't.Active', ''),
    (['ws.Range("B2").Select', 'wb.Worksheets.Add'], 't.Active', ''),
    (['t.Comment = "a" & vbLf & "b"'], 'Len(t.Comment)', 'comment_line.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, t As ListObject', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"', '    ws.Activate',
             '    ws.Range("A1:C1").Value = Array("Name", "Qty", "Price")',
             '    ws.Range("A2:C2").Value = Array("a", 1, 2.5)', '    ws.Range("A3:C3").Value = Array("b", 2, 3.5)',
             '    ws.Range("A4:C4").Value = Array("c", 3, 4.5)',
             '    Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {read}')
        if saved:
            lines += [f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook', '        wb.Close False',
                      f'        Set wb = Workbooks.Open(folder & "\\{saved}")', '        Set ws = wb.Worksheets(1)',
                      '        Set t = ws.ListObjects(1)', f'        Tried = Tried & " @ " & {read}']
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = Tried & "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    parts[saved] = package.read('xl/tables/table1.xml').decode('utf-8')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, _, saved, answer in record['cases']:
        print(statements[-1:] or '-', '->', answer)
    for name, part in parts.items():
        print(name, part[part.index('<table '):])


if __name__ == '__main__':
    main()
