"""Print settings as they read from a file Excel saved: each sheet of one workbook changes one setting, Excel saves it and opens it
again, and every PageSetup property of every sheet is read. The replay opens the same file in the model, and opens the model's own
save of it, so a setting either loses shows as a difference.

    python scripts/measure_page_files.py

writes tests/fixtures/page_files.json, with the workbook in it, which tests/test_excel_page_files.py replays.
"""
from __future__ import annotations

import base64
import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fixture_workbook import without_save_path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/page_files.json'

READ = ['AlignMarginsHeaderFooter', 'BlackAndWhite', 'BottomMargin', 'CenterFooter', 'CenterHeader', 'CenterHorizontally', 'CenterVertically', 'DifferentFirstPageHeaderFooter', 'Draft', 'FirstPageNumber',
        'FitToPagesTall', 'FitToPagesWide', 'FooterMargin', 'HeaderMargin', 'LeftFooter', 'LeftHeader', 'LeftMargin', 'OddAndEvenPagesHeaderFooter', 'Order', 'Orientation', 'PaperSize', 'PrintArea', 'PrintComments',
        'PrintErrors', 'PrintGridlines', 'PrintHeadings', 'PrintNotes', 'PrintTitleColumns', 'PrintTitleRows', 'RightFooter', 'RightHeader', 'RightMargin', 'ScaleWithDocHeaderFooter', 'TopMargin', 'Zoom']
PARTS = ['EvenPage.LeftHeader.Text', 'EvenPage.CenterFooter.Text', 'FirstPage.RightHeader.Text', 'FirstPage.LeftFooter.Text']
CHANGES = [
    '', 'p.AlignMarginsHeaderFooter = False', 'p.BlackAndWhite = True', 'p.BottomMargin = 40.5', 'p.CenterFooter = "cf &P of &N"', 'p.CenterHeader = "&""Arial,Bold""&12ch"', 'p.CenterHorizontally = True', 'p.CenterVertically = True',
    'p.DifferentFirstPageHeaderFooter = True', 'p.DifferentFirstPageHeaderFooter = True: p.FirstPage.RightHeader.Text = "first": p.FirstPage.LeftFooter.Text = "ff"', 'p.Draft = True', 'p.FirstPageNumber = 7', 'p.FirstPageNumber = 0',
    'p.Zoom = False', 'p.Zoom = False: p.FitToPagesWide = 2: p.FitToPagesTall = 3', 'p.Zoom = False: p.FitToPagesWide = False', 'p.Zoom = False: p.FitToPagesTall = False', 'p.FooterMargin = 11.25', 'p.HeaderMargin = 12.75', 'p.LeftFooter = "lf"', 'p.LeftHeader = "lh&D"', 'p.LeftMargin = 33.3',
    'p.OddAndEvenPagesHeaderFooter = True', 'p.OddAndEvenPagesHeaderFooter = True: p.EvenPage.LeftHeader.Text = "even": p.EvenPage.CenterFooter.Text = "ef"', 'p.Order = xlOverThenDown', 'p.Orientation = xlLandscape', 'p.PaperSize = xlPaperA4', 'p.PaperSize = xlPaperLegal',
    'p.PrintArea = "$A$1:$B$2"', 'p.PrintArea = "A1:B2,D1:E3"', 'p.PrintComments = xlPrintSheetEnd', 'p.PrintComments = xlPrintInPlace', 'p.PrintErrors = xlPrintErrorsBlank', 'p.PrintErrors = xlPrintErrorsDash', 'p.PrintErrors = xlPrintErrorsNA', 'p.PrintGridlines = True', 'p.PrintHeadings = True', 'p.PrintNotes = True',
    'p.PrintTitleColumns = "$A:$B"', 'p.PrintTitleRows = "$1:$2"', 'p.PrintTitleColumns = "$C:$C": p.PrintTitleRows = "$3:$3"', 'p.RightFooter = "rf"', 'p.RightHeader = "rh&&x"', 'p.RightMargin = 21.5', 'p.ScaleWithDocHeaderFooter = False', 'p.TopMargin = 60.125', 'p.Zoom = 75', 'p.Zoom = 400',
]


def _source(statements: list[str]) -> str:
    changes = ''.join(f'    Case {index}\n        {statement}\n' for index, statement in enumerate(statements))
    return ('Private Sub Change(ByVal which As Long, p As PageSetup)\n    Select Case which\n' + changes + '    End Select\nEnd Sub\n'
            'Public Sub Made(ByVal path As String, ByVal count As Long)\n    Dim wb As Workbook, i As Long\n    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
            '    For i = 0 To count - 1\n        If i > 0 Then wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)\n        wb.Worksheets(i + 1).Name = "S" & i\n        wb.Worksheets(i + 1).Range("A1").Value = 1\n        Change i, wb.Worksheets(i + 1).PageSetup\n    Next\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\nEnd Sub\n')


READER = '''Private Function One(o As Object, ByVal member As String) As String
    On Error GoTo Failed
    Dim v As Variant
    v = CallByName(o, member, VbGet)
    One = TypeName(v) & ":" & CStr(v)
    Exit Function
Failed:
    One = "E" & Err.Number
End Function

Private Function Part(p As Object, ByVal family As String, ByVal place As String) As String
    On Error GoTo Failed
    Part = CStr(CallByName(CallByName(CallByName(p, family, VbGet), place, VbGet), "Text", VbGet))
    Exit Function
Failed:
    Part = "E" & Err.Number
End Function

Public Function ReadAll(ByVal wb As Object, ByVal members As String) As String
    Dim ws As Object, p As Object, s As String, m As Variant
    For Each ws In wb.Worksheets
        Set p = ws.PageSetup
        s = s & "@@@" & ws.Name
        For Each m In Split(members, ",")
            s = s & "~" & One(p, m)
        Next
        s = s & "~" & Part(p, "EvenPage", "LeftHeader") & "~" & Part(p, "EvenPage", "CenterFooter") & "~" & Part(p, "FirstPage", "RightHeader") & "~" & Part(p, "FirstPage", "LeftFooter")
    Next
    ReadAll = Mid$(s, 4)
End Function

Public Function Opened(ByVal path As String, ByVal members As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Opened = ReadAll(wb, members)
    wb.Close False
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        path = Path(folder) / 'page_files.xlsx'
        excel.new_document()
        made = excel.run_vba(_source(CHANGES), 'Made', (str(path), len(CHANGES)), timeout=170.0)
        assert made.ok, made
        result = excel.run_vba(READER, 'Opened', (str(path), ','.join(READ)), timeout=120.0)
        assert result.ok, result
        data = without_save_path(path.read_bytes())
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'read': READ, 'changes': CHANGES, 'reader': READER, 'answer': str(result.value), 'file': base64.b64encode(data).decode('ascii')}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(len(CHANGES), len(data), flush=True)


if __name__ == '__main__':
    main()
