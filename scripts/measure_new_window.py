"""Workbook.NewWindow and Window.NewWindow: a workbook with several windows, on a shown Excel, whose placement is put back after each row.

Rival rules. A window's sheet, selection and view: shared with the workbook's other windows, or each its own from when it was made. Closing one:
the numbers stay, or close up. Saving: one workbookView, or one for each window. Each row runs through a Workbook and through an Object.

    python scripts/measure_new_window.py

writes tests/fixtures/new_window.json, which tests/test_excel_new_window.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAError  # noqa: E402

OUT = ROOT / 'tests/fixtures/new_window.json'
# Each uses wb, a new workbook with sheets A and B; s holds the answer. L(w) reads a window's caption, number and index.
ROWS = {
    'captions': 'wb.NewWindow\ns = wb.Windows.Count & "|" & Application.Windows.Count - before & "|" & L(wb.Windows(1)) & "|" & L(wb.Windows(2)) & "|" & L(ActiveWindow) & "|" & wb.Name',
    'returned': 'Set w = wb.NewWindow\ns = TypeName(w) & "|" & L(w) & "|" & (w Is ActiveWindow) & "|" & TypeName(w.Parent) & "|" & w.Parent.Name',
    'from_window': 'Set w = wb.Windows(1).NewWindow\ns = wb.Windows.Count & "|" & L(w)',
    'three': 'wb.NewWindow\nwb.NewWindow\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2)) & "|" & L(wb.Windows(3))',
    'own_caption': 'wb.Windows(1).Caption = "Mine"\nwb.NewWindow\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2))',
    'caption_set_after': 'wb.NewWindow\nwb.Windows(2).Caption = "Mine"\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2)) & "|" & wb.Name',
    'lookup': 'wb.NewWindow\nn = wb.Name\ns = Application.Windows(n & "  -  1").WindowNumber & "|" & wb.Windows(n & "  -  2").WindowNumber',
    'lookup_plain': 'wb.NewWindow\ns = wb.Windows(wb.Name).WindowNumber',
    'lookup_after_close': 'wb.NewWindow\nwb.Windows(1).Close\ns = wb.Windows(wb.Name).WindowNumber',
    'sheet_copied': 'wb.Sheets("B").Activate\nwb.NewWindow\ns = wb.Windows(1).ActiveSheet.Name & "|" & wb.Windows(2).ActiveSheet.Name',
    'sheet_own': 'wb.NewWindow\nwb.Sheets("B").Activate\ns = wb.Windows(1).ActiveSheet.Name & "|" & wb.Windows(2).ActiveSheet.Name & "|" & wb.ActiveSheet.Name & "|" & L(ActiveWindow)',
    'sheet_after_switch': 'wb.NewWindow\nwb.Sheets("B").Activate\nwb.Windows(2).Activate\ns = wb.ActiveSheet.Name & "|" & L(ActiveWindow) & "|" & wb.Windows(1).ActiveSheet.Name & "|" & wb.Windows(2).ActiveSheet.Name',
    'selection_copied': 'wb.Sheets("A").Range("C3").Select\nwb.NewWindow\ns = wb.Windows(1).RangeSelection.Address & "|" & wb.Windows(2).RangeSelection.Address',
    'selection_own': 'wb.NewWindow\nwb.Sheets("A").Range("D4").Select\ns = wb.Windows(1).RangeSelection.Address & "|" & wb.Windows(2).RangeSelection.Address & "|" & Selection.Address & "|" & ActiveCell.Address',
    'zoom_copied': 'wb.Windows(1).Zoom = 150\nwb.NewWindow\ns = wb.Windows(1).Zoom & "|" & wb.Windows(2).Zoom',
    'zoom_own': 'wb.NewWindow\nwb.Windows(1).Zoom = 80\ns = wb.Windows(1).Zoom & "|" & wb.Windows(2).Zoom',
    'gridlines_own': 'wb.NewWindow\nwb.Windows(1).DisplayGridlines = False\ns = wb.Windows(1).DisplayGridlines & "|" & wb.Windows(2).DisplayGridlines',
    'headings_copied': 'wb.Windows(1).DisplayHeadings = False\nwb.NewWindow\ns = wb.Windows(1).DisplayHeadings & "|" & wb.Windows(2).DisplayHeadings',
    'tabs_own': 'wb.NewWindow\nwb.Windows(1).DisplayWorkbookTabs = False\ns = wb.Windows(1).DisplayWorkbookTabs & "|" & wb.Windows(2).DisplayWorkbookTabs',
    'freeze_copied': 'wb.Sheets("A").Range("B2").Select\nActiveWindow.FreezePanes = True\nwb.NewWindow\ns = wb.Windows(1).FreezePanes & "|" & wb.Windows(2).FreezePanes & "|" & wb.Windows(2).Panes.Count',
    'views_of_sheet': 'wb.NewWindow\ns = wb.Windows(1).SheetViews.Count & "|" & wb.Windows(2).SheetViews.Count',
    'close_front': 'wb.NewWindow\nr = wb.Windows(1).Close\ns = r & "|" & wb.Windows.Count & "|" & L(wb.Windows(1)) & "|" & Workbooks.Count - books',
    'close_back': 'wb.NewWindow\nr = wb.Windows(2).Close\ns = r & "|" & wb.Windows.Count & "|" & L(wb.Windows(1))',
    'close_middle': 'wb.NewWindow\nwb.NewWindow\nwb.Windows(2).Close\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2))',
    'close_number_one': 'wb.NewWindow\nwb.NewWindow\nFor i = 1 To 3\nIf wb.Windows(i).WindowNumber = 1 Then Exit For\nNext\nwb.Windows(i).Close\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2))',
    'number_after_close': 'wb.NewWindow\nwb.NewWindow\nwb.Windows(2).Close\nSet w = wb.NewWindow\ns = L(w) & "|" & L(wb.Windows(2)) & "|" & L(wb.Windows(3))',
    'held_after_close': 'Set w = wb.NewWindow\nwb.Windows(2).Close\nwb.NewWindow\nwb.Windows(1).Close\ns = w.WindowNumber',
    'last_closes_book': 'wb.NewWindow\nwb.Windows(1).Close\nwb.Windows(1).Close\ns = Workbooks.Count - books',
    'workbook_close': 'wb.NewWindow\nwb.Close False\ns = Workbooks.Count - books & "|" & Application.Windows.Count - before',
    'activate_next': 'wb.NewWindow\nr = ActiveWindow.ActivateNext\ns = r & "|" & L(ActiveWindow) & "|" & L(Application.Windows(1))',
    'activate_previous': 'wb.NewWindow\nwb.NewWindow\nr = ActiveWindow.ActivatePrevious\ns = r & "|" & L(ActiveWindow) & "|" & L(wb.Windows(1))',
    'hidden_copied': 'wb.Windows(1).Visible = False\nSet w = wb.NewWindow\ns = wb.Windows.Count & "|" & w.Visible & "|" & L(w)',
    'hidden_other': 'wb.NewWindow\nwb.Windows(2).Visible = False\ns = wb.Windows(1).Visible & "|" & wb.Windows(2).Visible & "|" & L(ActiveWindow)',
    'protected': 'wb.Protect Windows:=True\nwb.NewWindow\ns = wb.Windows.Count',
    'protect_with_two': 'wb.NewWindow\nwb.Protect Windows:=True\ns = wb.Windows.Count & "|" & wb.ProtectWindows',
    'argument': 'wb.NewWindow 1\ns = wb.Windows.Count',
    'rename_by_save': 'wb.NewWindow\nwb.SaveAs folder & "\\Named.xlsx", xlOpenXMLWorkbook\ns = L(wb.Windows(1)) & "|" & L(wb.Windows(2)) & "|" & wb.Windows.Count',
    'saved_reopen': 'wb.Sheets("B").Activate\nwb.NewWindow\nwb.Sheets("A").Activate\nwb.SaveAs folder & "\\Two.xlsx", xlOpenXMLWorkbook\nwb.Close False\n'
                    'Set wb = Workbooks.Open(folder & "\\Two.xlsx", 0)\ns = wb.Windows.Count & "|" & L(wb.Windows(1)) & "|" & wb.Windows(1).ActiveSheet.Name',
    'saved_then_second_reopen': 'wb.NewWindow\nwb.SaveAs folder & "\\Two.xlsx", xlOpenXMLWorkbook\nwb.Close False\nSet wb = Workbooks.Open(folder & "\\Two.xlsx", 0)\n'
                                's = wb.Windows.Count & "|" & L(wb.Windows(wb.Windows.Count)) & "|" & wb.Windows(wb.Windows.Count).ActiveSheet.Name',
    'new_sheet_elsewhere': 'wb.NewWindow\nwb.Sheets.Add\ns = wb.Windows(1).ActiveSheet.Name & "|" & wb.Windows(2).ActiveSheet.Name',
    'sheet_deleted': 'wb.NewWindow\nwb.Sheets("B").Activate\nwb.Sheets("B").Delete\ns = wb.Windows(1).ActiveSheet.Name & "|" & wb.Windows(2).ActiveSheet.Name',
}


def source(binding: str) -> str:
    cases = ''.join(f'    Case "{name}"\n' + ''.join(f'        {line}\n' for line in body.split('\n')) for name, body in ROWS.items())
    return f'''Public Function Row{binding}(wb As {binding}, ByVal name As String, ByVal folder As String, ByVal before As Long, ByVal books As Long) As String
    Dim s As String, n As String, w As Object, r As Variant, i As Long
    On Error GoTo failed
    Select Case name
{cases}    End Select
    Row{binding} = s
    Exit Function
failed:
    Row{binding} = "E" & Err.Number
End Function
'''


DRIVER = '''
Public Function Tried(ByVal name As String, ByVal binding As String, ByVal folder As String) As String
    Dim wb As Workbook, before As Long, books As Long, one As Object, s As String, w As Double, h As Double, x As Double, y As Double, shown As Boolean, state As Long
    Application.DisplayAlerts = False
    shown = Application.Visible
    Application.Visible = True
    Application.ScreenUpdating = True
    state = Application.WindowState
    Application.WindowState = xlNormal
    w = Application.Width: h = Application.Height: x = Application.Left: y = Application.Top
    before = Application.Windows.Count
    books = Workbooks.Count
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "A"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "B"
    wb.Worksheets("A").Activate
    If binding = "Workbook" Then
        s = RowWorkbook(wb, name, folder, before, books)
    Else
        s = RowObject(wb, name, folder, before, books)
    End If
    On Error Resume Next
    For Each one In Workbooks
        If one.Name <> ThisWorkbook.Name Then
            If one.Name Like "Book*" Or one.Name Like "Sheet*" Or one.Name Like "Two*" Or one.Name Like "Named*" Then one.Close False
        End If
    Next
    Application.WindowState = xlNormal
    Application.Width = w: Application.Height = h: Application.Left = x: Application.Top = y
    Application.WindowState = state
    Application.ScreenUpdating = False
    Application.Visible = shown
    Tried = Replace(s, ThisWorkbook.Name, "<HOME>")
End Function
'''
MODEL = '''
Public Function InModel(ByVal name As String, ByVal binding As String, ByVal folder As String) As String
    Dim wb As Workbook, before As Long, books As Long
    Application.DisplayAlerts = False
    before = Application.Windows.Count
    books = Workbooks.Count
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "A"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "B"
    wb.Worksheets("A").Activate
    If binding = "Workbook" Then
        InModel = RowWorkbook(wb, name, folder, before, books)
    Else
        InModel = RowObject(wb, name, folder, before, books)
    End If
End Function
'''


def masked(answer: str) -> str:
    """A new workbook's number, which depends on how many came before."""
    return re.sub(r'\b(Sheet|Book)\d+', 'Book', answer)


def saved(path: Path) -> dict[str, list[str]]:
    with zipfile.ZipFile(path) as package:
        found = {'book': re.findall(r'<workbookView\b[^>]*>', package.read('xl/workbook.xml').decode('utf-8'))}
        for name in sorted(package.namelist()):
            if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name):
                found[name] = re.findall(r'<sheetView\b[^>]*>', package.read(name).decode('utf-8'))
    return found


def main() -> None:
    module = 'Private Function L(w As Object) As String\n    L = w.Caption & "#" & w.WindowNumber & "#" & w.Index\nEnd Function\n\n' + source('Workbook') + source('Object')
    cases: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for name in ROWS:
            for binding in ('Workbook', 'Object'):
                with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                    done = excel.run_vba(module + DRIVER, 'Tried', (name, binding, folder), timeout=40.0, line_numbers=False)
                    answer = masked(str(done.value).replace(folder, '<DIR>')) if done.ok else f'failed {done.outcome} {[one.message[:80] for one in done.dialogs]}'
                    files = {path.name: saved(path) for path in Path(folder).glob('*.xlsx')}
                app = ExcelApplication()
                app.add_workbook()
                app.interpreter.sandbox = False
                app.add_module(module + MODEL, name='Probe')
                with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                    try:
                        model = masked(str(app.run('Probe.InModel', name, binding, folder)).replace(folder, '<DIR>'))
                    except VBAError as failed:
                        model = f'refused {str(failed)[:60]}'
                cases.append({'name': name, 'binding': binding, 'excel': answer, 'files': files})
                print('same' if model == answer else 'DIFF', name, binding, answer[:120], '|', model[:50], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': module + MODEL, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
