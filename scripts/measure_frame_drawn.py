"""What makes Excel draw a chart in a frame, so that Chart.Export writes a picture of it rather than an empty file.

Each row makes a new workbook with a chart in a frame, or opens one saved with it, does one thing, and exports the chart; the
row records the file's size. Some rows export twice, to see whether drawing lasts, or whether a change undoes it.

Rival rules, written before measuring: the chart is drawn by (1) being made in this session, (2) its sheet being activated,
(3) ChartObject.Activate or Select, (4) Chart.Refresh, (5) any export of another chart, (6) time passing or DoEvents, so that the
state follows the screen and cannot be modelled. Predicted: 3 and 4 draw it, 1, 2 and 5 do not, and 6 does not.

    python scripts/measure_frame_drawn.py
    python scripts/measure_frame_drawn.py first

writes tests/fixtures/frame_drawn.json; the second adds each candidate as the first act of a fresh Excel. Found: drawing is the session's, not the
chart's. A fresh automated Excel exports every frame empty until a ChartObject is activated or selected, and every frame after; selecting a chart's
area draws that chart alone. Refresh, a title, a chart sheet and its export draw nothing. The model does not follow this (see _output).
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/frame_drawn.json'
#: Each row: how its chart is had ('new' or 'opened'), and what it does before the export, in o, the ChartObject.
ROWS = {
    'new_nothing': ('new', ''),
    'new_doevents': ('new', 'DoEvents'),
    'new_wait': ('new', 'Application.Wait Now + TimeSerial(0, 0, 2)'),
    'new_calculate': ('new', 'Application.Calculate'),
    'new_screen_updating': ('new', 'Application.ScreenUpdating = False\n    Application.ScreenUpdating = True'),
    'new_sheet_activate': ('new', 'w.Activate'),
    'new_other_then_back': ('new', 'b.Worksheets.Add.Activate\n    w.Activate'),
    'new_range_select': ('new', 'w.Range("H20").Select'),
    'new_object_activate': ('new', 'o.Activate'),
    'new_object_select': ('new', 'o.Select'),
    'new_chart_activate': ('new', 'o.Chart.Activate'),
    'new_chart_area_select': ('new', 'o.Chart.ChartArea.Select'),
    'new_refresh': ('new', 'o.Chart.Refresh'),
    'new_title': ('new', 'o.Chart.HasTitle = True'),
    'new_chart_sheet_made': ('new', 'b.Charts.Add.SetSourceData w.Range("A1:B5")\n    w.Activate'),
    'new_chart_sheet_exported': ('new', 'b.Charts.Add.SetSourceData w.Range("A1:B5")\n    b.Charts(1).Export folder & "\\sheet.png"\n    w.Activate'),
    'new_other_frame_exported': ('new', 'w.ChartObjects.Add(400, 10, 200, 150).Chart.SetSourceData w.Range("A1:B5")\n    w.ChartObjects(2).Activate\n    w.ChartObjects(2).Chart.Export folder & "\\other.png"\n    w.Range("A1").Select'),
    'new_export_twice': ('new', 'o.Chart.Export folder & "\\first.png"'),
    'new_refresh_then_data': ('new', 'o.Chart.Refresh\n    w.Range("B2").Value = 99'),
    'new_refresh_then_other_sheet': ('new', 'o.Chart.Refresh\n    b.Worksheets.Add.Activate'),
    'new_refresh_then_resize': ('new', 'o.Chart.Refresh\n    o.Width = 300'),
    'new_refresh_then_save': ('new', 'o.Chart.Refresh\n    b.SaveAs folder & "\\again.xlsx", xlOpenXMLWorkbook'),
    'opened_nothing': ('opened', ''),
    'opened_sheet_activate': ('opened', 'w.Activate'),
    'opened_object_activate': ('opened', 'o.Activate'),
    'opened_object_select': ('opened', 'o.Select'),
    'opened_refresh': ('opened', 'o.Chart.Refresh'),
    'opened_chart_activate': ('opened', 'o.Chart.Activate'),
    'opened_on_other_sheet': ('opened', 'b.Worksheets("Two").Activate\n    o.Chart.Refresh'),
    'opened_hidden_sheet_refresh': ('opened', 'w.Visible = xlSheetHidden\n    o.Chart.Refresh'),
}
SOURCE = r'''Public Function Make(ByVal path As String) As String
    Dim b As Workbook, w As Worksheet
    Application.DisplayAlerts = False
    Set b = Workbooks.Add(xlWBATWorksheet)
    Set w = b.Worksheets(1)
    w.Name = "One"
    w.Range("A1:A5").Value = 1
    w.Range("B1:B5").Value = 2
    w.ChartObjects.Add(150, 10, 240, 160).Chart.SetSourceData w.Range("A1:B5")
    b.Worksheets.Add(After:=w).Name = "Two"
    w.Activate
    b.SaveAs path, xlOpenXMLWorkbook
    b.Close False
    Make = "ok"
End Function

Private Function Book(ByVal how As String, ByVal path As String) As Workbook
    Dim b As Workbook, w As Worksheet
    If how = "opened" Then
        Set Book = Workbooks.Open(path, 0)
        Exit Function
    End If
    Set b = Workbooks.Add(xlWBATWorksheet)
    Set w = b.Worksheets(1)
    w.Name = "One"
    w.Range("A1:A5").Value = 1
    w.Range("B1:B5").Value = 2
    w.ChartObjects.Add(150, 10, 240, 160).Chart.SetSourceData w.Range("A1:B5")
    b.Worksheets.Add(After:=w).Name = "Two"
    Set Book = b
End Function
'''


def row_source(how: str, action: str, declared: str) -> str:
    body = ''.join(f'    {line}\n' for line in action.split('\n')) if action else ''
    return SOURCE + f'''
Public Function Row(ByVal path As String, ByVal folder As String) As String
    Dim b As Workbook, w As Worksheet, o As {declared}, e As Long
    Application.DisplayAlerts = False
    Set b = Book("{how}", path)
    Set w = b.Worksheets("One")
    Set o = w.ChartObjects(1)
    On Error Resume Next
{body}    e = Err.Number
    Err.Clear
    o.Chart.Export folder & "\\frame.png"
    Row = "E" & e & "|export=" & Err.Number
    On Error GoTo 0
    b.Close False
End Function
'''


#: Each as the first act of a fresh Excel, then a new chart exported with nothing done to it.
FIRST_ACTS = ['new_object_activate', 'new_object_select', 'new_chart_activate', 'new_chart_area_select', 'new_refresh', 'new_title', 'new_chart_sheet_made',
              'new_chart_sheet_exported', 'opened_object_activate', 'opened_refresh']


def first_acts() -> None:
    """The rows of ``cases`` ran in one session, in order, so a frame activated early drew every one after it. These split that: one fresh Excel each."""
    from pyvbaharness import ExcelSession, HarnessConfig

    record = json.loads(OUT.read_text(encoding='utf-8'))
    found: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        book = home / 'frame.xlsx'
        for number, name in enumerate(FIRST_ACTS):
            for binding, declared in (('Typed', 'ChartObject'), ('Late', 'Object')):
                sizes: list[int] = []
                with ExcelSession(HarnessConfig(exclusive=False)) as excel:
                    excel.new_document()
                    assert excel.run_vba(SOURCE, 'Make', (str(book),), timeout=60.0, line_numbers=False).ok
                    for step, row in enumerate((name, 'new_nothing')):
                        how, action = ROWS[row]
                        folder = home / f'f{number}_{binding}_{step}'
                        folder.mkdir()
                        excel.run_vba(row_source(how, action, declared if step == 0 else 'Object'), 'Row', (str(book), str(folder)), timeout=45.0, line_numbers=False)
                        sizes.append((folder / 'frame.png').stat().st_size if (folder / 'frame.png').exists() else -1)
                found.append({'first': name, 'binding': binding, 'itself_drawn': sizes[0] > 0, 'next_new_chart_drawn': sizes[1] > 0})
                print(name, binding, sizes, flush=True)
    record['first_acts'] = found
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, object] = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': []}
    cases: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        book = home / 'frame.xlsx'
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            made = excel.run_vba(SOURCE, 'Make', (str(book),), timeout=60.0, line_numbers=False)
            assert made.ok, made
            for name, (how, action) in ROWS.items():
                for binding, declared in (('Typed', 'ChartObject'), ('Late', 'Object')):
                    folder = home / f'{name}_{binding}'
                    folder.mkdir()
                    ran = excel.run_vba(row_source(how, action, declared), 'Row', (str(book), str(folder)), timeout=45.0, line_numbers=False)
                    answer = str(ran.value) if ran.ok else f'failed {ran.outcome} {[one.message[:90] for one in ran.dialogs]}'
                    sizes = {one.name: one.stat().st_size for one in sorted(folder.iterdir()) if one.suffix == '.png'}
                    cases.append({'row': name, 'binding': binding, 'how': how, 'action': action, 'excel': answer, 'drawn': sizes.get('frame.png', -1) > 0,
                                  'sizes': sizes})
                    print(name, binding, answer, sizes, flush=True)
    record['cases'] = cases
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    if sys.argv[1:] == ["first"]:
        first_acts()
    else:
        main()
