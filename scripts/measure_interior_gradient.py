"""Measure Interior.Gradient: linear and rectangular gradient fills set as a recorded macro sets them, what the
gradient, its colour stops and the Interior read, and the fill a save writes.

Each case runs on a new workbook, works on A1 and answers as text; some save the workbook and styles.xml's fills
are recorded.

    python scripts/measure_interior_gradient.py

writes tests/fixtures/interior_gradient.json, which tests/test_excel_interior_gradient.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/interior_gradient.json'
LINEAR = ['i.Pattern = xlPatternLinearGradient']
RECORDED = LINEAR + ['i.Gradient.Degree = 90', 'i.Gradient.ColorStops.Clear',
                     'Set st = i.Gradient.ColorStops.Add(0)', 'st.ThemeColor = xlThemeColorDark1',
                     'st.TintAndShade = 0', 'Set st = i.Gradient.ColorStops.Add(1)',
                     'st.ThemeColor = xlThemeColorAccent1', 'st.TintAndShade = 0']
RECT = ['i.Pattern = xlPatternRectangularGradient']
#: Each case: statements with i the Interior of A1 on a new workbook's sheet ws; the file saved, if any.
CASES: list[tuple[list[str], str]] = [
    ([], ''),
    (LINEAR, 'linear.xlsx'),
    (RECORDED, 'recorded.xlsx'),
    (RECT, 'rect.xlsx'),
    (RECT + ['i.Gradient.RectangleLeft = 0.5', 'i.Gradient.RectangleRight = 0.5', 'i.Gradient.RectangleTop = 0.5',
             'i.Gradient.RectangleBottom = 0.5', 'i.Gradient.ColorStops.Clear',
             'i.Gradient.ColorStops.Add(0).Color = RGB(255, 0, 0)',
             'i.Gradient.ColorStops.Add(1).Color = RGB(0, 0, 255)'], 'rect_centre.xlsx'),
    (['i.Color = RGB(255, 200, 0)'] + LINEAR, 'coloured_first.xlsx'),
    (LINEAR + ['i.Gradient.Degree = 45'], 'degree45.xlsx'),
    (LINEAR + ['i.Gradient.Degree = 400'], ''),
    (LINEAR + ['i.Gradient.Degree = -90'], ''),
    (LINEAR + ['i.Gradient.Degree = 12.5'], 'degree_fraction.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops.Add(0.5).Color = RGB(0, 255, 0)'], 'three.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops.Clear', 'i.Gradient.ColorStops.Add(1).Color = RGB(0, 0, 255)',
               'i.Gradient.ColorStops.Add(0).Color = RGB(255, 0, 0)'], 'reversed.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops.Clear'], 'cleared.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops.Clear', 'i.Gradient.ColorStops.Add(0.3).Color = RGB(1, 2, 3)'],
     'one_stop.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops.Add(1.5)'], ''),
    (LINEAR + ['i.Gradient.ColorStops.Add(-0.1)'], ''),
    (LINEAR + ['i.Gradient.ColorStops.Add(0).Color = RGB(9, 9, 9)'], 'duplicate.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops(1).Delete'], ''),
    (LINEAR + ['i.Gradient.ColorStops(2).Position = 0.25'], 'moved.xlsx'),
    (LINEAR + ['i.Gradient.ColorStops(1).TintAndShade = 0.5'], ''),
    (LINEAR + ['i.Gradient.ColorStops(1).ThemeColor = xlThemeColorAccent2',
               'i.Gradient.ColorStops(1).TintAndShade = -0.25'], 'theme_tint.xlsx'),
    (LINEAR + ['i.Pattern = xlSolid'], 'back_solid.xlsx'),
    (LINEAR + ['i.Color = RGB(10, 20, 30)'], 'colour_after.xlsx'),
    (LINEAR + ['i.ColorIndex = xlNone'], ''),
    (['ws.Range("A1:B2").Interior.Pattern = xlPatternLinearGradient', 'Set i = ws.Range("A1:B2").Interior'],
     'area.xlsx'),
    (RECT + ['i.Gradient.RectangleLeft = 2'], ''),
    (RECT + ['i.Gradient.Degree = 30'], ''),
    (LINEAR + ['i.Gradient.RectangleLeft = 0.2'], ''),
    (LINEAR + ['ws.Range("A1").Copy ws.Range("C3")', 'Set i = ws.Range("C3").Interior'], ''),
    (LINEAR + ['i.Gradient.ColorStops(1).Color = RGB(1, 2, 3)', 'i.Gradient.ColorStops(2).ThemeColor = 5'],
     'mixed.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Prop(ByVal o As Object, ByVal what As String) As String',
             '    Dim v As Variant', '    On Error Resume Next', '    v = CallByName(o, what, VbGet)',
             '    If Err.Number <> 0 Then Prop = "E" & Err.Number Else Prop = TypeName(v) & ":" & CStr(v)',
             'End Function', '',
             'Private Function Described(ByVal i As Interior) As String',
             '    Dim s As String, g As Object, st As Object, p As Variant',
             '    For Each p In Array("Pattern", "Color", "ColorIndex", "ThemeColor", "TintAndShade", "PatternColor")',
             '        s = s & " " & p & "=" & Prop(i, CStr(p))', '    Next', '    On Error Resume Next',
             '    Set g = i.Gradient', '    If Err.Number <> 0 Then Described = s & " G=E" & Err.Number: Exit Function',
             '    On Error GoTo 0', '    If g Is Nothing Then Described = s & " G=Nothing": Exit Function',
             '    s = s & " G=" & TypeName(g)',
             '    For Each p In Array("Degree", "RectangleLeft", "RectangleRight", "RectangleTop", "RectangleBottom")',
             '        s = s & " " & p & "=" & Prop(g, CStr(p))', '    Next',
             '    s = s & " stops=" & g.ColorStops.Count',
             '    For Each st In g.ColorStops',
             '        s = s & " [" & Prop(st, "Position") & " " & Prop(st, "Color") & " " & Prop(st, "ThemeColor") & '
             '" " & Prop(st, "TintAndShade") & "]"', '    Next', '    Described = s', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, i As Interior, st As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set i = ws.Range("A1").Interior', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    Tried = Mid(Described(i), 2)', 'Done:', '    wb.Close False', '    Exit Function',
              'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function',
              '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _fills(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        styles = package.read('xl/styles.xml').decode('utf-8')
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    return {'fills': (re.search(r'<fills\b.*?</fills>', styles, re.DOTALL) or [''])[0],
            'cellXfs': (re.search(r'<cellXfs\b.*?</cellXfs>', styles, re.DOTALL) or [''])[0],
            'sheetData': (re.search(r'<sheetData\b.*?(?:</sheetData>|/>)', sheet, re.DOTALL) or [''])[0]}


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        fills = {saved: _fills(Path(folder) / saved) for _, saved in CASES if saved and (Path(folder) / saved).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'fills': fills}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statements, saved, answer) in enumerate(record['cases']):
        print(index, (statements or [''])[-1][:50], '->', answer)
        if saved in fills:
            print('    ', re.findall(r'<fill>.*?</fill>', fills[saved]['fills'])[2:])


if __name__ == '__main__':
    main()
