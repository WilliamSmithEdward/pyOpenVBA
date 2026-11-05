"""The chart Font settings tests/fixtures/chart_fonts.json left unmeasured: what ChartArea.Font does to the other elements and they to
it, every value Size, FontStyle, Underline, Strikethrough, ColorIndex, Background, Subscript, Superscript, ThemeColor, TintAndShade
and ThemeFont take, on a chart ChartObjects.Add made and one AddChart2 made.

Each case has ch, a titled three-series column chart with a legend and s, its first series, with data labels. It records r, each
element's font as Say reads it, and the chart part Excel saves.

    python scripts/measure_chart_font_gaps.py

writes tests/fixtures/chart_font_gaps.json, which tests/test_excel_chart_font_gaps.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_font_gaps.json'
ELEMENTS = {'title': 'ch.ChartTitle.Font', 'legend': 'ch.Legend.Font', 'ticks': 'ch.Axes(2).TickLabels.Font', 'labels': 's.DataLabels.Font', 'area': 'ch.ChartArea.Font', 'axistitle': 'ch.Axes(2).AxisTitle.Font'}
READS = ('Name', 'Size', 'Bold', 'Italic', 'Underline', 'Strikethrough', 'Color', 'ColorIndex', 'FontStyle', 'Subscript', 'Superscript', 'Background', 'ThemeColor', 'TintAndShade', 'ThemeFont')
VALUES: dict[str, tuple[str, ...]] = {
    'Size': ('0.5', '1', '1.5', '7.25', '8.3', '10.7', '409', '409.5', '410', '499', '500', '0', '-1', '"x"', '"12"'),
    'FontStyle': ('"Bold"', '"bold"', '"Italic"', '"Bold Italic"', '"Regular"', '"Oblique"', '"Bold Oblique"', '"Normal"', '""', '"x"', '12'),
    'Underline': ('xlUnderlineStyleNone', 'xlUnderlineStyleSingle', 'xlUnderlineStyleDouble', 'xlUnderlineStyleSingleAccounting', 'xlUnderlineStyleDoubleAccounting', 'True', 'False', '0', '1', '99', '"x"'),
    'Strikethrough': ('True', 'False', '0', '1', '"x"'),
    'ColorIndex': ('1', '3', '56', '0', '57', '-1', 'xlColorIndexAutomatic', 'xlColorIndexNone', '"x"'),
    'Background': ('xlBackgroundAutomatic', 'xlBackgroundOpaque', 'xlBackgroundTransparent', '0', '1', '4', '"x"'),
    'Subscript': ('True', 'False', '"x"'),
    'Superscript': ('True', 'False', '"x"'),
    'ThemeColor': ('xlThemeColorAccent1', 'xlThemeColorDark1', 'xlThemeColorLight1', 'xlThemeColorHyperlink', '0', '13', '"x"'),
    'TintAndShade': ('0', '0.5', '-0.5', '1', '2', '"x"'),
    'ThemeFont': ('xlThemeFontMajor', 'xlThemeFontMinor', 'xlThemeFontNone', '0', '9', '"x"'),
    'Name': ('"Arial"', '""', '"No Such Font"', '12'),
    'Color': ('255', '0', '-1', '16777216', '"x"'),
    'Bold': ('True', 'False', '"x"'),
    'Italic': ('True', 'False', '"x"'),
}
CASES: list[str] = ['r = 0']
# Every value, on the chart title, whose fixed answers the others share; the undoing pairs on all of them.
for name, values in VALUES.items():
    CASES += [f'{ELEMENTS["title"]}.{name} = {value}: r = {ELEMENTS["title"]}.{name}' for value in values]
UNDOS = (('Underline', 'xlUnderlineStyleSingle', 'xlUnderlineStyleNone'), ('Strikethrough', 'True', 'False'), ('ColorIndex', '3', 'xlColorIndexAutomatic'), ('ColorIndex', '3', 'xlColorIndexNone'), ('Color', '255', '0'), ('Bold', 'True', 'False'), ('Italic', 'True', 'False'), ('Subscript', 'True', 'False'), ('Superscript', 'True', 'False'), ('Size', '14', '10'), ('Underline', 'xlUnderlineStyleDouble', 'xlUnderlineStyleSingle'))
for element, font in ELEMENTS.items():
    for name, value in (('Underline', 'xlUnderlineStyleDouble'), ('Underline', 'xlUnderlineStyleNone'), ('Strikethrough', 'False'), ('ColorIndex', 'xlColorIndexAutomatic'), ('ColorIndex', 'xlColorIndexNone'), ('Subscript', 'True'), ('Superscript', 'True'), ('Background', 'xlBackgroundOpaque'), ('ThemeColor', 'xlThemeColorAccent1'), ('TintAndShade', '0.5'), ('ThemeFont', 'xlThemeFontMajor'), ('Size', '7.25')):
        if element != 'title':
            CASES += [f'{font}.{name} = {value}: r = {font}.{name}']
    CASES += [f'{font}.{name} = {first}: {font}.{name} = {second}: r = {font}.{name}' for name, first, second in UNDOS]
CASES += [f'{ELEMENTS["title"]}.Subscript = True: {ELEMENTS["title"]}.Superscript = True: r = {ELEMENTS["title"]}.Subscript & " " & {ELEMENTS["title"]}.Superscript', f'{ELEMENTS["title"]}.ThemeColor = xlThemeColorAccent1: {ELEMENTS["title"]}.TintAndShade = 0.5: r = {ELEMENTS["title"]}.ThemeColor & " " & {ELEMENTS["title"]}.TintAndShade & " " & {ELEMENTS["title"]}.Color', f'{ELEMENTS["title"]}.Color = 255: {ELEMENTS["title"]}.ThemeColor = xlThemeColorAccent1: r = {ELEMENTS["title"]}.Color', f'{ELEMENTS["title"]}.ThemeColor = xlThemeColorAccent1: {ELEMENTS["title"]}.Color = 255: r = {ELEMENTS["title"]}.Color', f'{ELEMENTS["title"]}.Underline = xlUnderlineStyleDouble: {ELEMENTS["title"]}.Strikethrough = True: r = {ELEMENTS["title"]}.Underline & " " & {ELEMENTS["title"]}.Strikethrough']

# The chart area and the elements, each way round.
AREA = ELEMENTS['area']
for name, value in (('Size', '14'), ('Bold', 'True'), ('Bold', 'False'), ('Italic', 'True'), ('Color', '255'), ('Name', '"Arial"'), ('Underline', 'xlUnderlineStyleSingle'), ('Strikethrough', 'True'), ('ColorIndex', '3'), ('FontStyle', '"Italic"')):
    CASES += [f'{AREA}.{name} = {value}: r = {AREA}.{name}']
    for element, font in ELEMENTS.items():
        if element != 'area':
            other = '20' if name == 'Size' else 'False' if value == 'True' else 'True' if value == 'False' else '65280' if name == 'Color' else '"Courier New"' if name == 'Name' else 'xlUnderlineStyleNone' if name == 'Underline' else '5' if name == 'ColorIndex' else '"Bold"' if name == 'FontStyle' else 'False'
            CASES += [f'{font}.{name} = {other}: {AREA}.{name} = {value}: r = {font}.{name} & " " & {AREA}.{name}', f'{AREA}.{name} = {value}: {font}.{name} = {other}: r = {font}.{name} & " " & {AREA}.{name}']
CASES += [f'{AREA}.Size = 14: ch.HasTitle = False: ch.HasTitle = True: r = ch.ChartTitle.Font.Size', f'{AREA}.Size = 14: ch.HasLegend = False: ch.HasLegend = True: r = ch.Legend.Font.Size', f'{AREA}.Size = 14: s.HasDataLabels = False: s.HasDataLabels = True: r = s.DataLabels.Font.Size', f'{AREA}.Size = 14: ch.SeriesCollection(2).HasDataLabels = True: r = ch.SeriesCollection(2).DataLabels.Font.Size', f'{AREA}.Size = 14: ch.Axes(1).HasTitle = True: r = ch.Axes(1).AxisTitle.Font.Size', f'{AREA}.Size = 14: ch.HasDataTable = True: r = ch.DataTable.Font.Size', f'{AREA}.Size = 14: r = ch.Legend.LegendEntries(1).Font.Size', 'ch.Legend.LegendEntries(1).Font.Size = 20: ch.ChartArea.Font.Size = 14: r = ch.Legend.LegendEntries(1).Font.Size & " " & ch.Legend.Font.Size',
          f'{AREA}.Size = 14: ch.ChartTitle.Text = "Typed": r = ch.ChartTitle.Font.Size', 'ch.ChartTitle.Text = "Typed": ch.ChartArea.Font.Size = 14: r = ch.ChartTitle.Font.Size', 'ch.ChartTitle.Characters(1, 2).Font.Size = 30: ch.ChartArea.Font.Size = 14: r = ch.ChartTitle.Font.Size & " " & ch.ChartTitle.Characters(1, 2).Font.Size', f'{AREA}.Size = 14: ch.ChartArea.ClearFormats: r = ch.ChartArea.Font.Size & " " & ch.ChartTitle.Font.Size', f'{AREA}.Size = 14: ch.ClearToMatchStyle: r = ch.ChartArea.Font.Size', f'{AREA}.Size = 14: ch.ChartType = xlLine: r = ch.Legend.Font.Size', f'{AREA}.Size = 14: ch.ChartStyle = 5: r = ch.Legend.Font.Size', 'ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 14: r = ch.ChartArea.Font.Size & " " & ch.Legend.Font.Size']
# On a chart a recorded macro makes.
STYLED = 'ws.ChartObjects(1).Delete: Set ch = ws.Shapes.AddChart2(201, xlColumnClustered).Chart: ch.SetSourceData ws.Range("A1:D4"): ch.HasTitle = True: ch.ChartTitle.Text = "Mine": ch.HasLegend = True: ch.Axes(2).HasTitle = True: Set s = ch.SeriesCollection(1): s.HasDataLabels = True: '
CASES += [f'{STYLED}r = 0', f'{STYLED}{AREA}.Size = 14: r = {AREA}.Size', f'{STYLED}{AREA}.Bold = True: r = {AREA}.Bold', f'{STYLED}{AREA}.Color = 255: r = {AREA}.Color', f'{STYLED}ch.ChartTitle.Font.Size = 20: {AREA}.Size = 14: r = ch.ChartTitle.Font.Size', f'{STYLED}{AREA}.Size = 14: ch.ChartTitle.Font.Size = 20: r = ch.ChartTitle.Font.Size', f'{STYLED}ch.ChartTitle.Font.Underline = xlUnderlineStyleDouble: r = ch.ChartTitle.Font.Underline', f'{STYLED}ch.Legend.Font.ThemeColor = xlThemeColorAccent1: r = ch.Legend.Font.ThemeColor']

# Does a setting on the chart area clear only itself from the elements, or all that was set of them? Rivals split by another setting.
ACROSS = ('ch.ChartTitle.Font.Size = 20', 'ch.Legend.Font.Bold = True', 'ch.Axes(2).TickLabels.Font.Color = 255', 's.DataLabels.Font.Name = "Arial"', 'ch.Axes(2).AxisTitle.Font.Underline = xlUnderlineStyleSingle')
for prefix in ('', STYLED):
    for one in ACROSS:
        CASES += [f'{prefix}{one}: {AREA}.Italic = True: r = 0', f'{prefix}{one}: {AREA}.Size = 14: r = 0', f'{prefix}{one}: {AREA}.Bold = False: r = 0']
    CASES += [f'{prefix}{AREA}.Italic = True: ch.ChartTitle.Font.Size = 20: {AREA}.Bold = True: r = 0', f'{prefix}{AREA}.Size = 14: {AREA}.Bold = True: r = 0', f'{prefix}{AREA}.Size = 14: ch.ChartTitle.Font.Bold = False: {AREA}.Size = 16: r = 0']

MAKE = '''    Set ws = wb.Worksheets(1)
    ws.Range("A1:D1").Value = Array("", "Rev", "Cost", "Rate")
    ws.Range("A2:D2").Value = Array("a", 10, 4, 1)
    ws.Range("A3:D3").Value = Array("b", 20, 6, 2)
    ws.Range("A4:D4").Value = Array("c", 30, 8, 3)
    Set ch = ws.ChartObjects.Add(48, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D4")
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.ChartTitle.Text = "Mine"
    ch.HasLegend = True
    ch.Axes(2).HasTitle = True
    Set s = ch.SeriesCollection(1)
    s.HasDataLabels = True
'''


def probe() -> str:
    lines = ['Private Function Say(o As Object) As String', '    Dim s As String, v As Variant', '    On Error Resume Next']
    for name in READS:
        lines += ['    v = "?": Err.Clear', f'    v = o.{name}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & v & ","']
    lines += ['    Err.Clear', '    Say = s', 'End Function', '', 'Private Function Fonts(ch As Object) As String', '    Dim s As String, o As Object', '    On Error Resume Next', '    Dim sr As Object', '    Set sr = ch.SeriesCollection(1)']
    for element, font in ELEMENTS.items():
        lines += ['    Set o = Nothing: Err.Clear', f'    Set o = {font.replace("s.DataLabels", "sr.DataLabels")}', f'    If Err.Number Then s = s & "{element}=E" & Err.Number & ";" Else s = s & "{element}=" & Say(o) & ";"']
    lines += ['    Err.Clear', '    Fonts = s', 'End Function', '']
    for first in range(0, len(CASES), 60):
        lines += [f'Private Sub Act{first // 60}(ByVal which As Long, ws As Worksheet, ch As Object, s As Object, r As Variant)', '    Select Case which']
        for index, statement in enumerate(CASES[first:first + 60], first):
            lines += [f'    Case {index}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As Object, r As Variant, t As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', MAKE.rstrip('\n'), '    On Error GoTo Failed', '    Select Case which \\ 60']
    lines += [f'    Case {block}: Act{block} which, ws, ch, s, r' for block in range((len(CASES) + 59) // 60)]
    lines += ['    End Select', '    t = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    t = t & "|" & v & "|" & Fonts(ws.ChartObjects(1).Chart)', '    On Error GoTo 0', '    Tried = t', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function', 'Failed:', '    t = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_parts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


def batch(source: str, folder: str, first: int, last: int) -> list[str]:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        done = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0, line_numbers=False)
        if done.ok and not done.dialogs:
            return str(done.value).split('@@@')
    if first == last:
        return [f'STOPPED {done.outcome}']
    return [answer for index in range(first, last + 1) for answer in batch(source, folder, index, index)]


def main() -> None:
    source = probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for first in range(0, len(CASES), 40):
            answers += batch(source, folder, first, min(first + 39, len(CASES) - 1))
        parts = [chart_parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(len(CASES), 'cases', sum(1 for answer in answers if answer.startswith('STOPPED')), 'stopped')


if __name__ == '__main__':
    main()
