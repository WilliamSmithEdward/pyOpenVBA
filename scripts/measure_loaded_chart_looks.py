"""What Excel reads of a loaded chart's looks: fill, line, interior, border and font, element by element, through a typed Chart
and a late-bound Object, each read with its type and error. Read only; nothing is set after the book is opened.

On four books: the plain and dressed charts Excel saved and the chart openpyxl wrote (tests/fixtures/loaded_chart_*.xlsx), and a
themed one made here with theme colours, brightness, transparency, a gradient and a dashed line.

    python scripts/measure_loaded_chart_looks.py

writes tests/fixtures/loaded_chart_looks.json and the themed book beside it, which tests/test_excel_loaded_chart_looks.py replays.
"""

from __future__ import annotations

import datetime
import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
ABSOLUTE = re.compile(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', re.S)
BOOKS = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_openpyxl.xlsx', 'loaded_chart_themed.xlsx')
ELEMENTS = ('ChartArea', 'PlotArea', 'Legend', 'ChartTitle', 'SeriesCollection(1)', 'SeriesCollection(2)', 'SeriesCollection(1).Points(1)',
            'Axes(xlValue)', 'Axes(xlValue).MajorGridlines', 'Axes(xlCategory)')
FORMAT = ('Format.Fill.Visible', 'Format.Fill.Type', 'Format.Fill.ForeColor.RGB', 'Format.Fill.ForeColor.ObjectThemeColor',
          'Format.Fill.ForeColor.TintAndShade', 'Format.Fill.ForeColor.Brightness', 'Format.Fill.Transparency', 'Format.Fill.GradientStops.Count',
          'Format.Line.Visible', 'Format.Line.ForeColor.RGB', 'Format.Line.ForeColor.ObjectThemeColor', 'Format.Line.Weight', 'Format.Line.DashStyle',
          'Format.Line.Transparency')
CLASSIC = ('Interior.Color', 'Interior.ColorIndex', 'Interior.Pattern', 'Border.Color', 'Border.ColorIndex', 'Border.Weight', 'Border.LineStyle')
FONT = ('Font.Size', 'Font.Bold', 'Font.Italic', 'Font.Color', 'Font.Name')
READS = tuple(f'{element}.{member}' for element in ELEMENTS for member in (*FORMAT, *CLASSIC, *FONT))
#: Through a typed Chart these stop the module compiling, which ends the harness's run, so they are read late only: a PlotArea has no Font.
NOT_TYPED = frozenset(read for read in READS if read.startswith('PlotArea.Font.'))

MAKE = r'''
Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, c As Chart
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    c.HasTitle = True
    c.ChartTitle.Text = "Themed"
    With c.SeriesCollection(1).Format.Fill
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.Brightness = 0.4
        .Transparency = 0.3
    End With
    c.SeriesCollection(2).Format.Fill.TwoColorGradient msoGradientHorizontal, 1
    With c.SeriesCollection(2).Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(10, 20, 30)
        .Weight = 2.5
        .DashStyle = msoLineDash
    End With
    c.ChartArea.Format.Fill.ForeColor.RGB = RGB(250, 240, 230)
    c.PlotArea.Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground2
    c.Legend.Format.Fill.Visible = msoFalse
    c.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 0, 0)
    c.ChartArea.Font.Size = 12
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function
'''


def probe() -> str:
    """A function for each element through each binding, since one procedure for them all is too large for VBA to compile."""
    lines: list[str] = []
    calls: list[str] = []
    for holder, kind in (('t', 'Chart'), ('l', 'Object')):
        for number, element in enumerate(ELEMENTS):
            name = f'R{holder}{number}'
            calls.append(f'    s = s & {name}(t)')
            lines += [f'Private Function {name}(ByVal {holder} As {kind}) As String', '    Dim s As String, x As String', '    On Error Resume Next']
            for member in (*FORMAT, *CLASSIC, *FONT):
                read = f'{element}.{member}'
                if holder == 't' and read in NOT_TYPED:
                    lines += ['    s = s & "does not compile|"']
                    continue
                # A read that raises skips its whole statement, so the answer is made first and always added after.
                lines += ['    Err.Clear', '    x = ""', f'    x = TypeName({holder}.{read}) & " " & {holder}.{read}', '    s = s & x & " " & Err.Number & "|"']
            lines += [f'    {name} = s', 'End Function', '']
    lines += ['Public Function Read(ByVal path As String) As String', '    Dim wb As Workbook, t As Chart, s As String',
              '    Set wb = Workbooks.Open(path)', '    Set t = wb.Worksheets(1).ChartObjects(1).Chart', *calls, '    wb.Close False',
              '    Read = s', 'End Function']
    return '\n'.join(lines)


def without_folder(made: Path, out: Path) -> None:
    with zipfile.ZipFile(made) as package, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as target:
        for item in package.infolist():
            data = package.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                data = ABSOLUTE.sub('', data.decode('utf-8')).encode('utf-8')
            target.writestr(item, data)


def main() -> None:
    import tempfile

    folder = Path(tempfile.mkdtemp())
    source = probe()
    answers: dict[str, dict[str, list[str]]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        done = excel.run_vba(MAKE, 'Make', (str(folder / BOOKS[3]),), timeout=120.0, line_numbers=False)
        assert done.outcome == 'passed', done.message
    without_folder(folder / BOOKS[3], FIXTURES / BOOKS[3])
    for name in BOOKS:
        # Each book in a session of its own, so nothing one book did is still about when the next opens.
        with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
            done = excel.run_vba(source, 'Read', (str(FIXTURES / name),), timeout=300.0, line_numbers=False)
        assert done.outcome == 'passed', (name, done.message)
        parts = str(done.value).split('|')[:-1]
        half = len(READS)
        assert len(parts) == 2 * half, (name, len(parts))
        answers[name] = {'typed': parts[:half], 'late': parts[half:]}
        print(name, sum(1 for one in parts if one.endswith(' 0')), 'read of', len(parts), flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': source, 'reads': READS, 'answers': answers}
    (FIXTURES / 'loaded_chart_looks.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
