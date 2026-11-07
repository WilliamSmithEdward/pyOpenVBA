"""What Excel reads of a loaded chart's data labels, series by series and for a point, through a typed Chart and a late-bound Object,
each read with its type and error. Read only, on the four loaded books and two made here: one with labels set every way a macro might,
and one with labels turned, shadowed, aligned and read right to left, which the model reads as new labels only where the file says so.

    python scripts/measure_loaded_chart_labels.py

writes tests/fixtures/loaded_chart_labels.json and the two made books beside it, which tests/test_excel_loaded_chart_labels.py replays.
"""

from __future__ import annotations

import datetime
import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
ABSOLUTE = re.compile(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', re.S)
BOOKS = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_openpyxl.xlsx', 'loaded_chart_themed.xlsx',
         'loaded_chart_labelled.xlsx', 'loaded_chart_turned.xlsx')
OWNERS = ('SeriesCollection(1)', 'SeriesCollection(2)')
SERIES = ('HasDataLabels', 'DataLabels.ShowValue', 'DataLabels.ShowCategoryName', 'DataLabels.ShowSeriesName', 'DataLabels.ShowLegendKey',
          'DataLabels.Position', 'DataLabels.NumberFormat', 'DataLabels.Separator', 'DataLabels.Count', 'DataLabels.Font.Size',
          'DataLabels.Font.Bold', 'Points(1).HasDataLabel', 'Points(2).HasDataLabel', 'Points(1).DataLabel.Text', 'Points(2).DataLabel.Text',
          'Points(2).DataLabel.Caption', 'Points(2).DataLabel.ShowValue', 'Points(2).DataLabel.Position', 'Points(2).DataLabel.NumberFormat',
          'Points(2).DataLabel.AutoText', 'Points(2).DataLabel.Font.Size', 'DataLabels.Orientation', 'DataLabels.HorizontalAlignment',
          'DataLabels.VerticalAlignment', 'DataLabels.ReadingOrder', 'DataLabels.Shadow', 'DataLabels.AutoText', 'DataLabels.ShowRange',
          'Points(2).DataLabel.Orientation', 'Points(2).DataLabel.Shadow', 'Points(2).DataLabel.HorizontalAlignment')
READS = tuple(f'{owner}.{member}' for owner in OWNERS for member in SERIES)

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
    c.SeriesCollection(1).HasDataLabels = True
    c.SeriesCollection(1).DataLabels.ShowCategoryName = True
    c.SeriesCollection(1).DataLabels.NumberFormat = "0.0"
    c.SeriesCollection(1).DataLabels.Position = xlLabelPositionInsideEnd
    c.SeriesCollection(2).Points(2).HasDataLabel = True
    c.SeriesCollection(2).Points(2).DataLabel.Text = "P"
    c.SeriesCollection(2).Points(2).DataLabel.Font.Size = 12
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function

Public Function Turn(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, c As Chart
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    c.SeriesCollection(1).HasDataLabels = True
    c.SeriesCollection(1).DataLabels.Orientation = xlUpward
    c.SeriesCollection(1).DataLabels.Shadow = True
    c.SeriesCollection(1).Points(2).DataLabel.Orientation = 45
    c.SeriesCollection(2).HasDataLabels = True
    c.SeriesCollection(2).DataLabels.ReadingOrder = xlRTL
    c.SeriesCollection(2).DataLabels.HorizontalAlignment = xlLeft
    c.SeriesCollection(2).Points(2).DataLabel.Shadow = True
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function
'''


def probe() -> str:
    """A function for each owner through each binding, under VBA's procedure size limit. A read that raises skips its whole statement, so
    each answer is made first and always added after."""
    lines: list[str] = []
    calls: list[str] = []
    for holder, kind in (('t', 'Chart'), ('l', 'Object')):
        for number, owner in enumerate(OWNERS):
            name = f'R{holder}{number}'
            calls.append(f'    s = s & {name}(t)')
            lines += [f'Private Function {name}(ByVal {holder} As {kind}) As String', '    Dim s As String, x As String', '    On Error Resume Next']
            for member in SERIES:
                read = f'{owner}.{member}'
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
    folder = Path(tempfile.mkdtemp())
    source = probe()
    answers: dict[str, dict[str, list[str]]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        done = excel.run_vba(MAKE, 'Make', (str(folder / BOOKS[4]),), timeout=120.0, line_numbers=False)
        assert done.outcome == 'passed', done.message
    without_folder(folder / BOOKS[4], FIXTURES / BOOKS[4])
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        done = excel.run_vba(MAKE, 'Turn', (str(folder / BOOKS[5]),), timeout=120.0, line_numbers=False)
        assert done.outcome == 'passed', done.message
    without_folder(folder / BOOKS[5], FIXTURES / BOOKS[5])
    for name in BOOKS:
        with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
            done = excel.run_vba(source, 'Read', (str(FIXTURES / name),), timeout=300.0, line_numbers=False)
        assert done.outcome == 'passed', (name, done.message)
        parts = str(done.value).split('|')[:-1]
        assert len(parts) == 2 * len(READS), (name, len(parts))
        answers[name] = {'typed': parts[:len(READS)], 'late': parts[len(READS):]}
        print(name, sum(1 for one in parts if one.endswith(' 0')), 'read of', len(parts), flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': source, 'reads': READS, 'answers': answers}
    (FIXTURES / 'loaded_chart_labels.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
