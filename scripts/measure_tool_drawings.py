"""What Excel reads from a drawing that does not spell its namespace with xdr: as openpyxl writes it, the namespace the part's default,
and as another writer might, under a prefix of its own. Each read with its type. Needs Excel and openpyxl:

    python scripts/measure_tool_drawings.py

writes tests/fixtures/tool_drawings.json and the books beside it, which tests/test_excel_tool_drawings.py replays.
"""

from __future__ import annotations

import datetime
import json
import re
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
READS = ('s.Shapes.Count', 's.Shapes(1).Name', 's.Shapes(1).Type', 's.Shapes(1).HasChart', 's.Shapes(2).Name', 's.ChartObjects.Count', 's.ChartObjects(1).Name',
         's.ChartObjects(2).Name', 's.ChartObjects(1).Chart.ChartType', 's.ChartObjects(1).Chart.SeriesCollection.Count', 's.ChartObjects(1).Chart.SeriesCollection(1).Name',
         's.ChartObjects(1).Chart.SeriesCollection(1).Formula', 's.ChartObjects(1).Chart.HasTitle', 's.ChartObjects(1).Chart.HasLegend', 's.ChartObjects(2).Chart.ChartType',
         's.ChartObjects(2).Chart.SeriesCollection.Count', 's.ChartObjects(1).TopLeftCell.Address', 's.ChartObjects(2).TopLeftCell.Address')
CH = 's.ChartObjects(1).Chart'
#: What a part left without an element reads by that element's default, by each kind of chart.
READS += tuple(f'{chart}.{one}' for chart in (CH, 's.ChartObjects(2).Chart') for one in (
    'HasAxis(1)', 'HasAxis(2)', 'ChartGroups(1).VaryByCategories', 'SeriesCollection(1).InvertIfNegative', 'SeriesCollection(1).MarkerStyle',
    'SeriesCollection(1).Smooth'))
BOOKS = ('tool_drawing_default.xlsx', 'tool_drawing_prefixed.xlsx', 'tool_drawing_kinds.xlsx')
PROBE = ['Public Function Probe(ByVal path As String) As String', '    Dim wb As Workbook, s As Worksheet, r As String', '    Set wb = Workbooks.Open(path)', '    Set s = wb.Worksheets(1)',
         '    On Error Resume Next']
for read in READS:
    PROBE += ['    Err.Clear', f'    r = r & TypeName({read}) & " " & {read} & " " & Err.Number & "|"']
PROBE += ['    On Error GoTo 0', '    wb.Close False', '    Probe = r', 'End Function', '',
          'Public Function Moved(ByVal path As String, ByVal out As String) As String', '    Dim wb As Workbook, s As Worksheet', '    Set wb = Workbooks.Open(path)',
          '    Set s = wb.Worksheets(1)', '    s.Rows(1).Insert', f'    Moved = {CH}.SeriesCollection(1).Formula', '    Application.DisplayAlerts = False',
          '    wb.SaveAs out, xlOpenXMLWorkbook', '    wb.Close False', 'End Function']
SOURCE = '\n'.join(PROBE)


def by_openpyxl(out: Path) -> None:
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = 'One'
    for row in (('', 'Rev', 'Cost'), ('a', 10, 4), ('b', 20, 6)):
        sheet.append(row)
    for corner, last in (('E2', 3), ('E20', 2)):
        chart = BarChart()
        chart.add_data(Reference(sheet, min_col=2, max_col=last, min_row=1, max_row=3), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
        sheet.add_chart(chart, corner)
    book.save(out)


def other_kinds(out: Path) -> None:
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = 'One'
    for row in (('', 'Rev', 'Cost'), ('a', 10, 4), ('b', 20, 6)):
        sheet.append(row)
    for corner, chart in (('E2', LineChart()), ('E20', PieChart())):
        chart.add_data(Reference(sheet, min_col=2, max_col=2, min_row=1, max_row=3), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
        sheet.add_chart(chart, corner)
    book.save(out)


def anchor_rows(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return re.findall(r'<(?:\w+:)?row>(\d+)', package.read('xl/drawings/drawing1.xml').decode('utf-8'))


def references(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return re.findall(r'<(?:c:)?f>(.*?)</(?:c:)?f>', package.read('xl/charts/chart1.xml').decode('utf-8'))


def under_a_prefix(source: Path, out: Path) -> None:
    """The drawing's own elements under p: rather than the default."""
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as target:
        for item in package.infolist():
            data = package.read(item.filename)
            if item.filename == 'xl/drawings/drawing1.xml':
                text = re.sub(r'<(/?)([A-Za-z]\w*)(?=[\s/>])', r'<\1p:\2', data.decode('utf-8'))
                data = text.replace('<p:wsDr xmlns=', '<p:wsDr xmlns:p=', 1).encode('utf-8')
            target.writestr(item, data)


def main() -> None:
    by_openpyxl(FIXTURES / BOOKS[0])
    under_a_prefix(FIXTURES / BOOKS[0], FIXTURES / BOOKS[1])
    other_kinds(FIXTURES / BOOKS[2])
    answers = {}
    moved = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        for name in BOOKS:
            done = excel.run_vba(SOURCE, 'Probe', (str(FIXTURES / name),), timeout=120.0, line_numbers=False)
            assert done.outcome == 'passed', done.message
            answers[name] = done.value
            print(name, done.value)
            out = Path(tempfile.mkdtemp()) / 'out.xlsx'
            done = excel.run_vba(SOURCE, 'Moved', (str(FIXTURES / name), str(out)), timeout=120.0, line_numbers=False)
            assert done.outcome == 'passed', done.message
            moved[name] = [done.value, references(out), anchor_rows(out)]
            print(moved[name])
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': SOURCE, 'reads': READS, 'answers': answers, 'moved': moved}
    (FIXTURES / 'tool_drawings.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
