"""What Excel reads from a drawing that does not spell its namespace with xdr: as openpyxl writes it, the namespace the part's default,
and as another writer might, under a prefix of its own. Each read with its type. Needs Excel and openpyxl:

    python scripts/measure_tool_drawings.py

writes tests/fixtures/tool_drawings.json and the books beside it, which tests/test_excel_tool_drawings.py replays.
"""

from __future__ import annotations

import datetime
import json
import re
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
READS = ('s.Shapes.Count', 's.Shapes(1).Name', 's.Shapes(1).Type', 's.Shapes(1).HasChart', 's.Shapes(2).Name', 's.ChartObjects.Count', 's.ChartObjects(1).Name',
         's.ChartObjects(2).Name', 's.ChartObjects(1).Chart.ChartType', 's.ChartObjects(1).Chart.SeriesCollection.Count', 's.ChartObjects(1).Chart.SeriesCollection(1).Name',
         's.ChartObjects(1).Chart.SeriesCollection(1).Formula', 's.ChartObjects(1).Chart.HasTitle', 's.ChartObjects(1).Chart.HasLegend', 's.ChartObjects(2).Chart.ChartType',
         's.ChartObjects(2).Chart.SeriesCollection.Count', 's.ChartObjects(1).TopLeftCell.Address', 's.ChartObjects(2).TopLeftCell.Address')
BOOKS = ('tool_drawing_default.xlsx', 'tool_drawing_prefixed.xlsx')
PROBE = ['Public Function Probe(ByVal path As String) As String', '    Dim wb As Workbook, s As Worksheet, r As String', '    Set wb = Workbooks.Open(path)', '    Set s = wb.Worksheets(1)',
         '    On Error Resume Next']
for read in READS:
    PROBE += ['    Err.Clear', f'    r = r & TypeName({read}) & " " & {read} & " " & Err.Number & "|"']
PROBE += ['    On Error GoTo 0', '    wb.Close False', '    Probe = r', 'End Function']
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
    answers = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        for name in BOOKS:
            done = excel.run_vba(SOURCE, 'Probe', (str(FIXTURES / name),), timeout=120.0, line_numbers=False)
            assert done.outcome == 'passed', done.message
            answers[name] = done.value
            print(name, done.value)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': SOURCE, 'reads': READS, 'answers': answers}
    (FIXTURES / 'tool_drawings.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
