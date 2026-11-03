"""Shape.Adjustments and Shapes.AddConnector with ConnectorFormat: what each reads, the errors,
and the anchor each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E"
and the error number, and saves the workbook. The shape's whole anchor is recorded from the drawing part.

    python scripts/measure_shape_adjustments.py

writes tests/fixtures/shape_adjustments.json, which tests/test_excel_shape_adjustments.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_adjustments.json'
RECT = 'ws.Shapes.AddShape(1, 10, 10, 100, 40)'
#: msoShapeRoundedRectangle 5, msoShapeOctagon 6, IsoscelesTriangle 7, Parallelogram 2, Trapezoid 3, Donut 18, Cube 14,
#: Can 13, Smiley 17, BlockArc 20, RightArrow 33, LeftRightArrow 37, LineCallout1 109, Star5 92, Wave 103.
AUTO = (1, 2, 3, 5, 6, 7, 8, 9, 10, 13, 14, 15, 17, 18, 20, 33, 37, 61, 92, 103, 109, 139)
CASES = [
    *[(f'ws.Shapes.AddShape({kind}, 10, 10, 100, 40)', '') for kind in AUTO],
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments.Item(1) = 0.3'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.5'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 2'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = -1'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.123456'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.16667'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(2) = 0.3'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(0) = 0.3'),
    ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddShape(33, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddShape(33, 10, 10, 100, 40)', 'sh.Adjustments(2) = 0.3'),
    ('ws.Shapes.AddShape(33, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.3: sh.Adjustments(2) = 0.7'),
    ('ws.Shapes.AddShape(33, 10, 10, 100, 40)', 'sh.Adjustments(2) = 0.7: sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddShape(7, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25'),
    ('ws.Shapes.AddShape(2, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25'),
    ('ws.Shapes.AddShape(61, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25: sh.AutoShapeType = 1'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Adjustments(1) = 0.25: sh.AutoShapeType = 33'),
    ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'sh.Width = 200: sh.Adjustments(1) = 0.25'),
    # Connectors.
    *[(f'ws.Shapes.AddConnector({kind}, 10, 10, 110, 50)', '') for kind in (1, 2, 3)],
    ('ws.Shapes.AddConnector(0, 10, 10, 110, 50)', ''), ('ws.Shapes.AddConnector(4, 10, 10, 110, 50)', ''),
    ('ws.Shapes.AddConnector(-2, 10, 10, 110, 50)', ''),
    ('ws.Shapes.AddConnector(1, 110, 50, 10, 10)', ''), ('ws.Shapes.AddConnector(1, 110, 10, 10, 50)', ''),
    ('ws.Shapes.AddConnector(2, 110, 50, 10, 10)', ''), ('ws.Shapes.AddConnector(3, 110, 10, 10, 50)', ''),
    ('ws.Shapes.AddConnector(1, 10, 10, 10, 10)', ''),
    ('ws.Shapes.AddConnector(1, 10, 10, 110, 50)', 'sh.ConnectorFormat.Type = 2'),
    ('ws.Shapes.AddConnector(2, 10, 10, 110, 50)', 'sh.ConnectorFormat.Type = 3'),
    ('ws.Shapes.AddConnector(1, 10, 10, 110, 50)', 'sh.ConnectorFormat.Type = 9'),
    ('ws.Shapes.AddLine(10, 10, 110, 50)', ''), ('ws.Shapes.AddLine(10, 10, 110, 50)', 'sh.ConnectorFormat.Type = 2'),
    ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'sh.ConnectorFormat.Type = 2'),
    ('ws.Shapes.AddConnector(1, 10, 10, 110, 50)', 'sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddConnector(2, 10, 10, 110, 50)', 'sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddConnector(3, 10, 10, 110, 50)', 'sh.Adjustments(1) = 0.3'),
    ('ws.Shapes.AddConnector(1, 10, 10, 110, 50)', 'sh.Line.EndArrowheadStyle = 2'),
]
READS = [
    'sh.Type', 'sh.AutoShapeType', 'sh.Connector', 'sh.Name', 'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height',
    'sh.HorizontalFlip', 'sh.VerticalFlip', 'TypeName(sh)', 'TypeName(sh.Adjustments)', 'sh.Adjustments.Count',
    'sh.Adjustments(1)', 'sh.Adjustments.Item(1)', 'sh.Adjustments(2)', 'sh.Adjustments(3)', 'sh.Adjustments(4)',
    'TypeName(sh.ConnectorFormat)', 'sh.ConnectorFormat.Type', 'sh.ConnectorFormat.BeginConnected',
    'sh.ConnectorFormat.EndConnected', 'sh.Fill.Visible', 'sh.Line.Visible', 'sh.Line.Weight',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (shape, _) in enumerate(CASES):
        lines += [f'    Case {index}: Set sh = {shape}']
    lines += ['    End Select', '    Select Case which']
    for index, (_, statements) in enumerate(CASES):
        if statements:
            lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawn(path: Path) -> str:
    """The first shape as the drawing part holds it, without its place in the sheet."""
    with zipfile.ZipFile(path) as package:
        try:
            xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        except KeyError:
            return ''
    found = re.search(r'<xdr:(sp|cxnSp)\b.*?</xdr:\1>', xml, re.DOTALL)
    return found.group(0) if found else ''


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=400.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part]
                        for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    strip = re.compile(r'<xdr:style>.*?</xdr:style>|<a:creationId.*?/>|<a:ext uri=.*?</a:ext>')
    for case, answer, part in record['cases']:
        print(case, '->', answer)
        print('    ', strip.sub('', part)[:640])


if __name__ == '__main__':
    main()
