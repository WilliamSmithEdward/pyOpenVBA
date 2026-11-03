"""Shape.Adjustments, Shapes.AddConnector and ConnectorFormat, replayed from live Excel
(scripts/measure_shape_adjustments.py, scripts/measure_shape_adjustment_table.py, and their fixtures)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'shape_adjustments.json').read_text(encoding='utf-8'))
TABLE: dict[str, Any] = json.loads((FIXTURES / 'shape_adjustment_table.json').read_text(encoding='utf-8'))
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
INDEXES = range(len(RECORD['cases']))
ROWS = {int(row[0].split('|')[0]): row for row in TABLE['rows'] if not row[0].split('|')[1].startswith('E')}
#: Connector changes no case measured.
UNMEASURED = ['sh.ConnectorFormat.Type = 2: sh.ConnectorFormat.Type = 1',
              'sh.ConnectorFormat.Type = 2: sh.Adjustments(1) = 0.3: sh.ConnectorFormat.Type = 3']


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:(sp|cxnSp)\b.*?</xdr:\1>', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_adjustments')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='AdjustmentProbe')
    return str(app.run('AdjustmentProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    # Names, ids and the place in the sheet are other tests'; the shape's properties and geometry are this one's.
    strip = re.compile(r'<xdr:nv\w+>.*?</xdr:nv\w+>|<xdr:style>.*?</xdr:style>')
    case = RECORD['cases'][index]
    if case[2]:
        assert strip.sub('', drawn(run[1] / f'c{index}.xlsx')) == strip.sub('', case[2])


@pytest.mark.parametrize('number', sorted(ROWS))
def test_each_autoshape_has_excels_adjustments(number: int, tmp_path: Path) -> None:
    fields = ROWS[number][0].split('|')
    count = int(fields[3])
    # The defaults, then each value read back after adding 0.1, in order.
    source = ('Public Function Go() As String\n    Dim sh As Object, k As Long, d As Double, s As String\n'
              f'    Set sh = ActiveSheet.Shapes.AddShape({number}, 10, 10, 100, 40)\n'
              '    s = sh.Adjustments.Count\n    For k = 1 To sh.Adjustments.Count\n'
              '        s = s & "|" & sh.Adjustments(k)\n    Next\n'
              '    For k = 1 To sh.Adjustments.Count\n        d = sh.Adjustments(k)\n'
              '        sh.Adjustments(k) = d + 0.1\n        s = s & "|R" & k & "=" & sh.Adjustments(k)\n    Next\n'
              '    Go = s & "|T=" & sh.Type\nEnd Function\n')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Table')
    assert app.run('Table.Go') == '|'.join(fields[3:5 + 2 * count])
    path = tmp_path / 'book.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert ROWS[number][1] in xml
    assert ROWS[number][2] in xml


@pytest.mark.parametrize('number', [row[0].split('|')[0] for row in TABLE['rows'] if row[0].split('|')[1].startswith('E')])
def test_a_type_excel_refuses_is_refused(number: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As Long\n    On Error Resume Next\n'
                   f'    ActiveSheet.Shapes.AddShape {number}, 10, 10, 100, 40\n    Go = Err.Number\nEnd Function\n',
                   name='Refused')
    assert app.run('Refused.Go') == 1004


@pytest.mark.parametrize('statements', UNMEASURED)
def test_an_unmeasured_connector_change_is_refused(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddConnector(1, 10, 10, 110, 50)\n'
                   f'    {statements}\nEnd Sub\n', name='Unmeasured')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unmeasured.Go')


def test_a_duplicate_keeps_its_adjustments_and_changes_its_own() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim sh As Object, d As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(5, 10, 10, 100, 40)\n'
                   '    sh.Adjustments(1) = 0.25\n    Set d = sh.Duplicate\n'
                   '    Go = d.Adjustments(1)\n    d.Adjustments(1) = 0.4\n'
                   '    Go = Go & " " & sh.Adjustments(1)\nEnd Function\n', name='Copied')
    assert app.run('Copied.Go') == '0.25 0.25'
