"""LineFormat, FillFormat, ColorFormat, Adjustments and ConnectorFormat through a ShapeRange, replayed from live Excel
(scripts/measure_shape_range_more_format.py, tests/fixtures/shape_range_more_format.json)."""
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

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/shape_range_more_format.json').read_text(encoding='utf-8'))
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
INDEXES = range(len(RECORD['cases']))
IDS = [f'{index}: {RECORD["cases"][index][0]}' for index in INDEXES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_range_more_format')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='RangeProbe')
    return str(app.run('RangeProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=IDS)
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=IDS)
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL) == RECORD['cases'][index][2]


@pytest.mark.parametrize('statements', ['Dim v: v = r.Fill.ForeColor.SchemeColor', 'r.Fill.GradientStops(1).Position = 0.5'])
def test_a_range_refuses_what_only_one_shape_has_measured(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim r As Object\n'
                   '    ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40).Name = "a"\n'
                   '    ActiveSheet.Shapes.AddShape(1, 20, 20, 100, 40).Name = "b"\n'
                   '    Set r = ActiveSheet.Shapes.Range(Array("a", "b"))\n    r.Fill.TwoColorGradient 1, 1\n'
                   f'    {statements}\nEnd Sub\n', name='Ranged')
    with pytest.raises(VBAUnsupportedError):
        app.run('Ranged.Go')
