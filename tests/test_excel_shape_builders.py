"""Shapes.BuildFreeform with its FreeformBuilder, and Shapes.AddCallout, replayed from live Excel
(scripts/measure_shape_builders.py, tests/fixtures/shape_builders.json)."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_builders.json').read_text(encoding='utf-8'))
#: Excel smooths these by a rule not measured, so they report themselves. The probe's others run as they are.
SMOOTHED = [index for index, case in enumerate(RECORD['cases']) if 'msoSegmentCurve, msoEditingAuto' in case[0]]


def _source() -> str:
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for index in SMOOTHED:
        source = re.sub(rf'    Case {index}\n        [^\n]*\n', f'    Case {index}\n', source)
    return source


INDEXES = [index for index in range(len(RECORD['cases'])) if index not in SMOOTHED]


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        try:
            xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        except KeyError:
            return ''
    found = re.search(r'<xdr:(sp|cxnSp)\b.*?</xdr:\1>', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_builders')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='BuilderProbe')
    return str(app.run('BuilderProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=[RECORD['cases'][index][0] for index in INDEXES])
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    strip = re.compile(r'<xdr:nv\w+>.*?</xdr:nv\w+>')
    assert strip.sub('', drawn(run[1] / f'c{index}.xlsx')) == strip.sub('', RECORD['cases'][index][2])


def test_a_smoothed_curve_reports_itself() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim b As Object\n    Set b = ActiveSheet.Shapes.BuildFreeform(1, 10, 10)\n'
                   '    b.AddNodes msoSegmentCurve, msoEditingAuto, 40, 60\nEnd Sub\n', name='Smooth')
    with pytest.raises(VBAUnsupportedError):
        app.run('Smooth.Go')
