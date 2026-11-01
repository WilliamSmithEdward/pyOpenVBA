"""ShapeRange properties replayed from live Excel (scripts/measure_shape_range_properties.py,
tests/fixtures/shape_range_properties.json): what a range of one or two shapes reads, what setting each does to its
shapes, and the errors either raises."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/shape_range_properties.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def answers(tmp_path_factory: pytest.TempPathFactory) -> list[str]:
    folder = tmp_path_factory.mktemp('shape_range_properties')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='RangeProbe')
    return str(app.run('RangeProbe.Probe', str(folder))).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
