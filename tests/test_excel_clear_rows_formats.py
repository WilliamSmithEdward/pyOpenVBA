"""Clearing whole rows on a sheet whose columns carry a format, as live Excel does it (issue #256;
scripts/measure_clear_rows_formats.py, tests/fixtures/clear_rows_formats.json).

Three formatted columns or more give each cleared row a format of the default, fewer a cell at each formatted
column; rows cleared to the last row take the columns' formats away, the rows above keeping them. Clearing down to
the last row after Cells.Font.Name made a cell at every position of every row, which ran the model out of memory.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/clear_rows_formats.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def _summary(xml: str) -> dict[str, object]:
    """A saved sheet read as the probe's _sheet_summary read Excel's: its cols, how many rows and cells it keeps,
    and its dimension."""
    return {'cols': re.findall(r'<col\b[^>]*/>', xml), 'row_count': len(re.findall(r'<row\b', xml)),
            'cell_count': len(re.findall(r'<c\b', xml)), 'dimension': re.findall(r'<dimension\b[^>]*/>', xml)}


@pytest.fixture(scope='module')
def answers(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('cleared')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='Clearing')
    return str(app.run('Clearing.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case[1] for case in RECORD['cases']])
def test_each_clear_as_excel_answered(answers: tuple[list[str], Path], index: int) -> None:
    assert answers[0][index] == RECORD['cases'][index][2]


@pytest.mark.parametrize('saved', list(RECORD['sheets']))
def test_the_saved_rows_and_cells_as_excel_kept_them(answers: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(answers[1] / saved) as package:
        mine = _summary(package.read('xl/worksheets/sheet1.xml').decode('utf-8'))
    wanted = RECORD['sheets'][saved]
    for key in ('cols', 'row_count', 'cell_count', 'dimension'):
        assert mine[key] == wanted[key], key
