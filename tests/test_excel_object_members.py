"""Small members of pivot, shape, pane and conditional-format objects, replayed from live Excel
(scripts/measure_object_members.py, tests/fixtures/object_members.json): Application, Parent, PTCondition, CalcFor
and the pivot objects' default members; Validation.IMEMode, which this Excel, with no East Asian input method,
leaves at 0 and leaves out of the file; and a cell's Font.Background and Interior.InvertIfNegative."""
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
    (Path(__file__).parent / 'fixtures/object_members.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
EXPECTED = [answer for _, answer in RECORD['reads']] + [edit[3] for edit in RECORD['edits']]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('object_members')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='MembersProbe')
    return str(app.run('MembersProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(EXPECTED)))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == EXPECTED[index]


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_ime_mode_is_left_out_of_the_file(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<dataValidation\b[^>]*>', sheet)
    assert found is not None and 'imeMode' not in found.group(0)
    assert 'imeMode' not in RECORD['parts'][saved]
