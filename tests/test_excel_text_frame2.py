"""TextFrame2's margins, AutoSize, anchors, NoTextRotation, Orientation, WordWrap and DeleteText, replayed from
live Excel (scripts/measure_text_frame2.py, tests/fixtures/text_frame2.json): what a rectangle and a text box
answer after each setting, the errors, and each shape's saved txBody."""
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
    (Path(__file__).parent / 'fixtures/text_frame2.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def text_bodies(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.S)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('text_frame2')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='TextFrame2Probe')
    return str(app.run('TextFrame2Probe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_text_layout_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert text_bodies(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
