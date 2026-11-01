"""Range.ApplyOutlineStyles replayed from live Excel (scripts/measure_outline_styles.py,
tests/fixtures/outline_styles.json): what it answers, which cells take which outline style, and the fonts, cell style
formats, cell formats and cell styles the saved workbook holds."""
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
    (Path(__file__).parent / 'fixtures/outline_styles.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def style_parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        styles = package.read('xl/styles.xml').decode('utf-8')
    parts: list[str] = []
    for tag in ('fonts', 'cellStyleXfs', 'cellXfs', 'cellStyles'):
        found = re.search(f'<{tag}\\b.*?</{tag}>', styles, re.S)
        parts.append(found.group() if found else '')
    return parts


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('outline_styles')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='OutlineProbe')
    return str(app.run('OutlineProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_stylesheet_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert style_parts(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
