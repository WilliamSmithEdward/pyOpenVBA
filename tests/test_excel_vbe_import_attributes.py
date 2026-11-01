"""Procedure attributes through VBComponent.Export and VBComponents.Import, replayed from live Excel
(scripts/measure_vbe_import_attributes.py, tests/fixtures/vbe_import_attributes.json): the exported file, the
imported module's lines, which leave the attributes out, and the module the importing workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.excel import ExcelFile

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/vbe_import_attributes.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[str, Path]:
    folder = tmp_path_factory.mktemp('vbe_import_attributes')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ImportProbe')
    return str(app.run('ImportProbe.Probe', str(folder))), folder


def test_the_imported_lines_leave_the_attributes_out(run: tuple[str, Path]) -> None:
    assert run[0] == RECORD['answer']


def test_the_export_carries_the_attributes(run: tuple[str, Path]) -> None:
    assert (run[1] / 'm.bas').read_bytes().decode('latin-1') == RECORD['exported']


def test_the_import_keeps_them_with_their_procedures(run: tuple[str, Path]) -> None:
    assert ExcelFile(run[1] / 'b.xlsm').get_module('Module1') == RECORD['module']
