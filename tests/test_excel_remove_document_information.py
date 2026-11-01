"""Workbook.RemoveDocumentInformation replayed from live Excel (scripts/measure_remove_document_information.py,
tests/fixtures/remove_document_information.json): what the document properties read after each removal, and the
core and app properties, custom properties and workbookPr the saved file holds."""
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
    (Path(__file__).parent / 'fixtures/remove_document_information.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def saved(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        core = package.read('docProps/core.xml').decode('utf-8')
        app = package.read('docProps/app.xml').decode('utf-8')
        custom = package.read('docProps/custom.xml').decode('utf-8') if 'docProps/custom.xml' in package.namelist() \
            else ''
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    core = re.sub(r'(<dcterms:(?:created|modified)\b[^>]*>)[^<]*', r'\1TIME', core)
    found = re.search(r'<workbookPr\b[^>]*/>', workbook)
    return [core, app, custom, found.group() if found else '']


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('remove_document_information')
    app = ExcelApplication()
    app.application.vba_set('UserName', 'William Smith')
    app.add_workbook()
    app.add_module(SOURCE, name='RemovalProbe')
    return str(app.run('RemovalProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_file_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    core, app, custom, workbook = saved(run[1] / f'c{index}.xlsx')
    assert [core, app, workbook] == [RECORD['cases'][index][2], RECORD['cases'][index][3], RECORD['cases'][index][5]]
    assert bool(custom) == bool(RECORD['cases'][index][4])
