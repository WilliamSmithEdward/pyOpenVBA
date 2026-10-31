"""PageSetup.PrintNotes replayed from live Excel (scripts/measure_print_notes.py, tests/fixtures/print_notes.json):
what it reads beside PrintComments, what a set leaves there, and the pageSetup's cellComments a save writes."""
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
    (Path(__file__).parent / 'fixtures/print_notes.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('print_notes')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='NotesProbe')
    return str(app.run('NotesProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][2]


def _settings(page_setup: str) -> dict[str, str]:
    """The pageSetup's settings for comments, errors and page order; a default one is left out."""
    return dict(re.findall(r'\b(cellComments|errors|pageOrder)="([^"]*)"', page_setup))


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_each_save_writes_the_settings_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    """Excel's pageSetup also names the printer's settings part, which the model does not write."""
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<pageSetup\b[^>]*>', sheet)
    assert _settings(found.group(0) if found else '') == _settings(RECORD['parts'][saved])
