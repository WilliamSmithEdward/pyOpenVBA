"""Workbook settings the workbook part keeps, replayed from live Excel (scripts/measure_workbook_options.py,
tests/fixtures/workbook_options.json): DisplayDrawingObjects, DisplayInkComments, RemovePersonalInformation,
EnableAutoRecover and ReadOnlyRecommended, whether setting each leaves the workbook unsaved, and the fileSharing,
workbookPr and fileRecoveryPr a save writes."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/workbook_options.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: The elements of the workbook part the model writes for these settings.
OWNED = re.compile(r'<(?:fileSharing|workbookPr|fileRecoveryPr)\b')


def test_each_case_as_excel_answered(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='OptionsProbe')
    got = str(app.run('OptionsProbe.Probe', str(tmp_path))).split('|')
    assert got == [answer for _, _, answer in RECORD['cases']]
    for name, elements in RECORD['parts'].items():
        with zipfile.ZipFile(tmp_path / name) as package:
            workbook = package.read('xl/workbook.xml').decode('utf-8')
        mine = re.findall(r'<(?:fileSharing|workbookPr|fileRecoveryPr)\b[^>]*>', workbook)
        assert mine == [element for element in elements if OWNED.match(element)], name
