"""ChartObject's Visible, Locked, PrintObject, Placement, BringToFront and SendToBack, and ChartObjects' settings,
replayed from live Excel (scripts/measure_chart_object_more.py, tests/fixtures/chart_object_more.json): what each
case answers, the shapes' order, and each saved anchor's opening tag, cNvPr and clientData."""
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
    (Path(__file__).parent / 'fixtures/chart_object_more.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def anchor_settings(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found: list[str] = []
    for anchor in re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S):
        parts = [re.search(pattern, anchor) for pattern in
                 (r'<xdr:twoCellAnchor\b[^>]*>', r'<xdr:cNvPr\b[^>]*>', r'<xdr:clientData\b[^>]*/>')]
        found.append(''.join(part.group(0) for part in parts if part))
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_object_more')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ChartObjectProbe')
    return str(app.run('ChartObjectProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_anchor_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert anchor_settings(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
