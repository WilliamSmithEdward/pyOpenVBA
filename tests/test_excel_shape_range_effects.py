"""A ShapeRange's Shadow, Glow, SoftEdge and Reflection, replayed from live Excel (scripts/measure_shape_range_effects.py,
tests/fixtures/shape_range_effects.json)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_range_effects.json').read_text(encoding='utf-8'))
#: A preset shadow reports itself on one shape as well, so these are left out of the replay.
PRESET = [index for index, case in enumerate(RECORD['cases']) if 'Shadow.Type' in case[0]]
INDEXES = [index for index in range(len(RECORD['cases'])) if index not in PRESET]


def _source() -> str:
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for index in PRESET:
        source = re.sub(rf'    Case {index}\n        [^\n]*\n', f'    Case {index}\n', source)
    return source


IDS = [f'{index}: {RECORD["cases"][index][0]}' for index in INDEXES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('shape_range_effects')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='EffectProbe')
    return str(app.run('EffectProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', INDEXES, ids=IDS)
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=IDS)
def test_each_shape_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL) == RECORD['cases'][index][2]
