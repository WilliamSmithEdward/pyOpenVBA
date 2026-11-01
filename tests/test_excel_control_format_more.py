"""ControlFormat's DropDownLines, Enabled, ListIndex, LockedText and PrintObject, replayed from live Excel
(scripts/measure_control_format_more.py, tests/fixtures/control_format_more.json): what each case answers, and the
settings the control's saved ctrlProp part and VML client data hold for them."""
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
    (Path(__file__).parent / 'fixtures/control_format_more.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def control_parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        found = [package.read(name).decode('utf-8') for name in names if name.startswith('xl/ctrlProps/')]
        for name in names:
            if name.endswith('.vml'):
                found += re.findall(r'<x:ClientData\b.*?</x:ClientData>', package.read(name).decode('utf-8'), re.S)
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('control_format_more')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ControlProbe')
    return str(app.run('ControlProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


def settings(parts: list[str]) -> dict[str, str | None]:
    """What these properties write: the ctrlProp's lockText, dropLines and sel, and the VML's Disabled, PrintObject,
    LockText, DropLines and Sel. The rest of the control's markup, its anchor in cells among it, is not this test's."""
    properties, vml = parts
    found: dict[str, str | None] = {}
    for name in ('lockText', 'dropLines', 'sel'):
        value = re.search(rf'\b{name}="([^"]*)"', properties)
        found[name] = value.group(1) if value else None
    found['Disabled'] = 'yes' if '<x:Disabled/>' in vml else None
    for name in ('PrintObject', 'LockText', 'DropLines', 'Sel'):
        value = re.search(rf'<x:{name}>([^<]*)</x:{name}>', vml)
        found[name] = value.group(1) if value else None
    return found


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_control_setting_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert settings(control_parts(run[1] / f'c{index}.xlsx')) == settings(RECORD['cases'][index][2])
