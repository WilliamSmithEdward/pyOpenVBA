"""Small properties of PivotTable, PivotCache, PivotField and PivotItem over a worksheet range, replayed from live
Excel (scripts/measure_pivot_properties.py, tests/fixtures/pivot_properties.json): what each reads on a new table,
what a set does, and the attributes each save writes against a save with nothing set."""
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
    (Path(__file__).parent / 'fixtures/pivot_properties.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
IGNORED = ('xr:uid', 'refreshedDate', 'refreshedBy', 'cacheId', 'r:id')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], list[str], Path]:
    folder = tmp_path_factory.mktemp('pivot_properties')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='PropertiesProbe')
    reads, *answers = str(app.run('PropertiesProbe.Probe', str(folder))).split('@@@')
    return reads.split('~'), answers, folder


@pytest.mark.parametrize('index', range(len(RECORD['reads'])),
                         ids=[f'{key}.{name}' for key, name, _ in RECORD['reads']])
def test_each_read_as_excel_answered(run: tuple[list[str], list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['reads'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['sets'])))
def test_each_set_as_excel_answered(run: tuple[list[str], list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['sets'][index][1]


def _attributes(path: Path) -> dict[str, dict[str, str]]:
    found: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith(('xl/pivotTables/pivotTable', 'xl/pivotCache/pivotCacheDefinition')):
                text = package.read(name).decode('utf-8')
                root = re.search(r'<(?:pivotCacheDefinition|pivotTableDefinition)\b[^>]*>', text)
                kept = dict(re.findall(r'([\w:]+)="([^"]*)"', root.group(0) if root else ''))
                for index, field in enumerate(re.findall(r'<pivotField\b[^>]*>', text)):
                    kept.update({f'field{index}.{key}': value
                                 for key, value in re.findall(r'([\w:]+)="([^"]*)"', field)})
                found[name.rsplit('/', 1)[-1]] = {key: value for key, value in kept.items()
                                                  if key not in IGNORED and not key.startswith('xmlns')}
    return found


@pytest.mark.parametrize('index', [index for index, (_, _, change) in enumerate(RECORD['sets']) if change is not None])
def test_each_save_changes_what_excels_did(run: tuple[list[str], list[str], Path], index: int) -> None:
    plain = _attributes(run[2] / f'c{len(RECORD["sets"])}.xlsx')
    found: dict[str, dict[str, list[str | None]]] = {}
    for part, now in _attributes(run[2] / f'c{index}.xlsx').items():
        before = plain.get(part, {})
        differ = {key: [before.get(key), now.get(key)] for key in sorted(set(before) | set(now))
                  if before.get(key) != now.get(key)}
        if differ:
            found[part] = differ
    assert found == RECORD['sets'][index][2]
