"""Range.Cut through a table, replayed from live Excel (scripts/measure_table_cut.py, tests/fixtures/table_cut.json): what
each case reads, the cells of its sheets and where the table part puts the table."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_list_unlist import formats
from test_excel_list_unlist import cells as cells_of

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/table_cut.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Cutting over hidden cells, which the model refuses.
LEFT_OUT: tuple[str, ...] = ('o.Range.AutoFilter',)
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


def cells(sheet: str) -> str:
    if '</sheetData>' not in sheet:
        return ''
    found = cells_of(sheet)
    return '' if found == '<sheetData>' else found


def table_says(part: str) -> tuple[str, ...]:
    """Where the part puts the table and what it names the columns."""
    found = re.search(r'<table\b[^>]*?\bref="([^"]+)"', part)
    assert found is not None
    filtered = re.search(r'<autoFilter ref="([^"]+)"', part)
    names = re.findall(r'<tableColumn [^>]*?\bname="([^"]*)"', part)
    return (found.group(1), filtered.group(1) if filtered else '', *names)


def saved(path: Path) -> dict[str, Any]:
    """The cells of each sheet, the numbering of their text aside, and what its tables say."""
    found: dict[str, Any] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if re.fullmatch(r'xl/worksheets/sheet\d\.xml', name):
                found[name] = _unnumbered(cells(package.read(name).decode('utf-8')))
            elif re.fullmatch(r'xl/tables/table\d+\.xml', name):
                found[name] = table_says(package.read(name).decode('utf-8'))
            elif name == 'xl/styles.xml':
                found[name] = formats(package.read(name).decode('utf-8'))
    return found


def _unnumbered(block: str) -> str:
    """Cells without the row's spans and with their shared text's numbers gone, which depend on what else the file
    holds."""
    return re.sub(r' spans="[^"]*"|(?<=t="s"><v>)\d+', '', block)


def wanted(held: dict[str, str]) -> dict[str, Any]:
    def said(name: str, text: str) -> Any:
        return (table_says(text) if 'tables/' in name else formats(text) if name == 'xl/styles.xml'
                else _unnumbered(cells(text)))

    return {name: said(name, text) for name, text in held.items()}


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('table_cut')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='CutProbe')
    return {index: str(app.run('CutProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert saved(run[1] / f'c{index}.xlsx') == wanted(RECORD['cases'][index][2])
