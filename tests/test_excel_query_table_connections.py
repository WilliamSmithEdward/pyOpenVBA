"""The connection behind a query table over a text file, replayed from live Excel
(scripts/measure_query_table_connections.py, tests/fixtures/query_table_connections.json)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/query_table_connections.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused because each of these stopped Excel when measured: setting a delimiter or a URL through TextConnection, and
#: naming a connection its table left behind.
STOPPED = [index for index, case in enumerate(RECORD['cases']) if case[1].startswith('STOPPED')]
#: Refused at the save: Excel's file then ties the first table to the new table's connection.
UNSAVED = [index for index, name in enumerate(NAMES) if 'q.WorkbookConnection: c.Delete: Set o = ' in name and 'Refresh False' in name]
#: Not built: the one text member a query table lacks, and a left-behind connection read from a file.
UNBUILT = [index for index, name in enumerate(NAMES) if 'TextFileHeaderRow' in name or 'book.Connections(1).TextConnection' in name]
REFUSALS = STOPPED + UNSAVED + UNBUILT
#: Answered as Excel does; saving a connection Connections.Add2 added reports itself, as it does without a query table.
ADDED = [index for index, name in enumerate(NAMES) if 'Connections.Add2' in name]
KEPT = [index for index in range(len(NAMES)) if index not in REFUSALS and index not in ADDED]


def _source(kept: set[int]) -> str:
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case ') and not stripped.endswith(' which, folder, ws, r'):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        if os.sep != '\\' and 'folder &' in line:
            line = line.replace('\\', os.sep)
        written.append(line)
    return '\n'.join(written)


def _spelled(text: str, folder: Path) -> str:
    """The folder as the fixture spells it, with the separator after it as recorded on Windows."""
    return text.replace(str(folder), '<folder>').replace(f'<folder>{os.sep}', '<folder>\\')


def _plain(text: str) -> str:
    """A part without Excel's random ids and the namespaces only they need."""
    text = re.sub(r'<\?xml.*?\?>\s*', '', text)
    text = re.sub(r' xmlns:(?!x15)\w+="[^"]*"', '', text)
    text = re.sub(r' mc:Ignorable="[^"]*"', '', text)
    return re.sub(r' xr16:uid="[^"]*"', '', text)


def _numbered_apart(found: dict[str, str]) -> dict[str, object]:
    """What parts say whatever Excel numbers a book's tables, which differs from save to save (see test_excel_query_tables_saved)."""
    said: dict[str, object] = {name: text for name, text in found.items() if not name.startswith('xl/queryTables/')}
    said['tables'] = sorted(text for name, text in found.items() if name.startswith('xl/queryTables/'))
    return said


def _parts(path: Path, folder: Path) -> dict[str, str]:
    found: dict[str, str] = {}
    if not path.exists():
        return found
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('xl/queryTables/') or name == 'xl/connections.xml':
                found[name] = _plain(_spelled(package.read(name).decode('utf-8'), folder))
            elif name == 'xl/workbook.xml':
                found['names'] = ''.join(re.findall(r'<definedName .*?</definedName>', package.read(name).decode('utf-8')))
    return found


def _recorded(index: int) -> dict[str, str]:
    saved = RECORD['cases'][index][2]
    found = {name: _plain(text) for name, text in saved.items() if name.startswith('xl/queryTables/') or name == 'xl/connections.xml'}
    if saved:
        found['names'] = ''.join(re.findall(r'<definedName .*?</definedName>', saved['xl/workbook.xml']))
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('query_table_connections')
    for name, text in RECORD['files'].items():
        (folder / name).write_bytes(text.encode('ascii'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(set(KEPT)), name='QueryProbe')
    return {index: _spelled(str(app.run('QueryProbe.Probe', str(folder), index, index)), folder) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[f'{index}: {NAMES[index]}' for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[f'{index}: {NAMES[index]}' for index in KEPT])
def test_each_case_saves_the_parts_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    assert _numbered_apart(_parts(run[1] / f'c{index}.xlsx', run[1])) == _numbered_apart(_recorded(index))


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_stops_excel_or_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    for name, text in RECORD['files'].items():
        (tmp_path / name).write_bytes(text.encode('ascii'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='QueryProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('QueryProbe.Probe', str(tmp_path), index, index)


@pytest.mark.parametrize('index', ADDED, ids=[NAMES[index] for index in ADDED])
def test_a_connection_added_sorts_and_names_with_the_tables(tmp_path: Path, index: int) -> None:
    for name, text in RECORD['files'].items():
        (tmp_path / name).write_bytes(text.encode('ascii'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    source = '\n'.join(line for line in _source({index}).split('\n') if '.SaveAs ' not in line)
    app.add_module(source, name='QueryProbe')
    assert _spelled(str(app.run('QueryProbe.Probe', str(tmp_path), index, index)), tmp_path) == RECORD['cases'][index][1]


def test_excel_loses_the_first_sheets_table_every_time() -> None:
    """The quirk the model copies, repeated in the fixture: three runs each way, with the same result."""
    for repeat in (NAMES[-10:-7], NAMES[-7:-4], NAMES[-4:-1]):
        assert len(set(repeat)) == 1
    answers = [case[1].split('|')[1] for case in RECORD['cases'][-10:-4]]
    assert answers == ['0 0 0'] * 3 + ['1 0 1'] * 3
