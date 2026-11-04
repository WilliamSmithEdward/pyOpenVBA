"""A workbook Excel saved with query tables, opened in the model (scripts/make_query_table_workbook.py,
tests/fixtures/query_table_saved.xlsx), and what Excel does with one (scripts/measure_query_tables_saved.py,
tests/fixtures/query_tables_saved.json)."""
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

FIXTURES = Path(__file__).parent / 'fixtures'
SAVED = FIXTURES / 'query_table_saved.xlsx'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'query_tables_saved.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What the model answers differently and refuses: Connections once the tables have changed.
REFUSED = ('.Connections.Count', 'w2.Parent.Connections')
REFUSALS = [index for index, name in enumerate(NAMES) if any(text in name for text in REFUSED)]
#: Left out: the order of a saved sheet's tables, which goes with Excel's numbering of them and changes from save to save; and
#: deleting a table on a copied sheet, after which Excel's file loses the original's same-named table and name, though the
#: original goes on answering that it has them. The model keeps the original's.
LEFT_OUT = ('QueryTables(1).Name & ', 'w2.QueryTables(1).Name', 'w2.QueryTables(2).Name', 'w2.QueryTables("ExternalData_1").Delete')
KEPT = [index for index in range(len(NAMES)) if index not in REFUSALS and not any(text in NAMES[index] for text in LEFT_OUT)]
PARTS = ('xl/connections.xml',)


def _source(kept: set[int]) -> str:
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


def _folder(folder: Path, change: str = '') -> Path:
    """The workbook Excel saved, in this folder with the files it reads, its C:\\data spelled as this folder."""
    for name, text in RECORD['files'].items():
        (folder / name).write_bytes(text.encode('ascii'))
    target = folder / 'saved.xlsx'
    with zipfile.ZipFile(SAVED) as package, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as written:
        for name in package.namelist():
            data = package.read(name)
            if name == 'xl/connections.xml':
                data = data.decode('utf-8').replace('C:\\data\\', f'{folder}{os.sep}').replace(change, change.replace('type="6"', 'type="1"') if change else change).encode('utf-8')
            written.writestr(name, data)
    return target


def _spelled(text: str, folder: Path) -> str:
    """The folder as the fixture spells it, with the separator after it as recorded on Windows."""
    return text.replace(str(folder), '<folder>').replace(f'<folder>{os.sep}', '<folder>\\')


def _plain(text: str) -> str:
    text = re.sub(r'<\?xml.*?\?>\s*', '', text)
    text = re.sub(r' xmlns:\w+="[^"]*"', '', text)
    text = re.sub(r' mc:Ignorable="[^"]*"', '', text)
    return re.sub(r' xr16:uid="[^"]*"', '', text)


def _said(parts: dict[str, str]) -> dict[str, object]:
    """What parts say whatever Excel numbers them: Excel's own numbering of a sheet's tables differs from run to run."""
    tables = sorted(_plain(text) for name, text in parts.items() if name.startswith('xl/queryTables/'))
    found: dict[str, object] = {'tables': tables, 'connections': _plain(parts.get('xl/connections.xml', ''))}
    found['names'] = sorted(re.findall(r'<definedName .*?</definedName>', parts.get('xl/workbook.xml', '')))
    found['related'] = len(re.findall(r'queryTable\d+\.xml', parts.get('xl/worksheets/_rels/sheet1.xml.rels', '')))
    return found


def _read(path: Path, folder: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: _spelled(package.read(name).decode('utf-8'), folder) for name in package.namelist()
                if name.startswith('xl/queryTables/') or name in ('xl/connections.xml', 'xl/workbook.xml', 'xl/worksheets/_rels/sheet1.xml.rels')}


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    _folder(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='SavedProbe')
    assert _spelled(str(app.run('SavedProbe.Probe', index, str(tmp_path))), tmp_path) == RECORD['cases'][index][1]
    assert _said(_read(tmp_path / f'c{index}.xlsx', tmp_path)) == _said(RECORD['cases'][index][2])
    alone = {name.removeprefix('copy/'): text for name, text in RECORD['cases'][index][2].items() if name.startswith('copy/')}
    if alone:
        assert _said(_read(tmp_path / f'o{index}.xlsx', tmp_path)) == _said(alone)


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    _folder(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='SavedProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('SavedProbe.Probe', index, str(tmp_path))


def _opened(path: Path) -> tuple[ExcelApplication, Any]:
    app = ExcelApplication(sandbox=False)
    return app, app.open_workbook(str(path))


def test_an_edit_elsewhere_leaves_the_query_table_parts_as_excel_wrote_them(tmp_path: Path) -> None:
    source = _folder(tmp_path)
    app, book = _opened(source)
    app.add_module('Public Sub Go()\n    Workbooks(1).Worksheets(1).Range("A21").Value = "edited"\nEnd Sub\n', name='Go')
    app.run('Go.Go')
    out = tmp_path / 'out.xlsx'
    book.save_as(str(out))
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(out) as after:
        for name in ('xl/connections.xml', 'xl/queryTables/queryTable1.xml', 'xl/queryTables/queryTable2.xml'):
            assert before.read(name) == after.read(name)
        assert after.read('xl/worksheets/_rels/sheet1.xml.rels').decode().count('queryTable"') == 2


def test_connections_read_while_the_tables_say_what_the_file_does(tmp_path: Path) -> None:
    app, _ = _opened(_folder(tmp_path))
    app.add_module('Public Function Go() As String\n    Go = Workbooks(1).Connections.Count & " " & Workbooks(1).Connections(1).Name\nEnd Function\n', name='Go')
    assert app.run('Go.Go') == '2 d'


def test_a_table_over_anything_else_stays_refused_and_is_saved_untouched(tmp_path: Path) -> None:
    source = _folder(tmp_path, 'name="s" type="6"')
    app, book = _opened(source)
    app.add_module('Public Function Go() As Long\n    Go = Workbooks(1).Worksheets(1).QueryTables.Count\nEnd Function\n', name='Go')
    with pytest.raises(VBAUnsupportedError, match='queryTable parts'):
        app.run('Go.Go')
    out = tmp_path / 'out.xlsx'
    book.save_as(str(out))
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(out) as after:
        for name in ('xl/connections.xml', 'xl/queryTables/queryTable1.xml', 'xl/queryTables/queryTable2.xml'):
            assert before.read(name) == after.read(name)
