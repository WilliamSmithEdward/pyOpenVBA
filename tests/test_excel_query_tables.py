"""Query tables over local text files, replayed from live Excel (scripts/measure_query_tables.py,
tests/fixtures/query_tables.json): what each case answers, the sheet's cells, and the parts it saves."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/query_tables.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused, which each case below checks.
REFUSED = ('q.PostText = "x"', 'q.Connection = "URL;', 'q.TextFilePromptOnRefresh = True', '"URL;http', 'q.TextFilePlatform = 1252', 'q.TextFilePlatform = 65001',
           'TypeName(q.Parameters)', 'TypeName(q.Sort)', 'TypeName(q.WorkbookConnection)', 'q.WorkbookConnection.Name',
           'Parent.Connections', 'ws.Copy')
#: What is measured and not modelled, so not replayed: widths after AdjustColumnWidth, which a 144-DPI display decides; what a
#: refresh does to formats in its way; a name changed by hand; a table's cells all deleted; and a connection Excel saves for a
#: table deleted before any refresh.
LEFT_OUT = ('.ColumnWidth', 'Interior.ColorIndex', 'A2").NumberFormat', 'ws.Names(1).Delete', 'ws.Names(1).RefersTo =', 'ws.Rows("1:3").Delete', 'r = q.Delete & ')
#: Connections answers again once the table has gone.
ANSWERED = 'q.Delete: r = ws.Names.Count'
REFUSALS = [index for index, name in enumerate(NAMES) if any(text in name for text in REFUSED) and not any(text in name for text in LEFT_OUT) and ANSWERED not in name]
KEPT = [index for index, name in enumerate(NAMES) if index not in REFUSALS and not any(text in name for text in LEFT_OUT)]


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
        if os.sep != '\\':
            line = line.replace('\\', os.sep) if 'folder &' in line else line
        written.append(line)
    return '\n'.join(written)


def _plain(text: str) -> str:
    """A part without Excel's random ids and the namespaces only they need."""
    text = re.sub(r'<\?xml.*?\?>\s*', '', text)
    text = re.sub(r' xmlns:\w+="[^"]*"', '', text)
    text = re.sub(r' mc:Ignorable="[^"]*"', '', text)
    return re.sub(r' xr16:uid="[^"]*"', '', text)


def _parts(path: Path, folder: Path) -> dict[str, str]:
    found: dict[str, str] = {}
    if not path.exists():
        return found
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('xl/queryTables/') or name == 'xl/connections.xml' or name.endswith('sheet1.xml.rels'):
                found[name] = _plain(package.read(name).decode('utf-8').replace(str(folder), '<folder>'))
            elif name == 'xl/workbook.xml':
                found['names'] = ''.join(re.findall(r'<definedName .*?</definedName>', package.read(name).decode('utf-8')))
    return found


def _recorded(index: int) -> dict[str, str]:
    saved = RECORD['cases'][index][2]
    found = {name: _plain(text) for name, text in saved.items() if name.startswith('xl/queryTables/') or name == 'xl/connections.xml' or name.endswith('sheet1.xml.rels')}
    if saved:
        found['names'] = ''.join(re.findall(r'<definedName .*?</definedName>', saved['xl/workbook.xml']))
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('query_tables')
    for name, text in RECORD['files'].items():
        (folder / name).write_bytes(text.encode('ascii'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(set(KEPT)), name='QueryProbe')
    return {index: str(app.run('QueryProbe.Probe', str(folder), index, index)).replace(str(folder), '<folder>') for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_saves_the_parts_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    assert _parts(run[1] / f'c{index}.xlsx', run[1]) == _recorded(index)


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    for name, text in RECORD['files'].items():
        (tmp_path / name).write_bytes(text.encode('ascii'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='QueryProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('QueryProbe.Probe', str(tmp_path), index, index)


def test_refreshing_is_refused_in_the_sandbox(tmp_path: Path) -> None:
    (tmp_path / 'd.csv').write_bytes(b'a\r\n')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Go(ByVal path As String)
    ActiveSheet.QueryTables.Add("TEXT;" & path, ActiveSheet.Range("A1")).Refresh False
End Sub
''', name='Go')
    with pytest.raises(Exception, match='outside the engine'):
        app.run('Go.Go', str(tmp_path / 'd.csv'))
