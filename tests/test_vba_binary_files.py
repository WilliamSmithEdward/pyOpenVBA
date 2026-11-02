"""VBA's Binary and Random files, replayed from live Excel (scripts/measure_binary_files.py,
tests/fixtures/binary_files.json): what each case says, and the bytes of every file it leaves."""
from __future__ import annotations

import json
import os
import stat
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter import VBASandboxError
from test_vba_file_statements import files_of

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/binary_files.json').read_text(encoding='utf-8'))


def seed(folder: Path, name: str) -> None:
    for text in RECORD['seeds'][name]:
        _, file_name, *data = text.split(' ')
        target = folder / file_name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(bytes.fromhex(data[0] if data else ''))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    application = ExcelApplication(sandbox=False)
    application.add_workbook()
    application.add_module(RECORD['source'], name='BinaryProbe')
    return application


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case[0] for case in RECORD['cases']])
def test_each_case_as_excel_answered(app: ExcelApplication, tmp_path: Path, index: int) -> None:
    name, answer, files = RECORD['cases'][index]
    folder = tmp_path / 'case'
    folder.mkdir()
    seed(folder, name)
    try:
        got = str(app.run('BinaryProbe.Probe', str(folder), os.sep, index))
        assert got.replace(str(folder), 'FOLDER') == answer
        assert files_of(folder) == files
    finally:
        for path in folder.rglob('*'):
            os.chmod(path, stat.S_IREAD | stat.S_IWRITE | (stat.S_IEXEC if path.is_dir() else 0))


@pytest.mark.parametrize('statement', ['Get #1, , x', 'Put #1, , x', 'Lock #1', 'Unlock #1'])
def test_the_sandbox_refuses_it(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    {statement}\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


@pytest.mark.parametrize('body', [
    'Dim x() As Integer\n    ReDim x(1, 1)\n    Open path For Random As #1 Len = 30\n    Put #1, , x',
    'Dim x As Variant\n    Open path For Binary As #1\n    Put #1, , CreateObject("Scripting.Dictionary")',
])
def test_what_was_not_measured_is_refused(body: str, tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go(path)\n    {body}\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go', str(tmp_path / 'a.bin'))
