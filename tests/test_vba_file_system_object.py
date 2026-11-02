"""Scripting.FileSystemObject, replayed from live Excel (scripts/measure_file_system_object.py,
tests/fixtures/file_system_object.json): what each case says, and the files its folder ends up with."""
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

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/file_system_object.json').read_text(encoding='utf-8'))
#: What only Windows can answer: a drive letter, and its own folders.
WINDOWS_ONLY = {'root', 'drive exists', 'special folder'}
KEPT = [index for index, case in enumerate(RECORD['cases']) if os.name == 'nt' or case[0] not in WINDOWS_ONLY]


def masked(text: str, folder: Path) -> str:
    """The answer with the folder and its parent taken out, and one kind of separator, as the probe has them."""
    for path, word in ((str(folder), 'FOLDER'), (str(folder.parent), 'ROOT')):
        text = text.replace(path, word).replace(path.replace('\\', '\\u005C'), word)
    return text.replace('\\u005C', '/')


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
    application.add_module(RECORD['source'], name='FsoProbe')
    return application


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(app: ExcelApplication, tmp_path: Path, index: int) -> None:
    name, answer, files = RECORD['cases'][index]
    folder = tmp_path / 'case'
    folder.mkdir()
    seed(folder, name)
    try:
        got = str(app.run('FsoProbe.Probe', str(folder), os.sep, index))
        assert masked(got, folder) == answer
        assert files_of(folder) == files
    finally:
        for path in folder.rglob('*'):
            os.chmod(path, stat.S_IREAD | stat.S_IWRITE | (stat.S_IEXEC if path.is_dir() else 0))


@pytest.mark.parametrize('member', ['FileExists("a")', 'CreateTextFile("a")', 'OpenTextFile("a")', 'GetFile("a")',
                                    'GetFolder("a")', 'CreateFolder("a")', 'DeleteFile("a")', 'DeleteFolder("a")',
                                    'CopyFile("a", "b")',
                                    'MoveFile("a", "b")', 'GetTempName', 'GetSpecialFolder(2)'])
def test_the_sandbox_refuses_it(member: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    x = CreateObject("Scripting.FileSystemObject").{member}\nEnd Sub\n',
                   name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


def test_the_path_functions_stay_open_in_the_sandbox() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go()\n    Go = CreateObject("Scripting.FileSystemObject").GetBaseName("a.txt")\n'
                   'End Function\n',
                   name='Probe')
    assert app.run('Probe.Go') == 'a'


@pytest.mark.parametrize('read', ['f.ShortName', 'f.ShortPath', 'f.Type', 'f.Drive', 'fso.Drives',
                                  'fso.GetDrive("C:")'])
def test_what_comes_from_the_machine_is_refused(read: str, tmp_path: Path) -> None:
    (tmp_path / 'a.txt').write_bytes(b'a')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go(path)\n    Set fso = CreateObject("Scripting.FileSystemObject")\n'
                   f'    Set f = fso.GetFile(path)\n    x = {read}\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go', str(tmp_path / 'a.txt'))
