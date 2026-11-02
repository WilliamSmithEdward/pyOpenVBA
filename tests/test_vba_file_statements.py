"""VBA's text files and file system statements, replayed from live Excel (scripts/measure_file_statements.py,
tests/fixtures/file_statements.json): what each case says, and the files its folder ends up with."""
from __future__ import annotations

import json
import os
import stat
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError
from pyopenvba.interpreter import VBASandboxError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/file_statements.json').read_text(encoding='utf-8'))


def seed(folder: Path, name: str) -> None:
    for text in RECORD['seeds'][name]:
        _, file_name, *data = text.split(' ')
        target = folder / file_name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(bytes.fromhex(data[0] if data else ''))


def files_of(folder: Path) -> str:
    found: list[str] = []
    for path in sorted(folder.rglob('*'), key=lambda one: one.as_posix().casefold()):
        relative = path.relative_to(folder).as_posix()
        found.append(f'{relative}/' if path.is_dir() else f'{relative}={path.read_bytes().hex()}')
    return ' '.join(found)


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    application = ExcelApplication(sandbox=False)
    application.add_workbook()
    application.add_module(RECORD['source'], name='FileProbe')
    return application


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case[0] for case in RECORD['cases']])
def test_each_case_as_excel_answered(app: ExcelApplication, tmp_path: Path, index: int) -> None:
    name, answer, files = RECORD['cases'][index]
    folder = tmp_path / 'case'
    folder.mkdir()
    seed(folder, name)
    try:
        got = str(app.run('FileProbe.Probe', str(folder), os.sep, index))
        assert got.replace(str(folder), 'FOLDER') == answer
        assert files_of(folder) == files
    finally:
        for path in folder.rglob('*'):
            os.chmod(path, stat.S_IREAD | stat.S_IWRITE | (stat.S_IEXEC if path.is_dir() else 0))


@pytest.mark.parametrize('statement', ['Kill "a"', 'Open "a" For Output As #1', 'x = FreeFile', 'x = Dir("*")', 'MkDir "a"',
                                       'x = Environ("PATH")', 'Print #1, "a"', 'Close', 'x = CurDir', 'x = EOF(1)'])
def test_the_sandbox_refuses_it(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    {statement}\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


@pytest.mark.parametrize('statement', ['ChDrive "D"'])
def test_what_is_not_implemented_yet_says_so(statement: str, tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    {statement}\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go')
    assert not (tmp_path / 'a').exists()


def test_line_input_into_a_number_does_not_compile(tmp_path: Path) -> None:
    """Excel refuses it when it compiles the procedure, measured as a Compile error dialog."""
    (tmp_path / 'a.txt').write_bytes(b'1\r\n')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go(path)\n    Dim k As Long\n    Open path For Input As #1\n    Line Input #1, k\nEnd Sub\n',
                   name='Probe')
    with pytest.raises(VBACompileError):
        app.run('Probe.Go', str(tmp_path / 'a.txt'))


def test_a_name_that_is_also_a_statement_still_works(tmp_path: Path) -> None:
    """Name, Close and Input stay ordinary members everywhere else."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go()\n    ActiveSheet.Name = "Fresh"\n    With ActiveSheet\n        .Name = "Mine"\n'
                   '    End With\n    Go = ActiveSheet.Name\nEnd Function\n', name='Probe')
    assert app.run('Probe.Go') == 'Mine'
    assert not list(tmp_path.iterdir())
