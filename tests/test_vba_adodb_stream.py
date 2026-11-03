"""ADODB.Stream, replayed from live Excel (scripts/measure_adodb_stream.py, tests/fixtures/adodb_stream.json): what each
case says, and the files its folder ends up with."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/adodb_stream.json').read_text(encoding='utf-8'))
#: The charsets that answer by the machine, which the model refuses.
MACHINE = {'text ansi', 'text x user', 'load auto', 'load auto all'}
KEPT = [index for index, case in enumerate(RECORD['cases']) if case[0] not in MACHINE]


def masked(text: str, folder: Path) -> str:
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
    application.add_module(RECORD['source'], name='StreamProbe')
    return application


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(app: ExcelApplication, tmp_path: Path, index: int) -> None:
    name, answer, files = RECORD['cases'][index]
    folder = tmp_path / 'case'
    folder.mkdir()
    seed(folder, name)
    try:
        got = str(app.run('StreamProbe.Probe', str(folder), os.sep, index))
        assert masked(got, folder) == answer
        assert files_of(folder) == files
    finally:
        for path in folder.rglob('*'):
            os.chmod(path, stat.S_IREAD | stat.S_IWRITE | (stat.S_IEXEC if path.is_dir() else 0))


def _module(body: str, *, sandbox: bool = True) -> ExcelApplication:
    app = ExcelApplication(sandbox=sandbox)
    app.add_workbook()
    app.add_module('Public Function Go()\n    Set st = CreateObject("ADODB.Stream")\n'
                   '    st.Charset = "utf-8"\n    st.Open\n'
                   f'{body}\nEnd Function\n', name='Probe')
    return app


def test_in_memory_it_works_in_the_sandbox() -> None:
    """Nothing here reaches out, so charset conversion needs no leave."""
    app = _module('    st.WriteText "a"\n    Go = st.Size')
    assert app.run('Probe.Go') == 4


@pytest.mark.parametrize('statement', ['st.SaveToFile "a.txt"', 'st.LoadFromFile "a.txt"'])
def test_the_sandbox_refuses_what_reaches_out(statement: str) -> None:
    with pytest.raises(VBASandboxError):
        _module(f'    {statement}').run('Probe.Go')


@pytest.mark.parametrize('name', ['_autodetect', '_autodetect_all', 'x-user-defined', 'nonsense'])
def test_a_charset_not_known_is_refused(name: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        _module(f'    st.Close\n    st.Charset = "{name}"').run('Probe.Go')
