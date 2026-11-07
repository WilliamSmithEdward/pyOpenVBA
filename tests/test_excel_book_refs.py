"""A formula naming a workbook the workbook had no link to, replayed from live Excel (scripts/measure_book_refs.py,
tests/fixtures/book_refs.json): what it reads, shows and lists with its source open and closed, what a save writes, and
what Excel read of the model's save. `excel_reads_model` is Excel's reading of the model's save, recorded when the
fixture was measured."""
from __future__ import annotations

import base64
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/book_refs.json').read_text(encoding='utf-8'))
CASES: list[list[str]] = RECORD['cases']
#: Excel names a new workbook by a count of its session, Sheet13 there; the model's count is its own.
NEW_BOOK = re.compile(r'(?<=\[)Sheet\d+(?=\])|(?<=links )Sheet\d+|(?<=Target=")Sheet\d+(?=")')


def _masked(text: str) -> str:
    return NEW_BOOK.sub('<book>', text)


def _clean(text: str, folder: Path) -> str:
    """The folder written <folder>. Off Windows a path in it has a slash after it, where Excel's has a backslash."""
    text = text.replace(f"file:///{str(folder).lstrip('/')}", 'file:///<folder>')
    text = re.sub(re.escape(str(folder)), '<folder>', text, flags=re.IGNORECASE).replace('<folder>/', '<folder>\\')
    # A link to a workbook above the folder it is saved in is written from the root with no drive.
    return re.sub(re.escape(os.path.splitdrive(str(folder))[1].replace('\\', '/')), '/<folder>', text, flags=re.IGNORECASE)


def _ran(folder: Path, index: int) -> str:
    for name, data in RECORD['files'].items():
        (folder / name).parent.mkdir(parents=True, exist_ok=True)
        (folder / name).write_bytes(base64.b64decode(data))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    return _clean(str(app.run('Probe.Tried', index, str(folder))), folder)


@pytest.mark.parametrize('index', range(len(CASES)), ids=[str(at) for at in range(len(CASES))])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    assert _masked(_ran(tmp_path, index)) == _masked(CASES[index][1])


@pytest.mark.parametrize('index', sorted(map(int, RECORD['excel_reads_model'])), ids=str)
def test_excel_read_the_model_save_as_its_own(index: int) -> None:
    assert _masked(RECORD['excel_reads_model'][str(index)]) == _masked(CASES[index][1].rpartition('||')[2])


@pytest.mark.skipif(os.name != 'nt', reason="Excel's targets are Windows paths: a drive, and \\ inside file:///")
@pytest.mark.parametrize('index', sorted(map(int, RECORD['saved'])), ids=str)
def test_a_save_writes_the_links_as_excel_does(tmp_path: Path, index: int) -> None:
    _ran(tmp_path, index)
    host = next(one for one in (tmp_path / 'Host.xlsx', tmp_path / 'Sub/Host.xlsx') if one.is_file())
    with zipfile.ZipFile(host) as saved:
        parts = {name: _clean(saved.read(name).decode('utf-8'), tmp_path) for name in saved.namelist()}
    wanted: dict[str, str] = RECORD['saved'][str(index)]
    links = sorted(name for name in parts if 'externalLink' in name)
    assert links == sorted(name for name in wanted if 'externalLink' in name)
    for name in links:
        assert _masked(parts[name]) == _masked(wanted[name]), name
    assert re.findall(r'<f>[^<]*</f>', parts['xl/worksheets/sheet1.xml']) == re.findall(r'<f>[^<]*</f>', wanted['xl/worksheets/sheet1.xml'])
    for name, pattern in (('xl/workbook.xml', r'<externalReference '), ('[Content_Types].xml', r'externalLink\d+\.xml"'),
                          ('xl/_rels/workbook.xml.rels', r'/externalLink"')):
        assert len(re.findall(pattern, parts[name])) == len(re.findall(pattern, wanted[name])), name


def test_under_the_sandbox_a_closed_workbook_is_not_read(tmp_path: Path) -> None:
    """The file is read past the gate that UpdateLink's is, which the sandbox shuts before anything is looked for."""
    (tmp_path / 'Alpha.xlsx').write_bytes(base64.b64decode(RECORD['files']['Alpha.xlsx']))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Put(ByVal folder As String)\n'
                   '    ActiveSheet.Range("A1").Formula = "=\'" & folder & "\\[Alpha.xlsx]Sheet1\'!A1"\nEnd Sub\n', name='Writer')
    with pytest.raises(VBASandboxError):
        app.run('Writer.Put', str(tmp_path))
    assert app.workbook.linked_books == {}


def test_a_name_alone_is_not_looked_for_under_the_sandbox() -> None:
    """A workbook named by its file alone would be looked for in the current folder: the sandbox refuses that too."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Put()\n    ActiveSheet.Range("A1").Formula = "=[NotHere.xlsx]Sheet1!A1"\nEnd Sub\n', name='Writer')
    with pytest.raises(VBASandboxError):
        app.run('Writer.Put')
    assert app.workbook.linked_books == {}


#: Each link's target, its workbook's folder, and the folder and full name the link gives: a POSIX folder joined with /, a Windows one with \\.
#: The full name is what a linked file is read and checked against granted folders by (see _hooks.link_gate).
SPELLINGS = [
    # Off Windows: a link saved relative to its workbook, read back in a POSIX folder, stays a POSIX folder (cases 15 and 25 on Linux).
    ('Sub/Deep.xlsx', '/tmp/x', '/tmp/x/Sub', '/tmp/x/Sub/Deep.xlsx'),
    ('../Alpha.xlsx', '/tmp/x/Sub', '/tmp/x', '/tmp/x/Alpha.xlsx'),
    ('Alpha.xlsx', '/tmp/x', '/tmp/x', '/tmp/x/Alpha.xlsx'),
    ('/tmp/y/Alpha.xlsx', '/tmp/x', '/tmp/y', '/tmp/y/Alpha.xlsx'),
    # On Windows the same links, as before.
    ('Sub/Deep.xlsx', 'C:\\x', 'C:\\x\\Sub', 'C:\\x\\Sub\\Deep.xlsx'),
    ('../Alpha.xlsx', 'C:\\x\\Sub', 'C:\\x', 'C:\\x\\Alpha.xlsx'),
    ('Alpha.xlsx', 'C:\\x', 'C:\\x', 'C:\\x\\Alpha.xlsx'),
    ('/Users/y/Alpha.xlsx', 'C:\\x', 'C:\\Users\\y', 'C:\\Users\\y\\Alpha.xlsx'),
    ('file:///C:\\y\\Alpha.xlsx', 'C:\\x', 'C:\\y', 'C:\\y\\Alpha.xlsx'),
]


@pytest.mark.parametrize(('target', 'here', 'folder', 'whole'), SPELLINGS)
@pytest.mark.parametrize('system', ['nt', 'posix'])
def test_a_link_s_folder_and_full_name_keep_the_spelling_of_its_workbook_s(target: str, here: str, folder: str, whole: str, system: str,
                                                                            monkeypatch: pytest.MonkeyPatch) -> None:
    """The answers go by the spellings alone, so they are the same as if each system ran them. On Linux the full name once came out as
    /tmp/x\\Alpha.xlsx, a file named x\\Alpha.xlsx in /tmp, so UpdateLink was refused a file in its granted folder."""
    import ntpath
    import posixpath

    from pyopenvba.apps.excel._external_links import LinkedBook

    linked = LinkedBook(1, target, ('Sheet1',), ({},), {})
    with monkeypatch.context() as patched:
        patched.setattr(os, 'name', system)
        patched.setattr(os, 'path', ntpath if system == 'nt' else posixpath)
        found, named = linked.folder(here), linked.path(here)
    assert (found, named) == (folder, whole)
