"""ExcelApplication.allow_link_reads: the folders a macro under the sandbox may read linked workbooks' files in, checked as
set_hook's folders are (apps/excel/_hooks.is_inside). The workbooks are those of tests/fixtures/book_refs.json."""
from __future__ import annotations

import base64
import json
import os
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/book_refs.json').read_text(encoding='utf-8'))
MACROS = '''Public Function Put(ByVal formula As String) As String
    ActiveSheet.Range("A1").Formula = formula
    Put = ActiveSheet.Range("A1").Formula & "=" & CStr(ActiveSheet.Range("A1").Value)
End Function

Public Function Changed(ByVal before As String, ByVal after As String) As String
    ActiveWorkbook.ChangeLink before, after, xlLinkTypeExcelLinks
    Changed = ActiveSheet.Range("A1").Formula
End Function

Public Function Updated(ByVal name As String) As String
    ActiveWorkbook.UpdateLink name, xlExcelLinks
    Updated = CStr(ActiveSheet.Range("A1").Value)
End Function

Public Function Status(ByVal name As String) As Long
    Status = ActiveWorkbook.LinkInfo(name, xlLinkInfoStatus)
End Function
'''


def _books(folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    for name in ('Alpha.xlsx', 'Other.xlsx'):
        (folder / name).write_bytes(base64.b64decode(RECORD['files'][name]))
    return folder


def _app(*folders: Path) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MACROS, name='Links')
    if folders:
        app.allow_link_reads(folders)
    return app


def _alpha(folder: Path) -> str:
    return f"='{folder}{os.sep}[Alpha.xlsx]Sheet1'!A1"


def test_a_granted_folder_is_read(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books')
    assert str(_app(folder).run('Links.Put', _alpha(folder))).endswith('=1')


def test_a_folder_not_granted_is_refused_and_nothing_is_made(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books')
    app = _app(tmp_path / 'elsewhere')
    with pytest.raises(VBASandboxError) as raised:
        app.run('Links.Put', _alpha(folder))
    assert str(folder) not in str(raised.value)
    assert app.workbook.linked_books == {}


def test_a_folder_sharing_the_granted_one_s_name_as_a_prefix_is_refused(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books2')
    with pytest.raises(VBASandboxError):
        _app(tmp_path / 'books').run('Links.Put', _alpha(folder))


def test_a_link_inside_the_granted_folder_out_of_it_is_refused(tmp_path: Path) -> None:
    away = _books(tmp_path / 'away')
    granted = tmp_path / 'granted'
    granted.mkdir()
    try:
        (granted / 'link').symlink_to(away, target_is_directory=True)
    except OSError:
        pytest.skip('this system lets no one make a link')
    with pytest.raises(VBASandboxError):
        _app(granted).run('Links.Put', _alpha(granted / 'link'))


@pytest.mark.skipif(sys.platform != 'win32', reason='a junction is a Windows link')
def test_a_junction_inside_the_granted_folder_out_of_it_is_refused(tmp_path: Path) -> None:
    """A junction needs no privilege, so this runs wherever Windows does."""
    away = _books(tmp_path / 'away')
    granted = tmp_path / 'granted'
    granted.mkdir()
    if sys.platform == 'win32':
        import _winapi

        _winapi.CreateJunction(str(away), str(granted / 'joined'))
    with pytest.raises(VBASandboxError):
        _app(granted).run('Links.Put', _alpha(granted / 'joined'))


def test_change_link_update_link_and_status_read_inside_the_granted_folder(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books')
    app = _app(folder)
    app.run('Links.Put', _alpha(folder))
    alpha, other = str(folder / 'Alpha.xlsx'), str(folder / 'Other.xlsx')
    assert str(app.run('Links.Updated', alpha)) == '1'
    # Excel's backslash before the file, on every system, as the formula shows a link.
    assert str(app.run('Links.Changed', alpha, other)) == f"='{folder}\\[Other.xlsx]Elsewhere'!A1"
    assert int(str(app.run('Links.Status', other))) != 1  # xlLinkStatusMissingFile: the file was read, not refused


def test_change_link_to_a_folder_not_granted_is_refused(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books')
    away = _books(tmp_path / 'away')
    app = _app(folder)
    app.run('Links.Put', _alpha(folder))
    with pytest.raises(VBASandboxError):
        app.run('Links.Changed', str(folder / 'Alpha.xlsx'), str(away / 'Other.xlsx'))


def test_each_call_replaces_the_folders_and_none_takes_them_away(tmp_path: Path) -> None:
    folder = _books(tmp_path / 'books')
    app = _app(tmp_path / 'elsewhere')
    app.allow_link_reads([folder])
    app.run('Links.Put', _alpha(folder))
    app.allow_link_reads(())
    with pytest.raises(VBASandboxError):
        app.run('Links.Put', _alpha(folder).replace('Alpha', 'Other').replace('Sheet1', 'Elsewhere'))


def test_a_name_alone_is_looked_for_in_the_folder_the_application_was_made_in(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """A macro under the sandbox cannot change the current folder, so a workbook named by its file alone is looked for in the
    folder this process was in when the application was made, and read only if that folder is granted."""
    folder = _books(tmp_path / 'books')
    monkeypatch.chdir(folder)
    app = _app(folder)
    monkeypatch.chdir(tmp_path)
    assert str(app.run('Links.Put', '=[Alpha.xlsx]Sheet1!A1')).endswith('=1')
    refused = _app(tmp_path / 'elsewhere')
    with pytest.raises(VBASandboxError):
        refused.run('Links.Put', '=[Alpha.xlsx]Sheet1!A1')


def test_folders_must_be_a_list() -> None:
    with pytest.raises(TypeError):
        ExcelApplication().allow_link_reads('C:\\books')


@pytest.mark.parametrize('opening', ['Workbooks.Open folder & "\\Alpha.xlsx"', 'ActiveWorkbook.OpenLinks folder & "\\Alpha.xlsx"'])
def test_a_granted_folder_does_not_let_a_macro_open_a_workbook(tmp_path: Path, opening: str) -> None:
    """The grant is for reading a link's file: opening a workbook keeps its own gate, which only the link read's token passes."""
    folder = _books(tmp_path / 'books')
    app = _app(folder)
    app.run('Links.Put', _alpha(folder))
    app.add_module(f'Public Sub Opened(ByVal folder As String)\n    {opening}\nEnd Sub\n', name='Opener')
    with pytest.raises(VBASandboxError):
        app.run('Opener.Opened', str(folder))
    assert [book.name for book in app.workbooks()] == [app.workbook.name]


@pytest.mark.parametrize('route', ['/..\\away\\', '\\..\\away\\', '/sub\\..\\..\\away\\', '/./..\\away/'])
def test_a_backslashed_way_out_of_the_granted_folder_is_refused(tmp_path: Path, route: str) -> None:
    """Off Windows a backslash is a character of a name: "grant/..\\away" checked as given is one file inside grant, which a read turning
    its backslashes round finds in away. The gate checks the path the read opens (_hooks.final_path), so each is refused on every system;
    on Linux and macOS the first once read away's Alpha.xlsx."""
    grant, away = tmp_path / 'grant', _books(tmp_path / 'away')
    grant.mkdir()
    (grant / 'sub').mkdir()
    app = _app(grant)
    with pytest.raises(VBASandboxError):
        app.run('Links.Put', f"='{grant}{route}[Alpha.xlsx]Sheet1'!A1")
    assert app.workbook.linked_books == {}
    assert (away / 'Alpha.xlsx').is_file()


def test_a_grant_reached_through_a_link_keeps_the_path_as_given(tmp_path: Path) -> None:
    """As macOS reaches /var through /private/var: the gate checks the resolved path, and the formula, LinkSources and ChangeLink keep the one the
    macro gave, as Excel does."""
    real = tmp_path / 'real'
    (real / 'books').mkdir(parents=True)
    try:
        (tmp_path / 'link').symlink_to(real, target_is_directory=True)
    except OSError:
        pytest.skip('this system lets no one make a link')
    given = _books(tmp_path / 'link' / 'books')
    app = _app(given)
    assert str(app.run('Links.Put', _alpha(given))) == f"='{given}\\[Alpha.xlsx]Sheet1'!A1=1"
    assert str(app.run('Links.Changed', str(given / 'Alpha.xlsx'), str(given / 'Other.xlsx'))) == f"='{given}\\[Other.xlsx]Elsewhere'!A1"
    assert str(real) not in str(app.workbook.linked_books[1].target)


@pytest.mark.parametrize('system', ['nt', 'posix'])
def test_the_gate_checks_the_one_path_the_read_opens(system: str, monkeypatch: pytest.MonkeyPatch) -> None:
    """final_path turns a macro's path into the system's once, backslashes and .. included, and link_gate hands that same string to the read."""
    import ntpath
    import posixpath

    from pyopenvba.apps.excel import _hooks

    given, wanted = {'nt': ('C:\\grant\\..\\away\\Alpha.xlsx', 'C:\\away\\Alpha.xlsx'),
                     'posix': ('/grant/..\\away\\Alpha.xlsx', '/away/Alpha.xlsx')}[system]
    with monkeypatch.context() as patched:
        patched.setattr(os, 'name', system)
        patched.setattr(os, 'sep', '\\' if system == 'nt' else '/')
        patched.setattr(os, 'path', ntpath if system == 'nt' else posixpath)
        found = _hooks.final_path(given)
    assert (found.replace('/', '\\') if system == 'nt' else found) == wanted
    app = ExcelApplication()
    _, final = _hooks.link_gate(app.application, 'Reading a linked workbook', str(Path('grant') / 'Alpha.xlsx'))
    assert final == _hooks.final_path(str(Path('grant') / 'Alpha.xlsx'))


def test_the_loader_takes_no_token_but_the_gate_s() -> None:
    from pyopenvba.apps.excel._hooks import LinkReadPassed, passed

    assert not passed(True)
    assert not passed(LinkReadPassed())
