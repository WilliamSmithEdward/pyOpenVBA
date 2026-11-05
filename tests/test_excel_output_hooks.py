"""set_hook, ExportAsFixedFormat and PrintOut. What the model answers itself is replayed from live Excel (scripts/measure_output_members.py, tests/fixtures/output_members.json), and a small fake
stands in for a tool: no PDF library, and no output claimed to match Excel."""
from __future__ import annotations

import datetime as _dt
import io
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from types import MappingProxyType
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest, is_inside
from pyopenvba.exceptions import VBAError, VBADialogError, VBARuntimeError, VBASandboxError, VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'output_members.json').read_text(encoding='utf-8'))
CASES: list[dict[str, Any]] = [one for one in RECORD['cases'] if one['member'] in ('ExportAsFixedFormat', 'PrintOut')]
OWNERS = {'Worksheet': 'wb.Worksheets(1)', 'Workbook': 'wb', 'Range': 'wb.Worksheets(1).Range("A1:B3")', 'Chart': 'wb.Charts(1)', 'Sheets': 'wb.Sheets', 'Window': 'wb.Windows(1)'}
#: The model has no pages to count, so From past the last one goes to the function, where Excel finds nothing to print on these.
NO_PAGES = {('ExportAsFixedFormat', 'from_past', 'Worksheet'), ('ExportAsFixedFormat', 'from_past', 'Range')}


def _book(folder: Path) -> tuple[ExcelApplication, list[HookRequest]]:
    """The workbook Excel saved for these, with a chart sheet, its current folder the test's."""
    copy = folder / 'book.xlsx'
    copy.write_bytes((FIXTURES / 'output_members.xlsx').read_bytes())
    app = ExcelApplication.open(copy, with_vba=False)
    app.interpreter.sandbox = False
    app.add_module(f'Public Sub Here()\n    ChDir "{folder}"\nEnd Sub\n', name='Here')
    app.run('Here.Here')
    heard: list[HookRequest] = []
    return app, heard


def _row(case: dict[str, Any], folder: Path) -> str:
    declared = case['owner'] if case['binding'] == 'Typed' else 'Object'
    return (f'Public Function Row(ByVal folder As String) As String\n    Dim wb As Workbook, o As {declared}, s As String\n    Set wb = ActiveWorkbook\n    Set o = {OWNERS[case["owner"]]}\n'
            f'    On Error Resume Next\n    {case["call"]}\n    Row = "E" & Err.Number\nEnd Function\n')


def _prepared(folder: Path) -> None:
    (folder / 'sub').mkdir()
    (folder / 'there.pdf').write_bytes(b'before')
    for name in ('ro.pdf', 'ro.prn'):
        (folder / name).write_bytes(b'before')
        (folder / name).chmod(0o444)


@pytest.mark.parametrize('case', CASES, ids=[f'{one["member"]} {one["owner"]} {one["binding"]} {one["row"]}' for one in CASES])
def test_each_call_answers_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    if case['excel'].startswith('failed'):
        pytest.skip('a form or a compile error in Excel')
    if (case['member'], case['row'], case['owner']) in NO_PAGES:
        pytest.xfail('The model has no pages, so this goes to the function')
    if (case['row'], case['owner'], case['binding']) == ('type_omitted', 'Range', 'Typed'):
        pytest.xfail('It runs to 450 in Excel; the compile step takes it from the library as not optional')
    _prepared(tmp_path)
    try:
        app, heard = _book(tmp_path)
        app.set_hook(case['member'], heard.append)
        app.add_module(_row(case, tmp_path), name='Row')
        got = str(app.run('Row.Row', str(tmp_path)))
    finally:
        for name in ('ro.pdf', 'ro.prn'):
            (tmp_path / name).chmod(0o666)
    assert got == case['excel'].split('~')[0]
    # The function hears of what Excel would have gone on to do, and of nothing it refused.
    assert bool(heard) == (got == 'E0')
    if case['member'] == 'ExportAsFixedFormat' and case['files'] and heard:
        made = heard[0].path
        assert made is not None
        assert made.relative_to(tmp_path).as_posix() == case['files'][0].removesuffix(' before')


def _fill(book: Any) -> None:
    book.sheets_[0].vba_get('Range', ['A1']).vba_set('Value', 1)


def _asks(request: HookRequest) -> None:
    request.saved_bytes()


def _keeping(got: list[bytes]) -> Any:
    def keep(request: HookRequest) -> None:
        got.append(request.saved_bytes())
    return keep


def _macro(app: ExcelApplication, body: str, *args: object) -> object:
    app.add_module(f'Public Function Go(ByVal folder As String) As Variant\n{body}\nEnd Function\n', name='Macro')
    return app.run('Macro.Go', *args)


def _sandboxed() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    _fill(app.workbook)
    return app


EXPORT = 'ActiveWorkbook.Worksheets(1).ExportAsFixedFormat Type:=0, Filename:={}, OpenAfterPublish:=False'


def test_with_no_function_it_reports_itself_and_nothing_is_written(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    with pytest.raises(VBAUnsupportedError, match='set_hook'):
        _macro(app, EXPORT.format('folder & "\\x.pdf"'), str(tmp_path))
    assert not list(tmp_path.iterdir())


def test_a_macro_argument_error_comes_first_even_with_no_function() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    _fill(app.workbook)
    with pytest.raises(VBARuntimeError, match='5'):
        _macro(app, 'Dim o As Object\nSet o = ActiveWorkbook.Worksheets(1)\no.ExportAsFixedFormat Type:=2, OpenAfterPublish:=False', '')


def _inside_folder(base: Path) -> dict[str, tuple[str, bool]]:
    inside, sibling = base / 'out', base / 'out2'
    inside.mkdir()
    sibling.mkdir()
    return {'in': (str(inside / 'x.pdf'), True), 'nested file in sub is refused unless there': (str(inside / 'nope' / 'x.pdf'), True),
            'dotdot': (str(inside / '..' / 'x.pdf'), False), 'sibling sharing the prefix': (str(sibling / 'x.pdf'), False), 'parent': (str(base / 'x.pdf'), False),
            'the folder itself': (str(inside), False)}


def test_is_inside_goes_by_whole_parts(tmp_path: Path) -> None:
    for label, (path, wanted) in _inside_folder(tmp_path).items():
        assert is_inside(path, (str(tmp_path / 'out'),)) is wanted, label


@pytest.mark.skipif(os.name != 'nt', reason='case on a disk that ignores it, and UNC and device paths, are Windows')
def test_is_inside_on_windows(tmp_path: Path) -> None:
    folder = tmp_path / 'out'
    folder.mkdir()
    assert is_inside(str(folder / 'x.pdf').upper(), (str(folder),))
    for path in (r'\\server\share\x.pdf', r'\\.\C:\x.pdf', r'\\?\C:\x.pdf', str(folder / 'NUL.pdf'), str(folder / 'x.pdf:stream')):
        assert not is_inside(path, (str(folder),)), path


def test_a_link_out_of_the_folder_is_refused(tmp_path: Path) -> None:
    folder, away = tmp_path / 'out', tmp_path / 'away'
    folder.mkdir()
    away.mkdir()
    try:
        (folder / 'link').symlink_to(away, target_is_directory=True)
    except OSError:
        pytest.skip('this system lets no one make a link')
    assert not is_inside(folder / 'link' / 'x.pdf', (str(folder),))


@pytest.mark.skipif(sys.platform != 'win32', reason='a junction is a Windows link')
def test_a_junction_out_of_the_folder_is_refused(tmp_path: Path) -> None:
    """A junction needs no privilege, so this runs wherever Windows does, unlike the link above."""
    folder, away = tmp_path / 'out', tmp_path / 'away'
    folder.mkdir()
    away.mkdir()
    if sys.platform == 'win32':
        import _winapi

        _winapi.CreateJunction(str(away), str(folder / 'joined'))
    assert not is_inside(folder / 'joined' / 'x.pdf', (str(folder),))
    assert is_inside(folder / 'joined' / 'x.pdf', (str(away),))


@pytest.mark.parametrize('name', ['"..\\away.pdf"', 'folder & "\\..\\away.pdf"', '"rel.pdf"', '""'])
def test_under_the_sandbox_the_function_hears_of_no_path_outside(tmp_path: Path, name: str) -> None:
    folder = tmp_path / 'out'
    folder.mkdir()
    app = _sandboxed()
    heard: list[HookRequest] = []
    app.set_hook('ExportAsFixedFormat', heard.append, folders=[folder])
    with pytest.raises(VBASandboxError):
        _macro(app, EXPORT.format(name), str(folder))
    assert not heard


def test_under_the_sandbox_with_no_folder_named_no_path_passes(tmp_path: Path) -> None:
    app = _sandboxed()
    heard: list[HookRequest] = []
    app.set_hook('ExportAsFixedFormat', heard.append)
    with pytest.raises(VBASandboxError):
        _macro(app, EXPORT.format('folder & "\\x.pdf"'), str(tmp_path))
    assert not heard


def test_under_the_sandbox_a_path_inside_reaches_the_function(tmp_path: Path) -> None:
    app = _sandboxed()
    heard: list[HookRequest] = []
    app.set_hook('ExportAsFixedFormat', heard.append, folders=[tmp_path])
    _macro(app, EXPORT.format('folder & "\\x"'), str(tmp_path))
    assert heard[0].path == tmp_path.resolve() / 'x.pdf' or heard[0].path == tmp_path / 'x.pdf'
    assert heard[0].owner == 'Worksheet'
    assert heard[0].arguments['OpenAfterPublish'] is False


def test_a_print_with_no_file_runs_under_the_sandbox_and_a_file_is_gated(tmp_path: Path) -> None:
    app = _sandboxed()
    heard: list[HookRequest] = []
    app.set_hook('PrintOut', heard.append)
    _macro(app, 'ActiveWorkbook.Worksheets(1).PrintOut Copies:=3, ActivePrinter:="Anyone"', '')
    assert heard[0].path is None
    assert dict(heard[0].arguments) == {'Copies': 3, 'ActivePrinter': 'Anyone'}
    heard.clear()
    with pytest.raises(VBASandboxError):
        app.add_module('Public Sub Two(ByVal folder As String)\n    ActiveWorkbook.Worksheets(1).PrintOut PrintToFile:=True, PrToFileName:=folder & "\\p.prn"\nEnd Sub\n', name='Two')
        app.run('Two.Two', str(tmp_path))
    assert not heard


def test_arguments_cannot_be_changed() -> None:
    app = _sandboxed()
    heard: list[HookRequest] = []
    app.set_hook('PrintOut', heard.append)
    _macro(app, 'ActiveWorkbook.Worksheets(1).PrintOut Copies:=3', '')
    assert isinstance(heard[0].arguments, MappingProxyType)
    with pytest.raises(TypeError):
        heard[0].arguments['Copies'] = 4  # type: ignore[index]  # pyright: ignore[reportIndexIssue]


def test_the_hook_runs_with_the_sheet_active_and_it_goes_back_even_if_it_raises() -> None:
    app = _sandboxed()
    app.add_module('Public Sub Second()\n    ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(1)).Range("A1").Value = 1\n    ActiveWorkbook.Worksheets(1).Activate\nEnd Sub\n', name='Second')
    app.run('Second.Second')
    seen: list[str] = []

    def fails(request: HookRequest) -> None:
        active = request.workbook.active_sheet
        assert active is not None
        seen.append(active.name)
        raise OSError('out of paper')

    app.set_hook('PrintOut', fails)
    source = 'Public Function Go() As String\n    On Error Resume Next\n    ActiveWorkbook.Worksheets(2).PrintOut\n    Go = Err.Number & "|" & ActiveSheet.Name\nEnd Function\n'
    app.add_module(source, name='Go')
    first = app.workbook.sheets_[0].name
    assert str(app.run('Go.Go')) == f'1004|{first}'
    assert seen == [app.workbook.sheets_[1].name]


def test_what_a_function_raises() -> None:
    app = _sandboxed()

    def raises_this(failure: BaseException) -> Any:
        def function(request: HookRequest) -> None:
            del request
            raise failure
        return function

    source = 'Public Function Go() As String\n    On Error Resume Next\n    ActiveWorkbook.Worksheets(1).PrintOut\n    Go = Err.Number\nEnd Function\n'
    app.add_module(source, name='Go')
    app.set_hook('PrintOut', raises_this(ValueError('no')))
    assert str(app.run('Go.Go')) == '1004'
    from pyopenvba.interpreter._values import error

    app.set_hook('PrintOut', raises_this(error(70, 'denied')))
    assert str(app.run('Go.Go')) == '70'
    app.set_hook('PrintOut', raises_this(VBADialogError('nobody there')))
    with pytest.raises(VBADialogError):
        app.run('Go.Go')
    app.set_hook('PrintOut', raises_this(KeyboardInterrupt()))
    with pytest.raises(KeyboardInterrupt):
        app.run('Go.Go')


def test_a_function_that_asks_for_its_own_member_is_refused() -> None:
    app = _sandboxed()
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets(1).PrintOut\nEnd Sub\n', name='Go')
    calls: list[int] = []

    def again(request: HookRequest) -> None:
        del request
        calls.append(1)
        app.run('Go.Go')

    app.set_hook('PrintOut', again)
    with pytest.raises(VBAUnsupportedError, match='its own hook'):
        app.run('Go.Go')
    assert calls == [1]


def _no_print(book: object) -> bool:
    del book
    return False


def test_cancel_stops_the_function(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    book = app.workbook
    from pyopenvba.apps.excel import _events

    _fill(book)
    original = _events.before_print
    setattr(_events, 'before_print', _no_print)
    heard: list[HookRequest] = []
    try:
        app.set_hook('PrintOut', heard.append)
        app.set_hook('ExportAsFixedFormat', heard.append)
        book.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
        with pytest.raises(VBARuntimeError, match='Document not saved'):
            book.sheets_[0].ExportAsFixedFormat(0, str(tmp_path / 'x.pdf'), OpenAfterPublish=False)  # pyright: ignore[reportAttributeAccessIssue]
    finally:
        setattr(_events, 'before_print', original)
    assert not heard


def test_a_print_to_a_file_with_no_name_takes_a_queued_answer_or_reports_itself(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    _fill(app.workbook)
    heard: list[HookRequest] = []
    app.set_hook('PrintOut', heard.append)
    with pytest.raises(VBADialogError):
        _macro(app, 'ActiveWorkbook.Worksheets(1).PrintOut PrintToFile:=True', '')
    assert not heard
    app.answer_dialogs(str(tmp_path / 'p.prn'))
    app.run('Macro.Go', '')
    assert heard[0].path == tmp_path / 'p.prn'


def test_saved_bytes_change_nothing_and_leave_nothing(tmp_path: Path) -> None:
    source = FIXTURES / 'old_collections.xlsx'
    copy = tmp_path / 'book.xlsx'
    copy.write_bytes(source.read_bytes())
    app = ExcelApplication.open(copy, with_vba=False)
    app.interpreter.sandbox = False
    book = app.workbook
    events: list[str] = []
    before = (book.vba_get('Saved'), book.vba_get('Path'), book.vba_get('FullName'), book.vba_get('Name'), copy.read_bytes())
    got: list[bytes] = []
    app.set_hook('PrintOut', _keeping(got))
    book.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    assert got[0][:2] == b'PK'
    assert (book.vba_get('Saved'), book.vba_get('Path'), book.vba_get('FullName'), book.vba_get('Name'), copy.read_bytes()) == before
    assert not events
    again = tmp_path / 'again.xlsx'
    again.write_bytes(got[0])
    assert ExcelApplication.open(again, with_vba=False).workbook.sheets_[0].name == book.sheets_[0].name


def test_saved_bytes_of_a_workbook_never_saved() -> None:
    app = ExcelApplication()
    app.add_workbook()
    got: list[bytes] = []
    app.set_hook('PrintOut', _keeping(got))
    app.workbook.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    assert got[0][:2] == b'PK'
    assert app.workbook.vba_get('Path') == ''


def _saved_book(tmp_path: Path, *, ask: bool, name: str) -> bytes:
    """A workbook from a file, a value changed, its saved bytes asked for or not, and then saved."""
    copy = tmp_path / f'{name}.xlsx'
    copy.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication.open(copy, with_vba=False)
    book = app.workbook
    cast(Any, book.sheets_[0].vba_get('Range', ['H30'])).vba_set('Formula', '=1+1')
    if ask:
        app.set_hook('PrintOut', _asks)
        book.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    app.freeze_clock(_FROZEN)
    target = tmp_path / f'{name}_saved.xlsx'
    app.save(target)
    return target.read_bytes()


_FROZEN = _dt.datetime(2026, 1, 2, 3, 4, 5)


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return {name: package.read(name) for name in package.namelist() if name != 'docProps/core.xml'}


def test_saved_bytes_leave_the_next_save_as_it_would_have_been(tmp_path: Path) -> None:
    """The calculation chain among them."""
    assert _parts(_saved_book(tmp_path, ask=True, name='asked')) == _parts(_saved_book(tmp_path, ask=False, name='not'))


def _noting(ran: list[str], name: str) -> Any:
    def note(*_: object, **__: object) -> bool:
        ran.append(name)
        return True
    return note


def test_saved_bytes_run_no_event(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _events

    app = _sandboxed()
    ran: list[str] = []
    for name in ('before_save', 'after_save'):
        monkeypatch.setattr(_events, name, _noting(ran, name))
    app.set_hook('PrintOut', _asks)
    app.workbook.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    assert not ran


def test_saved_bytes_leave_no_file_even_when_they_fail(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _io

    monkeypatch.setattr(tempfile, 'tempdir', str(tmp_path))
    app = _sandboxed()
    app.set_hook('PrintOut', _asks)
    app.workbook.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    assert not list(tmp_path.iterdir())

    def fails(*_: object, **__: object) -> None:
        raise OSError('no room')

    monkeypatch.setattr(_io, 'save_workbook', fails)
    with pytest.raises(VBARuntimeError):
        app.workbook.sheets_[0].PrintOut()  # pyright: ignore[reportAttributeAccessIssue]
    assert not list(tmp_path.iterdir())


def test_set_hook_takes_only_the_names_it_knows() -> None:
    app = ExcelApplication()
    with pytest.raises(ValueError, match='SendMail'):
        app.set_hook('SendMail', print)
    with pytest.raises(TypeError):
        app.set_hook('PrintOut', 'not a function')
    app.set_hook('PrintOut', print)
    app.set_hook('PrintOut', None)
    assert 'PrintOut' not in app.interpreter.hooks


def test_vbaerror_is_the_one_that_goes_on_as_it_is() -> None:
    assert issubclass(VBADialogError, VBAError)


PRINT_RETURNS: list[dict[str, Any]] = json.loads((FIXTURES / 'print_returns.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', PRINT_RETURNS, ids=lambda one: f'{one["owner"]} {one["cancel"]}')
def test_what_printout_returns(case: dict[str, Any], tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """True, and False where BeforePrint cancels, with or without a function."""
    from pyopenvba.apps.excel import _events

    app, heard = _book(tmp_path)
    app.set_hook('PrintOut', heard.append)
    if case['cancel']:
        monkeypatch.setattr(_events, 'before_print', _no_print)
    app.add_module(f'Public Function Row(ByVal folder As String) As String\n    Dim wb As Workbook, o As Object, r As Variant\n    Set wb = ActiveWorkbook\n    Set o = {OWNERS[case["owner"]]}\n'
                   '    r = o.PrintOut(PrintToFile:=True, PrToFileName:=folder & "\\p.prn")\n    Row = TypeName(r) & " " & CStr(r)\nEnd Function\n', name='Row')
    assert str(app.run('Row.Row', str(tmp_path))) == case['answer']
    assert bool(heard) != case['cancel']


EMBEDDED: dict[str, Any] = json.loads((FIXTURES / 'output_embedded_chart.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', [one for one in EMBEDDED['cases'] if not one['excel'].startswith('failed')], ids=lambda one: f'{one["member"]} {one["binding"]} {one["row"]}')
def test_a_chart_in_a_frame_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    _prepared(tmp_path)
    try:
        app, heard = _book(tmp_path)
        app.set_hook(case['member'], heard.append)
        declared = 'Chart' if case['binding'] == 'Typed' else 'Object'
        app.add_module(f'Public Function Row(ByVal folder As String) As String\n    Dim wb As Workbook, o As {declared}, s As String\n    Set wb = ActiveWorkbook\n'
                       f'    wb.Worksheets(1).ChartObjects.Add(10, 10, 200, 100).Chart.SetSourceData wb.Worksheets(1).Range("A1:A5")\n    Set o = {EMBEDDED["get"]}\n'
                       f'    On Error Resume Next\n    {case["call"]}\n    Row = "E" & Err.Number\nEnd Function\n', name='Row')
        got = str(app.run('Row.Row', str(tmp_path)))
    finally:
        for name in ('ro.pdf', 'ro.prn'):
            (tmp_path / name).chmod(0o666)
    assert got == case['excel'].split('~')[0]
    assert bool(heard) == (got == 'E0')


def test_a_chart_in_a_frame_returns_as_excel_did(tmp_path: Path) -> None:
    app, heard = _book(tmp_path)
    app.set_hook('PrintOut', heard.append)
    app.add_module('Public Function Row() As String\n    Dim o As Object, r As Variant\n    ActiveWorkbook.Worksheets(1).ChartObjects.Add(10, 10, 200, 100).Chart.SetSourceData ActiveWorkbook.Worksheets(1).Range("A1:A5")\n'
                   f'    Set o = ActiveWorkbook.{EMBEDDED["get"].removeprefix("wb.")}\n    r = o.PrintOut\n    Row = TypeName(r) & " " & CStr(r)\nEnd Function\n', name='Row')
    assert str(app.run('Row.Row')) == EMBEDDED['returns']['False']
    assert len(heard) == 1
