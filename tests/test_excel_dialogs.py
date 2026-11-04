"""Excel's dialogs answered from the caller's queue, replayed from live Excel (scripts/measure_dialogs.py,
tests/fixtures/dialogs.json), and what happens with no answer."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBADialogError, VBASandboxError, VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/dialogs.json').read_text(encoding='utf-8'))
DRIVEN = re.compile(r'"drive\|(cancel|ok)(?:\|([^"]*))?"')
#: The caller's answer is what the dialog returned, so a saving dialog's filter adds no extension to it.
EXTENSIONS = ('FileDialog(2): x.Title = "drive|ok|', '"drive|ok|<DIR>\\new", FileFilter')
#: A folder that is not there: Excel moves the name to the user's Documents. pyOpenVBA does not look.
FOLDERS = ('name shared',)
#: Which workbook Execute saves on a saving dialog stayed unsettled, so it reports itself.
UNSETTLED = ('execute save', 'execute save names')


def _book(tmp_path: Path) -> ExcelApplication:
    (tmp_path / 'sub').mkdir(exist_ok=True)
    (tmp_path / 'pick.txt').write_text('x', encoding='ascii')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.interpreter.files.folder = str(tmp_path)
    return app


def _module(statement: str) -> str:
    return ('Private Function Shown(ByVal v As Variant) As String\n    On Error Resume Next\n    If IsError(v) Then\n        Shown = "Error:" & CStr(v)\n'
            '    ElseIf IsObject(v) Then\n        Shown = "object " & TypeName(v)\n    ElseIf IsArray(v) Then\n        Shown = "array " & TypeName(v)\n    Else\n'
            '        Shown = TypeName(v) & ":" & CStr(v)\n    End If\nEnd Function\n'
            'Public Function Tried() As String\n    Dim r As Variant, x As Object, y As Object, wb As Workbook\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
            '    wb.Worksheets(1).Range("A1").Value = 7\n    On Error GoTo Failed\n'
            f'    {statement}\n    Tried = "ok " & Shown(r)\nAnswered:\n    On Error Resume Next\n    wb.Close False\n    Exit Function\nFailed:\n'
            '    Tried = "Err " & Err.Number & " " & Err.Description\n    Resume Answered\nEnd Function\n')


def _answers(statement: str, folder: str, *, retried: bool) -> list[object]:
    """What the driver did, in the queue's terms; a form that turned its text down was cancelled the second time."""
    queue: list[object] = []
    for found in DRIVEN.finditer(statement):
        typed = found.group(2)
        queue.append(False if found.group(1) == 'cancel' else None if typed is None else typed.replace('<DIR>', folder).replace('@DIR@', folder))
    if retried or not queue:
        queue.append(False)
    return queue


@pytest.mark.skipif(sys.platform != 'win32', reason='the paths and the current folder are Windows ones')
def test_each_case_in_order_as_excel_answered(tmp_path: Path) -> None:
    """In one application and in order, as the file dialog keeps its state from one case to the next."""
    app = _book(tmp_path)
    folder = str(tmp_path)
    wrong: list[tuple[int, str, str, str]] = []
    for index, (statement, answer, _alerts) in enumerate(RECORD['cases']):
        shown = statement.replace('Set x = Application.FileDialog', 'FileDialog')
        if any(mark in shown for mark in EXTENSIONS) or answer.startswith('breaks'):
            continue
        app.interpreter.answers[:] = _answers(statement, folder, retried=True)
        name = f'Case{index}'
        app.add_module(_module(statement.replace('<DIR>', folder)), name=name)
        try:
            got = str(app.run(f'{name}.Tried')).replace(folder, '<DIR>')
        except VBAUnsupportedError as failed:
            got = f'unsupported: {failed}'
        if got != answer and not (answer.startswith('Err 424') and got.startswith('Err 424')):  # the wording is the interpreter's
            wrong.append((index, statement, answer, got))
    assert not wrong, f'{len(wrong)} wrong: {wrong[:25]}'


@pytest.mark.skipif(sys.platform != 'win32', reason='the paths and the current folder are Windows ones')
@pytest.mark.parametrize('name', [name for name in RECORD['fresh'] if name not in FOLDERS + UNSETTLED])
def test_a_new_session_as_excel_answered(tmp_path: Path, name: str) -> None:
    statement, answer = RECORD['fresh'][name]
    app = _book(tmp_path)
    folder = str(tmp_path)
    if 'book.xlsx' in statement:
        made = ExcelApplication(sandbox=False)
        made.add_workbook()
        made.save(tmp_path / 'book.xlsx')
    body = statement.replace('@DIR@', folder)
    app.interpreter.answers[:] = _answers(statement, folder, retried=False) if name != 'execute pickers' else [str(tmp_path / 'book.xlsx'), folder]
    source = ('Private Function d(ByVal k As Long) As Object\n    Set d = Application.FileDialog(k)\nEnd Function\nPublic Function Tried() As String\n'
              '    Dim s As String, k As Long, n As Long, m As Long, x As Object, y As Variant, r As Variant\n    On Error GoTo Failed\n'
              f'    {body}\n    Tried = s\n    Exit Function\nFailed:\n    Tried = s & "Err " & Err.Number & " " & Err.Description\nEnd Function\n')
    app.add_module(source, name='Fresh')
    assert str(app.run('Fresh.Tried')).replace(folder, '<DIR>') == answer


@pytest.mark.parametrize('call', ['InputBox("p")', 'Application.InputBox("p")', 'Application.GetOpenFilename()', 'Application.GetSaveAsFilename()',
                                  'Application.FileDialog(1).Show', 'Application.Dialogs(1).Show'])
def test_no_answer_stops_the_run_past_on_error(call: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Function Tried()\n    On Error Resume Next\n    Tried = {call}\nEnd Function\n', name='M')
    with pytest.raises(VBADialogError, match='answer_dialogs'):
        app.run('M.Tried')


def test_msgbox_with_no_answer_still_takes_its_default_button() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Tried()\n    Tried = MsgBox("p", vbYesNo + vbDefaultButton2)\nEnd Function\n', name='M')
    assert app.run('M.Tried') == 7


def test_a_refused_answer_with_nothing_after_it_names_what_was_refused() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Tried()\n    Tried = Application.InputBox("p", Type:=1)\nEnd Function\n', name='M')
    app.answer_dialogs('abc')
    with pytest.raises(VBADialogError, match="turned down 'abc'"):
        app.run('M.Tried')


def test_a_refused_answer_takes_the_next_and_the_transcript_has_one_dialog() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Tried()\n    Tried = Application.InputBox("p", "t", Type:=1)\nEnd Function\n', name='M')
    app.answer_dialogs('abc', '12')
    assert app.run('M.Tried') == 12.0
    assert app.dialogs == [('Application.InputBox', 'p', 't', 1)]


def test_answers_of_the_wrong_kind_are_refused_clearly() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function A()\n    A = Application.GetOpenFilename()\nEnd Function\nPublic Function B()\n    B = Application.Dialogs(1).Show\nEnd Function\n', name='M')
    app.answer_dialogs(5)
    with pytest.raises(VBADialogError, match='GetOpenFilename'):
        app.run('M.A')
    app.answer_dialogs(True)
    with pytest.raises(Exception, match='what OK does'):
        app.run('M.B')


def test_inputbox_cancel_and_ok_on_the_default() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Tried()\n    Tried = "[" & InputBox("p", "t", "d") & "]"\nEnd Function\n', name='M')
    app.answer_dialogs(False, None, 'typed')
    assert [app.run('M.Tried') for _ in range(3)] == ['[]', '[d]', '[typed]']
    assert app.dialogs[0] == ('InputBox', 'p', 't', 0)


def test_several_files_come_in_a_1_based_array_and_need_multi_select() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function A()\n    Dim v As Variant\n    v = Application.GetOpenFilename(MultiSelect:=True)\n    A = LBound(v) & UBound(v) & v(2)\nEnd Function\n'
                   'Public Function B()\n    With Application.FileDialog(1)\n        .Show\n        B = .SelectedItems.Count & .SelectedItems(2)\n    End With\nEnd Function\n'
                   'Public Function C()\n    C = Application.GetOpenFilename()\nEnd Function\n', name='M')
    app.answer_dialogs(['a.xlsx', 'b.xlsx'], ['c.xlsx', 'd.xlsx'], ['e', 'f'])
    assert app.run('M.A') == '12b.xlsx'
    assert app.run('M.B') == '2d.xlsx'
    with pytest.raises(VBADialogError, match='one path'):
        app.run('M.C')


def test_a_chosen_file_is_only_text_until_something_opens_it(tmp_path: Path) -> None:
    """The sandbox lets a macro choose a path and stops it there."""
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.save(tmp_path / 'book.xlsx')
    app = ExcelApplication()
    app.add_workbook()
    chosen = str(tmp_path / 'book.xlsx')
    app.add_module('Public Function Chose()\n    Chose = Application.GetOpenFilename()\nEnd Function\n'
                   'Public Sub Opened()\n    Workbooks.Open Application.GetOpenFilename()\nEnd Sub\n'
                   'Public Sub Executed()\n    With Application.FileDialog(1)\n        .Show\n        .Execute\n    End With\nEnd Sub\n', name='M')
    app.answer_dialogs(chosen, chosen, chosen)
    assert app.run('M.Chose') == chosen
    with pytest.raises(VBASandboxError):
        app.run('M.Opened')
    with pytest.raises(VBASandboxError):
        app.run('M.Executed')


def test_a_name_with_no_folder_needs_the_current_folder_which_the_sandbox_keeps() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Named()\n    Application.FileDialog(1).InitialFileName = "abc"\nEnd Sub\n', name='M')
    with pytest.raises(VBASandboxError):
        app.run('M.Named')
