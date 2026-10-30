"""A UserForm is reported as unsupported, never as a runtime error On Error can trap (issue #251): Load and Unload,
and any use of the form's name, Show and its properties among them. A project's own procedure named Load still runs.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

PROBE = '''Public Function Tried(ByVal which As Long) As String
    On Error Resume Next
    Select Case which
    Case 0: Load frmTest
    Case 1: frmTest.Show
    Case 2: Tried = frmTest.Caption
    Case 3: Unload frmTest
    End Select
    Tried = "trapped " & Err.Number
End Function
'''


@pytest.fixture(scope='module')
def book(tmp_path_factory: pytest.TempPathFactory) -> Path:
    target = tmp_path_factory.mktemp('forms') / 'forms.xlsm'
    with ExcelFile.create_new(target) as workbook:
        workbook.add_form('frmTest')
        workbook.save()
    return target


@pytest.mark.parametrize('which', range(4), ids=['Load', 'Show', 'Caption', 'Unload'])
def test_a_userform_is_reported_unsupported(book: Path, which: int) -> None:
    app = ExcelApplication.open(book)
    app.add_module(PROBE, name='Probe')
    with pytest.raises(VBAUnsupportedError, match='UserForm'):
        app.run('Probe.Tried', which)


def test_a_procedure_named_load_still_runs() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Dim n As Long\nPublic Function Tried() As Long\nLoad 3\nTried = n\nEnd Function\n'
                   'Private Sub Load(ByVal x As Long)\nn = x\nEnd Sub\n', name='Own')
    assert app.run('Own.Tried') == 3
