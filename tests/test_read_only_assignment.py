"""Assigning to a member the model only reads: a gap where the real class lets a macro set it, VBA's error 438 where
the type library marks it read-only."""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._inventory import member_settable

SETUP = ('    Dim co As Object\n'
         '    Set co = ActiveSheet.ChartObjects.Add(10, 10, 200, 100)\n'
         '    co.Chart.SetSourceData ActiveSheet.Range("A1:A2")\n')


def _run(statement: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Probe() As String\n    On Error GoTo Failed\n' + SETUP + f'    {statement}\n'
                   '    Probe = "set"\n    Exit Function\nFailed:\n    Probe = "E" & Err.Number\nEnd Function\n',
                   name='Assign')
    return app.run('Assign.Probe')


def test_the_inventory_knows_which_properties_a_macro_can_set() -> None:
    assert member_settable('ChartObject', 'ProtectChartObject', 'Excel')
    assert member_settable('Workbook', 'password', 'Excel')
    assert not member_settable('Workbook', 'PasswordEncryptionKeyLength', 'Excel')
    assert not member_settable('Shape', 'ID', 'Excel')


@pytest.mark.parametrize('statement', ['co.ProtectChartObject = True', 'ActiveWorkbook.Password = "p"',
                                       'co.Chart.ProtectData = True'])
def test_a_property_excel_lets_a_macro_set_is_reported_as_a_gap(statement: str) -> None:
    with pytest.raises(VBAUnsupportedError, match='Setting'):
        _run(statement)


@pytest.mark.parametrize('statement', ['ActiveWorkbook.PasswordEncryptionKeyLength = 5',
                                       'ActiveSheet.Shapes(1).ID = 5'])
def test_a_read_only_property_raises_438(statement: str) -> None:
    assert _run(statement) == 'E438'
