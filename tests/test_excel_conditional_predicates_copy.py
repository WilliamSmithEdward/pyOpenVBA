"""Predicate settings and differential styles survive worksheet copies and saves."""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Workbook

CASES = [
    ('Type:=9,String:="foo",TextOperator:=0', '11000'),
    ('Type:=9,String:="foo",TextOperator:=1', '00111'),
    ('Type:=9,String:="fo",TextOperator:=2', '11000'),
    ('Type:=9,String:="oo",TextOperator:=3', '10000'),
    ('Type:=10', '00010'),
    ('Type:=13', '11101'),
    ('Type:=16', '00001'),
    ('Type:=17', '11110'),
]
INSPECT = '''Function InspectPredicate() As String
Dim s As Worksheet,c As Range,f As FormatCondition
For Each s In Worksheets
Set f=s.Range("A1").FormatConditions(1)
InspectPredicate=InspectPredicate & s.Name & ":" & f.Type & ":" & f.Formula1 & ":" & f.AppliesTo.Address & ":"
For Each c In s.Range("A1:A5")
If c.DisplayFormat.Font.Bold Then
InspectPredicate=InspectPredicate & "1"
Else
InspectPredicate=InspectPredicate & "0"
End If
Next
InspectPredicate=InspectPredicate & ":" & f.NumberFormat & "|"
Next
End Function'''


def copied_predicate(arguments: str, external: bool, path: Path) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub MakePredicate()
Dim f As FormatCondition
ActiveSheet.Name="Data"
Range("A1").Value="foo"
Range("A2").Value="Food"
Range("A3").Value="bar"
Range("A5").Value=CVErr(2042)
Set f=Range("A1:A5").FormatConditions.Add(''' + arguments + ''')
f.Font.Bold=True
f.NumberFormat="0.00 ""hit"""
End Sub''', name='SetupPredicate')
    app.run('MakePredicate')
    app.add_module('Sub CopyPredicate()\n' + ('ActiveSheet.Copy' if external else
                   'ActiveSheet.Copy After:=ActiveSheet') + '\nEnd Sub', name='CopyPredicateModule')
    app.run('CopyPredicate')
    app.add_module(INSPECT, name='PredicateInspection')
    if external:
        target = app.application.ActiveWorkbook()
        assert isinstance(target, Workbook)
        app.add_module(INSPECT, name='PredicateInspection', workbook=target)
    result = app.run('InspectPredicate')
    assert isinstance(result, str)
    app.save(path)
    assert ExcelApplication.open(path).run('InspectPredicate') == result
    return result


@pytest.mark.parametrize('external', [False, True])
@pytest.mark.parametrize(('arguments', 'bits'), CASES)
def test_predicate_copy_and_roundtrip(arguments: str, bits: str, external: bool, tmp_path: Path) -> None:
    result = copied_predicate(arguments, external, tmp_path / 'predicate-copy.xlsm')
    pieces = result.rstrip('|').split('|')
    assert len(pieces) == (1 if external else 2)
    for piece in pieces:
        assert piece.endswith(':$A$1:$A$5:' + bits + ':0.00 "hit"')
    if not external:
        assert pieces[1].removeprefix('Data (2):') == pieces[0].removeprefix('Data:')


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize(('arguments', 'bits'), CASES)
def test_native_reopens_copied_predicates(arguments: str, bits: str, tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'native-predicate-copy.xlsm'
    expected = copied_predicate(arguments, True, path)
    assert ':' + bits + ':' in expected
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('PredicateInspection.InspectPredicate') == expected
        assert not any(event.get('dialogs') for event in excel.events)
