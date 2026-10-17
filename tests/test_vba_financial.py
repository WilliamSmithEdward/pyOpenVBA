"""Native VBA financial arithmetic, convergence and argument boundaries."""
from __future__ import annotations
import importlib.util
import json
import os
from pathlib import Path
from typing import Any
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/financial.json').read_text(encoding='utf-8'))

@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_financial(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    if case.get('compile_error'):
        with pytest.raises(VBACompileError):
            app.add_module(case['source'], name='Harness')
            app.run('Harness.Probe')
    else:
        app.add_module(case['source'], name='Harness')
        actual = app.run('Harness.Probe')
        if isinstance(case['result'], (int, float)):
            assert actual == pytest.approx(case['result'], rel=2e-13, abs=1e-14)
        else:
            assert actual == case['result']


def test_wrong_financial_array_fails_before_body_side_effects() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Probe()\nRange("A1").Value2="executed"\nDim values(2) As Integer\n'
                   'Dim ignored As Double\nignored=IRR(values)\nEnd Sub', name='Harness')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    cell = app.evaluate('Range("A1")')
    assert isinstance(cell, VBAObject)
    assert cell.vba_get('Value2') is EMPTY


def test_user_function_can_shadow_financial_library() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Rate(ByVal text As String) As String\nRate=text\nEnd Function\n'
                   'Function Probe() As String\nProbe=Rate("mine")\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 'mine'


def test_edited_array_type_invalidates_financial_argument_check() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public values(2) As Double', name='Other')
    app.add_module('Function Probe() As Double\nOther.values(0)=-100\nOther.values(1)=60\n'
                   'Other.values(2)=60\nProbe=IRR(Other.values)\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == pytest.approx(0.1306623862918074)
    app.add_module('Public values(2) As Integer', name='Other')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    app.add_module('Public values(2) As Double', name='Other')
    assert app.run('Harness.Probe') == pytest.approx(0.1306623862918074)


STORE = '''Sub StoreProbe()
Dim values(2) As Double
values(0)=-100
values(1)=60
values(2)=60
Range("A1").Value2=Rate(12,-100,1000)
Range("B1").Value2=IRR(values)
Range("C1").Value2=MIRR(values,0.1,0.12)
Range("D1").Value2=NPV(0.1,values)
Range("E1").Value2=IPmt(0.01,1,12,1000)
Range("F1").Value2=PPmt(0.01,1,12,1000)
Range("G1").Formula="=CashRate(12,-100,1000)"
End Sub
Function CashRate(ByVal periods As Double, ByVal paid As Double, ByVal present As Double) As Double
CashRate=Rate(periods,paid,present)
End Function
Function ReadValue(ByVal Index As Long) As Double
ReadValue=Cells(1,Index).Value2
End Function
'''
EXPECTED = [0.029228540769133997, 0.1306623862918074, 0.12782977438973497,
            3.756574004507888, -10.0, -78.84878867834168, 0.029228540769133997]


def _save_financial_project(path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(STORE, name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)


def test_financial_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'financial.xlsm'
    _save_financial_project(path)
    reopened = ExcelApplication.open(path)
    assert [reopened.run('Harness.ReadValue', i) for i in range(1, 8)] == pytest.approx(EXPECTED, rel=2e-13)
    reopened.run('Harness.StoreProbe')
    assert [reopened.run('Harness.ReadValue', i) for i in range(1, 8)] == pytest.approx(EXPECTED, rel=2e-13)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_financial_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'financial.xlsm'
    _save_financial_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert [excel.run_raw('Harness.ReadValue', i) for i in range(1, 8)] == pytest.approx(EXPECTED, rel=2e-13)
        excel.run_raw('Harness.StoreProbe')
        assert [excel.run_raw('Harness.ReadValue', i) for i in range(1, 8)] == pytest.approx(EXPECTED, rel=2e-13)
        assert not any(event.get('dialogs') for event in excel.events)
