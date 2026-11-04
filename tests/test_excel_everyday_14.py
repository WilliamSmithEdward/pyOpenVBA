"""A fourteenth sweep of everyday macros, replayed against what live Excel answered (scripts/measure_excel_everyday_14.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/excel_everyday_14.json').read_text(encoding='utf-8'))['probes']
#: What each probe the model cannot answer yet needs.
GAPS: dict[str, str] = {
    'chart_sheet': 'moving a chart to a chart sheet',
}


def _answer(macro: str) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    try:
        app.add_module(macro, name='Harness')
        return app.run('Harness.Probe')
    except VBACompileError:
        return 'does not compile'
    except VBARuntimeError as failure:
        return {'error': failure.number}


@pytest.mark.parametrize('probe', [
    pytest.param(probe, marks=pytest.mark.xfail(raises=VBAUnsupportedError, strict=True, reason=GAPS[probe['name']]))
    if probe['name'] in GAPS else probe for probe in PROBES], ids=lambda probe: probe['name'])
def test_an_everyday_macro(probe: dict[str, Any]) -> None:
    expected = 'does not compile' if 'compile' in probe else {'error': probe['error']} if 'error' in probe else probe['result']
    assert _answer(probe['macro']) == expected


def test_each_gap_is_a_probe() -> None:
    assert set(GAPS) <= {probe['name'] for probe in PROBES}


def test_a_saved_workbooks_name_takes_workbook_parameter_up_to_what_the_model_keeps(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.save(tmp_path / 'saved.xlsx')
    app.add_module('''Function Probe(ByVal value As Variant) As Long
    Dim n As Object
    Set n = ActiveWorkbook.Names.Add("Mac", "=Sheet1!$A$1")
    On Error Resume Next
    n.WorkbookParameter = value
    Probe = Err.Number
End Function''', name='Harness')
    assert app.run('Harness.Probe', False) == 0
    assert app.run('Harness.Probe', 'x') == 13
    with pytest.raises(VBAUnsupportedError):
        app.run('Harness.Probe', True)
