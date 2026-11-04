"""A fifteenth sweep of everyday macros, WorksheetFunction called by every kind of argument, replayed against what live Excel answered
(scripts/measure_excel_everyday_15.py). The answers are compared call by call."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/excel_everyday_15.json').read_text(encoding='utf-8'))['probes']

@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_each_call_as_excel_answered(probe: dict[str, Any]) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    got = str(app.run('Harness.Probe')).split(';')
    differing = [(call, wanted, answer) for call, wanted, answer in zip(probe['calls'], probe['result'].split(';'), got)
                 if wanted != answer]
    assert not differing


@pytest.mark.parametrize(('declared', 'expected'), [('Object', 'Double'), ('Variant', 'Double'), ('WorksheetFunction', 'String')])
def test_what_a_function_answers_follows_the_variable_it_is_called_through(declared: str, expected: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Function Probe() As String
    Dim f As {declared}
    Set f = Application.WorksheetFunction
    Probe = TypeName(f.Bin2Dec("101"))
End Function''', name='Harness')
    assert app.run('Harness.Probe') == expected
