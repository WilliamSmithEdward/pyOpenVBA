"""A member name in brackets, typed and through an Object, against live Excel (scripts/measure_vba_brackets.py)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

ANSWERS: dict[str, str] = json.loads((Path(__file__).parent / 'fixtures/vba_brackets.json').read_text(encoding='utf-8'))
BODY = """Function Probe() As String
    Dim ws As @W@, r As @R@, ch As @C@, ss As @S@, v As Variant
    Set ws = Worksheets.Add
    Set r = ws.Range("B2"): Set ss = Worksheets
    r.Value = 5
    On Error Resume Next
    @L@
    Probe = Err.Number & "|" & TypeName(v)
End Function
"""


@pytest.mark.parametrize('key', sorted(key for key, answer in ANSWERS.items() if not answer.startswith('stopped')))
def test_a_bracketed_name_answers_as_excel_did(key: str) -> None:
    line, binding = key.rsplit(' ', 1)
    names = ('Worksheet', 'Range', 'Chart', 'Sheets') if binding == 'typed' else ('Object',) * 4
    source = BODY.replace('@W@', names[0]).replace('@R@', names[1]).replace('@C@', names[2]).replace('@S@', names[3]).replace('@L@', line.replace(chr(92) + 'n', chr(10)))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    if ANSWERS[key].startswith('compile: '):
        with pytest.raises(VBACompileError):
            app.add_module(source, name='Harness')
            app.run('Harness.Probe')
        return
    app.add_module(source, name='Harness')
    assert app.run('Harness.Probe') == ANSWERS[key]
