"""GradientStops.Insert2, For Each and Application, replayed from live Excel (scripts/measure_shape_gradient_stops.py,
tests/fixtures/shape_gradient_stops.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_gradient_stops.json').read_text(encoding='utf-8'))
#: A stop inserted with both a transparency and a brightness reports itself.
BOTH = 'st.Insert2 16711680, 0.5, 0.5, 2, 0.4'
CASES = [case for case in RECORD['cases'] if case[0] != BOTH]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


def _source(statements: str) -> str:
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, st As Object, o As Object, v2 As Variant, s As String, v As String, s2 As Variant', '    Set ws = ActiveSheet',
             f'    {RECORD["make"].rstrip(": ")}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in RECORD['reads']:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    return '\n'.join(lines) + '\n'


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(case[0]), name='Stops')
    names = ['status', *RECORD['reads']]
    assert dict(zip(names, str(app.run('Stops.Go')).split('~'), strict=True)) == dict(zip(names, case[1].split('~'), strict=True))
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<a:gradFill\b.*?</a:gradFill>', xml, re.DOTALL)[0] == case[2]


def test_a_stop_with_both_reports_itself() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(BOTH).replace('On Error GoTo Failed', "' no handler"), name='Stops')
    with pytest.raises(VBAUnsupportedError):
        app.run('Stops.Go')
