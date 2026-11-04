"""TextFrame2.TextRange.Font, replayed from live Excel (scripts/measure_shape_font2.py, tests/fixtures/shape_font2.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_font2.json').read_text(encoding='utf-8'))
F = 'sh.TextFrame2.TextRange.Font.'
UNBUILT = ('Fill.', 'Line.', 'Shadow.', 'Glow.', 'Reflection.', 'NameFarEast', 'NameComplexScript', 'NameOther', 'Embossed', 'Engraved', 'WordArtformat',
           'Highlight.')
READS = [read for read in RECORD['reads'] if not read.replace(F, '').startswith(UNBUILT)]
PLACES = [RECORD['reads'].index(read) + 1 for read in READS]
CASES = [case for case in RECORD['cases'] if not any(F + one in case[0] for one in UNBUILT)]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    shape, statements = case[0].split(': ', 1)
    lines = ['Public Function Go() As String', '    Dim ws As Object, sh As Object, s As String, v As String', '    Set ws = ActiveSheet',
             f'    Set sh = {shape}', '    On Error GoTo Failed', f'    {statements}', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Fonts')
    parts = case[1].split('|')
    assert app.run('Fonts.Go') == '|'.join([parts[0], *(parts[place] for place in PLACES)])
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    assert re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0] == case[2]


def test_a_font_fill_reports_itself() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 40)\n'
                   f'    {F}Fill.Patterned 5\nEnd Sub\n', name='Unbuilt')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unbuilt.Go')
