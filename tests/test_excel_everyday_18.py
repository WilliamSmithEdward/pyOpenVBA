"""An eighteenth sweep: Range as everyday macros use it, each group bound late and typed, replayed against live Excel
(scripts/measure_excel_everyday_18.py). A line records its error number, the type of its answer and the answer."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

PROBES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/excel_everyday_18.json').read_text(encoding='utf-8'))['probes']
SEPARATOR = '¦'


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: f"{probe['name']}-{probe['binding']}")
def test_a_group_of_lines_answers_as_excel_did(probe: dict[str, Any]) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(probe['macro'], name='H')
    found = str(app.run('H.Probe')).split(SEPARATOR)
    wanted = str(probe['answer']).split(SEPARATOR)
    assert len(found) == len(wanted)
    assert not [(asked, excel, model) for asked, excel, model in zip(probe['asked'], wanted, found) if excel != model]


REFUSED = sorted({probe['asked'][at].strip(' /') for probe in PROBES for at in probe['not_compiling']})


@pytest.mark.parametrize('asked', REFUSED)
def test_what_did_not_compile_typed_does_not_compile_here(asked: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Function Probe()\n    Dim ws As Worksheet, v As Variant\n    Set ws = ActiveSheet\n    v = {asked}\nEnd Function\n', name='H')
    with pytest.raises(VBACompileError):
        app.run('H.Probe')


@pytest.mark.parametrize('declared', ['Object', 'Range'])
@pytest.mark.parametrize('kind', [11, 12])
def test_the_xml_a_range_reads_as_is_a_stated_gap(declared: str, kind: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Function Probe()\n    Dim r As {declared}\n    Set r = Range("A1")\n    Probe = r.Value({kind})\nEnd Function\n', name='H')
    with pytest.raises(VBAUnsupportedError):
        app.run('H.Probe')


def test_a_whole_sheet_of_blanks_shows_one_text_without_reading_each_cell() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function Probe()\n    Probe = TypeName(Cells.Text) & "|" & TypeName(Rows(30).Text)\n    Range("A1").Value = "x"\n'
                   '    Probe = Probe & "|" & TypeName(Cells.Text) & "|" & Columns(2).Text\nEnd Function\n', name='H')
    assert app.run('H.Probe') == 'String|String|Null|'


@pytest.mark.parametrize('sized', ['Resize()', 'Resize(1)', 'Resize(1, 1)', 'Resize(, 1)'])
def test_resize_takes_one_area(sized: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Function Probe()\n    On Error Resume Next\n    Dim a As String\n    a = Range("A1,C3").{sized}.Address\n    Probe = Err.Number\nEnd Function\n', name='H')
    assert app.run('H.Probe') == 1004
