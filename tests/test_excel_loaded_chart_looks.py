"""What Excel reads of a loaded chart's looks, through a typed Chart and a late-bound Object (scripts/measure_loaded_chart_looks.py,
tests/fixtures/loaded_chart_looks.json). Until the model reads a loaded chart's spPr and txPr, each read gives Excel's answer or is
refused: a wrong read of a loaded chart is worse than a refusal."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_looks.json').read_text(encoding='utf-8'))


def source(reads: list[str], kind: str) -> str:
    lines: list[str] = []
    for number, read in enumerate(reads):
        lines += [f'Public Function R{number}() As String', f'    Dim c As {kind}, x As String',
                  '    Set c = ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart', '    On Error Resume Next',
                  f'    x = TypeName(c.{read}) & " " & c.{read}', f'    R{number} = x & " " & Err.Number', 'End Function']
    return '\n'.join(lines)


@pytest.mark.parametrize('binding', ['typed', 'late'])
@pytest.mark.parametrize('name', RECORD['answers'])
def test_each_look_read_is_excels_or_refused(name: str, binding: str) -> None:
    reads: list[str] = RECORD['reads']
    wanted: list[str] = RECORD['answers'][name][binding]
    app = ExcelApplication.open(FIXTURES / name, with_vba=False)
    asked = [(read, want) for read, want in zip(reads, wanted) if want != 'does not compile']
    app.add_module(source([read for read, _ in asked], 'Chart' if binding == 'typed' else 'Object'), name='P')
    wrong: list[str] = []
    for number, (read, want) in enumerate(asked):
        try:
            got = str(app.run(f'P.R{number}'))
        except VBAUnsupportedError:
            continue
        if got != want:
            wrong.append(f'{read}: Excel {want!r}, model {got!r}')
    assert wrong == []
