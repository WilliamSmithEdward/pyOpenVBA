"""What Excel reads of a loaded chart's data labels, through a typed Chart and a late-bound Object (scripts/measure_loaded_chart_labels.py,
tests/fixtures/loaded_chart_labels.json). Each read gives Excel's answer or is refused. A series' dLbls and its points' dLbl are read
where they hold what the model writes; a turned, shadowed or right-to-left label, or anything else, leaves the series' labels and its
points' reads refused. The turned book is a held-out round: its reads were predicted refused before Excel was measured."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_labels.json').read_text(encoding='utf-8'))


def source(reads: list[str], kind: str) -> str:
    lines: list[str] = []
    for number, read in enumerate(reads):
        lines += [f'Public Function R{number}() As String', f'    Dim c As {kind}, x As String',
                  '    Set c = ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart', '    On Error Resume Next',
                  f'    x = TypeName(c.{read}) & " " & c.{read}', f'    R{number} = x & " " & Err.Number', 'End Function']
    return '\n'.join(lines)


@pytest.mark.parametrize('binding', ['typed', 'late'])
@pytest.mark.parametrize('name', RECORD['answers'])
def test_each_label_read_is_excels_or_refused(name: str, binding: str) -> None:
    reads: list[str] = RECORD['reads']
    wanted: list[str] = RECORD['answers'][name][binding]
    app = ExcelApplication.open(FIXTURES / name, with_vba=False)
    app.add_module(source(reads, 'Chart' if binding == 'typed' else 'Object'), name='P')
    wrong: list[str] = []
    for number, (read, want) in enumerate(zip(reads, wanted)):
        try:
            got = str(app.run(f'P.R{number}'))
        except VBAUnsupportedError:
            continue
        if got != want:
            wrong.append(f'{read}: Excel {want!r}, model {got!r}')
    assert wrong == []
