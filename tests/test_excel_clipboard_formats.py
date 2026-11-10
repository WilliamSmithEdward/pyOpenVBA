"""Application.ClipboardFormats, replayed from live Excel (scripts/measure_clipboard_formats.py,
tests/fixtures/clipboard_formats.json): the formats after each kind of copy, typed, late bound and by Index."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/clipboard_formats.json').read_text(encoding='utf-8'))
CASES: list[list[str]] = RECORD['cases']
#: What the model reports rather than does: copying a shape that is not a chart.
REFUSED = {index for index, (statement, _) in enumerate(CASES) if 'AddShape(1, 10, 10, 50, 50).Copy' in statement and 'CopyPicture' not in statement}
#: The Windows clipboard calls the measurement empties it with, which the model's clipboard does not need.
_EMPTIED = ('Private Sub Emptied()\n    If OpenClipboard(0) <> 0 Then\n        EmptyClipboard\n        CloseClipboard\n    End If\nEnd Sub\n')


def _source() -> str:
    source = '\n'.join(line for line in RECORD['source'].splitlines() if not line.startswith('Private Declare PtrSafe'))
    return source.replace(_EMPTIED, 'Private Sub Emptied()\nEnd Sub\n')


def _id(index: int) -> str:
    return CASES[index][0]


def _drawn(*_: object) -> None:
    """The picture CopyPicture would draw, which these cases never look at."""


def _app() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook('CopyPicture', _drawn)
    app.add_module(_source(), name='P')
    return app


#: Cells copied straight after a bitmap picture or a chart area's Copy: Excel answered the cells' list in some runs and the earlier
#: copy's in others, and the model says so.
RACING = {index for index, (statement, _) in enumerate(CASES)
          if re.search(r'(CopyPicture xlScreen, xlBitmap|ChartArea\.Copy): ws\.Range\("[A-Z]\d"\)\.(Copy|Cut)(: ws\.Paste .*)?$', statement)}
_CELLS = 'Variant() 1 23:'
_EARLIER = ('Variant() 1 3: Double 9 Double 11 Double 45', 'Variant() 1 18:')


@pytest.mark.parametrize('index', sorted(RACING), ids=_id)
def test_excel_answers_either_list_after_a_racing_copy(index: int) -> None:
    typed = CASES[index][1].split('|')[2]
    assert typed.startswith(_CELLS) or typed.startswith(_EARLIER)
    with pytest.raises(VBAUnsupportedError, match='answers either'):
        _app().run('P.Probe', index)


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if index not in REFUSED | RACING], ids=_id)
def test_the_clipboard_holds_what_excel_said(index: int) -> None:
    assert str(_app().run('P.Probe', index)) == CASES[index][1]


@pytest.mark.parametrize('index', sorted(REFUSED), ids=_id)
def test_what_the_model_does_not_copy_reports_itself(index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _app().run('P.Probe', index)
