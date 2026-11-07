"""What a too-narrow column shows does not hang on the widths it had before (scripts/measure_narrow_history.py): each sequence of widths, replayed."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / 'fixtures/narrow_history.json').read_text('utf-8'))


@pytest.mark.parametrize('case', RECORD['cases'], ids=[f"{one['cell']}-{one['font']}{one['size']:g}-{one['pass']}-{one['widths']}" for one in RECORD['cases']])
def test_each_sequence_shows_what_excel_showed(case: dict[str, Any]) -> None:
    assign, shape = RECORD['cells'][case['cell']]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'].replace('%SET%', assign.replace('A1.', 'A1").', 1)), name='Probe')
    if case['widths'].startswith('0,'):
        # A column of no width is hidden, whose Range.Text the model reports rather than answers.
        with pytest.raises(VBAUnsupportedError, match='hidden column'):
            app.run('Probe.Go', case['widths'], case['font'], case['size'], shape)
        return
    assert str(app.run('Probe.Go', case['widths'], case['font'], case['size'], shape)).split('|')[1:] == case['shown']


def test_the_record_was_measured_at_96_dpi() -> None:
    assert RECORD['standard_height'] == '15'
