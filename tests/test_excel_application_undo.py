"""Members of Application a macro meets with no screen, replayed from live Excel (scripts/measure_application_undo.py,
tests/fixtures/application_undo.json): what each answers or raises."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/application_undo.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: what broke the run in Excel, and what the model has no object for or depends on this machine.
LEFT_OUT = ('app.Repeat', 'MultiThreadedCalculation', 'AlwaysUseClearType', 'WindowsForPens', 'app.AddIns',
            'RecentFiles', 'AutoCorrect', 'DisplayXMLSourcePane', 'Speech', 'SpellingOptions', 'DefaultWebOptions',
            'DefaultPivotTableLayoutOptions', 'LanguageSettings', 'AutoRecover', 'UsedObjects', 'ClipboardFormats',
            'FileConverters', 'RegisteredFunctions', 'COMAddIns', 'CommandBars', 'app.VBE', 'Hwnd', 'Hinstance',
            'Height', 'Width', 'IsEmpty(app.PreviousSelections)')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run() -> dict[int, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='UndoProbe')
    return {index: str(app.run('UndoProbe.Probe', index, index)) for index in KEPT}


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: dict[int, str], index: int) -> None:
    assert run[index] == RECORD['cases'][index][1]


def test_repeat_raises_as_excel_does() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Function Try() As Long\n    On Error Resume Next\n    Application.Repeat\n'
                   '    Try = Err.Number\nEnd Function\n', name='Trial')
    assert app.run('Trial.Try') == 1004
