"""Held-out setter rounds for twelve Workbook settings (scripts/measure_workbook_setters_heldout.py): every kind of value, typed and late,
each set on a new workbook or one already set, then read, saved, opened again and read again, with the model's answers written before live
Excel's. Each round's predictions stay as they were written, with their digest and miss count; the model then replays every case."""
from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
FOLDER = ROOT / 'tests/fixtures/workbook_setters_heldout'
_SPEC = importlib.util.spec_from_file_location('measure_workbook_setters_heldout', ROOT / 'scripts/measure_workbook_setters_heldout.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
#: Each round's predictions, as their digest was recorded before Excel answered (.git/excel-coordination/BOARD.md).
DIGESTS = {'round1': 'd2992aded0f70a1eda855da749bc67be4003dea7241c17ed912e76c31a5b8fe4',
           'round2': 'f47b8541af04de39f3dd25b45b0fdd47290d454146a26471a7ef19c81ba6e79f'}
#: Round1 missed 38 of 480: DisplayDrawingObjects past the three values a file names (refused on save), bound late (true, fractions,
#: past a Long, text and ranges), Null on three flags (refused openly, where Excel says 94 typed and 13 late), and text, a Currency, a
#: Date and a range bound late on Date1904, PrecisionAsDisplayed and SaveLinkValues, which Excel refuses with 1004.
MISSES = {'round1': 38, 'round2': 7}
#: A Range given late to these, which Excel refuses with 1004: the model hands the setter the range's value, so it takes it. Fixing that is
#: a change to how every late set receives an object, not to these settings.
GAPS = frozenset({'Date1904', 'DisplayDrawingObjects', 'PrecisionAsDisplayed', 'SaveLinkValues'})
RANGE = 'ws.Range("A1,C1")'
ROUNDS = sorted(path.name.removesuffix('_excel.json') for path in FOLDER.glob('*_excel.json'))


def _load(round_name: str, kind: str) -> dict[str, Any]:
    return json.loads((FOLDER / f'{round_name}_{kind}.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('round_name', ROUNDS)
def test_the_predictions_are_as_they_were_written(round_name: str) -> None:
    text = (FOLDER / f'{round_name}_predictions.json').read_bytes().replace(b'\r\n', b'\n')
    assert hashlib.sha256(text).hexdigest() == DIGESTS[round_name]


@pytest.mark.parametrize('round_name', ROUNDS)
def test_the_miss_count_is_the_one_recorded(round_name: str) -> None:
    predicted, excel = _load(round_name, 'predictions'), _load(round_name, 'excel')
    assert sum(a != b for declared in ('Workbook', 'Object')
               for a, b in zip(predicted['answers'][declared], excel['answers'][declared])) == MISSES[round_name]


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
@pytest.mark.parametrize('round_name', ROUNDS)
def test_every_case_answers_as_excel_did(round_name: str, declared: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    excel = _load(round_name, 'excel')
    every = excel['cases']
    found: list[str] = []
    for at in range(0, len(every), MEASURE.CHUNK):
        folder = tmp_path / f'chunk{at}'
        folder.mkdir()
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(MEASURE.macro(declared, every[at:at + MEASURE.CHUNK]), name='H')
        found += str(app.run('H.Probe', str(folder))).split(MEASURE.SEPARATOR)
    differ = [(case['setting'], case['value'], case['first'], mine, theirs)
              for case, mine, theirs in zip(every, found, excel['answers'][declared]) if mine != theirs]
    known = [one for one in differ if declared == 'Object' and one[1] == RANGE and one[0] in GAPS]
    assert [one for one in differ if one not in known] == []
    # Each named gap still differs, so one that closes leaves GAPS.
    if declared == 'Object':
        assert {one[0] for one in known} == {case['setting'] for case in every if case['value'] == RANGE and case['setting'] in GAPS}
