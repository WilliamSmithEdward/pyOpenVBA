"""Data validation through inserts and deletes, replayed against Excel's answers (scripts/measure_x14_validation_shifts.py):
rules kept in the x14 extension and plain ones in an Excel-saved book, each saved and its validation elements compared as the file
spells them; then three narrower looks: a range's pieces through a shift, and where a rule an edit changes, or whose formula it
changes, goes among the sheet's rules."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

_spec = importlib.util.spec_from_file_location("measure_x14_validation_shifts",
                                              Path(__file__).resolve().parent.parent / "scripts/measure_x14_validation_shifts.py")
assert _spec is not None and _spec.loader is not None
probe: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(probe)

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('x14_validation_shifts', 'x14_validation_shifts2', 'x14_validation_shifts3', 'x14_validation_shifts4',
                                  'x14_validation_shifts5', 'x14_validation_shifts6', 'x14_validation_shifts7',
                                  'x14_validation_shifts8', 'x14_validation_shifts9', 'x14_validation_shifts10') if (FIXTURES / f'{stem}.json').exists())


def _load(stem: str) -> dict[str, Any]:
    return json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))


BOOK_CASES = [(stem, call, answer) for stem in ROUNDS for call, answer in zip(_load(stem)['cases'], _load(stem)['answers'], strict=True)]
NARROW_CASES = [(stem, record['source'], call, answer) for stem in ('validation_pieces', 'validation_list_order', 'validation_formula_order')
                for record in (_load(stem),) for call, answer in zip(record['cases'], record['answers'], strict=True)]


@pytest.mark.parametrize('stem, call, answer', BOOK_CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(BOOK_CASES)])
def test_each_edit_leaves_the_book_s_validation_as_excel_saved_it(stem: str, call: list[Any], answer: str, tmp_path: Path) -> None:
    del stem
    assert probe.model_answer(call, tmp_path) == answer


@pytest.mark.parametrize('stem, source, call, answer', NARROW_CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(NARROW_CASES)])
def test_each_edit_leaves_the_rules_as_excel_saved_them(stem: str, source: str, call: list[Any], answer: str, tmp_path: Path) -> None:
    del stem
    assert probe.model_narrow(source, call, tmp_path) == answer


def test_every_case_was_answered() -> None:
    for stem in (*ROUNDS, 'validation_pieces', 'validation_list_order', 'validation_formula_order'):
        assert not any(str(answer).startswith('ERROR') for answer in _load(stem)['answers']), stem


def test_the_base_book_keeps_its_cross_sheet_lists_in_the_extension() -> None:
    import zipfile

    with zipfile.ZipFile(probe.BASE) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    assert xml.count('<x14:dataValidation ') == 2 and xml.count('<dataValidation ') == 2
