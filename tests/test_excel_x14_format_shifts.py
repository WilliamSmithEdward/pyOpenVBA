"""Conditional formats the Excel 2010 extension holds alone, cell-value and formula rules reading another sheet, through inserts and
deletes, replayed against Excel's answers (scripts/measure_x14_format_shifts.py): each book saved and its conditional formats
compared as the file spells them. Where the model refuses, it is for a reason the module names."""
from __future__ import annotations

import importlib.util
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

_spec = importlib.util.spec_from_file_location("measure_x14_format_shifts",
                                              Path(__file__).resolve().parent.parent / "scripts/measure_x14_format_shifts.py")
assert _spec is not None and _spec.loader is not None
probe: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(probe)

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('x14_format_shifts', 'x14_format_shifts_bar', 'x14_format_shifts2') if (FIXTURES / f'{stem}.json').exists())
#: What the model may refuse here: a data bar reshaped beside the extension's rules, and remaking rules alongside others.
REFUSALS = ('beside conditional formats the Excel 2010 extension holds alone', 'where one is made anew',
            'several conditional formatting ranges would be made anew')


def _load(stem: str) -> dict[str, Any]:
    return json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))


CASES = [(stem, call, answer, record.get('base', 'x14_format_base.xlsx')) for stem in ROUNDS for record in (_load(stem),)
         for call, answer in zip(record['cases'], record['answers'], strict=True)]


@pytest.mark.parametrize('stem, call, answer, base', CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_edit_leaves_the_formats_as_excel_saved_them(stem: str, call: list[Any], answer: str, base: str, tmp_path: Path) -> None:
    del stem
    try:
        found = probe.model_answer(call, tmp_path, FIXTURES / base)
    except VBAUnsupportedError as refusal:
        assert any(reason in str(refusal) for reason in REFUSALS), str(refusal)
        return
    assert found == answer


@pytest.mark.parametrize('base', ['x14_format_base.xlsx', 'x14_format_base_bar.xlsx'])
def test_an_untouched_save_keeps_the_extension_as_the_file_had_it(base: str, tmp_path: Path) -> None:
    saved = tmp_path / 'out.xlsx'
    app = ExcelApplication(sandbox=False)
    book: Any = app.open_workbook(str(FIXTURES / base))
    book.SaveCopyAs(str(saved))

    def extension(path: Path) -> str:
        with zipfile.ZipFile(path) as package:
            xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
        found = re.search(r'<extLst>.*</extLst>', xml, re.DOTALL)
        assert found is not None
        return found.group(0)

    assert extension(saved) == extension(FIXTURES / base)


def test_the_rules_read_as_rules() -> None:
    app = ExcelApplication(sandbox=False)
    book: Any = app.open_workbook(str(FIXTURES / 'x14_format_base.xlsx'))
    sheet = book.sheets_[0]
    assert not sheet.conditional_foreign
    kinds = sorted((rule.root.get('type'), rule.extension_only) for rule in sheet.conditional_formats)
    assert kinds == [('cellIs', False), ('cellIs', True), ('expression', True)]


def test_every_case_was_answered() -> None:
    for stem in ROUNDS:
        assert not any(answer.startswith('ERROR') for answer in _load(stem)['answers']), stem
