"""A property set on a whole old collection of one kind, Worksheet.Buttons.Caption = "x" and the like, typed and late, replayed from live
Excel (scripts/measure_old_collection_lets.py, tests/fixtures/old_collection_lets.json): what the set answers, what each item reads after,
and what each reads once the workbook is saved and opened again. DrawingObjects, which mixes kinds, is replayed elsewhere."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._old_controls import OldControls  # pyright: ignore[reportPrivateUsage]
from pyopenvba.apps.excel._old_drawings import Lines  # pyright: ignore[reportPrivateUsage]

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_lets.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_old_collection_lets', Path(__file__).parents[1] / 'scripts/measure_old_collection_lets.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
row_source: Any = getattr(_script, 'row_source')
#: Sets the model does not yet answer as Excel did, each with why.
NOT_YET: dict[str, str] = {
    'Buttons.AutoSize': "the item setter needs the caption's text width",
    'Buttons.Formula': 'Excel answers 1004 and changes nothing; the item setter is not measured',
    **{f'{kind}.Formula': 'Excel links no item, answering 1004 on Arcs, Ovals and TextBoxes and nothing on Rectangles, by no rule found yet'
       for kind in ('Arcs', 'Ovals', 'Rectangles', 'TextBoxes')},
    'Ovals.AutoSize': "the drawing's item setter reports itself",
    'Pictures.Formula': "the picture's item setter reports itself",
    'Pictures.Shadow': "the picture's item setter reports itself",
    'Rectangles.AutoSize': "the drawing's item setter reports itself",
    'TextBoxes.AutoSize': "the drawing's item setter reports itself",
}
CASES = [case for case in RECORD['cases'] if case['collection'] != 'DrawingObjects']
REPLAYED = [case for case in CASES if f"{case['collection']}.{case['member']}" not in NOT_YET]


def _collections() -> dict[str, type[OldControls]]:
    found: dict[str, type[OldControls]] = {}
    waiting: list[type[OldControls]] = [OldControls]
    while waiting:
        for one in waiting.pop().__subclasses__():
            waiting.append(one)
            if one.vba_type_name:
                found[one.vba_type_name] = one
    return found


@pytest.mark.parametrize('case', REPLAYED, ids=lambda case: f"{case['collection']}.{case['member']} {case['binding']}")
def test_each_set_as_excel_answered(case: dict[str, str], tmp_path: Path) -> None:
    book = tmp_path / 'book.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(row_source(case), name='Probe')
    assert str(app.run('Probe.Row', str(book), str(tmp_path / 'saved.xlsx'))) == case['excel']


ROLES: list[dict[str, str]] = json.loads((FIXTURES / 'button_roles.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', ROLES, ids=lambda case: f"{case['binding']} {case['role'] or 'across'} {case['case']}")
def test_a_role_is_held_by_one_button_at_a_time(case: dict[str, str], tmp_path: Path) -> None:
    """scripts/measure_button_roles.py: DefaultButton, CancelButton and HelpButton each clear the same role on the sheet's other buttons."""
    book = tmp_path / 'book.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case['source'], name='Probe')
    assert str(app.run('Probe.Row', str(book))) == case['excel']


def test_not_yet_names_measured_sets_only() -> None:
    measured = {f"{case['collection']}.{case['member']}" for case in CASES}
    assert set(NOT_YET) <= measured, set(NOT_YET) - measured


def test_a_registered_setter_is_one_that_replays() -> None:
    """A collection lists a setter (spread_sets) only for a member measured on it whose every case replays."""
    measured = {(case['collection'], case['member']) for case in CASES}
    found = _collections()
    assert found['Lines'] is Lines and len(found) == 12, sorted(found)
    for name, kind in found.items():
        for member in kind.spread_sets:
            assert (name, member) in measured, (name, member)
            assert f'{name}.{member}' not in NOT_YET, (name, member)
