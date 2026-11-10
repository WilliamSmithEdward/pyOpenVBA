"""A name renamed to its own spelling in another case or to text holding !, by Name and NameLocal, and ValidWorkbookParameter
on books made, saved and reopened, as live Excel answered (scripts/measure_name_rename.py, tests/fixtures/name_rename.json).
The one rename Excel takes with a ! keeps the whole text as the bare name, which the model reports rather than holds."""
from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_name_rename', ROOT / 'scripts/measure_name_rename.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/name_rename.json').read_text(encoding='utf-8'))['cases']
UNSUPPORTED = 'model VBAUnsupportedError: a name renamed to a text holding ! is not implemented'


def _taken_whole(case: dict[str, Any]) -> bool:
    """Whether Excel took a text holding ! as the bare name: the name it reads holds two !, or a workbook's holds one."""
    if case['group'] != 'bang':
        return False
    read = case['late'].split(' || ')[1].split(' | ')
    return read[1].count('!') == 2 or (read[3] == 'Workbook' and '!' in read[1]) or _own_text(case)


def _own_text(case: dict[str, Any]) -> bool:
    """A sheet's own name given its own text, which changed nothing here and took the whole text in name round 55101."""
    return case['make'].startswith('Set n = ws2.Names.Add("Loc"') and case['call'].endswith('= "Data!Loc"')


@pytest.fixture(scope='module')
def folder() -> Any:
    with tempfile.TemporaryDirectory() as held:
        PROBE.rounds.books(Path(held))
        yield held


@pytest.mark.parametrize('case', CASES, ids=[f'{at}-{case["group"]}' for at, case in enumerate(CASES)])
def test_each_case_as_excel_answered(case: dict[str, Any], folder: str) -> None:
    for tag, typed in (('late', False), ('typed', True)):
        answer = PROBE.model(case, typed, folder)
        if _taken_whole(case):
            assert answer == UNSUPPORTED
        else:
            assert answer == case[tag]


def test_the_probe_reached_every_rule() -> None:
    assert sum(_taken_whole(case) for case in CASES) >= 10
    assert {case['group'] for case in CASES} == {'case', 'bang', 'valid', 'round'}
