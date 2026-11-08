"""Which characters a defined name may hold, as live Excel answered Names.Add for every UTF-16 unit from U+0080 up, first and after
the first (scripts/measure_name_characters_bmp.py), for pairs beyond the Basic Multilingual Plane
(scripts/measure_name_surrogate_pairs.py), for 53 characters after a first letter (scripts/measure_name_later_characters.py), and
how Names.Count counts names in scripts the model cannot sort (scripts/measure_names_count_scripts.py)."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel._names import valid

ROOT = Path(__file__).parents[1]
FIXTURES = ROOT / 'tests' / 'fixtures'
SWEEP: dict[str, Any] = json.loads((FIXTURES / 'name_characters_bmp.json').read_text(encoding='utf-8'))
UNITS = [point for lo, hi in SWEEP['done'] for point in range(lo, hi + 1)]


def _unit(point: int) -> str:
    return point.to_bytes(2, 'little').decode('utf-16-le', 'surrogatepass')


@pytest.mark.parametrize('first', [True, False], ids=['first', 'after the first'])
def test_every_unit_as_excel_takes_it(first: bool) -> None:
    refused = set(SWEEP['first'] if first else SWEEP['later'])
    wrong = [f'U+{point:04X}' for point in UNITS if valid(_unit(point) + 'q' if first else 'q_' + _unit(point)) == (point in refused)]
    assert wrong == []


def test_the_sweep_covered_every_unit_from_0080() -> None:
    assert sorted(UNITS) == list(range(0x80, 0x10000))


def test_characters_beyond_the_plane_as_excel_takes_them() -> None:
    record = json.loads((FIXTURES / 'name_surrogate_pairs.json').read_text(encoding='utf-8'))
    for case in record['cases']:
        character = chr(int(case['point'][2:], 16))
        assert valid('q_' + character) == (case['later_excel'] == 0), case['point']
        assert valid(character + 'q') == (case['first_excel'] == 0), case['point']


def test_after_a_first_letter_as_excel_takes_it() -> None:
    """With "a" first a digit of another script makes a cell's address, as A1 is, so Excel refuses it there."""
    record = json.loads((FIXTURES / 'name_later_characters.json').read_text(encoding='utf-8'))
    for case in record['cases']:
        assert valid('a' + chr(int(case['point'][2:], 16))) == (case['excel'] == 0), case['point']


@pytest.mark.parametrize('fixture', ['name_address_digits.json', 'name_address_digit_forms.json'])
def test_digits_of_other_scripts_make_an_address_as_excel_reads_them(fixture: str) -> None:
    """Seventeen scripts' digits stand for ASCII digits when a name is read as a reference: A1, XFD1048576, R1C1, RC3."""
    cases: list[dict[str, Any]] = json.loads((FIXTURES / fixture).read_text(encoding='utf-8'))['cases']
    for case in cases:
        name = case['name'] if 'name' in case else case['prefix'] + chr(int(case['point'][2:], 16))
        assert valid(name) == (case['excel'] == 0), case


def _probe(name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, ROOT / 'scripts' / f'{name}.py')
    assert spec is not None and spec.loader is not None
    module: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_names_count_counts_names_in_every_script(tmp_path: Path) -> None:
    probe = _probe('measure_names_count_scripts')
    record = json.loads((FIXTURES / 'names_count_scripts.json').read_text(encoding='utf-8'))
    for case in record['cases']:
        assert probe.model(case['typed'], case['reopen'], str(tmp_path)) == case['excel']
