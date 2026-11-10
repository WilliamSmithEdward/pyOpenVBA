"""Custom XML parts and their nodes, replayed against live Excel (scripts/measure_custom_xml.py): what a macro reads,
selects and changes, the events the workbook's parts raise, and what a save writes under customXml/.

Two things are normalised on both sides: a part's ID, a new GUID each time, and the line breaks of an answer, which
the harness hands back as line feeds where VBA's own Len counted carriage returns too (see the 'returns' case, read
with the carriage returns marked). The workbook's built-in parts, whose XML is its document properties, report
themselves.
"""
from __future__ import annotations

import importlib.util
import json
import re
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('custom_xml', 'custom_xml_b', 'custom_xml_c', 'custom_xml1', 'custom_xml2',
                                  'custom_xml3', 'custom_xml4')
               if (FIXTURES / f'{stem}.json').exists())
GUID = re.compile(r'\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}')
REFUSED = ('built in',)


def _script() -> ModuleType:
    path = Path(__file__).parents[1] / 'scripts/measure_custom_xml.py'
    spec = importlib.util.spec_from_file_location('measure_custom_xml', path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()


def _cases() -> list[tuple[str, int, list[Any]]]:
    found: list[tuple[str, int, list[Any]]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, number, case) for number, case in enumerate(record['cases'])]
    return found


CASES = _cases()


def _answer(text: str) -> str:
    return GUID.sub('{ID}', text.replace('\r\n', '\n'))


def _files(found: dict[str, str]) -> dict[str, str]:
    return {key: GUID.sub('{ID}', value) for key, value in found.items()}


@pytest.mark.parametrize('stem, number, case', CASES, ids=[f'{one[0]}-{one[2][0]}' for one in CASES])
def test_each_case_as_excel_answered_and_saved(stem: str, number: int, case: list[Any], tmp_path: Path) -> None:
    record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
    calls = [tuple(one[:3]) for one in record['cases']]
    answer, saved = SCRIPT.model_answer(calls, number, tmp_path)
    if answer is None:
        assert any(word in case[0] for word in REFUSED), case[0]
        return
    assert _answer(answer) == _answer(case[3])
    assert _files(saved) == _files(case[4])


@pytest.mark.parametrize('stem', [stem for stem in ROUNDS if (FIXTURES / f'{stem}_predictions.json').exists()])
def test_each_held_out_round_was_predicted_before_excel_answered(stem: str) -> None:
    record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
    predicted = json.loads((FIXTURES / f'{stem}_predictions.json').read_text('utf-8'))
    assert [case['call'] for case in predicted['cases']] == [one[:3] for one in record['cases']]
