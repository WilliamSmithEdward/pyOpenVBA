"""A rule's Font set through each range that reaches it, its Size and Name, and its NumberFormat after codes Excel refuses or
takes empty, as live Excel answered (scripts/measure_format_condition_font.py, tests/fixtures/format_condition_font.json).
Bold, Italic, Strikethrough and Underline set on a rule reached through other cells than its own report themselves."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_format_condition_font', ROOT / 'scripts/measure_format_condition_font.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/format_condition_font.json').read_text(encoding='utf-8'))['cases']
STYLES = ('Font.Bold', 'Font.Italic', 'Font.Strikethrough', 'Font.Underline')
OWN = ('ws.Range("D1:D5").FormatConditions(1)', 'ws.Range("D1:D5").FormatConditions(2)')


def _reported(case: dict[str, Any]) -> bool:
    return case['member'] in STYLES and case['reach'] not in OWN


@pytest.mark.parametrize('case', CASES, ids=[f'{at}-{case["member"]}' for at, case in enumerate(CASES)])
def test_each_case_as_excel_answered(case: dict[str, Any]) -> None:
    for tag, typed in (('late', False), ('typed', True)):
        answer = PROBE.model(case, typed)
        if _reported(case):
            assert answer.startswith('model VBAUnsupportedError: ' + case['member'] + ' set on a rule reached through other cells')
        else:
            assert answer == case[tag]


def test_excel_reads_back_another_answer_where_the_model_reports() -> None:
    """Each reported case is one where the object read what it read before, or the rule split: none the model could copy."""
    assert all(case['late'].split(' | ')[1] != case['late'].split(' | ')[2] or case['late'].split(' | ')[3] != '1'
               for case in CASES if _reported(case))
