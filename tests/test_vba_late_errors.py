"""A call bound late whose arguments do not fit the member, against live Excel (scripts/measure_vba_late_errors.py)."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.interpreter._type_library import find_class, late_error

ANSWERS: dict[str, str] = json.loads((Path(__file__).parent / 'fixtures/vba_late_errors.json').read_text(encoding='utf-8'))['answers']
CLASSES = {'r': 'Range', 'ws': 'Worksheet', 'wb': 'Workbook', 'f': 'Font', 'sh': 'Shape', 'lo': 'ListObject', 'ch': 'Chart', 'ss': 'Sheets', 'n': 'Name', 'se': 'Series', 'wi': 'Window', 'pt': 'PivotTable', 'cm': 'Comment', 'hl': 'Hyperlink', 'lc': 'ListColumn', 'co': 'ChartObject', 'it': 'Interior', 'ax': 'Axis', 'va': 'Validation', 'ps': 'PageSetup', 'wbs': 'Workbooks'}


def said(key: str) -> int | None:
    head, rest = key.split(': ', 1)
    line = rest.split(' ', 1)[1]
    found = find_class('excel', CLASSES[head])
    assert found is not None
    match = re.search(rf'{head}\.(\w+)(?:\((.*?)\))?', line)
    assert match is not None
    parts = [part.strip() for part in (match.group(2) or '').split(',') if part.strip()]
    named = [part.split(':=')[0] for part in parts if ':=' in part]
    return late_error(found, match.group(1), len(parts) - len(named), named)


#: Value properties that answer otherwise than others of their class. Nothing in the type library marks them, so they are known gaps.
UNEXPLAINED = {'pt: args_on_no_args v = pt.allocationweightexpression(1)', 'va: args_on_no_args v = va.errormessage(1)', 'ch: args_on_no_args v = ch.codename(1)', 'va: args_on_no_args v = va.ignoreblank(1)',
               'wb: args_on_function v = wb.exclusiveaccess(1)', 'wb: args_on_no_args v = wb.author(1)', 'ws: args_on_no_args v = ws.codename(1)', 'ws: args_on_no_args v = ws.creator(1)',
               'n: args_on_no_args v = n.category(1)', 'n: args_on_no_args v = n.categorylocal(1)', 'ws: args_on_no_args v = ws.name(1)'}


def test_the_known_gaps_are_still_gaps() -> None:
    assert {key for key in UNEXPLAINED if str(said(key)) == ANSWERS[key]} == set()


#: Where the rule leaves the member to answer, the object's own lookup decides: its answer is checked by replay elsewhere.
@pytest.mark.parametrize('key', sorted(key for key in ANSWERS if said(key) is not None and key not in UNEXPLAINED))
def test_a_late_call_that_does_not_fit_fails_as_excel_failed(key: str) -> None:
    assert str(said(key)) == ANSWERS[key]
