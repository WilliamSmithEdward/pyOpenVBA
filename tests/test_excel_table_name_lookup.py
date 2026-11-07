"""ListObjects(text) finds a table by its Name, in any case, and not by a DisplayName that differs from it, on a new book and after
Excel's own save and reopen (tests/fixtures/table_name_lookup.json, scripts/measure_table_name_lookup.py). Held-out round 51101 found
the model looking the other way."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures' / 'table_name_lookup.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['cases'], ids=lambda case: f'{case["name"]} reopened {case["reopen"]}')
def test_each_lookup_as_excel_answered(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Lookup')
    assert app.run('Lookup.Probe', str(tmp_path / 'n.xlsx'), case['name'], case['reopen']) == case['answer']


def test_the_names_that_differ_are_not_found_by_what_they_show() -> None:
    differing = [case for case in RECORD['cases'] if case['name'] != 'Plain']
    assert differing and all(case['answer'].endswith('|E9|E9') for case in differing)
