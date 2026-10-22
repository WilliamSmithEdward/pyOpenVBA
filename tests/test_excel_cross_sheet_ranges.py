"""Ranges asked for across two sheets, measured by scripts/measure_cross_sheet_ranges.py.

Worksheet.Range takes only its own sheet's cells: a Range, a qualified
reference or a defined name on another sheet is error 1004. Application.Range
takes cells on any one sheet, but not corners on two. Union and Intersect of
ranges on two sheets are error 1004.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/cross_sheet_ranges.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_a_range_across_two_sheets(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']
