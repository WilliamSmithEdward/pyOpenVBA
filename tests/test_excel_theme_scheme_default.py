"""ThemeColorScheme's default member, typed and late, as live Excel answered (scripts/measure_theme_scheme_default.py): an index straight on
the scheme is Colors(index), and one out of range is -2147024809, where Colors itself says 5."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_theme_scheme_default', ROOT / 'scripts/measure_theme_scheme_default.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/theme_scheme_default.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('declared', ['ThemeColorScheme', 'Object'])
def test_the_default_member_answers_as_excel_did(declared: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURE.macro(declared), name='H')
    found = str(app.run('H.Probe')).split(MEASURE.SEPARATOR)
    assert dict(zip(RECORD['cases'], found)) == dict(zip(RECORD['cases'], RECORD['answers'][declared]))
