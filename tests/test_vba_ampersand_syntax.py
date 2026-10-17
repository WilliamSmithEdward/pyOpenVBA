"""Native ampersand syntax, Long calls and stored-variable type characters."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/ampersand_syntax.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_ampersand_syntax(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    if case.get('compile_error'):
        with pytest.raises(VBACompileError):
            app.add_module(case['source'], name='Harness')
            app.run('Harness.Probe')
    else:
        app.add_module(case['source'], name='Harness')
        assert app.run('Harness.Probe') == case['result']
