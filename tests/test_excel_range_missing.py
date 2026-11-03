"""Range and WorksheetFunction members a macro can reach with no screen, replayed against what live Excel answered
(scripts/measure_range_missing.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.xpath import XPathError, select

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_missing.json').read_text(encoding='utf-8'))
#: What the model does not have yet, by what each case asks for.
GAPS = {'FlashFill': 'Range.FlashFill', 'ADODB.Recordset': 'a disconnected ADODB.Recordset'}
KEPT = [index for index, (case, _) in enumerate(RECORD['cases']) if not any(text in case for text in GAPS)]
LEFT = [index for index in range(len(RECORD['cases'])) if index not in KEPT]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='MissingProbe')
    return made


@pytest.mark.parametrize('index', KEPT)
def test_as_excel_answered(app: ExcelApplication, index: int) -> None:
    assert str(app.run('MissingProbe.Missing', index)) == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', LEFT)
def test_what_is_left_says_so(index: int) -> None:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='MissingProbe')
    with pytest.raises(VBAUnsupportedError):
        made.run('MissingProbe.Missing', index)


@pytest.mark.parametrize('document', [
    '<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>',
    '<!DOCTYPE a SYSTEM "file:///etc/passwd"><a/>',
    '<!DOCTYPE lolz [<!ENTITY a "a"><!ENTITY b "&a;&a;">]><l>&b;</l>',
])
def test_a_document_with_a_doctype_is_refused_before_any_entity_is_read(document: str) -> None:
    with pytest.raises(XPathError):
        select(document, '//a')
