"""The XML map events, BeforeXmlImport, AfterXmlImport, BeforeXmlExport and AfterXmlExport on a workbook and on Application,
replayed from live Excel (scripts/measure_xml_map_events.py, tests/fixtures/xml_map_events.json): their order, what they are
given, and a Cancel, which fails the call."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/xml_map_events.json').read_text(encoding='utf-8'))
CASES: list[list[str]] = RECORD['cases']


def _id(index: int) -> str:
    return CASES[index][0]


@pytest.mark.parametrize('index', range(len(CASES)), ids=_id)
def test_the_events_come_as_excel_raised_them(index: int, tmp_path: Path) -> None:
    (tmp_path / 's.xsd').write_text(RECORD['schema'], encoding='utf-8')
    (tmp_path / 'd.xml').write_text(RECORD['data'], encoding='utf-8')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['log'], name='EventLog')
    app.add_module(RECORD['sink'], name='ExcelSink', kind='class')
    app.add_module(RECORD['source'], name='ProbeModule')
    answer = str(app.run('ProbeModule.Probe', str(tmp_path), index)).replace(str(tmp_path), '<folder>')
    assert answer.replace('\r\n', '\n') == CASES[index][1]
