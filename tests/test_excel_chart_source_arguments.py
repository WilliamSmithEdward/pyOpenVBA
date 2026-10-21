"""Exact independent native SourceData argument traces; deferred cases stay explicit."""
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/chart_source_arguments'
SCALARS: dict[str, Any] = json.loads((ROOT / 'chart-source-arguments-native.json').read_text(encoding='utf8'))
REMAINING: dict[str, Any] = json.loads((ROOT / 'chart-source-arguments-remaining-native.json').read_text(encoding='utf8'))
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf8'))
CONTROLS: list[dict[str, Any]] = [probe for probe in REMAINING['probes'] if probe['name'] in MANIFEST['admitted_remaining_controls']]
SCALAR_EXPRESSIONS = {'source-42': '42', 'source-0': '0', 'source-text': '"A1:D4"',
                      'source-empty': 'Empty', 'source-null': 'Null'}
SCALAR_BODIES = dict(re.findall(r'CASE:([^\n]+)\n(.*?)(?=CASE:|PROBE-ERROR:|$)', SCALARS['raw_result'], re.DOTALL))


@pytest.mark.parametrize('control', CONTROLS, ids=lambda control: control['name'])
def test_complete_native_chart_source_argument_control(control: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(control['source'], name='SourceArgs')
    assert app.run('Probe') == control['raw_result']


@pytest.mark.parametrize('name', SCALAR_EXPRESSIONS)
def test_completed_scalar_case_projects_entire_native_before_result_after_body(name: str) -> None:
    # Only the case driver is narrowed: all helpers and complete case bodies
    # are retained. The disruptive Nothing case is never treated as passing.
    helpers = SCALARS['source'].split('Function Probe() As String', 1)[0]
    source = helpers + ('Function Probe() As String\nDim c As Object,before As String,answer As String\n'
                        'Set c=MakeChart(xlColumnClustered,True)\n'
                        'c.SetSourceData ActiveSheet.Range("A1:D4"),xlColumns\n'
                        'before=SourceState(c)\nanswer=SourceAttempt(c,' + SCALAR_EXPRESSIONS[name]
                        + ',xlColumns)\nProbe="before=" & before & vbLf & "result=" & answer & vbLf & "after=" & SourceState(c) & vbLf\nEnd Function')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='SourceArgs')
    assert app.run('Probe') == SCALAR_BODIES[name]


def test_raw_evidence_and_deferred_diagnostics_are_not_normalized() -> None:
    for name, expected in MANIFEST['raw_sha256'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected
    assert len(CONTROLS) == 15 and len(SCALAR_EXPRESSIONS) == 5
    for control in REMAINING['probes']:
        assert control['status']['ok'] and not control['status']['dialogs']
        assert hashlib.sha256(control['source'].encode()).hexdigest() == control['source_sha256']
    assert hashlib.sha256(SCALARS['source'].encode()).hexdigest() == SCALARS['source_sha256']
    assert not SCALARS['completed'] and 'PROBE-ERROR:-2147417848:' in SCALARS['raw_result']
    assert 'result=ok' in next(control['raw_result'] for control in REMAINING['probes'] if control['name'] == 'source-union')


@pytest.mark.parametrize('name,expression', [('source-union', 'Union(Range("A1:B4"),Range("D1:D4"))'),
                                            ('source-nothing', 'missing')])
def test_deferred_source_forms_refuse_without_claiming_native_error_fidelity(name: str, expression: str) -> None:
    helpers = SCALARS['source'].split('Function Probe() As String', 1)[0]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(helpers + '\nFunction Probe() As String\nDim c As Object,missing As Object\n'
                   'Set c=MakeChart(xlColumnClustered,True)\nProbe=SourceAttempt(c,' + expression
                   + ',xlColumns)\nEnd Function', name='SourceArgs')
    assert name in MANIFEST['deferred']
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe')
