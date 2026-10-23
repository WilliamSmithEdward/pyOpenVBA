"""VBScript.RegExp and FileSystemObject's path functions, replayed against live Excel
(scripts/measure_scripting_objects.py).

Each probe answers one line per case, %% apart; the cases are compared one by one.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/scripting_objects.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}


def _run(source: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    return app.run('Harness.Probe')


def _labels() -> dict[str, list[str]]:
    """What each probe's cases are, read from its macro, for the tests' names."""
    out: dict[str, list[str]] = {}
    for name in ('regexp_first', 'regexp_second'):
        out[name] = re.findall(r'Run\((.*)\)$', PROBES[name]['macro'], re.MULTILINE)
    macro = PROBES['fso_paths']['macro']
    out['fso_paths'] = [*(f'parts {path}' for path in re.findall(r'o\.GetBaseName\((.*?)\) & "\|"', macro)),
                        *(f'join {pair}' for pair in re.findall(r'o\.BuildPath\((.*)\)$', macro, re.MULTILINE)),
                        'TypeName']
    out['regexp_object'] = ['new', 'types', 'empty sub-match', 'item past the end', 'sub-match past the end',
                            'Test(Null)', 'Replace with Null', 'Replace of numbers']
    return out


CASES = [(name, index, label) for name, labels in _labels().items() for index, label in enumerate(labels)]


@pytest.fixture(scope='module')
def replayed() -> dict[str, list[str]]:
    return {name: str(_run(probe['macro'])).split('%%') for name, probe in PROBES.items()}


@pytest.mark.parametrize(('name', 'index'), [(name, index) for name, index, _ in CASES],
                         ids=[f'{name}-{label}' for name, _, label in CASES])
def test_a_case_answers_as_excel_does(replayed: dict[str, list[str]], name: str, index: int) -> None:
    assert replayed[name][index] == PROBES[name]['result'].split('%%')[index]


def test_each_probe_has_its_cases() -> None:
    assert {name: len(labels) for name, labels in _labels().items()} == {
        name: len(probe['result'].split('%%')) for name, probe in PROBES.items()}


def test_early_binding() -> None:
    assert _run('Function Probe() As String\nDim re As New RegExp, fso As New FileSystemObject, m As Match\n'
                're.Pattern = "b(.)"\nFor Each m In re.Execute("abc")\nProbe = m.SubMatches(0) & '
                'fso.GetExtensionName("x.csv")\nNext\nEnd Function\n') == 'ccsv'


def test_the_file_system_stays_out_of_reach() -> None:
    with pytest.raises(VBAUnsupportedError, match='real file system'):
        _run('Function Probe() As Boolean\nProbe = CreateObject("Scripting.FileSystemObject").FileExists("a.txt")\n'
             'End Function\n')
