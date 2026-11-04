"""What VBA's compiler refuses on a call bound early, read from the type library, against what Excel refused
(scripts/measure_vba_compile_kinds.py)."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter._type_library import CallShape, find_class, refusal

ANSWERS: dict[str, str] = json.loads((Path(__file__).parent / 'fixtures/vba_compile_kinds.json').read_text(encoding='utf-8'))['answers']
CLASSES = {'r': 'Range', 'ws': 'Worksheet', 'wb': 'Workbook', 'f': 'Font', 'sh': 'Shape', 'lo': 'ListObject', 'a': 'Application', 'ch': 'Chart', 'ss': 'Sheets', 'n': 'Name',
           'se': 'Series', 'wi': 'Window', 'pt': 'PivotTable', 'cm': 'Comment', 'hl': 'Hyperlink', 'lc': 'ListColumn', 'co': 'ChartObject', 'it': 'Interior', 'ax': 'Axis', 'va': 'Validation', 'ps': 'PageSetup', 'wbs': 'Workbooks'}
#: Excel stopped on a modal dialog of its own, so that nothing was learned.
ASKED = sorted(key for key, answer in ANSWERS.items() if not answer.startswith('stopped'))


def shape_of(arguments: str) -> CallShape:
    if not arguments:
        return CallShape()
    parts = [part.strip() for part in arguments.split(',')]
    named = tuple(part.split(':=')[0] for part in parts if ':=' in part)
    positional = [part for part in parts if ':=' not in part]
    return CallShape(len(positional), named, frozenset(range(len(positional))))


def predicted(key: str) -> str | None:
    head, rest = key.split(': ', 1)
    family, line = rest.split(' ', 1)
    found = find_class('excel', CLASSES[head])
    assert found is not None
    match = re.search(rf'{head}\.(\w+)(?:\((.*?)\)| (\d))?', line)
    assert match is not None
    use = {'read_only': 'assign', 'assign_method': 'assign', 'with_result': 'with'}.get(family, 'value')
    return refusal('excel', found, match.group(1), use, shape_of(match.group(2) or match.group(3) or ''))


@pytest.mark.parametrize('key', ASKED)
def test_a_call_is_refused_as_excels_compiler_refused_it(key: str) -> None:
    answer = ANSWERS[key]
    said = predicted(key)
    # The harness keeps the first 70 characters of what the dialog says.
    assert (said or 'compiled')[:70] == answer


@pytest.mark.parametrize('key', [key for key in ASKED if ANSWERS[key] != 'compiled'])
def test_the_compiler_refuses_it_in_a_macro(key: str) -> None:
    line = key.split(': ', 1)[1].split(' ', 1)[1]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    declared = ', '.join(f'{name} As {type_}' for name, type_ in CLASSES.items())
    app.add_module(f'Function Probe() As String\n    Dim {declared}, v As Variant, o As Object\n    If False Then\n{line}\n    End If\n    Probe = "compiled"\nEnd Function', name='Harness')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')


@pytest.mark.parametrize('key', [key for key in ASKED if ANSWERS[key] == 'compiled'])
def test_the_compiler_lets_it_through_in_a_macro(key: str) -> None:
    line = key.split(': ', 1)[1].split(' ', 1)[1]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    declared = ', '.join(f'{name} As {type_}' for name, type_ in CLASSES.items())
    app.add_module(f'Function Probe() As String\n    Dim {declared}, v As Variant, o As Object\n    If False Then\n{line}\n    End If\n    Probe = "compiled"\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 'compiled'
