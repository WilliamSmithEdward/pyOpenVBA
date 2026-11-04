"""What a member answers through a variable of its class, where the fixtures measured that beside the late answer they replay.

Each probe is the late one, with its variable typed: the first procedure of a case declares `o As Object` or `c As Object`.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
TYPED = {'member_audit': ('o', 0), 'format_properties': ('c', None)}
CASES = [(name, int(number)) for name in TYPED for number in json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))['early']]


def typed(source: str, number: int, variable: str, type_name: str) -> str:
    start = source.index(f'Function S{number}()')
    end = source.index('End Function', start)
    body = re.sub(rf'\b{variable} As Object', f'{variable} As {type_name}', source[start:end], count=1)
    return source[:start] + body + source[end:]


@pytest.mark.parametrize(('name', 'number'), CASES, ids=[f'{name}-{number}' for name, number in CASES])
def test_a_typed_variable_answers_as_recorded(name: str, number: int) -> None:
    record = json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))
    case = record['cases'][number]
    variable, _ = TYPED[name]
    type_name = case[0] if name == 'member_audit' else 'Range'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(typed(record['source'], number, variable, type_name), name='Replay')
    if 'Setup' in record['source']:
        app.run('Replay.Setup')
    try:
        found = str(app.run(f'Replay.S{number}'))
    except VBAUnsupportedError:
        pytest.skip('reported as a gap')
    assert found.split('|')[0] == record['early'][str(number)].split('|')[0]
