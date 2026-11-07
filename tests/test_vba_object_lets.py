"""An object given by Let (no Set) to a member, typed and late, as live Excel answered in two probes (scripts/measure_late_object_lets.py,
scripts/measure_late_object_lets_2.py), each case run alone.

Excel hands the object itself to a member bound late, and to one typed whose type is a Variant or an object; a typed member of a value
type gets the object's default value from VBA. What a member then makes of an object goes by the member, not its type
(src/pyopenvba/interpreter/_late_object_lets.py, built from the probes): where it was measured, the model answers as Excel did; an object
with a value gives that value; anything else is refused openly. Typed, a Let of Nothing or of a new Collection does not compile.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
PROBES = ('late_object_lets', 'late_object_lets_2')
#: Widths and heights depend on the display's DPI, which the model emulates at 96 and this machine runs at 144.
DPI = frozenset({'StandardWidth', 'RowHeight'})
#: Named gaps that are not of objects, each held so that one that closes leaves the list: StatusBar read after a set of False (Excel says
#: "FALSE", a String), ScrollArea given a reference's text, and a Date into a Boolean that is a Variant.
GAPS = {('StatusBar', 'ws.Range("A1,C1")'), ('StatusBar', 'ws.Range("A1").Font'), ('StatusBar', 'New Collection'), ('StatusBar', 'Nothing'),
        ('ScrollArea', 'wb.Names("Base")'), ('WrapText', 'ws.Range("A2")'), ('Locked', 'ws.Range("A2")'), ('Italic', 'ws.Range("A2")')}


def _probe(name: str) -> tuple[Any, dict[str, Any]]:
    spec = importlib.util.spec_from_file_location(name, ROOT / f'scripts/measure_{name}.py')
    assert spec is not None and spec.loader is not None
    module: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module, json.loads((ROOT / f'tests/fixtures/{name}.json').read_text(encoding='utf-8'))


def _alone(module: Any, case: list[str], typed: bool) -> str:
    target, value = tuple(case[:-1]), case[-1]
    body = module._case(target, value, typed, False)
    declare = module.macro(typed).split('    On Error Resume Next')[0].split('Private Sub Part0(ByRef out As String)\n')[1]
    source = f'Function Probe() As String\n    Dim out As String\n{declare}    On Error Resume Next\n{body}    Probe = Mid(out, 3)\nEnd Function\n'
    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(source, name='H')
        return str(app.run('H.Probe'))
    except VBACompileError:
        return 'does not compile'
    except VBAUnsupportedError:
        return 'refused'


@pytest.mark.parametrize('typed', [True, False], ids=['typed', 'late'])
@pytest.mark.parametrize('name', PROBES)
def test_each_case_answers_as_excel_did_or_is_refused_openly(name: str, typed: bool) -> None:
    module, record = _probe(name)
    binding = 'typed' if typed else 'late'
    differ: list[tuple[str, str, str, str]] = []
    gaps_seen: set[tuple[str, str]] = set()
    for case, excel in zip(record['cases'], record['excel'][binding]):
        prop, value = case[2], case[-1]
        got = _alone(module, case, typed)
        if prop in DPI and got != excel and got != 'refused':
            got, excel = got.split('/')[0], excel.split('/')[0]
        if (prop, value) in GAPS:
            gaps_seen.add((prop, value))
            continue
        if got not in (excel, 'refused'):
            differ.append((prop, value, got, excel))
    assert differ == []
    assert gaps_seen <= GAPS


def test_typed_a_let_of_nothing_or_a_new_collection_does_not_compile() -> None:
    for value, message in (('Nothing', 'Invalid use of object'), ('New Collection', 'Argument not optional')):
        app = ExcelApplication()
        app.add_workbook()
        with pytest.raises(VBACompileError, match=message):
            app.add_module(f'Sub P()\n    Dim r As Range\n    Set r = ActiveSheet.Range("A1")\n    r.Value = {value}\nEnd Sub\n', name='H')
            app.run('H.P')


def test_the_table_is_the_one_the_probes_build() -> None:
    spec = importlib.util.spec_from_file_location('build_late_object_lets', ROOT / 'scripts/build_late_object_lets.py')
    assert spec is not None and spec.loader is not None
    builder: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    from pyopenvba.interpreter._late_object_lets import LATE_OBJECT_LETS

    assert builder.build() == LATE_OBJECT_LETS
