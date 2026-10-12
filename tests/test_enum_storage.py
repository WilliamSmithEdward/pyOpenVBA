"""A variable, parameter, return, array or record field typed as an enum is a Long."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/enum_storage.json').read_text(encoding='utf-8'))['probes']

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_enum_storage(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


def palette_app(source: str) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Enum Shade\nLight = 2\nEnd Enum', name='Palette')
    app.add_module(source, name='Harness')
    return app


def test_a_public_enum_from_another_module_and_a_private_one_are_longs() -> None:
    # Measured live on 2026-10-12: 'Long|2' for each half.
    app = palette_app('Private Enum Tone\nOne = 1\nEnd Enum\nFunction Probe() As String\n'
                      'Dim a As Shade, b As Tone\na = 1.5\nb = 2.5\n'
                      'Probe = TypeName(a) & "|" & a & "|" & TypeName(b) & "|" & b\nEnd Function')
    assert app.run('Harness.Probe') == 'Long|2|Long|2'


def test_an_enum_type_qualified_by_its_module_is_not_a_type() -> None:
    # Measured live: "Compile error: User-defined type not defined", while
    # Palette.Light and Palette.Shade.Light both read 2.
    app = palette_app('Function Probe() As String\nDim a As Palette.Shade\na = 1\nProbe = TypeName(a)\nEnd Function')
    with pytest.raises(VBACompileError, match='User-defined type not defined'):
        app.run('Harness.Probe')
    app = palette_app('Function Probe() As String\nProbe = CStr(Palette.Light) & "|" & CStr(Palette.Shade.Light)\nEnd Function')
    assert app.run('Harness.Probe') == '2|2'
