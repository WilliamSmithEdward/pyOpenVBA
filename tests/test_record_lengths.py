"""Native 64-bit VBA record file widths and in-memory padding."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/record_lengths.json').read_text(encoding='utf-8'))['probes']

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_lengths(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


def test_fixed_string_fields_preserve_width_through_record_copy() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Private Const Width As Long = 3\nPrivate Type RecordData\nText As String * Width\nEnd Type\nFunction Probe() As String\nDim first As RecordData, second As RecordData\nfirst.Text = "abcd"\nsecond = first\nfirst.Text = "x"\nProbe = first.Text & "|" & second.Text & "|" & Len(first) & "|" & LenB(first)\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 'x  |abc|3|6'
