"""Workbook.SetPasswordEncryptionOptions and the four PasswordEncryption reads, as live Excel answered them once, typed and late
(tests/fixtures/workbook_encryption_options.json). That probe is retired, since naming the smartcard provider raised smartcard prompts on the
machine, so the cases come from the fixture and the macro is built here. The smartcard provider's cases are refused openly: what Excel
answered there depended on the prompt, not on a rule."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/workbook_encryption_options.json').read_text(encoding='utf-8'))
READS = ('PasswordEncryptionProvider', 'PasswordEncryptionAlgorithm', 'PasswordEncryptionKeyLength', 'PasswordEncryptionFileProperties')
HARDWARE = 'Smart Card'


def _macro(declared: str, case: str) -> str:
    reads = ''.join(f'''    Err.Clear: v = Empty
    v = wb.{read}
    s = s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
''' for read in READS)
    return f'''Function Probe() As String
    Dim wb As {declared}, v As Variant, s As String
    Set wb = Workbooks.Add
    On Error Resume Next
    wb.SetPasswordEncryptionOptions {case}
    s = "e" & Err.Number
{reads}    Probe = s
End Function
'''


def _run(declared: str, case: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_macro(declared, case), name='H')
    return app.run('H.Probe')


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
def test_every_case_answers_as_excel_did(declared: str) -> None:
    cases = [(case, answer) for case, answer in zip(RECORD['cases'], RECORD['answers'][declared]) if HARDWARE not in case]
    differ = [(case, answer, model) for case, answer in cases if (model := _run(declared, case)) != answer]
    assert differ == []


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if HARDWARE in case], ids=lambda case: case[-12:])
def test_the_smartcard_provider_is_refused_openly(case: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run('Object', case)
