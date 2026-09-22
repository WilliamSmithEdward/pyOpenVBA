"""Native conflict defaults and explicit headless deconfliction policies."""
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

RECORDS = json.loads((Path(__file__).parent / "fixtures/range_copy_names.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_name_copy(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim src As Object, dst As Object, nm As Object\n' + record['body'] + 'End Function', name='Probe')
    assert app.run('Report') == record['reported']


@pytest.mark.parametrize("local", [False, True])
def test_rename_dependency_chain_and_roundtrip(tmp_path: Path, local: bool) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = app.sheet(1)
    names = source if local else app
    names.add_named_range('Amount', '=Sheet1!$A$1', visible=False, comment='copied')
    names.add_named_range('Total', '=Amount')
    source.set_value('B1', '=Total+Amount+LEN("Amount")')
    destination = app.add_workbook()
    target = app.sheet(1)
    target.set_value('A1', 7)
    app.add_named_range('amount', '=99')
    app.add_named_range('Amount_2', '=88')
    app.add_named_range('Total', '=77')
    source.copy_range('B1', target, 'D3', name_conflict='rename')
    assert target.value('D3') == 20
    assert app.named_range('amount').refers_to == '=99'
    copied_names = target.named_ranges() if local else app.named_ranges()
    copied = next(n for n in copied_names if n.name.endswith('Amount_3'))
    assert copied.comment == 'copied' and copied.visible is False
    assert source.value('B1') == 6
    path = tmp_path / 'names.xlsx'
    app.save(path, workbook=destination)
    reopened = ExcelApplication.open(path, with_vba=False)
    assert reopened.sheet(1).value('D3') == 20


@pytest.mark.parametrize('policy', ['error', 'rename'])
def test_failure_is_atomic(policy: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = app.sheet(1)
    app.add_named_range('FirstName', '=7')
    app.add_named_range('Conflict', '=Sheet1!A1')
    source.set_value('A1', '=Conflict+FirstName')
    app.add_workbook()
    target = app.sheet(1)
    app.add_named_range('Conflict', '=99')
    target.set_value('C3', 123)
    before = app.named_ranges()
    # Conflict error or unsupported relative-name import both happen before mutation.
    with pytest.raises(VBARuntimeError if policy == 'error' else VBAUnsupportedError, match='conflict|relative'):
        source.copy_range('A1', target, 'C3', name_conflict=policy)
    assert app.named_ranges() == before
    assert target.value('C3') == 123


def test_external_reference_rejection_is_atomic() -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = app.sheet(1)
    source.set_value('A1', '=Sheet1!B1')
    app.add_workbook()
    target = app.sheet(1)
    target.set_value('C3', 42)
    with pytest.raises(VBAUnsupportedError, match='external-link'):
        source.copy_range('A1', target, 'C3')
    assert target.value('C3') == 42
