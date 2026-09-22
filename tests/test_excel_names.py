"""Named-range CRUD measured against Excel."""
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/names_crud.json").read_text())

@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_names(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim nm As Object, n As Long, answer As String\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    assert app.run("Report") == record["reported"]


def test_python_crud_scope_and_round_trip(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).set_value("A1", 10)
    app.sheet(1).set_value("B1", 20)
    app.add_named_range("Choice", "=Sheet1!$A$1", visible=False, comment="workbook")
    local = app.sheet(1).add_named_range("Choice", "=$B$1", comment="local")
    assert local.scope == "Sheet1"
    assert len(app.named_ranges()) == 2
    assert len(app.sheet(1).named_ranges()) == 1
    assert app.sheet(1).value("Choice") == 20
    app.sheet(1).update_named_range("Choice", new_name="LocalChoice", refers_to="=$A$1", visible=False)
    app.update_named_range("Choice", new_name="GlobalChoice", refers_to="=Sheet1!$B$1", comment="changed")
    path = tmp_path / "names.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    assert reopened.named_ranges() == app.named_ranges()
    assert reopened.sheet(1).value("LocalChoice") == 10
    assert reopened.sheet(1).value("GlobalChoice") == 20
    assert reopened.remove_named_range("GlobalChoice")
    assert not reopened.remove_named_range("GlobalChoice")
    assert reopened.sheet(1).remove_named_range("LocalChoice")
    reopened.save(path)
    assert ExcelApplication.open(path, with_vba=False).named_ranges() == []


def test_rename_updates_aliases_and_recalculates() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).set_value("A1", 3)
    app.add_named_range("FirstName", "=Sheet1!$A$1")
    app.add_named_range("AliasName", "=FirstName")
    app.sheet(1).set_value("B1", "=FirstName*2")
    assert app.sheet(1).value("B1") == 6
    app.update_named_range("FirstName", new_name="Renamed")
    assert app.named_range("AliasName").refers_to == "=Renamed"
    assert app.sheet(1).formula("B1") == "=Renamed*2"
    app.update_named_range("Renamed", refers_to="=Sheet1!$A$2")
    app.sheet(1).set_value("A2", 9)
    assert app.sheet(1).value("B1") == 18
    assert app.sheet(1).value("AliasName") == 9
    app.sheet(1).set_value("C1", "=AliasName*2")
    assert app.sheet(1).value("C1") == 18
    app.remove_named_range("Renamed")
    assert str(app.sheet(1).value("C1")) == "Error 2029"


def test_global_and_local_names_are_independently_editable() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_named_range("Choice", "=Sheet1!$A$1")
    app.sheet(1).add_named_range("Choice", "=$B$1")
    app.sheet(1).set_value("C1", "=Choice")
    app.update_named_range("Choice", new_name="GlobalChoice")
    assert app.sheet(1).formula("C1") == "=Choice"
    assert app.named_range("GlobalChoice").scope is None
    app.sheet(1).update_named_range("Choice", new_name="LocalChoice")
    assert app.sheet(1).formula("C1") == "=LocalChoice"


def test_named_multi_area_range_preserves_all_areas(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Range("A1:B2,D4:E5").Name = "Sections"
End Sub
Function Report() As String
Report = ThisWorkbook.Names("Sections").RefersToRange.Address
End Function''', name="Probe")
    app.run("Build")
    assert app.run("Report") == "$A$1:$B$2,$D$4:$E$5"
    path = tmp_path / "areas.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    assert reopened.named_range("Sections") == app.named_range("Sections")


def test_rename_keeps_control_bindings_working_after_reopen(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.set_value("J1", "one")
    sheet.set_value("J2", "two")
    app.add_named_range("Choices", "=Sheet1!$J$1:$J$2")
    sheet.add_named_range("Target", "=$H$2")
    sheet.add_form_control(6, name="Picker")
    sheet.update_control("Picker", linked_cell="Target", list_range="Choices")
    app.update_named_range("Choices", new_name="NewChoices")
    sheet.update_named_range("Target", new_name="NewTarget")
    control = sheet.shape("Picker").control
    assert control is not None
    assert control.linked_cell == "NewTarget"
    assert control.list_range == "NewChoices"
    path = tmp_path / "control_names.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False).sheet(1)
    reopened.set_control_value("Picker", 2)
    assert reopened.value("H2") == 2
    assert len(reopened.control_items("Picker")) == 2
