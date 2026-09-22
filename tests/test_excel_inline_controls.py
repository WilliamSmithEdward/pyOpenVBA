"""Replay inline list edits and multi-selection measurements from Excel."""
from __future__ import annotations

import json
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.shapes._xlsx import read_controls

FOLDER = Path(__file__).parent / "fixtures" / "shapes"
RECORDS = json.loads((FOLDER / "inline_controls.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("record", RECORDS, ids=[f'{one["kind"]}-{one["name"]}' for one in RECORDS])
def test_measured_inline_edits(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = ('Public Function Report() As String\nDim sh As Object, n As Long, i As Long, v As String\n'
              + str(record["body"]) + 'End Function\n')
    app.add_module(source, name="Probe")
    assert app.run("Report") == record["reported"]
    saved = cast("dict[str, str]", record["parts"])
    measured = read_controls('<control shapeId="1" r:id="rId1"/>',
                             {"rId1": saved["xl/ctrlProps/ctrlProp1.xml"]})[1]
    ours = app.sheet(1).shapes()[0].control
    assert ours is not None
    assert ours.items == measured.items
    assert ours.selection_mode == measured.selection_mode
    assert sorted(ours.selected_indices) == sorted(measured.selected_indices)


def make_list(kind: int = 6) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Public Sub Build()
Dim sh As Object
Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 60)
sh.Name = "Choices"
End Sub''', name="Builder")
    app.run("Build")
    return app


def test_python_inline_items_and_selections_round_trip(tmp_path: Path) -> None:
    app = make_list()
    sheet = app.sheet(1)
    for text in ("one", 'two & "quoted"', "three"):
        sheet.add_control_item("Choices", text)
    sheet.update_control_item("Choices", 1, "new one")
    sheet.set_control_selection_mode("Choices", 2)
    sheet.set_control_selection("Choices", [1, 3])
    path = tmp_path / "inline.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    sheet = app.sheet(1)
    assert sheet.control_items("Choices") == ["new one", 'two & "quoted"', "three"]
    control = sheet.shape("Choices").control
    assert control is not None and control.selected_indices == [1, 3]
    sheet.remove_control_item("Choices", 1)
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False).sheet(1)
    assert reopened.control_items("Choices") == ['two & "quoted"', "three"]
    control = reopened.shape("Choices").control
    assert control is not None and control.selected_indices == [2]
    sheet.clear_control_items("Choices")
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False).sheet(1)
    assert reopened.control_items("Choices") == []
    control = reopened.shape("Choices").control
    assert control is not None and control.selected_indices == []


def test_invalid_inline_edit_is_atomic() -> None:
    app = make_list()
    sheet = app.sheet(1)
    sheet.add_control_item("Choices", "one")
    for operation in (lambda: sheet.remove_control_item("Choices", 1, 2),
                      lambda: sheet.update_control_item("Choices", 1, "bad\x00")):
        with pytest.raises(ValueError):
            operation()
        assert sheet.control_items("Choices") == ["one"]


@pytest.mark.parametrize("mode,native", [(1, -4142), (2, -4154), (3, 3)])
def test_mode_getter_and_selected_flags_match_live_excel(mode: int, native: int) -> None:
    app = make_list()
    sheet = app.sheet(1)
    for text in ("one", "two", "three"):
        sheet.add_control_item("Choices", text)
    sheet.set_control_selection_mode("Choices", mode)
    if mode == 1:
        sheet.set_control_value("Choices", 2)
    else:
        sheet.set_control_selection("Choices", [1, 3])
    assert app.evaluate('ActiveSheet.Shapes("Choices").ControlFormat.MultiSelect') == native
    flags = [app.evaluate(f'ActiveSheet.Shapes("Choices").DrawingObject.Selected({index})')
             for index in (1, 2, 3)]
    assert flags == ([False, True, False] if mode == 1 else [True, False, True])
