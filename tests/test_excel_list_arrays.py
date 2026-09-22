"""Excel-measured whole-list and range-backed editing semantics."""
from __future__ import annotations

import json
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.shapes._xlsx import read_controls

FOLDER = Path(__file__).parent / "fixtures" / "shapes"
RECORDS = json.loads((FOLDER / "list_arrays.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("record", RECORDS,
                         ids=[f'{r["kind"]}-{r["bound"]}-{r["name"]}' for r in RECORDS])
def test_measured_list_arrays(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = ('Public Function Report() As String\n'
              'Dim sh As Object, n As Long, i As Long, v As Variant, a As String, e As Long\n'
              + str(record["body"]) + 'End Function\n')
    app.add_module(source, name="Probe")
    assert app.run("Report") == record["reported"]
    saved = cast("dict[str, str]", record["parts"])
    measured = read_controls('<control shapeId="1" r:id="rId1"/>',
                             {"rId1": saved["xl/ctrlProps/ctrlProp1.xml"]})[1]
    ours = app.sheet(1).shapes()[0].control
    assert ours is not None
    assert (ours.items, ours.value, ours.list_range) == (measured.items, measured.value, measured.list_range)


def test_python_replacement_is_validated_and_persists(tmp_path: Path) -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.update_control("Drop1", linked_cell="$H$2")
    sheet.set_control_value("Drop1", 2)
    with pytest.raises(ValueError):
        sheet.set_control_items("Drop1", ["x", "bad\x00"])
    control = sheet.shape("Drop1").control
    assert control is not None and control.list_range == "$J$1:$J$3"
    assert sheet.value("H2") == 2
    sheet.set_control_items("Drop1", ["x", "y"])
    assert sheet.value("H2") == 0
    assert sheet.value("J1") == "a"
    path = tmp_path / "replaced.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False).sheet(1)
    assert reopened.control_items("Drop1") == ["x", "y"]
    control = reopened.shape("Drop1").control
    assert control is not None and control.list_range == "" and control.value == 0
