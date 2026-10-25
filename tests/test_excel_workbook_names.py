"""The names Workbooks.Add and a sheet copied out give new workbooks, replayed against the in-memory model.

tests/fixtures/workbook_names.json is what scripts/measure_workbook_names.py
saw in live Excel: a workbook with no template is Book and the next of its
own numbers, one made with xlWBATWorksheet is Sheet and the next of its
own, neither number taken back when a workbook closes, and a sheet copied
into a new workbook makes a Book, whatever the sheet came from.
"""
from __future__ import annotations

import json
from pathlib import Path

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "workbook_names.json").read_text(encoding="utf-8"))


def test_new_workbooks_are_named_as_excel_names_them() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["code"], name="Names")
    assert app.run("Names") == RECORD["names"]
