"""A formula's last sum or difference snapped to zero, and numbers compared to fifteen digits, as Excel does.

tests/fixtures/zero_snap.json is what scripts/measure_zero_snap.py read
from live Excel: pairs of doubles a few steps of the last bit apart, and
the eight bytes of what formulas over each pair answered. The model is
held to every answer, a block and a formula at a time.
"""

from __future__ import annotations

import json
import struct
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import column_letter
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Range, Worksheet
from pyopenvba.interpreter._values import VBAErrorValue

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "zero_snap.json").read_text(encoding="utf-8"))

#: The row each block starts on, as the measurement laid them out.
FIRST_ROW = {"sweep": 1, "places": 2001, "contexts": 3001, "sums": 4001, "compare": 5001}

#: Answers the model does not give as Excel does yet, and why.
GAPS: dict[tuple[str, str], str] = {
    ("contexts", "array"): "Range.FormulaArray, one array formula over a block, is not implemented",
}


def _double(text: str) -> float:
    return struct.unpack(">d", bytes.fromhex(text))[0]


def _encoded(value: object) -> str:
    """An answer as the measurement's Hex64 writes it."""
    if isinstance(value, VBAErrorValue):
        return f"E{value.number}"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return "S:" + value
    assert isinstance(value, (int, float)), value
    return struct.pack(">d", float(value)).hex().upper()


def _range(sheet: Worksheet, reference: str) -> Range:
    target = sheet.vba_get("Range", [reference])
    assert isinstance(target, Range)
    return target


def _laid_out(block: str) -> tuple[ExcelApplication, Worksheet]:
    """A sheet holding one block's pairs, from its first row, the way the measurement wrote them."""
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    for row, one in enumerate(RECORD[block]["rows"], FIRST_ROW[block]):
        _range(sheet, f"A{row}").vba_set("Value", _double(one["a"]))
        _range(sheet, f"B{row}").vba_set("Value", _double(one["b"]))
    return app, sheet


def _answers(block: str) -> list[dict[str, str]]:
    app, sheet = _laid_out(block)
    first = FIRST_ROW[block]
    last = first + len(RECORD[block]["rows"]) - 1
    names = list(RECORD[block]["formulas"])
    if block == "contexts":
        # The array formula is a known gap; the names and Evaluate are not.
        lines = ["Sub Name_Gaps()"]
        for row in range(first, last + 1):
            refers = f"={sheet.name}!$A${row}-{sheet.name}!$B${row}"
            lines.append(f'ActiveWorkbook.Names.Add Name:="Gap_{row}", RefersTo:="{refers}"')
            lines.append(f'Range("D{row}").Formula = "=Gap_{row}"')
        app.add_module("\n".join([*lines, "End Sub"]) + "\n", name="Names")
        app.run("Name_Gaps")
        columns = {"name": 4}
    else:
        for column, formula in enumerate(RECORD[block]["formulas"].values(), 3):
            letter = column_letter(column)
            _range(sheet, f"{letter}{first}:{letter}{last}").vba_set("FormulaR1C1", formula)
        columns = {name: column for column, name in enumerate(names, 3)}
    out: list[dict[str, str]] = []
    for row in range(first, last + 1):
        answers: dict[str, str] = {}
        for name, column in columns.items():
            answers[name] = _encoded(_range(sheet, f"{column_letter(column)}{row}").vba_get("Value"))
        if block == "contexts":
            answers["ws.Evaluate"] = _encoded(app.evaluate(f'ActiveSheet.Evaluate("A{row}-B{row}")'))
            answers["Application.Evaluate"] = _encoded(
                app.evaluate(f'Application.Evaluate("={sheet.name}!A{row}-{sheet.name}!B{row}")'))
        out.append(answers)
    return out


_CACHE: dict[str, list[dict[str, str]]] = {}


def _cases() -> list[Any]:
    cases: list[Any] = []
    for block in FIRST_ROW:
        for name in RECORD[block]["formulas"]:
            marks = [pytest.mark.xfail(reason=GAPS[(block, name)], strict=True)] if (block, name) in GAPS else []
            cases.append(pytest.param(block, name, marks=marks, id=f"{block}-{name}"))
    return cases


@pytest.mark.parametrize(("block", "name"), _cases())
def test_the_answer_is_excels(block: str, name: str) -> None:
    if (block, name) in GAPS:
        pytest.fail(GAPS[(block, name)])
    if block not in _CACHE:
        _CACHE[block] = _answers(block)
    wrong = [f"{one['a']} {one['b']}: Excel {one['answers'][name]}, ours {got[name]}"
             for one, got in zip(RECORD[block]["rows"], _CACHE[block], strict=True)
             if got[name] != one["answers"][name]]
    assert not wrong, f"{len(wrong)} of {len(_CACHE[block])} differ, first: " + "; ".join(wrong[:3])


def test_the_pairs_are_stored_as_meant() -> None:
    """Each pair is two doubles a known number of steps apart, a the base."""
    for block in FIRST_ROW:
        for one in RECORD[block]["rows"]:
            left, right = _double(one["a"]), _double(one["b"])
            assert left != right and left == left and right == right
