"""Measure everyday Excel object-model idioms from VBA that the sweep found the
emulator getting wrong: sheet-name rules, a Range's default member with arguments,
the shape of Range.Value over empty cells and Error values read from cells.

Each probe works on a sheet of its own and is recorded as a value or a run-time
error number (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_vba_idioms.json'
PRE = 'Dim ws As Worksheet\nSet ws = Worksheets.Add\nOn Error Resume Next\n'
SHEET_NAMES = ["", "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b", "'ab", "ab'", "a'b", "History", "history",
               "x" * 31, "x" * 32, " lead", "trail ", "a.b", "a,b", "a;b", "a#b"]


def literal(text: str) -> str:
    return text.replace('"', '""')


CASES = {f"sheet_name_{index:02d}": PRE + f'ws.Name = "{literal(name)}"\nProbe = Err.Number & "|" & (ws.Name = "{literal(name)}")'
         for index, name in enumerate(SHEET_NAMES)}
CASES.update({
    "range_call_one": PRE + 'Probe = ws.Range("B2:C3")(1).Address & "|" & ws.Range("B2:C3")(4).Address & "|" & ws.Range("B2:C3")(5).Address',
    "range_call_two": PRE + 'Probe = ws.Range("B2:C3")(2, 2).Address & "|" & ws.Range("B2:C3")(0, 0).Address',
    "range_call_value": PRE + 'ws.Range("B2").Value = 7\nDim x As Variant\nx = ws.Range("B2:C3")(1)\nProbe = TypeName(x) & "|" & x',
    "cells_call": PRE + 'Probe = ws.Cells(2, 2)(2).Address & "|" & ws.Cells(1)(3).Address',
    "cells_item_index": PRE + 'Probe = ws.Range("B2:C3").Cells(2, 2).Address & "|" & ws.Range("B2:C3").Cells(3).Address & "|" & ws.Range("B2:C3")(1).Address & "|" & ws.Range("B2:C3").Item(2, 1).Address',
    "value_of_empty_cells": PRE + 'Dim v\nv = ws.Range("A1:B2").Value\nProbe = TypeName(v) & "|" & UBound(v, 1) & UBound(v, 2) & "|" & IsEmpty(v(2, 2))',
    "value_past_used_cells": PRE + 'ws.Range("A1").Value = 1\nDim v\nv = ws.Range("A1:C3").Value\nProbe = UBound(v, 1) & UBound(v, 2) & "|" & v(1, 1) & "|" & IsEmpty(v(3, 3)) & "|" & UBound(ws.Range("D5:E6").Value, 1)',
    "value_error_cells": PRE + 'ws.Range("A1").Formula = "=1/0"\nws.Range("A2").Formula = "=NA()"\nDim v\nv = ws.Range("A1:A2").Value\nProbe = IsError(v(1, 1)) & "|" & CStr(v(1, 1)) & "|" & (v(2, 1) = CVErr(xlErrNA)) & "|" & ws.Range("A1").Text',
    "listrow_range_item": PRE + 'ws.Range("A1:B2").Value = 1\nDim lo As ListObject\nSet lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)\nDim r As ListRow\nSet r = lo.ListRows.Add\nr.Range(1, 2).Value = "c"\nProbe = lo.ListRows.Count & "|" & r.Range(1, 2).Address & "|" & ws.Range("B3").Value',
    "match_not_found": PRE +'ws.Range("A1:A3").Value = 1\nDim v\nv = Application.Match("zz", ws.Range("A1:A3"), 0)\nProbe = IsError(v) & "|" & (v = CVErr(xlErrNA)) & "|" & TypeName(v)',
})


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
