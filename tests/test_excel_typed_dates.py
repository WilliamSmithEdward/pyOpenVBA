"""Typed dates and times: what scripts/measure_typed_dates.py saw live Excel make of them, replayed in the model.

tests/fixtures/typed_dates.json holds 911 strings written through
Range.Value, a few into a cell given a number format first: a space before
a date or a time, a month's name between slashes or dashes or beside a
number that cannot be its day, a date and a time either way round, what
may follow them, times with a space round a colon or a colon at their
end, where Excel's misreadings stop, and strings made at random from
those parts, which the rules were not fitted to. Each read -- TypeName and
CStr of Value, CStr of Value2, NumberFormat and Text -- is compared
exactly.

Excel misreads two kinds of string into numbers no rule here gives: a
space before a month's name and a time, in a string of nothing but
numbers, month's names, AM or PM and separators; and digits after a point
after AM or PM that end the string. The model refuses those, and every
other string is replayed.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication, _typing
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "typed_dates.json").read_text(
    encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]

HELPER = '''Private Function Typed(ws As Object, first As Long, t As Variant, f As Variant) As String
    Dim out As String, c As Object, i As Long
    For i = LBound(t) To UBound(t)
        Set c = ws.Cells(first + i, 1)
        If f(i) <> "" Then c.NumberFormat = f(i)
        c.Value = t(i)
        out = out & TypeName(c.Value) & "`" & CStr(c.Value) & "`" & CStr(c.Value2) & "`" & c.NumberFormat _
            & "`" & c.Text & "^"
    Next
    Typed = out
End Function
'''


def _refused(case: dict[str, Any]) -> bool:
    """Whether Excel misreads the string: a space, then a month's name and a time; or digits after AM or PM's point."""
    text: str = case["text"]
    words = [word.lower() for word in re.findall("[A-Za-z]+", text)]
    halves = [word in ("a", "am", "p", "pm") for word in words]
    months = [_typing.month_number(word) is not None for word in words]
    dated = text.strip(" ").lstrip("/-")[:1].isalnum() and all(
        half or month for half, month in zip(halves, months, strict=True))
    spaced = text.startswith(" ") and dated and any(months) and (":" in text or any(halves))
    return spaced or re.search(r" [AaPp][Mm]? ?\. ?\d+$", text.rstrip(" ")) is not None


REFUSED = [case for case in CASES if _refused(case)]
REPLAYED = [case for case in CASES if not _refused(case)]


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _module(cases: list[dict[str, Any]]) -> str:
    """The probe scripts/measure_typed_dates.py ran, for ``cases``: batches of strings, a row each."""
    head = ["Public Function Probe() As String", "Dim ws As Object, out As String",
            "Set ws = ActiveWorkbook.Worksheets.Add", 'ws.Columns("A").ColumnWidth = 40']
    body: list[str] = []
    for number, start in enumerate(range(0, len(cases), 100)):
        batch = cases[start:start + 100]
        body += [f"Private Function Batch{number}(ws As Object) As String",
                 f"Dim t(0 To {len(batch) - 1}) As String, f(0 To {len(batch) - 1}) As String"]
        for index, case in enumerate(batch):
            body.append(f"t({index}) = {case.get('expression') or _vba_text(case['text'])}")
            if "format" in case:
                body.append(f"f({index}) = {_vba_text(case['format'])}")
        body += [f"Batch{number} = Typed(ws, {start + 1}, t, f)", "End Function"]
        head.append(f"out = out & Batch{number}(ws)")
    return "\n".join([HELPER, *head, "Probe = out", "End Function", *body]) + "\n"


@pytest.fixture(scope="module")
def answers() -> list[list[str]]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(REPLAYED), name="Probe")
    with pytest.MonkeyPatch.context() as patch:
        # A date typed without a year falls in the year the probe ran.
        patch.setattr(_typing, "_this_year", lambda: RECORD["year"])
        answer = str(app.run("Probe"))
    return [part.split("`") for part in answer.split("^")[: len(REPLAYED)]]


@pytest.mark.parametrize("index", range(len(REPLAYED)), ids=[repr(case["text"]) for case in REPLAYED])
def test_a_typed_date_or_time_reads_as_excel_reads_it(answers: list[list[str]], index: int) -> None:
    assert answers[index] == REPLAYED[index]["answer"]


@pytest.mark.parametrize("case", REFUSED, ids=[repr(case["text"]) for case in REFUSED])
def test_a_string_excel_misreads_is_refused(case: dict[str, Any]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _typing.typed(case["text"], case.get("format", "General"))


def test_only_misread_strings_are_refused() -> None:
    # Excel made numbers of most of them; the rest it kept as text, which the model does not tell apart.
    numbers = [case for case in REFUSED if case["answer"][0] != "String"]
    assert (len(REFUSED), len(numbers)) == (50, 36)
