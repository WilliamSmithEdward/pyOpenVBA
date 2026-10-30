"""VBA under the ja-JP Regional format: ExcelApplication(locale="ja-JP").

tests/fixtures/vba_locale_live.json is what live VBA answered under en-US and
under ja-JP, the Regional format switched by the machine's owner
(scripts/measure_vba_locale_live.py). tests/fixtures/vba_locale.json is
oleaut32's answers with LCID 1041 over many more dates, codes and strings
(scripts/measure_vba_locale.py); VBA calls oleaut32 for these, and the live
run agreed with it wherever both answered.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter import _locale, _values
from pyopenvba.interpreter._format import vba_format
from pyopenvba.interpreter._values import VBADate, to_number

FIXTURES = Path(__file__).parent / "fixtures"
ORACLE: dict[str, Any] = json.loads((FIXTURES / "vba_locale.json").read_text(encoding="utf-8"))
LIVE: dict[str, Any] = json.loads((FIXTURES / "vba_locale_live.json").read_text(encoding="utf-8"))


@pytest.fixture
def japanese(monkeypatch: pytest.MonkeyPatch) -> Any:
    """ja-JP in force, and text with no year read in the year the fixture was measured in."""
    monkeypatch.setattr(_values, "_this_year", lambda: ORACLE["measured_year"])
    with _locale.active(_locale.JA_JP):
        yield


def _moment(parts: list[int]) -> VBADate:
    year, month, day, *clock = parts
    hour, minute, second = [*clock, 0, 0, 0][:3]
    return VBADate.from_datetime(dt.datetime(year, month, day, hour, minute, second))


def test_dates_format_as_oleaut32_formats_them(japanese: None) -> None:
    missed = [(moment, code, expected, vba_format(_moment(moment), code))
              for moment, code, expected in ORACLE["format_dates"]
              if vba_format(_moment(moment), code) != expected]
    assert not missed, missed[:5]


def test_numbers_format_as_oleaut32_formats_them(japanese: None) -> None:
    missed = [(number, code, expected, vba_format(number, code)) for number, code, expected in ORACLE["format_numbers"]
              if vba_format(number, code) != expected]
    assert not missed, missed[:5]


def test_cstr_writes_dates_as_oleaut32_does(japanese: None) -> None:
    missed = [(moment, expected, _values.date_text(_moment(moment))) for moment, expected in ORACLE["date_strings"]
              if _values.date_text(_moment(moment)) != expected]
    assert not missed, missed


def _read_date(text: str) -> float | str | None:
    try:
        found = _values.read_date_text(text)
    except VBAUnsupportedError:
        return "unsupported"
    return None if found is None else found.serial


def test_date_text_reads_as_oleaut32_reads_it(japanese: None) -> None:
    """Every string reads as oleaut32 reads it, or reports itself: a time
    before the date, and A, P or a word joined to AM or PM, which oleaut32
    takes or refuses by rules the model does not have."""
    missed: list[tuple[str, object, object]] = []
    unsupported: list[str] = []
    for text, expected in ORACLE["read_dates"]:
        found = _read_date(text)
        if found == "unsupported":
            unsupported.append(text)
        elif not ((found is None and expected is None) or
                  (isinstance(found, float) and expected is not None and abs(found - expected) < 1e-8)):
            missed.append((text, expected, found))
    assert not missed, missed[:5]
    assert len(unsupported) < len(ORACLE["read_dates"]) // 10, unsupported


def _number(text: str) -> float | None:
    try:
        return float(to_number(text))
    except VBARuntimeError:
        return None


def test_number_text_reads_as_oleaut32_reads_it(japanese: None) -> None:
    """Text with a currency sign or a full-width character reads as oleaut32
    reads it under ja-JP; other text reads as under en-US, whose reading was
    measured live (oleaut32 alone reads 12$$ and refuses -&H10)."""
    missed: list[tuple[str, object, object]] = []
    for text, expected in ORACLE["read_numbers"]:
        if text.isascii() and "\\" not in text and "$" not in text:
            with _locale.active(_locale.EN_US):
                expected = _number(text)
        if _number(text) != expected:
            missed.append((text, expected, _number(text)))
    assert not missed, missed[:5]


def _shown(answer: object) -> str:
    if isinstance(answer, dict):
        return f"Err {answer['error']}"
    assert isinstance(answer, list)
    return f"{answer[0]}:{answer[1]}"


def _probe(expressions: list[str]) -> str:
    lines = ["Private Function Shown(ByVal v As Variant) As String", '    Shown = TypeName(v) & ":" & CStr(v)',
             "End Function", "", "Private Function Tried(ByVal which As Long) As String", "    On Error GoTo Failed",
             "    Select Case which"]
    for index, expression in enumerate(expressions):
        lines += [f"    Case {index}", f"        Tried = Shown({expression})"]
    lines += ["    End Select", "    Exit Function", "Failed:", '    Tried = "Err " & Err.Number', "End Function", "",
              "Public Function Probe() As String", "    Dim s As String, i As Long",
              f"    For i = 0 To {len(expressions) - 1}", '        s = s & "@@@" & Tried(i)', "    Next",
              "    Probe = Mid(s, 4)", "End Function"]
    return "\n".join(lines) + "\n"


@pytest.mark.parametrize("locale", sorted(LIVE))
def test_vba_answers_as_live_vba_did_under_the_regional_format(locale: str) -> None:
    cases = [(label, payload, answer) for label, kind, payload, answer in LIVE[locale]["cases"] if kind == "expr"
             and "Now" not in str(payload)]
    app = ExcelApplication(locale=locale)
    app.add_workbook()
    app.add_module(_probe([str(payload) for _label, payload, _answer in cases]), name="Measure")
    found = str(app.run("Measure.Probe")).split("@@@")
    missed = [(label, _shown(answer), mine) for (label, _payload, answer), mine in zip(cases, found, strict=True)
              if _shown(answer) != mine]
    assert not missed, missed


def test_the_locale_is_named_and_checked() -> None:
    with pytest.raises(ValueError, match="unsupported locale"):
        ExcelApplication(locale="fr-FR")
    assert ExcelApplication(locale="JA-jp").application.locale is _locale.JA_JP


def test_the_locale_holds_for_the_run_only() -> None:
    app = ExcelApplication(locale="ja-JP")
    app.add_workbook()
    app.add_module('Function P() As String\n    P = CStr(DateSerial(2026, 1, 5))\nEnd Function\n', name="M")
    assert app.run("P") == "2026/01/05"
    assert app.evaluate("CStr(DateSerial(2026, 1, 5))") == "2026/01/05"
    assert _locale.current() is _locale.EN_US
    assert ExcelApplication().evaluate("CStr(DateSerial(2026, 1, 5))") == "1/5/2026"


def test_a_date_literal_reads_as_en_us_under_every_locale() -> None:
    app = ExcelApplication(locale="ja-JP")
    app.add_workbook()
    app.add_module('Function P() As String\n    P = Format(#3/4/2026#, "yyyy-mm-dd")\nEnd Function\n', name="M")
    assert app.run("P") == "2026-03-04"
