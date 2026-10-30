"""Excel's cells under the ja-JP Regional format: ExcelApplication(locale="ja-JP").

tests/fixtures/excel_locale_live.json is what live Excel showed and read
under en-US and under ja-JP, the Regional format switched by the machine's
owner (scripts/measure_excel_locale_live.py): every built-in number format
and more, each on a date, a time, numbers and text; text typed into a cell;
NumberFormatLocal; worksheet functions; Application.International. The same
probe runs here in the model, which must answer as Excel did.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.cells import UnsupportedFormulaError

ROOT = Path(__file__).parents[1]
LIVE: dict[str, Any] = json.loads((ROOT / "tests/fixtures/excel_locale_live.json").read_text(encoding="utf-8"))
_spec = importlib.util.spec_from_file_location("measure_vba_locale_live", ROOT / "scripts/measure_vba_locale_live.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)

#: Cases the model reports itself unable to answer, under every locale and under ja-JP: CELL("format"); a local
#: spelling NumberFormat would spell as a built-in format; a * fill, whose width comes from a font Japanese Excel
#: changes to; and a month's name read as a date by a formula.
UNSUPPORTED = ('=CELL("format",B1)',)
UNSUPPORTED_JA = ("NumberFormatLocal = 'm/d/yyyy'", "_(", '=DATEVALUE("5-Jan-2026")')
#: Gaps of the model's under en-US, which ja-JP shares: a comma right after a digit placeholder in a local
#: spelling, and TEXT with a named format.
KNOWN = ("NumberFormatLocal = '0,00'", '=TEXT(1234.5,"Currency")')


def _skipped(label: str, locale: str) -> bool:
    return label.startswith(UNSUPPORTED + KNOWN) or (locale != "en-US" and label.startswith(UNSUPPORTED_JA))


@pytest.mark.parametrize("locale", sorted(LIVE))
def test_cells_show_and_read_as_live_excel_did_under_the_regional_format(locale: str) -> None:
    chosen = [case for case in LIVE[locale]["cases"] if not _skipped(case[0], locale)]
    cases = [(label, kind, payload) for label, kind, payload, _answer in chosen]
    app = ExcelApplication(locale=locale)
    app.add_workbook()
    app.add_module(PROBE.probe(cases), name="Measure")
    found: list[str] = []
    for first in range(0, len(cases), 60):
        found += str(app.run("Measure.Probe", first, min(first + 59, len(cases) - 1))).split("@@@")
    missed = [(label, answer, PROBE.readable(mine, kind))
              for (label, kind, _payload, answer), mine in zip(chosen, found, strict=True)
              if PROBE.readable(mine, kind) != answer]
    assert not missed, missed[:3]
    assert len(chosen) >= len(LIVE[locale]["cases"]) - 40


def _one(locale: str, case: list[Any]) -> object:
    app = ExcelApplication(locale=locale)
    app.add_workbook()
    app.add_module(PROBE.probe([tuple(case[:3])]), name="Measure")
    return app.run("Measure.Probe", 0, 0)


@pytest.mark.parametrize("label", UNSUPPORTED_JA)
def test_what_japanese_excel_does_otherwise_reports_itself(label: str) -> None:
    case = next(case for case in LIVE["ja-JP"]["cases"] if case[0].startswith(label))
    with pytest.raises((VBAUnsupportedError, UnsupportedFormulaError)):
        _one("ja-JP", case)


def test_the_built_in_formats_keep_their_number_through_a_save(tmp_path: Path) -> None:
    """A built-in format is saved as its number, which Japanese Excel shows its own way, and reads back the same."""
    app = ExcelApplication(locale="ja-JP")
    book = app.add_workbook()
    app.add_module('Sub P()\n    Range("A1").Value = DateSerial(2026, 1, 5)\n    Range("A1").NumberFormat = "m/d/yyyy"\n'
                   'End Sub\nFunction Q() As String\n    Q = Range("A1").Text & "|" & Range("A1").NumberFormatLocal\n'
                   'End Function\n', name="M")
    app.run("P")
    saved = tmp_path / "book.xlsx"
    book.SaveAs(str(saved))
    again = ExcelApplication.open(saved, locale="ja-JP")
    again.add_module('Function Q() As String\n    Q = Range("A1").Text & "|" & Range("A1").NumberFormatLocal\n'
                     'End Function\n', name="M")
    assert app.run("Q") == again.run("Q") == "2026/01/05|yyyy/mm/dd"
    assert ExcelApplication.open(saved).evaluate('Range("A1").Text') == "1/5/2026"
