"""BeforeDelete, NewChart, the window events, BeforePrint, FollowHyperlink and NewWorkbook, in the order Excel raises
them among the events around them (scripts/measure_more_events.py)."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/more_events.json").read_text(encoding="utf-8"))

#: What the model does not have, so that what Excel raised there cannot replay: chart sheets, a workbook's second
#: window, and a file a printer driver writes.
OUT_OF_REACH = re.compile(r"chart_sheet|chart_to_sheet|delete_chart_sheet|chart_events_off|print_chart|window_new\b"
                          r"|window_new_then_first|print_through|print_empty_through")
ROWS = [row for row in MEASURED["rows"] if not OUT_OF_REACH.fullmatch(row["name"])
        and not OUT_OF_REACH.match(row["name"])]
#: What each row does to leave Excel's saved window as it was, which is the screen's and the model has none of.
LIVE_ONLY = re.compile(r"^(?:Dim wide .*|Application\.WindowState = xlNormal|wide = .*|high = .*|across = .*|down = .*"
                       r"|home\.Activate|ActiveWindow\.(?:Width = wide|Height = high|Left = across|Top = down))\n",
                       re.MULTILINE)
#: Excel's own window settling, which arrives at no set time and no macro asked for.
STRAY_RESIZE = re.compile(r"(?:book-sink-|book-own-|app-sink-)WindowResize:[^|]*\|")


def _names(text: str) -> str:
    """Book numbers count through Excel's session; each row is read with its first book as Book1."""
    numbers: dict[str, str] = {}

    def renamed(found: re.Match[str]) -> str:
        return numbers.setdefault(found.group(0), f"Book{len(numbers) + 1}")
    return re.sub(r"Book\d+", renamed, text)


def _run(action: str, folder: Path) -> str:
    app = ExcelApplication(sandbox=False)
    book = app.add_workbook()
    sheet = book.sheets_[0]
    app.add_module(MEASURED["log"], name="EventLog")
    app.add_module(MEASURED["sink"], name="ExcelSink", kind="class")
    app.add_module(MEASURED["sheet"], name=sheet.code_name, kind="document")
    app.add_module(MEASURED["book"], name=book.code_name, kind="document")
    # Each row's own file, as the rows run side by side and each one's probe deletes it.
    probe = LIVE_ONLY.sub("", MEASURED["probe"]).replace("{action}", action)
    probe = probe.replace('Environ("TEMP") & "\\pyopenvba_open_probe.xlsx"', f'"{folder.as_posix()}/pyopenvba_open_probe.xlsx"')
    app.add_module(probe, name="ProbeModule")
    return str(app.run("Probe"))


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
def test_more_events(row: dict[str, str], tmp_path: Path) -> None:
    expected = row["result"]
    if not row["name"].startswith(("resize_", "state_")):
        expected = STRAY_RESIZE.sub("", expected)
    assert _names(_run(row["action"], tmp_path)) == _names(expected)


def test_a_print_to_a_file_reports_itself() -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    with pytest.raises(VBAUnsupportedError, match="printer driver"):
        ExcelApplication().add_workbook().sheets_[0].PrintOut(PrintToFile=True)


def test_following_a_link_out_of_the_workbook_reports_itself() -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Go()\nActiveSheet.Hyperlinks.Add Range("A1"), "https://example.com/"\n'
                   'ActiveSheet.Hyperlinks(1).Follow\nEnd Sub\n', name="M")
    assert book.sheets_
    with pytest.raises(VBAUnsupportedError, match="another application"):
        app.run("Go")
