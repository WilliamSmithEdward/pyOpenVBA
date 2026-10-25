"""Workbook.BuiltinDocumentProperties and CustomDocumentProperties, replayed against the in-memory model.

tests/fixtures/document_properties/ holds what scripts/measure_document_properties.py saw in live Excel: a
workbook for each case, its properties set, added, deleted or linked, an error's number and description where a
step raised one, and every property's name, type and value read back; the same read once the workbook was saved
and opened again, and again where the opened file was saved once more; the docProps parts, package relationships
and content types of each file Excel saved; a workbook Excel saved with its docProps parts taken out, opened, read
and saved; and the machine's time zone as Windows gives it for each year, which Excel wrote its dates in UTC by.

Each case runs here with Excel's own function for it, in a model of its own, the time zone the measured one, the
user the author and the clock held still: what the measurement read as <author> and as now. What the model reads
back goes through the harness's own spelling first, which carried a carriage return home as a line feed and any
other control character as a space.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _document_properties as properties
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "document_properties"
RECORD: dict[str, Any] = json.loads((FIXTURES / "document_properties.json").read_text(encoding="utf-8"))
CASES: dict[str, dict[str, Any]] = {case["name"]: case for case in RECORD["cases"]}
#: The cases Excel's Build made, in its order: each one's Make function.
MADE = [name for name in CASES if name != "no_docprops"]
#: The cases the model does not make, and why.
UNSUPPORTED = {"chart_sheet": "a chart sheet, which the model does not make"}
#: Where the model's names, not its properties, differ from Excel's, measured here and not reproduced yet.
NAME_GAPS = {
    "names_add_context": "Names.Add refuses =Sales[Qty] only while the table's own sheet is active, and an "
                         "unqualified name the active sheet has a name of its own for changes that one; the model "
                         "refuses the reference anywhere and adds a workbook name",
}
#: Where a saved file differs for the same reasons, or for the order the model's names take.
FILE_GAPS = {
    **NAME_GAPS,
    "names_more": "a name beyond printable ASCII, éclair, sorts by its code point in the model, where Windows "
                  "sorts it beside eclair",
}
#: The user the model writes as the author, and the moment its clock is held at.
USER = "Ann Author"
NOW = dt.datetime(2026, 10, 25, 4, 21, 32)

def portable(code: str) -> str:
    """Excel's macro with its folder joined to a file by this system's separator, not Windows' alone."""
    return code.replace('folder & "\\', 'folder & "' + os.sep)


ONE = portable(RECORD["build"].split("Public Function Build", 1)[0]) + '''
Public Function One(index As Long, folder As String) As String
    Dim wb As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Author = wb.BuiltinDocumentProperties("Author").Value
    wb.Close False
    Select Case index
{cases}
    End Select
End Function
'''.replace("{cases}", "".join(f"        Case {index}: One = Make{index}(folder)\n" for index in range(len(MADE))))


def recorded_rule(year: int) -> properties.ZoneRule:
    """The measuring machine's time zone for a year, as Windows gave it."""
    for run in RECORD["zone"]:
        if run["years"][0] <= year <= run["years"][1]:
            return properties.ZoneRule(run["bias"], run["standard_bias"], run["daylight_bias"],
                                       tuple(run["standard"][:7]), tuple(run["daylight"][:7]))
    raise AssertionError(f"no rule recorded for {year}")


def carried(text: str) -> str:
    """Text as the harness carries a macro's answer home: a surrogate pair as its character too."""
    text = text.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return "".join(" " if ord(character) < 32 and character not in "\n\t" else character for character in text)


def readings(answer: str) -> dict[str, tuple[str, str, str | None]]:
    """Each case's reading before the save, once opened again, and once saved and opened a second time."""
    found: dict[str, tuple[str, str, str | None]] = {}
    for one in carried(answer).replace(USER, "<author>").split("~"):
        if "^" in one:
            name, rest = one.split("^", 1)
            session, reopened, *again = rest.split("@@")
            found[name] = (session, reopened, again[0] if again else None)
    return found


def docprops_types(text: str) -> str:
    """The content types the docProps parts are declared with, in their order: what of the part a save of
    properties writes, the other parts' being their own features'."""
    return "".join(re.findall(r'<Override PartName="/docProps/[^"]*"[^>]*/>', text))


def saved_parts(path: Path, now: int) -> dict[str, str]:
    """A file's docProps parts, package relationships and docProps content types, the author written <author> and
    a time within ten minutes of the clock <now>, as the measurement writes Excel's."""
    out: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if name.startswith("docProps/") or name in ("_rels/.rels", "[Content_Types].xml"):
                text = package.read(name).decode("utf-8").replace(f">{USER}<", "><author><")
                if name == "[Content_Types].xml":
                    text = docprops_types(text)

                def stamp(found: re.Match[str]) -> str:
                    if found.group().startswith("0000"):
                        return found.group()
                    when = properties._read_w3c(found.group())  # pyright: ignore[reportPrivateUsage]
                    return "<now>" if abs(when - now) < 600 else found.group()

                out[name] = re.sub(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ", stamp, text)
    return out


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, tuple[str, str, str | None]], Path, int]:
    """Each case made in the model with Excel's function for it, its readings and the folder of the files it saved,
    and the clock's moment in UTC."""
    folder = tmp_path_factory.mktemp("document_properties")
    found: dict[str, tuple[str, str, str | None]] = {}
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(properties, "zone_rule", recorded_rule)
        now = properties.to_utc(NOW)
        for index, name in enumerate(MADE):
            if name not in UNSUPPORTED:
                found.update(_run(name, ONE, "One", index, str(folder)))
        shutil.copy(FIXTURES / "no_docprops.xlsx", folder / "no_docprops.xlsx")
        found.update(_run("no_docprops", portable(RECORD["opened"]), "Opened", str(folder)))
    return found, folder, now


def _run(name: str, code: str, macro: str, *args: object) -> dict[str, tuple[str, str, str | None]]:
    """A case run in a model of its own; a failure stands in for its readings."""
    app = ExcelApplication()
    app.application.user_name = USER
    app.freeze_clock(NOW)
    app.add_module(code, name="Cases")
    try:
        return readings(str(app.run(macro, *args)))
    except Exception as exc:  # noqa: BLE001 -- one case's failure is that case's, not every case's.
        return {name: (f"failed: {exc!r}", "", None)}


def _gaps(names: list[str], gaps: dict[str, str]) -> list[Any]:
    return [pytest.param(name, marks=pytest.mark.xfail(strict=True, reason=gaps[name])) if name in gaps else name
            for name in names]


@pytest.mark.parametrize("name", _gaps([name for name in CASES if name not in UNSUPPORTED], NAME_GAPS))
def test_a_workbooks_properties_read_as_excels(replayed: tuple[dict[str, Any], Path, int], name: str) -> None:
    """Every property's name, type, value and link, and the errors the steps raise, as the macro reads them before
    the save, once the file is opened again and once it is saved and opened a second time."""
    session, reopened, again = replayed[0][name]
    case = CASES[name]
    assert session == case["session"]
    assert reopened == case["reopened"]
    assert again == case["again"]


#: The files Excel saved that the model saves too: all but the one without docProps, which the measurement made.
FILES = [name for name in RECORD["files"] if name != "no_docprops" and name not in UNSUPPORTED]


def excels_parts(name: str) -> dict[str, str]:
    found = dict(RECORD["files"][name])
    found["[Content_Types].xml"] = docprops_types(found["[Content_Types].xml"])
    return found


@pytest.mark.parametrize("name", _gaps(FILES, FILE_GAPS))
def test_a_save_writes_the_docprops_parts_as_excel_does(replayed: tuple[dict[str, Any], Path, int],
                                                        name: str) -> None:
    """core.xml, app.xml and custom.xml, the package's relationships and the parts' content types, byte for byte
    but for the author and the time of the run."""
    _found, folder, now = replayed
    assert saved_parts(folder / f"{name}.xlsx", now) == excels_parts(name)


def test_a_new_xlsx_declares_the_defaults_excels_does(replayed: tuple[dict[str, Any], Path, int]) -> None:
    """No bin part, which the macro-enabled template the model starts from declares, as an xlsm does."""
    with zipfile.ZipFile(replayed[1] / "new.xlsx") as package:
        mine = package.read("[Content_Types].xml").decode("utf-8")
    defaults = re.compile(r"<Default\b[^>]*/>")
    assert defaults.findall(mine) == defaults.findall(RECORD["files"]["new"]["[Content_Types].xml"])


def test_an_xlsm_declares_its_project(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Go()\nEnd Sub\n", name="Kept")
    app.save(tmp_path / "kept.xlsm")
    with zipfile.ZipFile(tmp_path / "kept.xlsm") as package:
        assert 'Extension="bin" ContentType="application/vnd.ms-office.vbaProject"' in \
            package.read("[Content_Types].xml").decode("utf-8")


def _saved_by(code: str, path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(code, name="Saver")
    app.run("Go", str(path))


def test_another_writers_value_type_keeps_custom_xml_as_it_is(tmp_path: Path) -> None:
    """A custom property of a type Excel does not write, here a vt:i8: the custom properties report themselves
    unsupported, and a save keeps the part byte for byte."""
    first, second = tmp_path / "first.xlsx", tmp_path / "second.xlsx"
    _saved_by('Public Sub Go(path As String)\nActiveWorkbook.CustomDocumentProperties.Add Name:="Big", '
              'LinkToContent:=False, Type:=msoPropertyTypeString, Value:="x"\nActiveWorkbook.SaveAs path, 51\n'
              "End Sub\n", first)
    with zipfile.ZipFile(first) as package:
        items = [(info, package.read(info.filename)) for info in package.infolist()]
    odd = b""
    with zipfile.ZipFile(first, "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in items:
            if info.filename == properties.CUSTOM_PART:
                data = odd = data.replace(b"<vt:lpwstr>x</vt:lpwstr>", b"<vt:i8>5000000000</vt:i8>")
            package.writestr(info, data)
    app = ExcelApplication.open(first)
    app.add_module("Public Function Count() As Long\nCount = ActiveWorkbook.CustomDocumentProperties.Count\n"
                   "End Function\nPublic Function Title() As String\nActiveWorkbook.BuiltinDocumentProperties"
                   '("Title").Value = "kept"\nTitle = ActiveWorkbook.BuiltinDocumentProperties("Title").Value\n'
                   "End Function\n", name="Reader")
    with pytest.raises(VBAUnsupportedError):
        app.run("Count")
    assert app.run("Title") == "kept"
    app.save(second)
    with zipfile.ZipFile(second) as package:
        assert package.read(properties.CUSTOM_PART) == odd
        assert b"<dc:title>kept</dc:title>" in package.read(properties.CORE_PART)


@pytest.mark.parametrize("text", ["2024-01-02T03:04:05Z", "2024-01-02T03:04:05.25Z", "2024-01-02T05:04:05+02:00",
                                  "2024-01-01T22:04:05-05:00"])
def test_another_writers_dates_read_as_the_moment_they_name(text: str) -> None:
    moment = properties._read_w3c(text)  # pyright: ignore[reportPrivateUsage]
    assert moment == properties._seconds(dt.datetime(2024, 1, 2, 3, 4, 5))  # pyright: ignore[reportPrivateUsage]


def test_a_file_dated_past_9999_reports_itself() -> None:
    """Excel writes 0000 for a year past 9999, and cannot open the file again (dates_overflow)."""
    with pytest.raises(VBAUnsupportedError):
        properties._read_w3c("0000-01-01T07:59:59Z")  # pyright: ignore[reportPrivateUsage]


def test_half_a_surrogate_pair_reports_itself(tmp_path: Path) -> None:
    """What Excel writes of one half of a pair is not measured; the two halves together are one character."""
    with pytest.raises(VBAUnsupportedError):
        _saved_by('Public Sub Go(path As String)\nActiveWorkbook.BuiltinDocumentProperties("Title").Value = "a" & '
                  "ChrW(&HD83D)\nActiveWorkbook.SaveAs path, 51\nEnd Sub\n", tmp_path / "half.xlsx")


def test_the_model_does_not_make_a_chart_sheet(tmp_path: Path) -> None:
    """Which Excel lists under a heading of its own, Charts, after the worksheets."""
    app = ExcelApplication()
    app.add_module(ONE, name="Cases")
    with pytest.raises(VBAUnsupportedError):
        app.run("One", MADE.index("chart_sheet"), str(tmp_path))


@pytest.mark.parametrize(("local", "utc"), [
    (dt.datetime(2026, 3, 8, 1, 59, 59), "2026-03-08T09:59:59Z"),  # standard time, to the last second
    (dt.datetime(2026, 3, 8, 2, 30), "2026-03-08T09:30:00Z"),  # the hour the clock skips is daylight time
    (dt.datetime(2026, 11, 1, 1, 30), "2026-11-01T08:30:00Z"),  # the hour it repeats is daylight time too
    (dt.datetime(2026, 11, 1, 2, 0), "2026-11-01T10:00:00Z"),
    (dt.datetime(2006, 3, 20, 12), "2006-03-20T20:00:00Z"),  # before 2007, the old rule
    (dt.datetime(1899, 7, 1, 12), "1899-07-01T19:00:00Z"),  # the earliest rule, back to 1601
    (dt.datetime(9999, 12, 31, 23, 59, 59), "0000-01-01T07:59:59Z"),  # past 9999, Excel's own spelling
])
def test_a_local_time_goes_to_utc_by_the_years_rule(local: dt.datetime, utc: str) -> None:
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(properties, "zone_rule", recorded_rule)
        assert properties._w3c(properties.to_utc(local)) == utc  # pyright: ignore[reportPrivateUsage]


def test_utc_comes_back_to_the_local_time_but_in_the_hour_skipped() -> None:
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(properties, "zone_rule", recorded_rule)
        for local in (dt.datetime(2026, 7, 1, 12), dt.datetime(2026, 1, 1), dt.datetime(2026, 11, 1, 1, 30)):
            assert properties.to_local(properties.to_utc(local)) == local
        assert properties.to_local(properties.to_utc(dt.datetime(2026, 3, 8, 2, 30))) == dt.datetime(2026, 3, 8, 1, 30)


def test_without_windows_the_local_zone_is_pythons() -> None:
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(properties, "zone_rule", None)
        local = dt.datetime(2024, 1, 2, 3, 4, 5)
        assert properties.to_local(properties.to_utc(local)) == local


@pytest.mark.skipif(sys.platform != "win32", reason="the time zone comes from Windows")
def test_windows_gives_this_machines_rule() -> None:
    if sys.platform == "win32":
        rule = properties._system_rule(2024)  # pyright: ignore[reportPrivateUsage]
        assert isinstance(rule, properties.ZoneRule)
        assert properties.zone_rule is properties._system_rule  # pyright: ignore[reportPrivateUsage]
        local = dt.datetime(2024, 7, 1, 12)
        assert properties.to_local(properties.to_utc(local)) == local
