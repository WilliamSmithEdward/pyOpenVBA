"""A conditional format's number format, replayed against the in-memory model.

tests/fixtures/conditional_number_formats/ is what
scripts/measure_conditional_number_formats.py saw in live Excel: one
workbook whose sheets each hold a rule given a number format, Excel's
built-in ones in the order its Custom list shows them, codes Excel spells
another way and custom ones, some also on a cell, each read back at once;
the workbook as Excel saved it, each rule read again from it, and the file
Excel saved again once a rule made after opening took a new format.
chars_*.xlsx hold every printable character in a rule's format and a
cell's, bare and escaped, after a digit, before one, between date codes
and after the text placeholder.

foreign.json and foreign/ are what
scripts/measure_conditional_number_format_files.py saw: copies of one
workbook whose rules' numFmts take ids and spellings Excel does not
write, each opened in Excel, read and saved again. refusals.json is what
scripts/measure_number_format_refusals.py saw of whole codes given to a
rule and a cell: which Excel refuses and how it reads back the rest.
"""

from __future__ import annotations

import html
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest
from _conditional_files import number_formats

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._number_format import from_file, normalized, to_file

FIXTURES = Path(__file__).parent / "fixtures" / "conditional_number_formats"
RECORD: dict[str, Any] = json.loads((FIXTURES / "conditional_number_formats.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
CHARACTERS: list[dict[str, Any]] = RECORD["characters"]
FOREIGN: dict[str, Any] = json.loads((FIXTURES / "foreign.json").read_text(encoding="utf-8"))
REFUSALS: list[dict[str, Any]] = json.loads((FIXTURES / "refusals.json").read_text(encoding="utf-8"))["codes"]
#: The character probes each workbook held, as the measurement made them: Excel holds about 250 custom formats.
CHUNK = 120
_CELLS = ("Excel numbers the cells' own custom formats from 164 as it opens a file and saves them so; the model "
          "writes a stylesheet no macro changed as the file had it")


def _vba(text: str) -> str:
    """A VBA string literal of ``text``, as the measurement wrote one."""
    return '"' + text.replace('"', '""') + '"' if text.isascii() else \
        " & ".join(f"ChrW({ord(char)})" if ord(char) > 127 else '"' + char.replace('"', '""') + '"' for char in text)


def _layout_code(index: int, setup: str, code: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim f As Object, out As String",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "0")',
             "f.Interior.Color = RGB(255, 0, 0)", "Err.Clear", f"f.NumberFormat = {_vba(code)}",
             'If Err.Number <> 0 Then out = "E" & Err.Number & ";"', "Err.Clear",
             'out = out & TypeName(f.NumberFormat) & ":" & f.NumberFormat',
             'If Err.Number <> 0 Then out = out & "R" & Err.Number', "On Error GoTo 0", f"Case{index} = out",
             "End Function"]
    return "\n".join(lines) + "\n"


def made_layouts(path: Path) -> list[str]:
    """Every layout made in one workbook as Excel made them and saved to ``path``: what each rule read back."""
    build = ["Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Set wb = ActiveWorkbook"]
    bodies: list[str] = []
    for index, layout in enumerate(LAYOUTS):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"]}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(_layout_code(index, layout["setup"], layout["code"]))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*build, "Build = out", "End Function"]) + "\n" + "".join(bodies), name="Builder")
    answers = str(app.run("Build")).split("|")[: len(LAYOUTS)]
    app.save(path)
    return answers


#: A copy's reader: each rule's NumberFormat and what its cell shows, and the cells' own formats.
FOREIGN_READER = RECORD["reader"] + (
    "Public Function Read() As String\nDim wb As Object\nSet wb = ActiveWorkbook\n"
    'Read = Shown(wb.Worksheets("s1")) & "~" & wb.Worksheets("s1").Range("B2").Text & "~" & '
    'Shown(wb.Worksheets("s2")) & "~" & wb.Worksheets("s2").Range("B2").Text & "~" & '
    'wb.Worksheets("s1").Range("A1").NumberFormat & "~" & wb.Worksheets("s1").Range("A2").NumberFormat\n'
    "End Function\n")


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    """Every layout made in one workbook as Excel made them, what each rule read back, and the file saved."""
    path = tmp_path_factory.mktemp("number_formats") / "made.xlsx"
    return made_layouts(path), path


@pytest.mark.parametrize("index", range(len(LAYOUTS)), ids=[layout["name"] for layout in LAYOUTS])
def test_a_rule_reads_its_number_format_back_as_excel_does(made: tuple[list[str], Path], index: int) -> None:
    assert made[0][index] == LAYOUTS[index]["answer"]


@pytest.fixture(scope="module")
def made_file(made: tuple[list[str], Path]) -> tuple[dict[str, str], list[str]]:
    return number_formats(made[1])


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_rule_saves_its_number_format_as_excel_does(made_file: tuple[dict[str, str], list[str]],
                                                      layout: dict[str, Any]) -> None:
    """A built-in format under its place in Excel's list, a custom one under the id it took when made, the ids the
    cells' formats take too, and every literal escaped."""
    assert made_file[0][layout["name"]] == layout["file"]


def test_the_cells_formats_share_the_rules_ids(made_file: tuple[dict[str, str], list[str]]) -> None:
    assert made_file[1] == RECORD["cells"]["file"]


@pytest.fixture(scope="module")
def reopened(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], tuple[dict[str, str], list[str]]]:
    """Excel's file opened in the model, each rule read, a rule made after opening, and the file saved again."""
    app = ExcelApplication.open(FIXTURES / "conditional_number_formats.xlsx", with_vba=False)
    app.add_module(RECORD["reader"] + "\n".join([
        "Public Function Read() As String", "Dim ws As Object, out As String, f As Object",
        "For Each ws In ActiveWorkbook.Worksheets", '    out = out & ws.Name & "^" & Shown(ws) & "|"', "Next",
        "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))",
        'ws.Name = "after_open"', 'Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "0")',
        'f.NumberFormat = "0.0000000"', 'out = out & "after_open^" & Shown(ws) & "|"', "Read = out",
        "End Function"]) + "\n", name="Reader")
    read = dict(one.split("^", 1) for one in str(app.run("Read")).split("|") if one)
    path = tmp_path_factory.mktemp("number_formats_again") / "again.xlsx"
    app.save(path)
    return read, number_formats(path)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_rule_read_from_a_file_reads_as_excel_does(reopened: tuple[dict[str, str], Any],
                                                     layout: dict[str, Any]) -> None:
    assert reopened[0][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_file_saved_again_renumbers_its_rules_formats_as_excel_does(reopened: tuple[Any, tuple[dict[str, str], Any]],
                                                                      layout: dict[str, Any]) -> None:
    """Opening a file numbers its rules' custom formats after the cells' own, in the order the file lists its
    dxfs, one code under one id."""
    assert reopened[1][0][layout["name"]] == layout["resaved"]


def test_a_rule_made_after_opening_takes_the_next_id(reopened: tuple[dict[str, str],
                                                                     tuple[dict[str, str], list[str]]]) -> None:
    assert (reopened[0]["after_open"], reopened[1][0]["after_open"]) == (RECORD["after_open"]["answer"],
                                                                         RECORD["after_open"]["resaved"])


@pytest.mark.xfail(strict=True, reason=_CELLS)
def test_a_file_saved_again_renumbers_its_cells_formats_as_excel_does(
        reopened: tuple[Any, tuple[dict[str, str], list[str]]]) -> None:
    assert reopened[1][1] == RECORD["cells"]["resaved"]


def _character_code(start: int, probes: list[str], *, name: str) -> str:
    lines = [f"Private Function {name}(ws As Object) As String",
             "Dim f As Object, out As String, a As String, b As String", "On Error Resume Next"]
    for row, code in enumerate(probes, start + 1):
        lines += ["Err.Clear", f'Set f = ws.Cells({row}, 2).FormatConditions.Add(xlCellValue, xlGreater, "0")',
                  f"f.NumberFormat = {_vba(code)}",
                  'If Err.Number <> 0 Then a = "E" & Err.Number Else a = f.NumberFormat', "Err.Clear",
                  f"ws.Cells({row}, 1).NumberFormat = {_vba(code)}",
                  f'If Err.Number <> 0 Then b = "E" & Err.Number Else b = ws.Cells({row}, 1).NumberFormat',
                  'out = out & a & "<=>" & b & "<<>>"']
    lines += ["On Error GoTo 0", f"{name} = out", "End Function"]
    return "\n".join(lines) + "\n"


def _within(text: str, section: str) -> str:
    """The ``section`` element of ``text``, or nothing where it holds none."""
    found = re.search(rf"<{section}\b.*?</{section}>", text, re.DOTALL)
    return found.group() if found else ""


def _character_formats(path: Path, count: int) -> list[tuple[str, str]]:
    """Each probe's row of the sheet "chars": its rule's numFmt and its cell's, as the file keeps them."""
    with zipfile.ZipFile(path) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
    dxfs = re.findall(r"<dxf>.*?</dxf>", _within(styles, "dxfs"))
    codes = {found.group(1): found.group(0) for found in
             re.finditer(r'<numFmt numFmtId="(\d+)"[^>]*/>', _within(styles, "numFmts"))}
    xfs = re.findall(r"<xf\b[^>]*?>", _within(styles, "cellXfs"))
    rules = {int(found.group(1)): int(found.group(2)) for found in re.finditer(
        r'<conditionalFormatting sqref="B(\d+)"><cfRule\b[^>]*?dxfId="(\d+)"', sheet)}
    cells = {int(found.group(1)): int(found.group(2)) for found in re.finditer(r'<c r="A(\d+)" s="(\d+)"', sheet)}
    out: list[tuple[str, str]] = []
    for row in range(1, count + 1):
        rule = re.search(r"<numFmt [^>]*/>", dxfs[rules[row]]) if row in rules else None
        identifier = re.search(r'numFmtId="(\d+)"', xfs[cells[row]]) if row in cells else None
        number = identifier.group(1) if identifier else ""
        out.append((rule.group() if rule else "", codes.get(number, f'<numFmt numFmtId="{number}"/>')))
    return out


@pytest.fixture(scope="module")
def characters(tmp_path_factory: pytest.TempPathFactory) -> list[tuple[str, str, str, str]]:
    """Each character probe in the model, a workbook to each batch as the measurement made them: what the rule and
    the cell read back, and the numFmt the saved file gives each."""
    folder = tmp_path_factory.mktemp("number_format_characters")
    codes = [one["code"] for one in CHARACTERS]
    out: list[tuple[str, str, str, str]] = []
    for chunk, start in enumerate(range(0, len(codes), CHUNK)):
        names: list[str] = []
        bodies: list[str] = []
        for first in range(start, min(start + CHUNK, len(codes)), 100):
            names.append(f"Chars{first}")
            bodies.append(_character_code(first - start, codes[first:min(first + 100, start + CHUNK)],
                                          name=f"Chars{first}"))
        app = ExcelApplication()
        app.add_workbook()
        app.add_module("Public Function Build() As String\nDim ws As Object, out As String\n"
                       'Set ws = ActiveWorkbook.Worksheets(1)\nws.Name = "chars"\n'
                       + "".join(f"out = out & {name}(ws)\n" for name in names)
                       + "Build = out\nEnd Function\n" + "".join(bodies), name="Builder")
        shown = [one.split("<=>") for one in str(app.run("Build")).split("<<>>")[:-1]]
        path = folder / f"chars_{chunk}.xlsx"
        app.save(path)
        spelled = _character_formats(path, len(shown))
        out += [(rule, cell, rule_file, cell_file) for (rule, cell), (rule_file, cell_file)
                in zip(shown, spelled, strict=True)]
    return out


@pytest.mark.parametrize("index", range(len(CHARACTERS)), ids=[one["code"] for one in CHARACTERS])
def test_a_character_in_a_format_reads_and_saves_as_excel_does(characters: list[tuple[str, str, str, str]],
                                                               index: int) -> None:
    """A rule's format and a cell's read back alike, and the file escapes every character either shows as it is."""
    want = CHARACTERS[index]
    assert characters[index] == (want["rule"], want["cell"], want["rule_file"], want["cell_file"])


@pytest.fixture(scope="module")
def refusals() -> list[tuple[str, str]]:
    """Each code of refusals.json given to a rule and to a cell in the model, a hundred to a workbook as the
    measurement gave them: what each reads back, or the error refusing it."""
    codes = [one["code"] for one in REFUSALS]
    out: list[tuple[str, str]] = []
    for start in range(0, len(codes), 100):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module("Public Function Build() As String\nBuild = Probe(ActiveWorkbook.Worksheets(1))\n"
                       "End Function\n" + _character_code(start, codes[start:start + 100], name="Probe"),
                       name="Builder")
        out += [(rule, cell) for rule, cell in (one.split("<=>") for one in str(app.run("Build")).split("<<>>")[:-1])]
    return out


@pytest.mark.parametrize("index", range(len(REFUSALS)), ids=[one["code"] for one in REFUSALS])
def test_a_code_is_refused_or_read_back_as_excel_does(refusals: list[tuple[str, str]], index: int) -> None:
    """Excel refuses a code mixing a number's, a date's and text's characters in one section, a section out of
    place, a bracket it does not read, and a letter it keeps for a code where it is none; an accented letter is
    the letter under the accent."""
    assert refusals[index] == (REFUSALS[index]["rule"], REFUSALS[index]["cell"])


@pytest.mark.parametrize("one", [one for one in CHARACTERS if not one["rule"].startswith("E")],
                         ids=lambda one: one["code"])
def test_a_code_read_back_from_its_file_spelling_is_the_same(one: dict[str, Any]) -> None:
    """What a file spells, read again, is what NumberFormat read: the two spellings stand for one code."""
    spelled = re.search(r'formatCode="([^"]*)"', one["rule_file"])
    assert spelled is not None
    assert (to_file(normalized(one["code"])), from_file(to_file(one["rule"]))) == (html.unescape(spelled.group(1)),
                                                                                   one["rule"])


@pytest.fixture(scope="module")
def foreign(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[str, tuple[dict[str, str], list[str]]]]:
    """Each copy another program might have written, opened in the model, read as Excel read it, and saved again."""
    folder = tmp_path_factory.mktemp("number_formats_foreign")
    out: dict[str, tuple[str, tuple[dict[str, str], list[str]]]] = {}
    for case in FOREIGN["cases"]:
        app = ExcelApplication.open(FIXTURES / "foreign" / f"{case['name']}.xlsx", with_vba=False)
        app.add_module(FOREIGN_READER, name="Reader")
        answer = str(app.run("Read"))
        out[case["name"]] = answer, number_formats(app.save(folder / f"{case['name']}.xlsx"))
    return out


@pytest.mark.parametrize("case", FOREIGN["cases"], ids=lambda case: case["name"])
def test_a_rule_another_program_wrote_reads_as_excel_reads_it(
        foreign: dict[str, tuple[str, Any]], case: dict[str, Any]) -> None:
    """Excel reads a rule's number format by its code alone: a built-in id over another code, or a cell's custom
    id over another code, names nothing."""
    assert foreign[case["name"]][0] == case["answer"]


@pytest.mark.parametrize("case", FOREIGN["cases"], ids=lambda case: case["name"])
def test_a_rule_another_program_wrote_is_saved_again_as_excel_saves_it(
        foreign: dict[str, tuple[str, tuple[dict[str, str], list[str]]]], case: dict[str, Any]) -> None:
    """Saved again, each rule's format takes Excel's id for its code and Excel's spelling of it."""
    assert foreign[case["name"]][1][0] == case["resaved"]


@pytest.mark.xfail(strict=True, reason=_CELLS)
def test_the_cells_formats_of_a_copy_are_saved_again_as_excel_saves_them(
        foreign: dict[str, tuple[str, tuple[dict[str, str], list[str]]]]) -> None:
    assert [foreign[case["name"]][1][1] for case in FOREIGN["cases"]] == [
        case["resaved_cells"] for case in FOREIGN["cases"]]
