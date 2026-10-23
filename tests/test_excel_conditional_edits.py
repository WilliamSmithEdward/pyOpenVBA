"""Conditional formats through inserted and deleted rows, columns and cells, replayed against the in-memory model.

tests/fixtures/conditional_edits/ is what scripts/measure_conditional_edits.py
saw in live Excel: layouts that put rules over the numbers r * 10 + c in
A1:J20 and make one structural edit, each dumped as every rule with its
priority, range and formulas or points, then the cells of B1:B16 and A2:J2
that show red; conditional_edits.xlsx, the workbook Excel saved of them all,
with each sheet as Excel read it again and the conditionalFormatting and
extLst it was saved with; and conditional_edits_resaved.xlsx, the same file
edited on a few sheets once Excel had opened it, and saved again.

Excel numbers a rule it makes anew from the workbook's highest priority,
which the layouts before each one raised: a layout replays beside a sheet
holding a rule of the priority Excel's workbook held as the layout began.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest
from _conditional_files import all_guids, formatting, guids

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "conditional_edits"
RECORD: dict[str, Any] = json.loads((FIXTURES / "conditional_edits.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
RESAVED: list[dict[str, Any]] = RECORD["resaved"]
#: What the model reports unsupported: a rule Excel splits, its cells left reading at different distances; one edit
#: making several ranges' rules anew, or one beside rules alike it cuts, which Excel joins; and a colour scale whose
#: formula point lies above its highest value, which the dump draws.
UNSUPPORTED = {"expr_rel_between", "expr_rel_delete", "shift_expr_part", "n_shift_expr_part", "r_two_groups",
               "rx_cut_bar_before", "rx_cut_top_before", "rx_two_cut_before", "x14_other_insert"}
_WRAPPED = "Excel wraps a rule's relative references round the sheet's edges, where the model makes them #REF!"
_JOINED = "Excel joins rules alike over ranges that make one block as it reads a file; the model keeps them apart"
_NAMING = "Excel keeps a rule whose formula names another sheet in its 2010 extension; the model keeps a cfRule"
#: Where the model answers otherwise than Excel, and why.
GAPS = {"order_expr": _WRAPPED}
#: Where the model's file keeps the rules otherwise than Excel's.
FILE_GAPS = {**GAPS, "cell_other_insert": _NAMING, "cell_other_delete": _NAMING}
#: Where the model reads Excel's file otherwise than Excel.
READ_GAPS = {**GAPS, "cell_other_insert": _NAMING, "cell_other_delete": _NAMING,
             **dict.fromkeys(("groups_meet", "n_groups_meet", "lx_two_cut", "same_side", "same_three", "same_overlap",
                              "same_between", "same_relative", "same_reversed"), _JOINED)}
#: Where an edit of Excel's file answers, or keeps the rules, otherwise than Excel.
RESAVED_GAPS = {"lx_two_cut": _JOINED}


def _priorities(answers: str) -> list[int]:
    return [int(found) for found in re.findall(r"\[\w+:Long:(\d+):", answers)]


def _highest() -> dict[str, int]:
    """The workbook's highest priority as each layout began in Excel: the most any sheet before it was left with."""
    highest, out = 0, dict[str, int]()
    for one in LAYOUTS:
        out[one["name"]] = highest
        highest = max([highest, *_priorities(one["answers"])])
    return out


HIGHEST = _highest()


def _report(layout: dict[str, Any]) -> str:
    name, highest = layout["name"], HIGHEST[layout["name"]]
    seed = ["Set s = ActiveWorkbook.Worksheets.Add(After:=ws)", 's.Name = "seed"',
            'Set f = s.Range("A1").FormatConditions.Add(xlCellValue, xlGreater, "0")', f"f.Priority = {highest}",
            "ws.Activate"] if highest else []
    lines = ["Public Function Report() As String",
             "Dim failed As String, r As Long, c As Long, f As Object, g As Object, h As Object, w2 As Object",
             "Dim ws As Object, s As Object", "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{name}"', *seed,
             "On Error Resume Next", "Err.Clear", *(RECORD["grid"] + layout["setup"]).splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "Err.Clear", *layout["edit"].splitlines(),
             'If Err.Number <> 0 Then failed = failed & "E" & Err.Number & ";"', "On Error GoTo 0",
             "ws.Activate", "Report = failed & Dump(ws)", "End Function"]
    return RECORD["helper"] + "\n".join(lines) + "\n"


def _replay(layout: dict[str, Any]) -> tuple[ExcelApplication, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_report(layout), name="Probe")
    return app, str(app.run("Report"))


def edited_layouts(folder: Path, size: int = 120) -> list[tuple[Path, list[dict[str, Any]]]]:
    """Every layout the model edits as Excel did, ``size`` to a workbook, as the model renames no more sheets than
    256, made one after another after a sheet whose rule holds the highest priority Excel's workbook held as each
    layout began, and saved in ``folder``: each file, and the layouts it holds."""
    supported = [one for one in LAYOUTS if one["name"] not in UNSUPPORTED and one["name"] not in FILE_GAPS]
    return [(folder / f"edits_{start}.xlsx", _edited(folder / f"edits_{start}.xlsx", supported[start:start + size]))
            for start in range(0, len(supported), size)]


def _edited(path: Path, made: list[dict[str, Any]]) -> list[dict[str, Any]]:
    # Every sheet is named first, helpers too, as a rename past a rule reading #REF! from a sheet, or past the cells
    # the layouts before have filled, is refused.
    helper = 'Set w2 = ws.Parent.Worksheets.Add(After:=ws)\nw2.Name = ws.Name & "_r"\n'
    named = [name for one in made for name in (one["name"], *((one["name"] + "_r",) if helper in one["setup"] else ()))]
    sheets = ["Public Sub Sheets()", "Dim s As Object", "Set s = ActiveWorkbook.Worksheets(1)", 's.Name = "seed"',
              's.Range("A1").FormatConditions.Add xlCellValue, xlGreater, "0"',
              *(f'ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))'
                f'.Name = "{name}"' for name in named), "End Sub"]
    cases = ["\n".join([f"Public Function Case{index}() As String",
                        "Dim failed As String, r As Long, c As Long, f As Object, g As Object, h As Object, w2 As Object",
                        "Dim ws As Object", f'Set ws = ActiveWorkbook.Worksheets("{one["name"]}")', "ws.Activate",
                        "On Error Resume Next", "Err.Clear",
                        *(RECORD["grid"] + one["setup"].replace(
                            helper, 'Set w2 = ws.Parent.Worksheets(ws.Name & "_r")\n')).splitlines(),
                        'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "Err.Clear",
                        *one["edit"].splitlines(), 'If Err.Number <> 0 Then failed = failed & "E" & Err.Number & ";"',
                        "On Error GoTo 0", "ws.Activate", f"Case{index} = failed & Dump(ws)", "End Function"]) + "\n"
             for index, one in enumerate(made)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(sheets) + "\nPublic Sub Seed(n As Long)\n"
                   'ActiveWorkbook.Worksheets("seed").Cells.FormatConditions(1).Priority = n\nEnd Sub\n'
                   + "".join(cases), name="Builder")
    app.run("Sheets")
    for index, one in enumerate(made):
        app.run("Seed", max(HIGHEST[one["name"]], 1))
        assert app.run(f"Case{index}") == one["answers"], one["name"]
    app.save(path)
    return made


def _marked(layouts: list[dict[str, Any]], unsupported: set[str], gaps: dict[str, str]) -> list[Any]:
    marked: list[Any] = []
    for one in layouts:
        if one["name"] in unsupported:
            marked.append(pytest.param(one, marks=pytest.mark.xfail(strict=True, raises=VBAUnsupportedError,
                                                                    reason="reported unsupported")))
        elif one["name"] in gaps:
            marked.append(pytest.param(one, marks=pytest.mark.xfail(strict=True, reason=gaps[one["name"]])))
        else:
            marked.append(one)
    return marked


@pytest.mark.parametrize("layout", _marked(LAYOUTS, UNSUPPORTED, GAPS), ids=lambda layout: layout["name"])
def test_an_edit_moves_the_rules_as_excel_does(layout: dict[str, Any]) -> None:
    assert _replay(layout)[1] == layout["answers"]


def test_one_workbook_takes_edit_after_edit(tmp_path: Path) -> None:
    """The layouts made one after another in a workbook, as Excel made them, each answering as Excel's did."""
    made = {one["name"] for _, layouts in edited_layouts(tmp_path) for one in layouts}
    assert made == {one["name"] for one in LAYOUTS} - UNSUPPORTED - set(FILE_GAPS)


@pytest.fixture(scope="module")
def excel_file() -> dict[str, str]:
    return formatting(FIXTURES / "conditional_edits.xlsx")


def test_the_record_holds_what_excels_file_holds(excel_file: dict[str, str]) -> None:
    assert all(re.sub(r'dxfId="\d+"', "", one["file"]) == re.sub(r'dxf="<dxf>.*?</dxf>"', "", excel_file[one["name"]])
               for one in LAYOUTS)


@pytest.mark.parametrize("layout", _marked(LAYOUTS, UNSUPPORTED, FILE_GAPS), ids=lambda layout: layout["name"])
def test_the_file_keeps_the_rules_as_excels_does(excel_file: dict[str, str], layout: dict[str, Any],
                                                 tmp_path: Path) -> None:
    app, _ = _replay(layout)
    path = tmp_path / "edited.xlsx"
    app.save(path)
    assert guids(formatting(path)[layout["name"]]) == guids(excel_file[layout["name"]])


@pytest.fixture(scope="module")
def excel_file_read() -> ExcelApplication:
    """The file Excel saved, opened in the model, with a reader for one sheet at a time."""
    app = ExcelApplication.open(FIXTURES / "conditional_edits.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read(name As String) As String\n"
                   "Read = Dump(ActiveWorkbook.Worksheets(name))\nEnd Function\n", name="Reader")
    return app


@pytest.mark.parametrize("layout", _marked(LAYOUTS, {"x14_other_insert"}, READ_GAPS),
                         ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: ExcelApplication, layout: dict[str, Any]) -> None:
    assert excel_file_read.run("Read", layout["name"]) == layout["reopened"]


@pytest.fixture(scope="module")
def resaved(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, object], dict[str, str]]:
    """Excel's file opened in the model and every sheet dumped, as Excel's was, which draws the colour scales over
    the cells the dump reads; then each sheet RESAVED names edited as Excel edited it, in its order, and saved
    again: what each sheet answered then, or the error the model reported, and the file."""
    app = ExcelApplication.open(FIXTURES / "conditional_edits.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read(name As String) As String\n"
                   "Read = Dump(ActiveWorkbook.Worksheets(name))\nEnd Function\n", name="Reader")
    for one in LAYOUTS:
        try:
            app.run("Read", one["name"])
        except VBAUnsupportedError:
            pass
    steps = ["\n".join([f"Public Function Resave{index}() As String", "Dim failed As String, ws As Object",
                        f'Set ws = ActiveWorkbook.Worksheets("{one["name"]}")', "On Error Resume Next", "ws.Activate",
                        "Err.Clear", *one["edit"].splitlines(),
                        'If Err.Number <> 0 Then failed = "E" & Err.Number & ";"', "On Error GoTo 0",
                        f"Resave{index} = failed & Dump(ws)", "End Function"]) + "\n"
             for index, one in enumerate(RESAVED)]
    app.add_module(RECORD["helper"] + "".join(steps), name="Changer")
    answers: dict[str, object] = {}
    for index, one in enumerate(RESAVED):
        try:
            answers[one["name"]] = app.run(f"Resave{index}")
        except VBAUnsupportedError as exc:
            answers[one["name"]] = exc
    path = tmp_path_factory.mktemp("resaved") / "resaved.xlsx"
    app.save(path)
    return answers, formatting(path)


#: Every GUID Excel's first file holds, which a save of it keeps.
KEPT = all_guids(json.dumps([one["file"] for one in LAYOUTS]))


@pytest.mark.parametrize("change", _marked(RESAVED, set(), RESAVED_GAPS), ids=lambda change: change["name"])
def test_an_edited_file_answers_as_excels_does(resaved: tuple[dict[str, object], dict[str, str]],
                                               change: dict[str, Any]) -> None:
    assert resaved[0][change["name"]] == change["answers"]


@pytest.fixture(scope="module")
def excel_resaved() -> dict[str, str]:
    return formatting(FIXTURES / "conditional_edits_resaved.xlsx")


@pytest.mark.parametrize("change", _marked(RESAVED, set(), RESAVED_GAPS), ids=lambda change: change["name"])
def test_an_edited_file_keeps_the_rules_as_excels_does(resaved: tuple[dict[str, object], dict[str, str]],
                                                       excel_resaved: dict[str, str], change: dict[str, Any]) -> None:
    assert guids(resaved[1][change["name"]], KEPT) == guids(excel_resaved[change["name"]], KEPT)
