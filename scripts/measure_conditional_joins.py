"""Which conditional formats Excel joins into one as it reads a file.

Every layout sits on a sheet of its own in one new workbook: the numbers
r * 10 + c in A1:J20 and rules a macro makes over them, which Excel keeps
apart while the workbook stays open. The workbook is saved three ways: as
.xlsx, as .xlsm with no code, and as .xlsm with a module, since Excel
reads a workbook holding a VBA project otherwise. Each file is opened
again and every sheet dumped as tests/fixtures/conditional_edits dumps
one -- each rule with its TypeName, priority, range and formulas, then
which cells of B1:B16 and of A2:J2 show red -- and then saved again, to
show what Excel writes of the rules it joined.

    python scripts/measure_conditional_joins.py

writes tests/fixtures/conditional_joins/: the three files as Excel saved
them, the three it saved again, and conditional_joins.json, which
tests/test_excel_conditional_joins.py replays.
"""

from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_conditional_edits import GRID, HELPER, rule  # noqa: E402
from measure_data_bars_icons import file_parts  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_joins"


def red(where: str) -> str:
    """A rule over ``where`` showing a red fill for numbers over 22."""
    return rule(where, "xlCellValue", "xlGreater", '"22"')


def coloured(where: str, colour: str, operator: str = "xlLess", value: str = '"50"') -> str:
    """Another rule over ``where``, in its own colour."""
    return (f'Set g = ws.Range("{where}").FormatConditions.Add(xlCellValue, {operator}, {value})\n'
            f"g.Interior.Color = {colour}\n")


def made(where: str, add: str, extra: str = "") -> str:
    """A rule ``add`` makes over ``where`` (an Add method and its arguments), with a red fill where it has one."""
    return (f'Set f = ws.Range("{where}").FormatConditions.{add}\n'
            'If TypeName(f) = "FormatCondition" Or TypeName(f) = "Top10" Or TypeName(f) = "AboveAverage" Or '
            'TypeName(f) = "UniqueValues" Then f.Interior.Color = RGB(255, 0, 0)\n' + extra)


#: name -> setup; each works on ws, the layout's own sheet, a rule in f, another in g.
LAYOUTS: dict[str, str] = {
    # Where the ranges lie.
    "j_apart": red("B2:B3") + red("B8:B9"),
    "j_side": red("B2:B3") + red("C2:C3"),
    "j_touch": red("B2:B3") + red("C3:C4"),
    "j_three": red("B2") + red("B3") + red("B4"),
    "j_chain": red("B2") + red("B4") + red("B3"),
    "j_overlap": red("B2:B4") + red("B3:B5"),
    "j_twice": red("B2:B3") + red("B2:B3"),
    "j_inside": red("B2:B5") + red("B3:B4"),
    "j_reversed": red("B4:B5") + red("B2:B3"),
    "j_columns": red("B2:C2") + red("D2:E2"),
    "j_block": red("B2:C3") + red("B4:C5"),
    "j_misaligned": red("B2:C3") + red("C4:D5"),
    "j_multi": red("B2:B3,D7") + red("B4:B5"),
    "j_multi_both": red("B2:B3,D7") + red("B4:B5,D8"),
    "j_multi_apart": red("B2:B3,D7") + red("B4:B5,F9"),
    # Rules between them.
    "j_between": red("B2:B3") + coloured("D2", "RGB(0, 0, 255)", "xlLess", '"5"') + red("B4:B5"),
    "j_cover": red("B2:B3") + coloured("B3:B4", "RGB(0, 0, 255)") + red("B4:B5"),
    "j_cover_bar": red("B2:B3") + 'ws.Range("B4:B5").FormatConditions.AddDatabar\n' + red("B4:B5"),
    "j_cover_receiving": red("B2:B3") + coloured("B2", "RGB(0, 0, 255)") + red("B4:B5"),
    "j_cover_lower": coloured("B4:B5", "RGB(0, 0, 255)") + red("B2:B3") + red("B4:B5"),
    "j_cover_higher": red("B2:B3") + red("B4:B5") + coloured("B4:B5", "RGB(0, 0, 255)"),
    "j_chain_survivor": red("B4") + red("B2") + red("B3"),
    "j_cover_moving_first": red("B4:B5") + coloured("B3:B4", "RGB(0, 0, 255)") + red("B2:B3"),
    "j_renumber": coloured("D2", "RGB(0, 0, 255)", "xlLess", '"5"') + red("B2:B3") + red("B4:B5") +
                  coloured("D3", "RGB(0, 255, 0)", "xlLess", '"6"'),
    "j_group": red("B2:B3") + coloured("B2:B3", "RGB(0, 0, 255)", "xlLess", '"10"') + red("B4:B5"),
    "j_group_both": red("B2:B3") + coloured("B2:B3", "RGB(0, 0, 255)", "xlLess", '"10"') + red("B4:B5") +
                    coloured("B4:B5", "RGB(0, 0, 255)", "xlLess", '"10"'),
    "j_group_order": red("B4:B5") + coloured("B4:B5", "RGB(0, 0, 255)", "xlLess", '"10"') + red("B2:B3"),
    # What the rules hold.
    "j_stop": red("B2:B3") + red("B4:B5") + "f.StopIfTrue = False\n",
    "j_fill": red("B2:B3") + coloured("B4:B5", "RGB(0, 0, 255)", "xlGreater", '"22"'),
    "j_font": red("B2:B3") + red("B4:B5") + "f.Font.Bold = True\n",
    "j_number_format": red("B2:B3") + red("B4:B5") + 'f.NumberFormat = "0.00"\n',
    "j_number_first": red("B2:B3") + 'f.NumberFormat = "0.00"\n' + red("B4:B5"),
    "j_number_both": red("B2:B3") + 'f.NumberFormat = "0.00"\n' + red("B4:B5") + 'f.NumberFormat = "0.000"\n',
    "j_number_custom": red("B2:B3") + red("B4:B5") + 'f.NumberFormat = "0.0"\n',
    "j_number_builtins": red("B2:B3") + 'f.NumberFormat = "0.00"\n' + red("B4:B5") + 'f.NumberFormat = "0"\n',
    "j_number_same": red("B2:B3") + 'f.NumberFormat = "0.000"\n' + red("B4:B5") + 'f.NumberFormat = "0.000"\n',
    "j_number_percent": red("B2:B3") + red("B4:B5") + 'f.NumberFormat = "0%"\n',
    "j_fill_one": red("B2:B3") + 'ws.Range("B4:B5").FormatConditions.Add xlCellValue, xlGreater, "22"\n',
    "j_border": red("B2:B3") + red("B4:B5") + "f.Borders(xlTop).LineStyle = xlContinuous\n",
    "j_italic": red("B2:B3") + red("B4:B5") + "f.Font.Italic = True\n",
    "j_font_colour": red("B2:B3") + red("B4:B5") + "f.Font.Color = RGB(0, 0, 255)\n",
    "j_unformatted": 'ws.Range("B2:B3").FormatConditions.Add xlCellValue, xlGreater, "22"\n'
                     'ws.Range("B4:B5").FormatConditions.Add xlCellValue, xlGreater, "22"\n',
    "j_stop_both": red("B2:B3") + "f.StopIfTrue = False\n" + red("B4:B5") + "f.StopIfTrue = False\n",
    "j_operator": red("B2:B3") + coloured("B4:B5", "RGB(255, 0, 0)", "xlLess", '"22"'),
    "j_value": red("B2:B3") + coloured("B4:B5", "RGB(255, 0, 0)", "xlGreater", '"23"'),
    "j_relative": rule("B2:B3", "xlExpression", first='"=B2>22"') + rule("B4:B5", "xlExpression", first='"=B4>22"'),
    "j_relative_off": rule("B2:B3", "xlExpression", first='"=B2>22"') +
                      rule("B4:B5", "xlExpression", first='"=B2>22"'),
    "j_absolute": rule("B2:B3", "xlExpression", first='"=$B$2>22"') +
                  rule("B4:B5", "xlExpression", first='"=$B$2>22"'),
    "j_between_op": rule("B2:B3", "xlCellValue", "xlBetween", '"20"', '"40"') +
                    rule("B4:B5", "xlCellValue", "xlBetween", '"20"', '"40"'),
    # Each kind of rule.
    "j_text": made("B2:B3", 'Add(xlTextString, String:="8", TextOperator:=xlContains)') +
              made("B4:B5", 'Add(xlTextString, String:="8", TextOperator:=xlContains)'),
    "j_begins": made("B2:B3", 'Add(xlTextString, String:="2", TextOperator:=xlBeginsWith)') +
                made("B4:B5", 'Add(xlTextString, String:="2", TextOperator:=xlBeginsWith)'),
    "j_blanks": made("B2:B3", "Add(xlBlanksCondition)") + made("B4:B5", "Add(xlBlanksCondition)"),
    "j_errors": made("B2:B3", "Add(xlErrorsCondition)") + made("B4:B5", "Add(xlErrorsCondition)"),
    "j_dates": made("B2:B3", "Add(xlTimePeriod, DateOperator:=xlYesterday)") +
               made("B4:B5", "Add(xlTimePeriod, DateOperator:=xlYesterday)"),
    "j_top": made("B2:B3", "AddTop10") + made("B4:B5", "AddTop10"),
    "j_average": made("B2:B3", "AddAboveAverage") + made("B4:B5", "AddAboveAverage"),
    "j_unique": made("B2:B3", "AddUniqueValues") + made("B4:B5", "AddUniqueValues"),
    "j_duplicates": made("B2:B3", "AddUniqueValues", "f.DupeUnique = xlDuplicate\n") +
                    made("B4:B5", "AddUniqueValues", "f.DupeUnique = xlDuplicate\n"),
    "j_scale": made("B2:B3", "AddColorScale(3)") + made("B4:B5", "AddColorScale(3)"),
    "j_bar": made("B2:B3", "AddDatabar") + made("B4:B5", "AddDatabar"),
    "j_icons": made("B2:B3", "AddIconSetCondition") + made("B4:B5", "AddIconSetCondition"),
}

#: The id Excel gives a number format a rule's format holds, for the built-in formats and a few others.
NUMBER_FORMATS = ("0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yyyy",
                  "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm",
                  "#,##0 ;(#,##0)", "#,##0 ;[Red](#,##0)", "#,##0.00;(#,##0.00)", "#,##0.00;[Red](#,##0.00)",
                  "mm:ss", "[h]:mm:ss", "mmss.0", "##0.0E+0", "@", "General", "0.0", "0.000%")
LAYOUTS.update({f"nf_{index}": red("B2:B3") + f'f.NumberFormat = "{code}"\n'
                for index, code in enumerate(NUMBER_FORMATS)})

#: The files, by their name in the record: the format each is saved in, and whether it holds a module.
FILES: dict[str, tuple[int, bool]] = {"xlsx": (51, False), "xlsm": (52, False), "xlsm_code": (52, True)}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, r As Long, c As Long, f As Object, g As Object",
             "On Error Resume Next", "Err.Clear", *(GRID + setup).splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             "ws.Activate", f'Case{index} = failed & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    paths = {name: OUT / f"conditional_joins.{'xlsx' if fmt == 51 else 'xlsm'}" if not code
             else OUT / "conditional_joins_code.xlsm" for name, (fmt, code) in FILES.items()}
    again = {name: path.with_name(path.stem + "_resaved" + path.suffix) for name, path in paths.items()}
    for one in (*paths.values(), *again.values()):
        if one.exists():
            one.unlink()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup))
    for name, (fmt, code) in FILES.items():
        if code:
            build += ['wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString "Sub Kept()" & vbLf & "End Sub"']
        build += [f'wb.SaveAs "{paths[name]}", {fmt}']
    build += ["wb.Close False"]
    for name, (fmt, _) in FILES.items():
        build += ['out = out & "<<PHASE>>"', f'Set wb = Workbooks.Open("{paths[name]}")', "For Each ws In wb.Worksheets",
                  '    out = out & ws.Name & "^" & Dump(ws) & "|"', "Next", f'wb.SaveAs "{again[name]}", {fmt}',
                  "wb.Close False"]
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    for one in (*paths.values(), *again.values()):
        strip_save_path(one)
    phases = str(result.value).split("<<PHASE>>")
    # An answer holding the separator would shift every answer after it onto the wrong layout.
    assert phases[0].count("|") == len(LAYOUTS), "an answer holds the layout separator"
    answers = phases[0].split("|")[: len(LAYOUTS)]
    reopened = {name: dict(one.split("^", 1) for one in phase.split("|") if one)
                for name, phase in zip(FILES, phases[1:], strict=True)}
    parts: dict[str, dict[str, str]] = {}
    for name in FILES:
        with zipfile.ZipFile(paths[name]) as package:
            parts[name] = file_parts(package)
            projects = "xl/vbaProject.bin" in package.namelist()
        with zipfile.ZipFile(again[name]) as package:
            parts[name + "_resaved"] = file_parts(package)
        print(f"{name}: {paths[name].name} holds a VBA project: {projects}")
    record = {"helper": HELPER, "grid": GRID, "files": {name: paths[name].name for name in FILES},
              "layouts": [{"name": name, "setup": setup, "answers": answer,
                           "reopened": {file: reopened[file].get(name, "") for file in FILES},
                           "file": {file: parts[file].get(name, "") for file in parts}}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    (OUT / "conditional_joins.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']}: {one['answers'][:300]}")
        for file in FILES:
            if one["reopened"][file] != one["answers"]:
                print(f"    {file}: {one['reopened'][file][:300]}")


if __name__ == "__main__":
    main()
