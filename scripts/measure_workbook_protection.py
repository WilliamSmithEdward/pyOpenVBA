"""Workbook protection from VBA, and in the file, as Excel has it.

Each case runs in a workbook of its own with two sheets, Alpha and Beta:
setup, then an action under On Error, then the error it raised and the
state it left -- whether the structure and windows are protected, the
sheets' names in order and each one's visibility. The cases cover what
Protect sets, the password rules, and which changes to the sheets a
protected structure refuses. Each configuration is then saved, its
workbookProtection element read out of the package, and the file opened
again to read the protection back.

    python scripts/measure_workbook_protection.py

writes tests/fixtures/workbook_protection.json, which tests/test_excel_workbook_protection.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "workbook_protection.json"

#: What every case reads back: the flags, and the sheets in order with their visibility.
STATE = 'wb.ProtectStructure & "," & wb.ProtectWindows & "," & Sheets_(wb)'

HELPER = '''Private Function Sheets_(wb As Object) As String
    Dim ws As Object, out As String
    For Each ws In wb.Worksheets
        out = out & ws.Name & ":" & ws.Visible & " "
    Next
    Sheets_ = Trim(out)
End Function
'''

#: (name, setup, action)
CASES: list[tuple[str, str, str]] = [
    ("protect", "", "wb.Protect"),
    ("protect_password", "", 'wb.Protect "pw"'),
    ("protect_windows", "", "wb.Protect Structure:=True, Windows:=True"),
    ("protect_structure_false", "", "wb.Protect Structure:=False"),
    ("protect_nothing", "", "wb.Protect Structure:=False, Windows:=False"),
    ("add_sheet", "wb.Protect", "wb.Worksheets.Add"),
    ("add_sheet_after", "wb.Protect", "wb.Worksheets.Add After:=wb.Worksheets(2)"),
    ("delete_sheet", "wb.Protect", 'wb.Worksheets("Beta").Delete'),
    ("rename_sheet", "wb.Protect", 'wb.Worksheets("Beta").Name = "Gamma"'),
    ("move_sheet", "wb.Protect", 'wb.Worksheets("Beta").Move Before:=wb.Worksheets("Alpha")'),
    ("copy_sheet", "wb.Protect", 'wb.Worksheets("Beta").Copy After:=wb.Worksheets("Beta")'),
    ("hide_sheet", "wb.Protect", 'wb.Worksheets("Beta").Visible = xlSheetHidden'),
    ("unhide_sheet", 'wb.Worksheets("Beta").Visible = xlSheetHidden: wb.Protect',
     'wb.Worksheets("Beta").Visible = xlSheetVisible'),
    ("very_hidden", "wb.Protect", 'wb.Worksheets("Beta").Visible = xlSheetVeryHidden'),
    ("write_cell", "wb.Protect", 'wb.Worksheets("Beta").Range("A1").Value = 1: '
                                 'out = out & wb.Worksheets("Beta").Range("A1").Value & ";"'),
    ("add_name", "wb.Protect", 'wb.Names.Add Name:="Here", RefersTo:="=Alpha!$A$1": out = out & wb.Names.Count & ";"'),
    ("activate_sheet", "wb.Protect", 'wb.Worksheets("Beta").Activate: out = out & wb.ActiveSheet.Name & ";"'),
    ("sheet_protect", "wb.Protect", 'wb.Worksheets("Beta").Protect: out = out & wb.Worksheets("Beta").ProtectContents & ";"'),
    ("unprotect", "wb.Protect", "wb.Unprotect"),
    ("unprotect_password", 'wb.Protect "pw"', 'wb.Unprotect "pw"'),
    ("unprotect_wrong", 'wb.Protect "pw"', 'wb.Unprotect "bad"'),
    ("unprotect_case", 'wb.Protect "pw"', 'wb.Unprotect "PW"'),
    ("unprotect_any", "wb.Protect", 'wb.Unprotect "anything"'),
    ("unprotect_unprotected", "", "wb.Unprotect"),
    ("protect_twice", "wb.Protect", "wb.Protect"),
    ("protect_twice_password", 'wb.Protect "pw"', 'wb.Protect "other"'),
    ("protect_again_windows", "wb.Protect", "wb.Protect Structure:=True, Windows:=True"),
    ("protect_again_password_windows", 'wb.Protect "pw"', 'wb.Protect "pw", Structure:=True, Windows:=True'),
    ("protect_again_other_windows", 'wb.Protect "pw"', 'wb.Protect "other", Structure:=True, Windows:=True'),
    ("unprotected_again_add", 'wb.Protect "pw": wb.Unprotect "pw"', "wb.Worksheets.Add"),
    ("windows_only_add", "wb.Protect Structure:=False, Windows:=True", "wb.Worksheets.Add"),
    # A second Protect with no Structure turns the protection off: how far that goes.
    ("protect_three_times", "wb.Protect: wb.Protect", "wb.Protect"),
    ("protect_password_twice", 'wb.Protect "pw"', 'wb.Protect "pw"'),
    ("protect_password_twice_then_add", 'wb.Protect "pw": wb.Protect "pw"', "wb.Worksheets.Add"),
    ("protect_structure_true_twice", "wb.Protect Structure:=True", "wb.Protect Structure:=True"),
    ("protect_plain_then_password", "wb.Protect", 'wb.Protect "pw"'),
    ("protect_plain_then_password_after", "wb.Protect: wb.Protect \"pw\"", "wb.Protect: out = out & "
                                                                           "wb.ProtectStructure & \";\""),
    ("protect_password_structure_false", 'wb.Protect "pw"', 'wb.Protect "pw", Structure:=False'),
    ("protect_plain_structure_false", "wb.Protect", "wb.Protect Structure:=False"),
    ("protect_password_then_unprotect_other", 'wb.Protect "pw": wb.Protect "pw", Structure:=True',
     'wb.Unprotect "other"'),
    ("protect_structure_false_unprotected", "", "wb.Protect Structure:=False: wb.Worksheets.Add"),
]

#: Configurations saved to a file: (name, what is done before the save).
FILES: list[tuple[str, str]] = [
    ("plain", "wb.Protect"),
    ("password", 'wb.Protect "pw"'),
    ("windows", "wb.Protect Structure:=True, Windows:=True"),
    ("windows_only", "wb.Protect Structure:=False, Windows:=True"),
    ("unprotected_again", 'wb.Protect "pw": wb.Unprotect "pw"'),
]


def fresh() -> list[str]:
    return ["Set wb = Workbooks.Add(xlWBATWorksheet)", 'wb.Worksheets(1).Name = "Alpha"',
            'wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Beta"']


#: Cases per procedure: a VBA procedure has a size limit.
BATCH = 10


def module(folder: Path) -> str:
    procedures: list[str] = []
    for start in range(0, len(CASES), BATCH):
        lines = [f"Private Function Batch{start // BATCH}() As String", "Dim out As String, wb As Object"]
        for name, setup, action in CASES[start:start + BATCH]:
            lines += [*fresh(), f'out = out & "{name}="', "On Error Resume Next", "Err.Clear"]
            if setup:
                lines += [setup, 'If Err.Number <> 0 Then out = out & "setup " & Err.Number & ";"', "Err.Clear"]
            lines += [action, 'out = out & "E" & Err.Number & ":" & Err.Description & ";"', "Err.Clear",
                      f"out = out & {STATE}", "On Error GoTo 0", 'out = out & "|"', 'wb.Unprotect "pw"',
                      "wb.Close False"]
        lines += [f"Batch{start // BATCH} = out", "End Function"]
        procedures.append("\n".join(lines).replace('wb.Unprotect "pw"',
                                                   'On Error Resume Next: wb.Unprotect "pw": On Error GoTo 0'))
    saves = ["Private Function Saves() As String", "Dim out As String, wb As Object"]
    for name, action in FILES:
        path = folder / f"{name}.xlsx"
        saves += [*fresh(), action, f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
                  f'Set wb = Workbooks.Open("{path}")', f'out = out & "{name}=" & {STATE} & "|"', "wb.Close False"]
    saves += ["Saves = out", "End Function"]
    probe = ["Public Function Probe() As String", "Dim out As String", "Application.DisplayAlerts = False",
             *[f"out = out & Batch{index}()" for index in range(len(procedures))],
             'out = out & "#" & Saves()', "Probe = out", "End Function"]
    return "\n".join([HELPER, *procedures, "\n".join(saves), *probe]) + "\n"


def element(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookProtection\b[^>]*/>", xml)
    if found is None:
        return {"element": "", "before": "", "after": ""}
    before = re.findall(r"<([A-Za-z]+)\b[^<]*>\s*$", xml[:found.start()])
    after = re.match(r"\s*<([A-Za-z]+)", xml[found.end():])
    return {"element": found.group(0), "before": before[-1] if before else "", "after": after.group(1) if after else ""}


def main() -> None:
    with tempfile.TemporaryDirectory() as temporary:
        folder = Path(temporary)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(module(folder), "Probe", timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        runtime, saved = str(result.value).split("#")
        answers = dict(part.split("=", 1) for part in runtime.split("|") if part)
        reread = dict(part.split("=", 1) for part in saved.split("|") if part)
        files = [{"name": name, "action": action, **element(folder / f"{name}.xlsx"), "read": reread[name]}
                 for name, action in FILES]
    record = {"helper": HELPER, "state": STATE,
              "cases": [{"name": name, "setup": setup, "action": action, "answer": answers[name]}
                        for name, setup, action in CASES], "files": files}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, _, _ in CASES:
        print(f"{name:32} {answers[name]}")
    for one in files:
        print(f"{one['name']:20} {one['element']}  ({one['before']} / {one['after']})  read {one['read']}")


if __name__ == "__main__":
    main()
