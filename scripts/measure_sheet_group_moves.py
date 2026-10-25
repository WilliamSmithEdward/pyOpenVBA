"""Where Excel puts a group of sheets it moves or copies beside another, and how it saves a group selected.

Moves: every two-sheet group of four sheets, every two- and three-sheet group of five, each in every array order,
Before and After every sheet outside it, and 200 seeded three- and four-sheet groups of six, most where the
array's first sheet stands past the place the anchor gives and some of the group before it. Each records the tab
order after the move and the active sheet. Copies: every two-sheet group of four, in both orders, Before and After
every sheet outside it, with the order and the active sheet. Saved: a workbook of four sheets saved with One and
Three selected, Three active, and each sheet's sheetView tabSelected and the workbookView's activeTab as the file
has them.

    python scripts/measure_sheet_group_moves.py

writes tests/fixtures/sheet_group_moves.json, which tests/test_excel_sheet_groups.py replays.
"""

from __future__ import annotations

import itertools
import json
import random
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "sheet_group_moves.json"
NAMES = ["One", "Two", "Three", "Four", "Five", "Six"]
BATCH = 40


def move_cases() -> list[tuple[int, list[str], str, bool]]:
    cases: list[tuple[int, list[str], str, bool]] = []
    for count, sizes in ((4, (2,)), (5, (2, 3))):
        sheets = NAMES[:count]
        for size in sizes:
            for group in itertools.permutations(sheets, size):
                for anchor in sheets:
                    if anchor not in group:
                        cases.extend((count, list(group), anchor, after) for after in (False, True))
    pick, sheets, seeded = random.Random(20261025), NAMES, 0
    while seeded < 200:
        group = pick.sample(sheets, pick.choice((3, 4)))
        anchor = pick.choice([name for name in sheets if name not in group])
        after = pick.random() < 0.5
        target = sheets.index(anchor) + (1 if after else 0)
        shifting = sheets.index(group[0]) > target and any(sheets.index(name) < target for name in group)
        if not shifting and pick.random() < 0.7:
            continue
        cases.append((6, group, anchor, after))
        seeded += 1
    return cases


def copy_cases() -> list[tuple[int, list[str], str, bool]]:
    sheets = NAMES[:4]
    return [(4, list(group), anchor, after) for group in itertools.permutations(sheets, 2) for anchor in sheets
            if anchor not in group for after in (False, True)]


def code(cases: list[tuple[int, list[str], str, bool]], verb: str) -> str:
    lines = ["Private Sub Fresh(wb As Object, count As Long)",
             "    Dim i As Long",
             '    wb.Worksheets(1).Name = "One"',
             "    For i = 2 To count",
             '        wb.Worksheets.Add(After:=wb.Worksheets(i - 1)).Name = Choose(i, "One", "Two", "Three", "Four", '
             '"Five", "Six")',
             "    Next",
             '    wb.Worksheets("One").Activate',
             "End Sub",
             "",
             "Private Function Order(wb As Object) As String",
             "    Dim s As Object, out As String",
             "    For Each s In wb.Sheets",
             '        out = out & s.Name & ","',
             "    Next",
             '    Order = out & ActiveSheet.Name & "|"',
             "End Function",
             ""]
    batches = [cases[start:start + BATCH] for start in range(0, len(cases), BATCH)]
    for number, batch in enumerate(batches):
        lines += [f"Private Function Part{number}() As String", "    Dim wb As Object, out As String"]
        for count, group, anchor, after in batch:
            listed = ", ".join(f'"{name}"' for name in group)
            lines += ["    Set wb = Workbooks.Add(xlWBATWorksheet)", f"    Fresh wb, {count}",
                      f'    wb.Sheets(Array({listed})).{verb} {"After" if after else "Before"}:=wb.Sheets("{anchor}")',
                      "    out = out & Order(wb)", "    wb.Close False"]
        lines += [f"    Part{number} = out", "End Function", ""]
    lines += ["Public Function Sweep() As String", "    Dim out As String",
              *(f"    out = out & Part{number}()" for number in range(len(batches))), "    Sweep = out", "End Function"]
    return "\n".join(lines) + "\n"


SAVED_CODE = '''Public Function Saved(path As String) As String
    Dim wb As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "One"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Three"
    wb.Worksheets.Add(After:=wb.Worksheets(3)).Name = "Four"
    wb.Sheets(Array("Three", "One")).Select
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    wb.Close False
    Saved = "saved"
End Function

Public Function Reopened(path As String) As String
    Dim wb As Object, s As Object, out As String
    Set wb = Workbooks.Open(path)
    out = ActiveSheet.Name & "|"
    For Each s In ActiveWindow.SelectedSheets
        out = out & s.Name & ","
    Next
    wb.Close False
    Reopened = out
End Function
'''


def tabs(path: Path) -> dict[str, object]:
    """Each sheet's tabSelected, in tab order, and the workbook's activeTab, as the saved file has them."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        selected = []
        for number in range(1, 5):
            sheet = package.read(f"xl/worksheets/sheet{number}.xml").decode("utf-8")
            view = re.search(r"<sheetView\b[^>]*>", sheet)
            assert view is not None
            selected.append('tabSelected="1"' in view.group(0))
    active = re.search(r'activeTab="(\d+)"', workbook)
    return {"tab_selected": selected, "active_tab": int(active.group(1)) if active else 0}


def rows(value: str, cases: list[tuple[int, list[str], str, bool]]) -> list[dict[str, object]]:
    found = []
    for (count, group, anchor, after), row in zip(cases, [part for part in value.split("|") if part], strict=True):
        *order, active = row.split(",")
        found.append({"sheets": count, "group": group, "anchor": anchor, "after": after, "order": order,
                      "active": active})
    return found


def main() -> None:
    moves, copies = move_cases(), copy_cases()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        moved = excel.run_vba(code(moves, "Move"), "Sweep", timeout=1800.0)
        assert moved.ok, moved.message
        copied = excel.run_vba(code(copies, "Copy"), "Sweep", timeout=1800.0)
        assert copied.ok, copied.message
        path = Path(folder) / "grouped.xlsx"
        saved = excel.run_vba(SAVED_CODE, "Saved", [str(path)], timeout=300.0)
        assert saved.ok, saved.message
        reopened = excel.run_vba(SAVED_CODE, "Reopened", [str(path)], timeout=300.0)
        assert reopened.ok, reopened.message
        file = tabs(path)
    record = {"moves": rows(str(moved.value), moves), "copies": rows(str(copied.value), copies),
              "saved_code": SAVED_CODE, "saved": file, "reopened": str(reopened.value)}
    OUT.write_text(json.dumps(record, indent=0) + "\n", encoding="utf-8")
    print(len(record["moves"]), "moves,", len(record["copies"]), "copies;", file, reopened.value)


if __name__ == "__main__":
    main()
