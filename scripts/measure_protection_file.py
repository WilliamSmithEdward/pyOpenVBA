"""How Excel saves sheet protection, and how it reads it back.

Each configuration protects a fresh workbook's sheet from VBA and saves
it as xlsx; the sheetProtection element Excel wrote, and the elements
either side of it, are read out of the package. Each saved file is then
opened again in Excel and its protection read back through the object
model, so the attributes are pinned in both directions. The password
cases give a salt and a hash to fit Excel's algorithm against.

    python scripts/measure_protection_file.py

writes tests/fixtures/protection_file.json, which tests/test_excel_protection_file.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "protection_file.json"

#: (name, what is done to the sheet before the save)
CONFIGS: list[tuple[str, str]] = [
    ("plain", "ws.Protect"),
    ("password", 'ws.Protect "pw"'),
    ("password_long", 'ws.Protect "Correct Horse 9!"'),
    ("contents_false", "ws.Protect Contents:=False"),
    ("drawing_scenarios_false", "ws.Protect DrawingObjects:=False, Scenarios:=False"),
    ("user_interface_only", "ws.Protect UserInterfaceOnly:=True"),
    ("allow_formatting_cells", "ws.Protect AllowFormattingCells:=True"),
    ("allow_everything", "ws.Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, "
                         "AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows:=True, "
                         "AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, "
                         "AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True"),
    ("select_unlocked", "ws.EnableSelection = xlUnlockedCells: ws.Protect"),
    ("select_none", "ws.EnableSelection = xlNoSelection: ws.Protect"),
    ("unprotected_again", 'ws.Protect "pw": ws.Unprotect "pw"'),
    ("unlocked_cell", 'ws.Range("A1").Locked = False: ws.Protect'),
    ("allow_everything_select_none", "ws.EnableSelection = xlNoSelection: ws.Protect \"pw\", "
                                     "DrawingObjects:=False, AllowFormattingCells:=True, "
                                     "AllowFormattingColumns:=True, AllowFormattingRows:=True, "
                                     "AllowInsertingColumns:=True, AllowInsertingRows:=True, "
                                     "AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, "
                                     "AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, "
                                     "AllowUsingPivotTables:=True"),
]

#: Passwords written with the legacy 16-bit hash some files carry instead of SHA-512, and what Unprotect is tried with.
LEGACY: list[tuple[str, str, str]] = [
    ("legacy_right", "pw", "pw"), ("legacy_wrong", "pw", "bad"), ("legacy_case", "pw", "PW"),
    ("legacy_long_right", "Correct Horse9!", "Correct Horse9!"), ("legacy_long_wrong", "Correct Horse9!", "Correct"),
]


def legacy_hash(password: str) -> str:
    """The 16-bit hash of ECMA-376's legacy password attribute, for an ASCII password."""
    value = 0
    for index, char in enumerate(password, 1):
        bits = ord(char) << index
        value ^= (bits & 0x7FFF) | (bits >> 15)
    return f"{value ^ len(password) ^ 0xCE4B:04X}"


def legacy_file(template: Path, path: Path, password: str) -> None:
    """``template`` with its sheet's protection swapped for one carrying the legacy hash of ``password``."""
    with zipfile.ZipFile(template) as source, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<sheetProtection\b[^>]*/>",
                              f'<sheetProtection password="{legacy_hash(password)}" sheet="1" objects="1" '
                              f'scenarios="1"/>'.encode("ascii"), data)
            target.writestr(item, data)


def legacy_module(folder: Path) -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "On Error Resume Next"]
    for name, _, attempt in LEGACY:
        lines += [f'Set wb = Workbooks.Open("{folder / (name + ".xlsx")}")', "Set ws = wb.Worksheets(1)", "Err.Clear",
                  f'ws.Unprotect "{attempt}"', f'out = out & "{name}=" & Err.Number & "," & ws.ProtectContents & "|"',
                  "Err.Clear", "wb.Close False"]
    lines += ["Probe = out", "End Function"]
    return "\n".join(lines) + "\n"

#: What a reopened file's sheet says about its protection.
READ = ('ws.ProtectContents & "," & ws.ProtectDrawingObjects & "," & ws.ProtectScenarios & "," & '
        'ws.ProtectionMode & "," & ws.EnableSelection & "," & ws.Protection.AllowFormattingCells & "," & '
        'ws.Protection.AllowInsertingRows & "," & ws.Protection.AllowSorting & "," & ws.Protection.AllowFiltering')


def module(folder: Path) -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False"]
    for name, action in CONFIGS:
        path = folder / f"{name}.xlsx"
        lines += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", action,
                  f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
                  f'Set wb = Workbooks.Open("{path}")', "Set ws = wb.Worksheets(1)",
                  f'out = out & "{name}=" & {READ} & "|"', "wb.Close False"]
    lines += ["Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def element(path: Path) -> dict[str, str]:
    """The sheetProtection element and its neighbours in the first sheet's XML."""
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found = re.search(r"<sheetProtection\b[^>]*/>", xml)
    if found is None:
        return {"element": "", "before": "", "after": ""}
    before = re.findall(r"</?([A-Za-z]+)[^<]*>\s*$", xml[:found.start()])
    after = re.match(r"\s*<([A-Za-z]+)", xml[found.end():])
    return {"element": found.group(0), "before": before[-1] if before else "",
            "after": after.group(1) if after else ""}


def main() -> None:
    with tempfile.TemporaryDirectory() as temporary:
        folder = Path(temporary)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(module(folder), "Probe", timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            for name, password, _ in LEGACY:
                legacy_file(folder / "plain.xlsx", folder / f"{name}.xlsx", password)
            legacy = excel.run_vba(legacy_module(folder), "Probe", timeout=600.0)
            assert legacy.ok, f"{legacy.outcome}: {legacy.message} {legacy.error}"
        read = dict(part.split("=", 1) for part in str(result.value).split("|") if part)
        cases = [{"name": name, "action": action, **element(folder / f"{name}.xlsx"), "read": read[name]}
                 for name, action in CONFIGS]
        opened = dict(part.split("=", 1) for part in str(legacy.value).split("|") if part)
        legacies = [{"name": name, "password": password, "hash": legacy_hash(password), "attempt": attempt,
                     "answer": opened[name]} for name, password, attempt in LEGACY]
    OUT.write_text(json.dumps({"read": READ, "cases": cases, "legacy": legacies}, indent=1) + "\n", encoding="utf-8")
    for case in cases:
        print(f"{case['name']:24} {case['element']}\n{'':24} {case['before']} / {case['after']}  read {case['read']}")
    for one in legacies:
        print(f"{one['name']:24} {one['hash']} Unprotect {one['attempt']!r}: {one['answer']}")


if __name__ == "__main__":
    main()
