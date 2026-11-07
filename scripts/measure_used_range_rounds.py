"""Held-out rounds for Worksheet.UsedRange, with SpecialCells(xlLastCell) and Range.End beside it: one to four random
edits on a new workbook (values, formats, every border on cells, ranges, unions, whole rows, whole columns and the
sheet, column XFD, heights, widths, hiding, merges, notes, clears and deletes), read typed or late, then saved as .xlsx,
reopened and read again.

    python scripts/measure_used_range_rounds.py SEED predict [count]   the model's answers and their digest, before Excel
    python scripts/measure_used_range_rounds.py SEED live              Excel's answers, each run under live.lock
    python scripts/measure_used_range_rounds.py SEED score [limit]

Rounds are kept in the folder SORT_ROUNDS names (default: the system temp folder) until they go into a fixture.
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]


def _shared_git() -> Path:
    """The repository's git folder, which a worktree's .git file points into."""
    git = ROOT / ".git"
    if git.is_file():
        return Path(git.read_text(encoding="utf-8").split(":", 1)[1].strip()).parents[1]
    return git


LOCK = _shared_git() / "excel-coordination" / "live.lock"
CELLS = ["A1", "B2", "C3", "D4", "E1", "A5", "H9", "C1048576", "XFD3", "XFC2", "B1048575"]
RANGES = ["A1:C3", "B2:D4", "C3:D3", "A1:A3", "B2,D3", "B2:C2,E5", "C3:E3"]
ROWS = ["ws.Rows(3)", 'ws.Rows("2:4")', "ws.Rows(1)", "ws.Rows(1048576)"]
COLUMNS = ["ws.Columns(3)", 'ws.Columns("B:D")', "ws.Columns(1)", 'ws.Columns("XFD")', 'ws.Columns("B:XFD")', 'ws.Columns("A:XFC")']
BORDERS = ["LineStyle = xlContinuous", "LineStyle = xlNone", "Weight = xlThick", "Weight = xlMedium", "LineStyle = xlDouble",
           "Color = 255", "ColorIndex = xlColorIndexNone", "ColorIndex = xlColorIndexAutomatic", "ThemeColor = 5"]
FORMATS = ["Interior.Color = 255", "Font.Bold = True", "Font.Bold = False", 'NumberFormat = "0.00"', "Interior.ColorIndex = xlNone",
           "HorizontalAlignment = xlCenter", "Font.Size = 16", "WrapText = True"]


def _target(r: random.Random) -> str:
    kind = r.random()
    if kind < 0.35:
        return f'ws.Range("{r.choice(CELLS)}")'
    if kind < 0.65:
        return f'ws.Range("{r.choice(RANGES)}")'
    if kind < 0.8:
        return r.choice(ROWS)
    if kind < 0.95:
        return r.choice(COLUMNS)
    return "ws.Cells"


def _edit(r: random.Random) -> str:
    kind = r.random()
    if kind < 0.3:
        return f"{_target(r)}.Borders({r.choice([7, 8, 9, 10, 11, 12, 5])}).{r.choice(BORDERS)}"
    if kind < 0.4:
        return f"{_target(r)}.Borders.{r.choice(BORDERS)}"
    if kind < 0.6:
        return f"{_target(r)}.{r.choice(FORMATS)}"
    if kind < 0.72:
        return f'ws.Range("{r.choice(CELLS)}").Value = {r.choice(["1", chr(34) + "a" + chr(34), chr(34) + "=1" + chr(34), chr(34) * 2])}'
    if kind < 0.8:
        return f'ws.Range("{r.choice(CELLS + RANGES)}").{r.choice(["ClearContents", "ClearFormats", "Clear"])}'
    if kind < 0.85:
        return r.choice(["ws.Rows(3).RowHeight = 30", "ws.Rows(5).Hidden = True", "ws.Columns(4).ColumnWidth = 20",
                         "ws.Columns(4).Hidden = True", 'ws.Range("C3").BorderAround xlContinuous, xlThick'])
    if kind < 0.9:
        return r.choice(['ws.Range("C3:D4").Merge', 'ws.Range("C3").AddComment "n"', 'ws.Range("B2:B3").Delete',
                         "ws.Rows(2).Delete", "ws.Columns(2).Delete", 'ws.Range("C3").Style = "Good"'])
    return f'ws.Range("{r.choice(CELLS)}").Interior.Color = 65535'


def make_case(r: random.Random) -> dict[str, Any]:
    return {"edits": [_edit(r) for _ in range(r.randint(1, 4))], "typed": r.random() < 0.5,
            "from": r.choice(["A1", "C3", "B2", "H9"])}


READS = ('s = s & ws.UsedRange.Address & "/" & ws.Cells.SpecialCells(xlLastCell).Address & "/" & ws.Range("{0}").End(xlDown).Address '
         '& "/" & ws.Range("{0}").End(xlToRight).Address & "/" & ws.Range("{0}").End(xlUp).Address & "/" & ws.UsedRange.Address')


def case_source(case: dict[str, Any], path: str) -> str:
    kind = "Worksheet" if case["typed"] else "Object"
    reads = READS.format(case["from"])
    lines = ["Public Function Probe(ByVal path As String) As String", f"Dim wb, ws As {kind}, s",
             "Application.DisplayAlerts = False", "On Error Resume Next",
             "Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)", 's = ""']
    for edit in case["edits"]:
        lines += [edit, 's = s & Err.Number & ",": Err.Clear']
    lines += [reads, 's = s & ";" & Err.Number & ";": Err.Clear', "wb.SaveAs path, 51", "wb.Close False",
              "Set wb = Workbooks.Open(path): Set ws = wb.Worksheets(1)", reads, 's = s & ";" & Err.Number', "wb.Close False",
              "Probe = s", "End Function", ""]
    return "\n".join(lines)


def model(case: dict[str, Any], path: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case_source(case, path), name="M")
    try:
        return str(app.run("M.Probe", path))
    except VBAUnsupportedError as exc:
        return f"REFUSED {exc}"


def _standard_height(excel: Any) -> str:
    """A new sheet's StandardHeight, which reads 15 only at 96 DPI."""
    source = ("Public Function Height() As String\nDim wb: Set wb = Workbooks.Add(xlWBATWorksheet)\n"
              "Height = CStr(wb.Worksheets(1).StandardHeight)\nwb.Close False\nEnd Function\n")
    answer = excel.run_vba(source, "Height", (), timeout=60.0, line_numbers=False)
    return str(answer.value) if answer.ok else "no answer"


def _lock() -> None:
    while True:
        try:
            os.mkdir(LOCK)
            return
        except FileExistsError:
            time.sleep(5)


def main() -> None:
    seed, step = int(sys.argv[1]), sys.argv[2]
    file = Path(os.environ.get("SORT_ROUNDS", tempfile.gettempdir())) / f"used_range_round_{seed}.json"
    folder = Path(tempfile.mkdtemp(prefix="used_range_"))
    if step == "predict":
        r = random.Random(seed)
        cases = [make_case(r) for _ in range(int(sys.argv[3]) if len(sys.argv) > 3 else 200)]
        for at, case in enumerate(cases):
            case["model"] = model(case, str(folder / f"{at}.xlsx"))
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
        print(len(cases), "digest", hashlib.sha256(json.dumps([c["model"] for c in cases]).encode()).hexdigest()[:16],
              "refused", sum(c["model"].startswith("REFUSED") for c in cases))
        return
    cases = json.loads(file.read_text(encoding="utf-8"))
    if step == "live":
        from pyvbaharness import ExcelSession, HarnessConfig

        todo = [at for at, case in enumerate(cases) if "excel" not in case]
        while todo:
            _lock()
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
                    excel.new_document()
                    # Rows a border or a font draws taller depend on the display: measure only at 96 DPI, StandardHeight 15.
                    if _standard_height(excel) != "15":
                        raise SystemExit("StandardHeight is not 15: the display is not at 96 DPI")
                    for at in list(todo)[:40]:
                        answer = excel.run_vba(case_source(cases[at], str(folder / f"{at}.xlsx")), "Probe",
                                               (str(folder / f"{at}.xlsx"),), timeout=60.0, line_numbers=False)
                        cases[at]["excel"] = str(answer.value) if answer.ok else f"no answer {answer.outcome}"
                        todo.remove(at)
                    if _standard_height(excel) != "15":
                        raise SystemExit("StandardHeight left 15 during the run: its answers are not kept")
            except Exception as failure:  # noqa: BLE001 - a case that ends Excel is recorded, and the rest go on
                cases[todo[0]]["excel"] = f"no answer, Excel ended: {str(failure)[:80]}"
                todo.pop(0)
            finally:
                os.rmdir(LOCK)
            file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    if step == "rescore":
        for at, case in enumerate(cases):
            case["model"] = model(case, str(folder / f"{at}.xlsx"))
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    wrong = [case for case in cases if case["model"] != case.get("excel")]
    print(seed, "differ", len(wrong), "of", len(cases), "refused", sum(case["model"].startswith("REFUSED") for case in cases))
    for case in wrong[:int(sys.argv[3]) if len(sys.argv) > 3 else 0]:
        print(" ", case["edits"], "T" if case["typed"] else "L", case["from"])
        print("     E", case.get("excel", "")[:300], "\n     M", case["model"][:300])


if __name__ == "__main__":
    main()
