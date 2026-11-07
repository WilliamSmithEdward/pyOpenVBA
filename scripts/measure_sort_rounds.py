"""Held-out rounds for the worksheet Sort rows: Sort's Header, MatchCase, Orientation, SortMethod and Rng, a SortField's Key,
Order, DataOption, SubField and ModifyKey, SortFields' hidden _Default and _NewEnum, and Worksheet.Sort. Each case is a new
workbook with random cells, fields added by Add or Add2, settings set (some to values Excel refuses), a range set or not,
then every setting read with its type, and Apply, once or twice, with the cells after. Half the cases are typed, half late.

    python scripts/measure_sort_rounds.py SEED predict [count]   the model's answers and their digest, before Excel
    python scripts/measure_sort_rounds.py SEED live              Excel's answers, each run under live.lock
    python scripts/measure_sort_rounds.py SEED score

Rounds are kept in the scratch folder given by SORT_ROUNDS (default: the system temp folder) until they go into
tests/fixtures/sort_rounds.json, which tests/test_excel_sort_rounds.py replays.
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

WORDS = ["apple", "Apple", "APPLE", "banana", "Banana", "cherry", "date", "b", "B", "a", "A", "zeta", "10", "9", "x1", "x10", "x2",
         "-", "", "ab", "aB"]
ODD = {"Header": ["3", '"x"', "Null", "True", "-1", "xlYes", "xlNo", "xlGuess"], "MatchCase": ['"x"', "Null", "1", "0", "2.5", '"True"'],
       "Orientation": ["0", "3", '"x"', "Null", "xlTopToBottom", "xlLeftToRight", "True"],
       "SortMethod": ["0", "3", '"x"', "Null", "xlPinYin", "xlStroke"]}
PLAIN = {"Header": ["xlYes", "xlNo", "xlGuess"], "MatchCase": ["True", "False"], "Orientation": ["xlTopToBottom", "xlLeftToRight"],
         "SortMethod": ["xlPinYin", "xlStroke"]}
FIELD_SETS = {"Order": ["xlAscending", "xlDescending", "1", "2", "0", "3", '"x"', "Null", "True", "1.5", '"2"'],
              "DataOption": ["xlSortNormal", "xlSortTextAsNumbers", "0", "1", "2", "-1", '"x"', "Null", "True", "Empty"],
              "SubField": ['"y"', '""', "5", "Null", "True", "Empty"]}
LETTERS = "ABCDEF"
def _shared_git() -> Path:
    """The repository's git folder, which a worktree's .git file points into."""
    git = Path(__file__).resolve().parents[1] / ".git"
    if git.is_file():
        return Path(git.read_text(encoding="utf-8").split(":", 1)[1].strip()).parents[1]
    return git


LOCK = _shared_git() / "excel-coordination" / "live.lock"


def _value(r: random.Random) -> str:
    kind = r.random()
    if kind < 0.35:
        return str(r.choice([1, 2, 3, 10, -4, 0, 2.5, 1e3, 7, 11, 100, 0.25]))
    if kind < 0.7:
        return '"' + r.choice(WORDS) + '"'
    if kind < 0.78:
        return r.choice(["True", "False"])
    if kind < 0.84:
        return '"=NA()"'
    if kind < 0.9:
        return r.choice(['"\'5"', '"\'12"', "#1/2/2020#", "#3/4/2019#"])
    return "Empty"


def make_case(r: random.Random) -> dict[str, Any]:
    rows, columns = r.randint(3, 9), r.randint(1, 4)
    cells = [[_value(r) for _ in range(columns)] for _ in range(rows)]
    if r.random() < 0.5:
        cells[0] = ['"' + h + '"' for h in r.sample(["Name", "Qty", "Code", "Kind", "Date"], columns)]
    left, top = r.choice(["A", "B"]), r.choice([1, 2])
    first = LETTERS.index(left)
    last = LETTERS[first + columns - 1]
    area = f"{left}{top}:{last}{top + rows - 1}"
    fields = []
    for _ in range(r.choice([0, 1, 1, 2, 2, 3])):
        column = LETTERS[first + r.randrange(columns)]
        key = r.choice([f"{column}{top}", f"{column}{top}:{column}{top + rows - 1}", f"{column}{top + 1}:{column}{top + rows - 1}",
                        f"{column}:{column}", f"{left}{top + r.randrange(rows)}:{last}{top + r.randrange(rows)}"])
        fields.append({"key": key, "order": r.choice([None, "xlAscending", "xlDescending", "1", "2"]),
                       "option": r.choice([None, "xlSortNormal", "xlSortTextAsNumbers", "0", "1"]),
                       "on": r.choice([None, None, "xlSortOnValues", "0"]), "add2": r.random() < 0.3, "sub": r.random() < 0.1})
    settings = {}
    for name in ("Header", "MatchCase", "Orientation", "SortMethod"):
        if r.random() < 0.55:
            settings[name] = r.choice(ODD[name] if r.random() < 0.25 else PLAIN[name])
    chosen = r.choice([area, area, area, f"{left}{top}:{last}{top + rows - 2}", left + str(top), None])
    sets = []
    if fields and r.random() < 0.45:
        for _ in range(r.randint(1, 2)):
            name = r.choice(sorted(FIELD_SETS))
            sets.append((r.randrange(len(fields)) + 1, name, r.choice(FIELD_SETS[name])))
    return {"cells": cells, "left": left, "top": top, "fields": fields, "settings": settings, "range": chosen, "sets": sets,
            "modify": r.random() < 0.15, "apply": r.random() < 0.8, "typed": r.random() < 0.5, "again": r.random() < 0.2, "area": area}


def case_source(case: dict[str, Any]) -> str:
    """The macro of one case: a function Probe answering every step's error and every read, with its type."""
    typed = case["typed"]
    lines = ["Public Function Probe() As String",
             "Dim wb, ws, s, v, k, c, n" + (", so As Sort, fs As SortFields, f As SortField" if typed else ", so, fs, f"),
             "Application.DisplayAlerts = False: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)"]
    first = LETTERS.index(case["left"])
    for i, row in enumerate(case["cells"]):
        for j, item in enumerate(row):
            if item != "Empty":
                lines.append(f'ws.Range("{LETTERS[first + j]}{case["top"] + i}").Value = {item}')
    lines += ['s = ""', "On Error Resume Next", "Set so = ws.Sort", "Set fs = so.SortFields", "fs.Clear", 's = s & "clear" & Err.Number & ";": Err.Clear']
    for n, field in enumerate(case["fields"]):
        args = [f'Key:=ws.Range("{field["key"]}")']
        if field["on"]:
            args.append(f"SortOn:={field['on']}")
        if field["order"]:
            args.append(f"Order:={field['order']}")
        if field["option"]:
            args.append(f"DataOption:={field['option']}")
        if field["sub"] and field["add2"]:
            args.append('SubField:="x"')
        lines.append(f"Set f = fs.{'Add2' if field['add2'] else 'Add'}({', '.join(args)})")
        lines.append(f's = s & "add{n}=" & Err.Number & ";": Err.Clear')
    for name, set_to in case["settings"].items():
        lines.append(f"so.{name} = {set_to}")
        lines.append(f's = s & "{name}=" & Err.Number & ";": Err.Clear')
    if case["range"]:
        lines.append(f'so.SetRange ws.Range("{case["range"]}")')
        lines.append('s = s & "range=" & Err.Number & ";": Err.Clear')
    for at, name, set_to in case.get("sets", []):
        lines.append(f"Set f = fs({at}): f.{name} = {set_to}")
        lines.append(f's = s & "set{at}{name[:3]}=" & Err.Number & ";": Err.Clear')
    if case["modify"] and case["fields"]:
        lines.append(f'fs(1).ModifyKey ws.Range("{case["left"]}{case["top"]}")')
        lines.append('s = s & "modify=" & Err.Number & ";": Err.Clear')
    for name in ("Header", "MatchCase", "Orientation", "SortMethod"):
        lines += [f"v = Empty: v = so.{name}", f's = s & "{name[:3]}=" & Err.Number & ":" & TypeName(v) & ":" & v & ";": Err.Clear']
    lines += ['k = "": k = so.Rng.Address', 's = s & "Rng=" & Err.Number & ":" & k & ";": Err.Clear',
              "v = Empty: v = fs.Count", 's = s & "Cnt=" & Err.Number & ":" & TypeName(v) & ":" & v & ";": Err.Clear',
              "n = 0", "For Each f In fs", "n = n + 1",
              'k = "": k = f.Key.Address: s = s & "F" & n & "key=" & Err.Number & ":" & k: Err.Clear',
              'v = Empty: v = f.Order: s = s & ",ord=" & Err.Number & ":" & TypeName(v) & ":" & v: Err.Clear',
              'v = Empty: v = f.DataOption: s = s & ",opt=" & Err.Number & ":" & TypeName(v) & ":" & v: Err.Clear',
              'v = Empty: v = f.SortOn: s = s & ",on=" & Err.Number & ":" & v: Err.Clear',
              'v = Empty: v = f.Priority: s = s & ",pri=" & Err.Number & ":" & v: Err.Clear',
              'v = Empty: v = f.SubField: s = s & ",sub=" & Err.Number & ":" & TypeName(v) & ":" & v: Err.Clear',
              'k = "": k = TypeName(f.SortOnValue): s = s & ",sov=" & Err.Number & ":" & k & ";": Err.Clear', "Next",
              's = s & "each=" & Err.Number & ":" & n & ";": Err.Clear',
              'k = "": k = fs(1).Key.Address: s = s & "def=" & Err.Number & ":" & k & ";": Err.Clear']
    if case["apply"]:
        lines += ["so.Apply", 's = s & "apply=" & Err.Number & ";": Err.Clear']
        if case["again"]:
            lines += ["so.Apply", 's = s & "again=" & Err.Number & ";": Err.Clear']
        lines += [f'For Each c In ws.Range("{case["area"]}").Cells', 's = s & c.Formula & "|"', "Next",
                  'k = "": k = so.Rng.Address: s = s & ";after=" & Err.Number & ":" & k: Err.Clear',
                  'v = Empty: v = fs.Count: s = s & ";cnt=" & v']
    lines += ["wb.Close False", "Probe = s", "End Function", ""]
    return "\n".join(lines)


def model(case: dict[str, Any]) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case_source(case), name="M")
    try:
        return str(app.run("M.Probe"))
    except VBAUnsupportedError as exc:
        return f"REFUSED {exc}"


def _lock() -> None:
    while True:
        try:
            os.mkdir(LOCK)
            return
        except FileExistsError:
            time.sleep(5)


def main() -> None:
    seed, step = int(sys.argv[1]), sys.argv[2]
    folder = Path(os.environ.get("SORT_ROUNDS", tempfile.gettempdir()))
    file = folder / f"sort_round_{seed}.json"
    if step == "predict":
        r = random.Random(seed)
        cases = [make_case(r) for _ in range(int(sys.argv[3]) if len(sys.argv) > 3 else 220)]
        for case in cases:
            case["model"] = model(case)
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
        print(len(cases), "digest", hashlib.sha256(json.dumps([c["model"] for c in cases]).encode()).hexdigest()[:16])
        return
    cases = json.loads(file.read_text(encoding="utf-8"))
    if step == "live":
        from pyvbaharness import ExcelSession, HarnessConfig

        todo = [case for case in cases if "excel" not in case]
        while todo:
            _lock()
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
                    excel.new_document()
                    for case in list(todo)[:25]:
                        answer = excel.run_vba(case_source(case), "Probe", (), timeout=60.0, line_numbers=False)
                        case["excel"] = str(answer.value) if answer.ok else f"no answer {answer.outcome}"
                        todo.remove(case)
            finally:
                os.rmdir(LOCK)
            file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    wrong = [case for case in cases if case["model"] != case.get("excel")]
    print(seed, "differ", len(wrong), "of", len(cases), "refused", sum(case["model"].startswith("REFUSED") for case in cases))


if __name__ == "__main__":
    main()
