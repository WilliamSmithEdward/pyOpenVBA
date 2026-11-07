"""Held-out rounds for WorksheetFunction.XLookup and XMatch: random keys and results on a sheet, looked up by numbers, text,
wildcards, blanks, logicals and arrays, in ranges and VBA arrays, down and across, with every match and search mode and
bad ones, if_not_found given, left empty or left out. Half the calls are typed, half late. Each answer is written with its
type, arrays item by item, a range by its address and values.

    python scripts/measure_xlookup_rounds.py SEED predict [count]   the model's answers and their digest, before Excel
    python scripts/measure_xlookup_rounds.py SEED live              Excel's answers, each run under live.lock
    python scripts/measure_xlookup_rounds.py SEED score [limit]

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

def _shared_git() -> Path:
    """The repository's git folder, which a worktree's .git file points into."""
    git = Path(__file__).resolve().parents[1] / ".git"
    if git.is_file():
        return Path(git.read_text(encoding="utf-8").split(":", 1)[1].strip()).parents[1]
    return git


LOCK = _shared_git() / "excel-coordination" / "live.lock"
WORDS = ["apple", "Apple", "banana", "cherry", "b", "B", "ab", "abc", "a*c", "a?c", "x~y", "10", "5", "", "zeta", "Zeta"]
PATTERNS = ['"a*"', '"?b"', '"*an*"', '"a~*c"', '"a?c"', '"~?"', '"b*"', '"*"', '"z*"', '"[ab]"', '"^a.c$"', '"an"']
HELPER = '''Private Function Show(v As Variant) As String
    Dim s As String, i As Long, j As Long
    If IsObject(v) Then
        Show = "R:" & v.Address(0, 0) & "=" & Show(v.Value)
    ElseIf IsArray(v) Then
        On Error Resume Next
        j = -1: j = UBound(v, 2)
        On Error GoTo 0
        If j = -1 Then
            For i = LBound(v) To UBound(v): s = s & Show(v(i)) & ",": Next
            Show = "A1[" & s & "]"
        Else
            For i = LBound(v, 1) To UBound(v, 1)
                For j = LBound(v, 2) To UBound(v, 2): s = s & Show(v(i, j)) & ",": Next
                s = s & ";"
            Next
            Show = "A2[" & s & "]"
        End If
    ElseIf IsError(v) Then
        Show = "E:" & CStr(v)
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''


def _item(r: random.Random) -> str:
    kind = r.random()
    if kind < 0.4:
        return str(r.choice([1, 2, 3, 5, 7, 10, -2, 0, 2.5, 100, 4, 6]))
    if kind < 0.75:
        return '"' + r.choice(WORDS) + '"'
    if kind < 0.82:
        return r.choice(["True", "False"])
    if kind < 0.88:
        return r.choice(['"=NA()"', '"=1/0"'])
    return "Empty"


def make_case(r: random.Random) -> dict[str, Any]:
    n = r.randint(2, 8)
    keys = [_item(r) for _ in range(n)]
    order = r.random()
    if order < 0.3:
        numbers = sorted(r.sample([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20], n))
        keys = [str(v) for v in (numbers if order < 0.2 else numbers[::-1])]
    results = [[_item(r) for _ in range(2)] for _ in range(n)]
    across = r.random() < 0.25
    look = r.random()
    if look < 0.3:
        value = str(r.choice([1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 0, 2.5, 25, -1]))
    elif look < 0.5:
        value = r.choice(keys) if keys[0] != "Empty" else '"apple"'
        value = '"apple"' if value == "Empty" else value
    elif look < 0.65:
        value = r.choice(PATTERNS)
    elif look < 0.72:
        value = r.choice(['"APPLE"', '"banana"', '"Banana "', '""'])
    elif look < 0.78:
        value = r.choice(["True", "False"])
    elif look < 0.84:
        value = r.choice(['ws.Range("H1")', 'ws.Range("H2")'])
    elif look < 0.92:
        value = r.choice(["Array(1, 3)", 'Array("apple", 5)', 'ws.Range("H3:H4")', "Array(2)"])
    else:
        value = r.choice(["CVErr(2042)", "Null", '"x"'])
    function = "XMatch" if r.random() < 0.35 else "XLookup"
    lookup = r.choice(["range", "range", "range", "array", "wide"])
    result = r.choice(["one", "one", "two", "short", "array"])
    nf = r.choice([None, None, None, '"nf"', "0", "", "Empty"])
    match = r.choice([None, None, "0", "-1", "1", "2", "3", "-2", '"1"', "Empty"])
    search = r.choice([None, None, "1", "-1", "2", "-2", "0", "3"])
    return {"keys": keys, "results": results, "across": across, "value": value, "function": function, "lookup": lookup,
            "result": result, "nf": nf, "match": match, "search": search, "typed": r.random() < 0.5,
            "extra": [r.choice(['"apple"', "3", "Empty", '"b"']), r.choice(['"apple"', "3", "Empty", '"b"']), r.choice(["3", '"apple"']), r.choice(["7", '"zeta"'])]}


def _arguments(case: dict[str, Any]) -> str:
    n = len(case["keys"])
    across = case["across"]

    def line(column: int, length: int) -> str:
        if across:
            return f'ws.Range(ws.Cells({column}, 1), ws.Cells({column}, {length}))'
        return f'ws.Range(ws.Cells(1, {column}), ws.Cells({length}, {column}))'

    lookup = {"range": line(1, n), "array": "Array(" + ", ".join("Empty" if k == "Empty" else k.replace('"=NA()"', "CVErr(2042)").replace('"=1/0"', "CVErr(2007)") for k in case["keys"]) + ")",
              "wide": 'ws.Range("A1:B2")'}[case["lookup"]]
    args = [case["value"], lookup]
    if case["function"] == "XLookup":
        two = (f'ws.Range(ws.Cells(2, 1), ws.Cells(3, {n}))' if across else f'ws.Range(ws.Cells(1, 2), ws.Cells({n}, 3))')
        result = {"one": line(2, n), "two": two, "short": line(2, max(1, n - 1)), "array": "Array(" + ", ".join(str(i * 10) for i in range(1, n + 1)) + ")"}[case["result"]]
        args += [result]
        tail = [case["nf"], case["match"], case["search"]]
    else:
        tail = [case["match"], case["search"]]
    # A call cannot end on an argument left empty: Excel's VBA refuses f(a, ) to compile.
    while tail and tail[-1] in (None, ""):
        tail.pop()
    args += ["" if part is None else part for part in tail]
    return ", ".join(args)


def case_source(case: dict[str, Any]) -> str:
    """One case: a function Probe answering the call's error and answer."""
    kind = "WorksheetFunction" if case["typed"] else "Object"
    lines = ["Public Function Probe() As String", f"Dim wb, ws, v, s, w As {kind}",
             "Application.DisplayAlerts = False: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)",
             "Set w = Application.WorksheetFunction"]
    across = case["across"]
    for i, key in enumerate(case["keys"]):
        place = f"ws.Cells(1, {i + 1})" if across else f"ws.Cells({i + 1}, 1)"
        if key != "Empty":
            lines.append(f"{place}.Value = {key}")
        for j, item in enumerate(case["results"][i]):
            spot = f"ws.Cells({j + 2}, {i + 1})" if across else f"ws.Cells({i + 1}, {j + 2})"
            if item != "Empty":
                lines.append(f"{spot}.Value = {item}")
    for k, item in enumerate(case["extra"]):
        if item != "Empty":
            lines.append(f'ws.Range("H{k + 1}").Value = {item}')
    lines += ["On Error Resume Next", f"v = Empty: v = w.{case['function']}({_arguments(case)})",
              's = Err.Number & "|" & Show(v)', "Err.Clear", "wb.Close False", "Probe = s", "End Function", ""]
    return HELPER + "\n".join(lines)


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
    file = Path(os.environ.get("SORT_ROUNDS", tempfile.gettempdir())) / f"xlookup_round_{seed}.json"
    if step == "predict":
        r = random.Random(seed)
        cases = [make_case(r) for _ in range(int(sys.argv[3]) if len(sys.argv) > 3 else 300)]
        for case in cases:
            case["model"] = model(case)
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
        print(len(cases), "digest", hashlib.sha256(json.dumps([c["model"] for c in cases]).encode()).hexdigest()[:16],
              "refused", sum(c["model"].startswith("REFUSED") for c in cases))
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
                    for case in list(todo)[:40]:
                        answer = excel.run_vba(case_source(case), "Probe", (), timeout=60.0, line_numbers=False)
                        case["excel"] = str(answer.value) if answer.ok else f"no answer {answer.outcome}"
                        todo.remove(case)
            except Exception as failure:  # noqa: BLE001 - a case that ends Excel is recorded, and the rest go on
                todo[0]["excel"] = f"no answer, Excel ended: {str(failure)[:80]}"
                todo.pop(0)
            finally:
                os.rmdir(LOCK)
            file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    if step == "rescore":
        for case in cases:
            case["model"] = model(case)
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    wrong = [case for case in cases if case["model"] != case.get("excel")]
    print(seed, "differ", len(wrong), "of", len(cases), "refused", sum(case["model"].startswith("REFUSED") for case in cases))
    for case in wrong[:int(sys.argv[3]) if len(sys.argv) > 3 else 0]:
        print(" ", case["function"], case["value"], case["lookup"], case.get("result"), "nf", case["nf"], "m", case["match"], "s", case["search"],
              "across" if case["across"] else "", "T" if case["typed"] else "L", case["keys"])
        print("     E", case.get("excel", "")[:200], "\n     M", case["model"][:200])


if __name__ == "__main__":
    main()
