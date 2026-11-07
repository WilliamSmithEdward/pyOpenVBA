"""Held-out rounds for Range.Value2: values written by Value2 and read back by it, on new workbooks, typed and late. Numbers,
Currency, Dates, logicals, errors, Empty, text Excel types as numbers, dates or times, text it keeps, formulas, and formulas
naming another workbook: open, closed and missing, by file name and by path. Each case writes one value or an array to one
cell or several, reads every cell back with its type, its formula and its text, reads Value2 of the whole range, and then,
where a source book was open, closes it and reads again.

The source books are tests/fixtures/book_refs.json's: Alpha.xlsx (1, 2, 3 in A1:A3, text, True and #N/A in B1:B3, a name
Rate), Other.xlsx and Sub\\Deep.xlsx, written to a folder each run. Paths in answers are written <folder>.

    python scripts/measure_value2_rounds.py SEED predict [count]   the model's answers and their digest, before Excel
    python scripts/measure_value2_rounds.py SEED live              Excel's answers, each run under live.lock
    python scripts/measure_value2_rounds.py SEED score [limit]

Rounds are kept in the folder SORT_ROUNDS names (default: the system temp folder) until they go into a fixture.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import random
import re
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
HELPER = '''Private Function Show(v As Variant) As String
    Dim s As String, i As Long, j As Long
    If IsArray(v) Then
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
    ElseIf IsNull(v) Then
        Show = "Null"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Cells(r As Object) As String
    Dim c As Object, s As String
    On Error Resume Next
    For Each c In r.Cells
        s = s & c.Address(0, 0) & "=" & Show(c.Value2) & "~" & c.Formula & "~" & c.Text & "|"
    Next
    Cells = s
End Function

Private Sub Shut(ByVal keep As Object)
    Dim b As Object, again As Boolean
    On Error Resume Next
    Do
        again = False
        For Each b In Workbooks
            If b.Name <> ThisWorkbook.Name And Not b Is keep Then
                b.Close False
                again = True
                Exit For
            End If
        Next
    Loop While again
End Sub
'''
PLAIN = ["0", "1.5", "-3", "1E+20", "123456789012345", "0.1 + 0.2", "1 / 3", "CCur(1.25)", "CCur(-0.0001)", "#1/2/2020#", "#12:30:00 PM#",
         "#3/4/1900#", "True", "False", "Empty", "CVErr(2042)", "CVErr(2007)", "CLng(7)", "CSng(2.5)", "CDec(1.1)", "CByte(200)"]
TEXT = ['"apple"', '""', '"\'quoted"', '"1/2"', '"12:30"', '"1e3"', '" 5"', '"$5"', '"50%"', '"1,234"', '"(3)"', '"TRUE"', '"#N/A"',
        '"3 1/2"', '"2020-01-02"', '"Jan 5"', '"5:"', '"-"', '"0012"', '"1.5E+400"']
LOCAL = ['"=1+1"', '"=A1"', '"=SUM(1,2)"', '"=NA()"', '"=1/0"', '"=\'Sheet1\'!B2"', '"=Nope!A1"', '"=LEN(""ab"")"', '"=TODAY()-TODAY()"',
         '"=Undefined+1"', '"=(1"', '"=1+"']
BOOKS = ['"=[Alpha.xlsx]Sheet1!A1"', '"=SUM([Alpha.xlsx]Sheet1!A1:A3)"', '"=[Alpha.xlsx]Sheet1!B1"', '"=[Alpha.xlsx]Sheet1!B3"',
         '"=[Alpha.xlsx]Sheet1!C9"', '"=Alpha.xlsx!Rate"', '"=[Alpha.xlsx]Nope!A1"', '"=[Alpha.xlsx]Sheet1!A2*2"',
         '"=\'" & folder & "\\[Alpha.xlsx]Sheet1\'!A1"', '"=\'" & folder & "\\[Alpha.xlsx]Sheet1\'!A3+1"', '"=\'" & folder & "\\Alpha.xlsx\'!Rate"',
         '"=\'" & folder & "\\[Nope.xlsx]Sheet1\'!A1"', '"=\'" & folder & "\\Sub\\[Deep.xlsx]Sheet1\'!A1"', '"=[Other.xlsx]Elsewhere!A1"']
TARGETS = ["A1", "B2", "A1:C1", "A1:A3", "B2:C3", "D4"]
#: Zeros of the 17 digit sets Excel types as ASCII digits (Arabic-Indic, Thai, Devanagari and the rest), and text in
#: ASCII digits to spell in them. Rounds from 14004 on write them, built with ChrW, as a module's source is ANSI.
ZEROS = [0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6, 0x0D66, 0x0E50, 0x0ED0,
         0x0F20, 0x1040, 0x17E0, 0x1810]
NATIVE = ["7", "12", "-12", "0", "1.5", "12.25", "(3)", "$5", "$12.50", "50%", "1.5%", "1,234", "1,234.5", "1,2", "12:30",
          "12:30:45 PM", "1/2", "1/2/2020", "2020-01-02", "Jan 5", "1e3", "3 1/2", "0012", " 5", "5:", "12a", "1-2-3"]


def _spelled(text: str, zero: int) -> str:
    """The VBA expression for text with its ASCII digits written in the set whose zero is given."""
    parts: list[str] = []
    for c in text:
        piece = f"ChrW({zero + int(c)})" if c.isdigit() else '"' + c.replace('"', '""') + '"'
        if parts and piece.startswith('"') and parts[-1].startswith('"'):
            parts[-1] = parts[-1][:-1] + piece[1:]
        else:
            parts.append(piece)
    return " & ".join(parts)


def _value(r: random.Random, native: bool = False) -> str:
    if native and r.random() < 0.3:
        return "(" + _spelled(r.choice(NATIVE), r.choice(ZEROS)) + ")"
    kind = r.random()
    if kind < 0.3:
        return r.choice(PLAIN)
    if kind < 0.6:
        return r.choice(TEXT)
    if kind < 0.75:
        return r.choice(LOCAL)
    return r.choice(BOOKS)


def make_case(r: random.Random, native: bool = False) -> dict[str, Any]:
    target = r.choice(TARGETS)
    if r.random() < 0.3:
        value = "Array(" + ", ".join(_value(r, native) for _ in range(r.randint(1, 4))) + ")"
    else:
        value = _value(r, native)
    opened = [name for name in ("Alpha.xlsx", "Other.xlsx") if r.random() < 0.3]
    case = {"target": target, "value": value, "opened": opened, "typed": r.random() < 0.5, "before": r.choice(["", "", '"x"', "5"])}
    if "ChrW" in value:
        case["native"] = True
    return case


def case_source(case: dict[str, Any]) -> str:
    """One case: a function Probe(folder) answering each step."""
    kind = "Range" if case["typed"] else "Object"
    lines = ["Public Function Probe(ByVal folder As String) As String", f"Dim h, ws, s, v, r As {kind}",
             "Application.DisplayAlerts = False: Application.AskToUpdateLinks = False"]
    lines += [f'Workbooks.Open folder & "\\{name}", 0' for name in case["opened"]]
    lines += ["Set h = Workbooks.Add(xlWBATWorksheet): Set ws = h.Worksheets(1)", f'Set r = ws.Range("{case["target"]}")']
    if case["before"]:
        lines.append(f"r.Value2 = {case['before']}")
    lines += ["On Error Resume Next", f"r.Value2 = {case['value']}", 's = "set=" & Err.Number & ";": Err.Clear',
              "v = Empty: v = r.Value2", 's = s & "read=" & Err.Number & ":" & Show(v) & ";" & Cells(r): Err.Clear']
    if case["opened"]:
        lines += ["Shut h", 'v = Empty: v = r.Value2', 's = s & ";closed=" & Err.Number & ":" & Show(v) & ";" & Cells(r): Err.Clear']
    lines += ["Shut Nothing", "h.Close False", "Probe = s", "End Function", ""]
    source = HELPER + "\n".join(lines)
    # A cell typed from other digits reads its NumberFormat in place of its Text, which the model refuses for them.
    return source.replace("& c.Text &", "& c.NumberFormat &") if case.get("native") else source


def books(folder: Path) -> None:
    """Alpha.xlsx, Other.xlsx and Sub\\Deep.xlsx, as tests/fixtures/book_refs.json holds them."""
    record = json.loads((ROOT / "tests/fixtures/book_refs.json").read_text(encoding="utf-8"))
    for name, data in record["files"].items():
        (folder / name).parent.mkdir(parents=True, exist_ok=True)
        (folder / name).write_bytes(base64.b64decode(data))


def clean(text: str, folder: Path) -> str:
    """The folder written <folder>, whatever its case. Off Windows, where Excel never ran, a path in it has a slash after
    it, where Excel's has a backslash, as a link path follows its folder's spelling (as tests/test_excel_book_refs.py
    reads it). On Windows nothing else is folded, so a slash the model wrote there would still differ."""
    found = re.sub(re.escape(str(folder)), "<folder>", text, flags=re.IGNORECASE)
    return found.replace("<folder>/", "<folder>\\") if os.sep == "/" else found


def model(case: dict[str, Any], folder: Path) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case_source(case), name="M")
    try:
        return clean(str(app.run("M.Probe", str(folder))), folder)
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
    file = Path(os.environ.get("SORT_ROUNDS", tempfile.gettempdir())) / f"value2_round_{seed}.json"
    folder = Path(tempfile.mkdtemp(prefix="value2_books_"))
    books(folder)
    if step == "predict":
        r = random.Random(seed)
        cases = [make_case(r, seed >= 14004) for _ in range(int(sys.argv[3]) if len(sys.argv) > 3 else 300)]
        for case in cases:
            case["model"] = model(case, folder)
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
                        answer = excel.run_vba(case_source(case), "Probe", (str(folder),), timeout=60.0, line_numbers=False)
                        case["excel"] = clean(str(answer.value), folder) if answer.ok else f"no answer {answer.outcome}"
                        todo.remove(case)
            except Exception as failure:  # noqa: BLE001 - a case that ends Excel is recorded, and the rest go on
                todo[0]["excel"] = f"no answer, Excel ended: {str(failure)[:80]}"
                todo.pop(0)
            finally:
                os.rmdir(LOCK)
            file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    if step == "rescore":
        for case in cases:
            case["model"] = model(case, folder)
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    wrong = [case for case in cases if case["model"] != case.get("excel")]
    print(seed, "differ", len(wrong), "of", len(cases), "refused", sum(case["model"].startswith("REFUSED") for case in cases))
    for case in wrong[:int(sys.argv[3]) if len(sys.argv) > 3 else 0]:
        print(" ", case["target"], case["value"], "opened", case["opened"], "before", case["before"], "T" if case["typed"] else "L")
        print("     E", case.get("excel", "")[:300], "\n     M", case["model"][:300])


if __name__ == "__main__":
    main()
