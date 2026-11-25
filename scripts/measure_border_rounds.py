"""Held-out rounds for Border and Borders: random stored lines in a book Excel opens, or random macro formats on a new
one, then a border read or set on a cell, a block, a union, whole rows, whole columns or the sheet, typed or late. Each
answer gives the border's members, UsedRange, and the four edges of eighteen cells.

    python scripts/measure_border_rounds.py SEED predict [count]   the model's answers and their digest, before Excel
    python scripts/measure_border_rounds.py SEED live              Excel's answers, each run under live.lock
    python scripts/measure_border_rounds.py SEED score [limit]

A round counts for the borders without UsedRange, which a row drawn taller can move with the display; the score gives
both, and the live step records StandardHeight. Rounds and their books are kept in the folder SORT_ROUNDS names (default:
the system temp folder) until they go into a fixture.
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import re
import sys
import tempfile
import time
import zipfile
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
#: The lines a book's cells store, by index; a round adds random ones up to 41.
BASE_LINES = [("thin", "FFFF0000"), ("thin", "FF00FF00"), ("thick", "FF0000FF"), ("dashed", "FF800080"),
              ("medium", "FF000000"), ("thin", "FF000500"), ("thin", "FF0A0000"), ("thin", "FF00020A"),
              ("double", "FFFFFF00"), ("hair", "FF00FFFF"), ("thin", "FF050000"), ("thin", "FF000200"),
              ("thin", "FF00000A"), ("thick", "FF0A0A00"), ("thick", "FF000046"), ("thin", "FF804000"),
              ("thin", "FF004080"), ("dashed", "FF001000"), ("dashed", "FF280000")]
STYLES = ["thin", "thick", "dashed", "medium", "double", "hair", "dotted", "mediumDashed", "dashDot", "slantDashDot"]
READS = ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3", "XFD1", "XFD2", "A1048576", "C1048576",
         "XFD1048576", "F6"]
SETUP_TARGETS = ['ws.Range("B2")', 'ws.Range("A1:C3")', 'ws.Rows(2)', 'ws.Columns(2)', 'ws.Rows("1:2")',
                 'ws.Columns("B:C")', 'ws.Range("C1048576")', 'ws.Range("XFD1")', 'ws.Rows(1)', 'ws.Rows(1048576)',
                 'ws.Columns("XFD")', 'ws.Range("A1")']
TARGETS = ['ws.Range("B2")', 'ws.Range("A1")', 'ws.Range("A1:B2")', 'ws.Range("B1:D3")', 'ws.Range("A1:A3")',
           'ws.Range("A1:C1")', 'ws.Rows(1)', 'ws.Rows(2)', 'ws.Rows("1:3")', 'ws.Rows(1048576)', 'ws.Columns(1)',
           'ws.Columns(2)', 'ws.Columns("B:D")', 'ws.Columns("XFD")', "ws.Cells", 'ws.Range("B2,D3")',
           'ws.Range("XFD1")', 'ws.Range("C1048576")', 'ws.Columns("B:XFD")', 'ws.Columns("A:XFC")']
CHANGES = ["Read", "Read", "Color = 255", "Color = 65280", "Color = 0", "ColorIndex = 3",
           "ColorIndex = xlColorIndexAutomatic", "ColorIndex = xlColorIndexNone", "Weight = xlThick",
           "Weight = xlHairline", "Weight = xlMedium", "LineStyle = xlDash", "LineStyle = xlLineStyleNone",
           "LineStyle = xlDouble", "LineStyle = xlContinuous", "ThemeColor = 5", "ThemeColor = 1"]
_NO_VALUE_TARGETS = ("ws.Rows(1048576).Value", 'ws.Columns("XFD").Value', ".Rows(2).Value", ".Rows(1).Value",
                     "Columns(2).Value", 'Rows("1:2").Value', 'Columns("B:C").Value')


def lines_for(seed: int) -> list[tuple[str, str]]:
    """The round's stored lines: the base ones and random ones up to 41."""
    r = random.Random(seed * 7919 + 1)
    return BASE_LINES + [(r.choice(STYLES), "FF%02X%02X%02X" % (r.randint(0, 255), r.randint(0, 255), r.randint(0, 255)))
                         for _ in range(len(BASE_LINES), 41)]


def make_case(r: random.Random, index: int) -> dict[str, Any]:
    case: dict[str, Any] = {"kind": "book" if r.random() < 0.55 else "macro", "index": index}
    if case["kind"] == "book":
        cells: dict[str, list[int | None]] = {}
        for _ in range(r.randint(1, 6)):
            address = r.choice("ABCD") + str(r.randint(1, 3))
            if r.random() < 0.08:
                address = r.choice(["XFD1", "XFD2", "A1048576", "C1048576"])
            cells[address] = [None if r.random() < 0.6 else r.randint(0, 40) for _ in range(4)]
        case["cells"] = cells
    else:
        setup = []
        for _ in range(r.randint(0, 4)):
            target = r.choice(SETUP_TARGETS)
            setup.append(r.choice([
                f"{target}.Value = 1", f"{target}.Font.Bold = True",
                f"{target}.Borders({r.choice([7, 8, 9, 10, 11, 12])}).Color = {r.choice([255, 65280, 16711680, 5, 512])}",
                f"{target}.Borders({r.choice([7, 8, 9, 10])}).Weight = {r.choice(['xlThick', 'xlThin', 'xlMedium', 'xlHairline'])}",
                f"{target}.Interior.Color = 65535", f"{target}.Borders.LineStyle = {r.choice(['xlDash', 'xlContinuous'])}"]))
        case["setup"] = [step for step in setup if not any(no in step for no in _NO_VALUE_TARGETS)]
    case.update(target=r.choice(TARGETS), border=r.choice([7, 8, 9, 10, 11, 12, 7, 8, 9, 10, 11, 12, 5, 6]),
                change=r.choice(CHANGES), typed=r.random() < 0.4)
    return case


def _side(name: str, line: int | None, lines: list[tuple[str, str]]) -> str:
    return f"<{name}/>" if line is None else f'<{name} style="{lines[line][0]}"><color rgb="{lines[line][1]}"/></{name}>'


def write_book(path: Path, cells: dict[str, list[int | None]], lines: list[tuple[str, str]]) -> None:
    """A one-sheet book whose cells store the lines given: {address: [left, right, top, bottom]} as line indexes."""
    borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>"]
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    rows: dict[int, list[tuple[str, int]]] = {}
    for address, sides in cells.items():
        borders.append("<border>" + "".join(_side(name, line, lines) for name, line in
                                            zip(("left", "right", "top", "bottom"), sides)) + "<diagonal/></border>")
        xfs.append(f'<xf numFmtId="0" fontId="0" fillId="0" borderId="{len(borders) - 1}" xfId="0" applyBorder="1"/>')
        found = re.fullmatch(r"([A-Z]+)(\d+)", address)
        assert found is not None
        rows.setdefault(int(found.group(2)), []).append((address, len(xfs) - 1))
    sheet = "".join(f'<row r="{row}">' + "".join(f'<c r="{a}" s="{s}"/>' for a, s in
                                                 sorted(rows[row], key=lambda one: (len(one[0]), one[0]))) + "</row>"
                    for row in sorted(rows))
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    parts = {
        "[Content_Types].xml": head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": head + f'<workbook xmlns="{main}" xmlns:r="{rels}"><sheets>'
        '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rels}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{rels}/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": head + f'<styleSheet xmlns="{main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font>'
        '</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/>'
        f'</fill></fills><borders count="{len(borders)}">' + "".join(borders) + '</borders><cellStyleXfs count="1"><xf '
        f'numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="{len(xfs)}">' + "".join(xfs)
        + "</cellXfs></styleSheet>",
        "xl/worksheets/sheet1.xml": head + f'<worksheet xmlns="{main}"><sheetData>' + sheet + "</sheetData></worksheet>",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            z.writestr(name, text)


def case_source(case: dict[str, Any], path: str) -> str:
    kind = "Border" if case["typed"] else "Object"
    lines = ["Public Function Probe() As String", f"Dim wb, ws, s, c, b As {kind}: Application.DisplayAlerts = False"]
    lines.append(f'Set wb = Workbooks.Open("{path}"): Set ws = wb.Worksheets(1)' if path
                 else "Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)")
    lines += case.get("setup", [])
    lines += ['s = ""', "On Error Resume Next", f'Set b = {case["target"]}.Borders({case["border"]})',
              's = Err.Number & ";": Err.Clear']
    if case["change"] != "Read":
        lines += [f'b.{case["change"]}', 's = s & Err.Number & ";": Err.Clear']
    lines.append('s = s & TypeName(b.LineStyle) & ":" & b.LineStyle & "," & TypeName(b.Weight) & ":" & b.Weight & "," & '
                 'TypeName(b.Color) & ":" & b.Color & "," & b.ColorIndex & "," & b.ThemeColor & ";": Err.Clear')
    lines.append('s = s & ws.UsedRange.Address & ";"')
    for cell in READS:
        lines.append(f'Set c = ws.Range("{cell}"): s = s & c.Borders(7).LineStyle & "," & c.Borders(7).Color & "," & '
                     'c.Borders(8).LineStyle & "," & c.Borders(8).Color & "," & c.Borders(9).LineStyle & "," & '
                     'c.Borders(9).Color & "," & c.Borders(10).LineStyle & "," & c.Borders(10).Color & "," & '
                     'c.Borders(8).Weight & "," & c.Borders(9).Weight & ";"')
    lines += ["wb.Close False", "Probe = s", "End Function", ""]
    return "\n".join(lines)


def model(case: dict[str, Any], path: str) -> str:
    sys.path.insert(0, str(ROOT / "src"))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(case_source(case, path), name="M")
    try:
        return str(app.run("M.Probe"))
    except VBAUnsupportedError as refused:
        return f"REFUSED {str(refused)[:110]}"


def without_used_range(answer: str) -> list[str]:
    return [part for part in answer.split(";") if not part.startswith("$")]


def _standard_height(excel: Any) -> str:
    source = "Public Function Height() As String\nHeight = CStr(ActiveSheet.StandardHeight)\nEnd Function\n"
    answer = excel.run_vba(source, "Height", (), timeout=60.0, line_numbers=False)
    return str(answer.value) if answer.ok else "no answer"


def _lock() -> None:
    """Take the live lock, one Excel at a time, and say who holds it in its owner file."""
    while True:
        try:
            os.mkdir(LOCK)
            break
        except FileExistsError:
            time.sleep(5)
    (LOCK / "owner").write_text(f"{time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())} measure_border_rounds "
                                f"{os.getpid()}\n", encoding="utf-8")


def _unlock() -> None:
    """Give back the lock this run took."""
    (LOCK / "owner").unlink(missing_ok=True)
    os.rmdir(LOCK)


def main() -> None:
    seed, step = int(sys.argv[1]), sys.argv[2]
    folder = Path(os.environ.get("SORT_ROUNDS", tempfile.gettempdir()))
    file = folder / f"border_round_{seed}.json"
    books = folder / f"border_round_{seed}_books"
    if step == "predict":
        r = random.Random(seed)
        books.mkdir(exist_ok=True)
        lines = lines_for(seed)
        cases = [make_case(r, at) for at in range(int(sys.argv[3]) if len(sys.argv) > 3 else 150)]
        for case in cases:
            path = str((books / f"{case['index']}.xlsx").resolve()) if case["kind"] == "book" else ""
            if path:
                write_book(Path(path), case["cells"], lines)
            case["model"] = model(case, path)
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
        print(len(cases), "digest", hashlib.sha256(json.dumps([c["model"] for c in cases]).encode()).hexdigest()[:16],
              "refused", sum(c["model"].startswith("REFUSED") for c in cases))
        return
    cases = json.loads(file.read_text(encoding="utf-8"))
    path_of = {case["index"]: str((books / f"{case['index']}.xlsx").resolve()) if case["kind"] == "book" else ""
               for case in cases}
    if step == "live":
        from pyvbaharness import ExcelSession, HarnessConfig

        todo = [at for at, case in enumerate(cases) if "excel" not in case]
        while todo:
            _lock()
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
                    excel.new_document()
                    height = _standard_height(excel)
                    for at in list(todo)[:40]:
                        case = cases[at]
                        answer = excel.run_vba(case_source(case, path_of[case["index"]]), "Probe", (), timeout=90.0,
                                               line_numbers=False)
                        case["excel"] = str(answer.value) if answer.ok else f"no answer {answer.outcome}"
                        case["standard_height"] = height
                        todo.remove(at)
            except Exception as failure:  # noqa: BLE001 - a case that ends Excel is recorded, and the rest go on
                cases[todo[0]]["excel"] = f"no answer, Excel ended: {str(failure)[:80]}"
                todo.pop(0)
            finally:
                _unlock()
            file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    if step == "rescore":
        for case in cases:
            case["model"] = model(case, path_of[case["index"]])
        file.write_text(json.dumps(cases, indent=0), encoding="utf-8")
    borders = [case for case in cases if without_used_range(case["model"]) != without_used_range(case.get("excel", ""))]
    whole = [case for case in cases if case["model"] != case.get("excel")]
    print(seed, "borders differ", len(borders), "with UsedRange", len(whole), "of", len(cases), "refused",
          sum(case["model"].startswith("REFUSED") for case in cases),
          "StandardHeight", sorted({case.get("standard_height", "?") for case in cases}))
    for case in whole[:int(sys.argv[3]) if len(sys.argv) > 3 else 0]:
        print(" ", case.get("cells") or case.get("setup"), case["target"], case["border"], case["change"],
              "T" if case["typed"] else "L")
        print("     E", case.get("excel", "")[:300], "\n     M", case["model"][:300])


if __name__ == "__main__":
    main()
