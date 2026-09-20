"""Ask Office what each msoShapeType actually is.

``Shapes.AddShape`` takes a number and writes a preset geometry, and
the two are not obviously related: the number is MsoAutoShapeType and
the preset is DrawingML's own name for the shape.  The mapping is not
guessable, and several of the names belong to more than one shape, so
this asks instead.

For each type it adds the shape, records the name the host gave it,
saves the file, and reads the ``a:prstGeom prst=`` back out of the
markup.  Word and PowerPoint are asked for the names too, because a
host names its own shapes: Excel's "Rounded Rectangle" is Word's
"Rectangle: Rounded Corners".

Run it on a Windows machine with Office installed:

    python scripts/measure_shape_types.py
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "shapes"
OUT = FOLDER / "shape_types.json"

#: The types pyOpenVBA claims to know, plus the ones next to them, so a
#: mistake by one is visible.
TYPES = list(range(1, 21)) + [92]


def build(types: list[int]) -> str:
    """A macro that adds one of each and says what it got."""
    lines = [
        "Public Function Report() As String",
        "    Dim out As String, sh As Object, t As Variant",
        f"    Dim types As Variant: types = Array({', '.join(str(one) for one in types)})",
        "    Dim i As Long, top As Double",
        "    top = 0",
        "    For i = LBound(types) To UBound(types)",
        "        On Error Resume Next",
        "        Set sh = Nothing",
        "        Set sh = TARGET.Shapes.AddShape(CLng(types(i)), 10, top, 40, 40)",
        "        If sh Is Nothing Then",
        '            out = out & CStr(types(i)) & vbTab & "!refused" & vbTab & "" & vbLf',
        "        Else",
        "            out = out & CStr(types(i)) & vbTab & sh.Name & vbTab & _",
        "                CStr(sh.AutoShapeType) & vbLf",
        '            sh.Name = "T" & CStr(types(i))',
        "        End If",
        "        On Error GoTo 0",
        "        top = top + 45",
        "    Next i",
        "    Report = out",
        "End Function",
    ]
    return "\n".join(lines)


def _rows(body: str) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for line in body.split("\n"):
        if not line.strip():
            continue
        pieces = line.rstrip("\r").split("\t")
        if len(pieces) < 3:
            continue
        out[pieces[0]] = {"name": pieces[1], "auto_shape_type": pieces[2]}
    return out


def _presets(path: Path, part_prefix: str, name_attribute: str) -> dict[str, str]:
    """The preset geometry of each shape, by the name it was given."""
    package = zipfile.ZipFile(path)
    out: dict[str, str] = {}
    for part in package.namelist():
        if not part.startswith(part_prefix) or not part.endswith(".xml"):
            continue
        text = package.read(part).decode("utf-8", errors="replace")
        for found in re.finditer(
            rf'{name_attribute}="(T\d+)".*?<a:prstGeom prst="([^"]+)"', text, re.DOTALL
        ):
            out.setdefault(found.group(1), found.group(2))
    return out


def measure_excel() -> dict[str, dict[str, str]]:
    from pyvbaharness import ExcelSession

    from pyopenvba.excel import ExcelFile

    FOLDER.mkdir(parents=True, exist_ok=True)
    seed = FOLDER / "_types_seed.xlsm"
    target = FOLDER / "_types.xlsm"
    ExcelFile.create_new(seed)
    code = build(TYPES).replace("TARGET", "ActiveWorkbook.Worksheets(1)")
    code += (
        "\nPublic Function BuildAndReport() As String\n"
        "    BuildAndReport = Report\n"
        f'    ActiveWorkbook.SaveAs "{target}", 52\n'
        "End Function\n"
    )
    with ExcelSession() as excel:
        excel.open_document(seed)
        result = excel.run_vba(code, "BuildAndReport", timeout=300.0)
        if not result.ok:
            raise SystemExit(f"excel: {result.outcome} {result.message}")
        rows = _rows(str(result.value or ""))
    presets = _presets(target, "xl/drawings/drawing", "name")
    for number, row in rows.items():
        row["preset"] = presets.get(f"T{number}", "")
    seed.unlink(missing_ok=True)
    target.unlink(missing_ok=True)
    return rows


#: Does the number in a new shape's name go back down when shapes are
#: deleted, and does a second sheet or slide start again?  Asked of each
#: host rather than assumed from another.
COUNTER = """
Public Function Counter() As String
    Dim sh As Object, first As String, second As String, third As String
    first = TARGET.Shapes.AddShape(1, 10, 10, 20, 20).Name
    second = TARGET.Shapes.AddShape(9, 40, 10, 20, 20).Name
    For Each sh In TARGET.Shapes
        sh.Delete
    Next sh
    third = TARGET.Shapes.AddShape(1, 10, 10, 20, 20).Name
    Counter = first & "|" & second & "|" & third & "|" & Elsewhere()
End Function
"""

#: The same again somewhere else in the same file, which says whether
#: the counter belongs to the sheet or to the workbook.
ELSEWHERE = {
    "excel": (
        "Private Function Elsewhere() As String\n"
        "    Dim s As Object\n"
        "    Set s = ActiveWorkbook.Worksheets.Add\n"
        "    Elsewhere = s.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        "End Function\n"
    ),
    "powerpoint": (
        "Private Function Elsewhere() As String\n"
        "    Dim sl As Object\n"
        "    Set sl = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, 12)\n"
        "    Elsewhere = sl.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        "End Function\n"
    ),
    "word": (
        "Private Function Elsewhere() As String\n"
        '    Elsewhere = "one story only"\n'
        "End Function\n"
    ),
}


def measure_names(host: str) -> dict[str, str]:
    """What one host calls each shape, which is not what another does."""
    import pyvbaharness

    opener = {"word": "WordSession", "powerpoint": "PowerPointSession"}[host]
    target = {
        "word": "ActiveDocument",
        "powerpoint": "ActivePresentation.Slides(1)",
    }[host]
    # A new presentation has no slide to draw on.
    preamble = (
        "Public Function Prepare() As String\n"
        "    If ActivePresentation.Slides.Count = 0 Then ActivePresentation.Slides.Add 1, 12\n"
        '    Prepare = "ready"\n'
        "End Function\n"
        if host == "powerpoint"
        else ""
    )
    code = preamble + build(TYPES).replace("TARGET", target)
    with getattr(pyvbaharness, opener)() as office:
        office.new_document()
        if preamble:
            office.run_vba(code, "Prepare", timeout=120.0)
        result = office.run_vba(code, "Report", timeout=300.0)
        if not result.ok:
            raise SystemExit(f"{host}: {result.outcome} {result.message}")
        return {number: row["name"] for number, row in _rows(str(result.value or "")).items()}


def measure_counter(host: str) -> str:
    """Whether a host's shape numbering goes back down after a delete."""
    import pyvbaharness

    opener = {
        "excel": "ExcelSession",
        "word": "WordSession",
        "powerpoint": "PowerPointSession",
    }[host]
    target = {
        "excel": "ActiveWorkbook.Worksheets(1)",
        "word": "ActiveDocument",
        "powerpoint": "ActivePresentation.Slides(1)",
    }[host]
    prepare = (
        "Public Function Prepare() As String\n"
        "    If ActivePresentation.Slides.Count = 0 Then ActivePresentation.Slides.Add 1, 12\n"
        '    Prepare = "ready"\n'
        "End Function\n"
        if host == "powerpoint"
        else ""
    )
    code = prepare + COUNTER.replace("TARGET", target) + ELSEWHERE[host]
    with getattr(pyvbaharness, opener)() as office:
        office.new_document()
        if prepare:
            office.run_vba(code, "Prepare", timeout=120.0)
        result = office.run_vba(code, "Counter", timeout=180.0)
        if not result.ok:
            raise SystemExit(f"{host} counter: {result.outcome} {result.message}")
        return str(result.value or "")


def main() -> int:
    try:
        import pyvbaharness  # noqa: F401
    except ImportError:
        print("pyvbaharness is not installed; this script needs it and live Office")
        return 2

    excel = measure_excel()
    word = measure_names("word")
    powerpoint = measure_names("powerpoint")
    counters = {one: measure_counter(one) for one in ("excel", "word", "powerpoint")}
    body = {
        "excel": excel,
        "word_names": word,
        "powerpoint_names": powerpoint,
        "counter_after_delete": counters,
    }
    OUT.write_text(json.dumps(body, indent=2) + "\n", encoding="utf-8")
    for number in sorted(excel, key=int):
        row = excel[number]
        print(
            f"{number:>3}  {row.get('preset', ''):<18} excel={row['name']:<24}"
            f" word={word.get(number, ''):<28} ppt={powerpoint.get(number, '')}"
        )
    print()
    for host, answer in counters.items():
        print(f"{host:12} add, add, delete both, add -> {answer}")
    print(f"\nwrote {OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
