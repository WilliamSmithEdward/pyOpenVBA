"""Typing: what a value written to a cell becomes, replayed against the in-memory model.

tests/fixtures/value_typing.json is what scripts/measure_value_typing.py saw
in live Excel: strings and computed values written to a fresh cell, a
number read through a set of formats, strings typed into cells that
already have a format, repeated writes to one cell, and arrays written to
a block. tests/fixtures/typing_formats.json is what
scripts/measure_typing_formats.py saw: every typed value into every
built-in format and some custom ones, one typing after another, what
keeps and clears a prefix character, what a Text cell makes of a write,
how far a typed time and fraction may run, and how Excel rewrites a
number format it is given. Each probe is rebuilt from its record, run
once in the model, and compared read by read.

Every read is compared exactly, Range.Text included, except Text through
a format with a ``*`` fill: what that shows depends on the cell's width in
pixels, which the model reports it cannot tell, so the replay reads a
placeholder there instead.
"""

from __future__ import annotations

import html
import json
import re
import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication, _typing
from pyopenvba.formula._display import tokens

FIXTURES = Path(__file__).parent / "fixtures"
TYPING: dict[str, Any] = json.loads((FIXTURES / "value_typing.json").read_text(encoding="utf-8"))
FORMATS: dict[str, Any] = json.loads((FIXTURES / "typing_formats.json").read_text(encoding="utf-8"))
SAVED: dict[str, Any] = json.loads((FIXTURES / "typing" / "typing.json").read_text(encoding="utf-8"))
STEP = "^"


#: What the replay reads in place of Range.Text through a filled format; the comparison drops it on both sides.
_PLACEHOLDER = "String:(text)"


def _read_lines(expressions: list[str], *, text: bool = True) -> list[str]:
    lines: list[str] = []
    for expression in expressions:
        read = '"(text)"' if expression == "c.Text" and not text else expression
        lines += ["Err.Clear", "v = Empty", f"v = {read}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return lines


def _fills(code: str) -> bool:
    """Whether a code pads a cell with a * fill, which makes Range.Text depend on the cell's width in pixels."""
    return any(kind == "pair" and text[0] == "*" for kind, text in tokens(code))


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _function(name: str, body: list[str]) -> list[str]:
    return [f"Private Function {name}(ws As Object) As String", "Dim out As String, v As Variant, c As Object",
            'Set c = ws.Range("B2")', "On Error Resume Next", *body, "On Error GoTo 0", f"{name} = out",
            "End Function"]


def _step(statement: str) -> str:
    return statement if statement.startswith(("c.", "ws.", "Set ")) else f"c.Value = {statement}"


def _typing_module(record: dict[str, Any]) -> str:
    """The module scripts/measure_value_typing.py ran, rebuilt from its record."""
    head = [record["helper"], "Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
            "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Columns("B:E").ColumnWidth = 40']
    body: list[str] = []
    for index, case in enumerate(record["cases"]):
        head += ["ws.Cells.Clear", f'out = out & Case{index}(ws) & "|"']
        body += _function(f"Case{index}", ["Err.Clear", f"c.Value = {case['text']}",
                                           'If Err.Number <> 0 Then out = "S" & Err.Number & ";"',
                                           *_read_lines(record["reads"])])
    head.append('out = out & "~FORMATS~"')
    for index, case in enumerate(record["formats"]):
        head.append(f'out = out & Format{index}(ws) & "|"')
        lines: list[str] = []
        for value in record["format_values"]:
            lines += ["c.Clear", "Err.Clear", f"c.Value = {value}", f"c.NumberFormat = {_vba_text(case['code'])}",
                      'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"',
                      *_read_lines(record["format_reads"], text=not _fills(case["code"])), f'out = out & "{STEP}"']
        body += _function(f"Format{index}", lines)
    head.append('out = out & "~PRESETS~"')
    for index, case in enumerate(record["presets"]):
        head += ["ws.Cells.Clear", f'out = out & Preset{index}(ws) & "|"']
        body += _function(f"Preset{index}", ["Err.Clear", f"c.NumberFormat = {_vba_text(case['code'])}",
                                             f"c.{case['writer']} = {case['text']}",
                                             'If Err.Number <> 0 Then out = "S" & Err.Number & ";"',
                                             *_read_lines(record["preset_reads"])])
    head.append('out = out & "~SEQUENCES~"')
    for index, case in enumerate(record["sequences"]):
        head += ["ws.Cells.Clear", f'out = out & Sequence{index}(ws) & "|"']
        lines = []
        for step in case["steps"]:
            lines += ["Err.Clear", "c.ClearContents" if step == "ClearContents" else f"c.Value = {step}",
                      'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"',
                      *_read_lines(record["sequence_reads"]), f'out = out & "{STEP}"']
        body += _function(f"Sequence{index}", lines)
    head.append('out = out & "~ARRAYS~"')
    for index, case in enumerate(record["arrays"]):
        head += ["ws.Cells.Clear", f'out = out & Array{index}(ws) & "|"']
        lines = ["Err.Clear", case["write"], 'If Err.Number <> 0 Then out = "S" & Err.Number & ";"']
        for reference in ("B2", "C2", "D2", "E2"):
            lines += [f'Set c = ws.Range("{reference}")', *_read_lines(record["array_reads"]), f'out = out & "{STEP}"']
        body += _function(f"Array{index}", lines)
    head += ["wb.Close False", "Probe = out", "End Function"]
    return "\n".join(head + body) + "\n"


def _formats_module(record: dict[str, Any]) -> str:
    """The module scripts/measure_typing_formats.py ran, rebuilt from its record."""
    loop = f'''Private Function Matrix(ws As Object, code As String, items As Variant) As String
    Dim out As String, c As Object, i As Long
    Set c = ws.Range("B2")
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        c.Clear
        Err.Clear
        c.NumberFormat = code
        c.Value = items(i)
        If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"
        out = out & TypeName(c.Value) & ";" & Show(c.Value2) & ";" & c.NumberFormat & "{STEP}"
    Next
    On Error GoTo 0
    Matrix = out
End Function

Private Function Chain(ws As Object, first As Variant, items As Variant) As String
    Dim out As String, c As Object, i As Long
    Set c = ws.Range("B2")
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        c.Clear
        Err.Clear
        c.Value = first
        c.Value = items(i)
        If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"
        out = out & TypeName(c.Value) & ";" & Show(c.Value2) & ";" & c.NumberFormat & "{STEP}"
    Next
    On Error GoTo 0
    Chain = out
End Function
'''
    head = [record["helper"], loop, "Public Function Probe() As String",
            "Dim wb As Object, ws As Object, out As String, items As Variant, i As Long",
            "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            "ws.Columns(2).ColumnWidth = 40", f"items = Array({', '.join(record['typed'])})"]
    body: list[str] = []
    for row in record["matrix"]:
        head.append(f'out = out & Matrix(ws, {_vba_text(row["code"])}, items) & "|"')
    head += ['out = out & "~CHAINS~"', "For i = LBound(items) To UBound(items)",
             'out = out & Chain(ws, items(i), items) & "|"', "Next", 'out = out & "~SEQUENCES~"']
    for index, case in enumerate(record["sequences"]):
        head += ["ws.Cells.Clear", f'out = out & Sequence{index}(ws) & "|"']
        lines: list[str] = []
        for step in case["steps"]:
            lines += ["Err.Clear", _step(step), 'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"',
                      *_read_lines(record["reads"]), f'out = out & "{STEP}"']
        body += _function(f"Sequence{index}", lines)
    head += ['out = out & "~NORMALIZED~" & Normalized(ws)', "wb.Close False", "Probe = out", "End Function"]
    lines = ["Private Function Normalized(ws As Object) As String", "Dim out As String, c As Object",
             'Set c = ws.Range("B2")', "On Error Resume Next"]
    for entry in record["normalized"]:
        lines += ["c.Clear", "Err.Clear", f"c.NumberFormat = {_vba_text(entry['code'])}",
                  'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"', f'out = out & c.NumberFormat & "{STEP}"']
    body += [*lines, "On Error GoTo 0", "Normalized = out", "End Function"]
    return "\n".join(head + body) + "\n"


def _run(module: str, year: int) -> str:
    """A probe module run in a fresh model, with a date typed without a year falling in the probe's year."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(module, name="Probe")
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(_typing, "_this_year", lambda: year)
        return str(app.run("Probe"))


@pytest.fixture(scope="module")
def typing_answers() -> dict[str, list[str]]:
    answer = _run(_typing_module(TYPING), TYPING["year"])
    typed, rest = answer.split("~FORMATS~", 1)
    formatted, rest = rest.split("~PRESETS~", 1)
    preset, rest = rest.split("~SEQUENCES~", 1)
    sequenced, arrayed = rest.split("~ARRAYS~", 1)
    return {"cases": typed.split("|"), "formats": formatted.split("|"), "presets": preset.split("|"),
            "sequences": sequenced.split("|"), "arrays": arrayed.split("|")}


@pytest.fixture(scope="module")
def formats_answers() -> dict[str, list[str]]:
    answer = _run(_formats_module(FORMATS), FORMATS["year"])
    matrix, rest = answer.split("~CHAINS~", 1)
    chains, rest = rest.split("~SEQUENCES~", 1)
    sequenced, normalized = rest.split("~NORMALIZED~", 1)
    return {"matrix": matrix.split("|"), "chains": chains.split("|"), "sequences": sequenced.split("|"),
            "normalized": normalized.split(STEP)}


#: Where one read's answer starts: every read is Show()'s TypeName:text, Empty, Null or Error:n, or E<n> for a
#: read that failed; S<n> leads when the write itself failed. A format holding ";" continues the read before it.
_HEAD = re.compile(r"(?:Empty|Null|Error:-?\d+|[ES]\d+|[A-Z][A-Za-z]*:)")


def _reads(answer: str) -> list[str]:
    parts: list[str] = []
    for piece in answer.split(";")[:-1]:
        if parts and not _HEAD.match(piece):
            parts[-1] += ";" + piece
        else:
            parts.append(piece)
    return parts


def _split(answer: str, reads: list[str]) -> tuple[list[str], list[str]]:
    """A step's answer as the failure of its write, if any, and one answer per read."""
    parts = _reads(answer)
    failed = [parts.pop(0)] if parts and re.fullmatch(r"S\d+", parts[0]) and len(parts) > len(reads) else []
    return failed, parts


def _steps(answer: str) -> list[str]:
    return [part for part in answer.split(STEP) if part]


def _compare(excel: list[str], model: list[str], reads: list[str]) -> Iterator[tuple[int, list[str], list[str]]]:
    for index, (want, got) in enumerate(zip(excel, model, strict=True)):
        (failed, wanted), (missed, found) = _split(want, reads), _split(got, reads)
        if "c.Text" in reads and len(found) == len(reads) and found[reads.index("c.Text")] == _PLACEHOLDER:
            # The model read a placeholder where Text depends on the width: compare the rest.
            wanted = [part for position, part in enumerate(wanted) if position != reads.index("c.Text")]
            found = [part for position, part in enumerate(found) if position != reads.index("c.Text")]
        if failed + wanted != missed + found:
            yield index, failed + wanted, missed + found


@pytest.mark.parametrize("index", range(len(TYPING["cases"])), ids=[case["text"] for case in TYPING["cases"]])
def test_a_written_value_is_typed_as_excel_types_it(typing_answers: dict[str, list[str]], index: int) -> None:
    case = TYPING["cases"][index]
    assert not list(_compare([case["answers"]], [typing_answers["cases"][index]], TYPING["reads"]))


@pytest.mark.parametrize("index", range(len(TYPING["formats"])), ids=[case["code"] for case in TYPING["formats"]])
def test_a_format_decides_what_value_reads_a_number_as(typing_answers: dict[str, list[str]], index: int) -> None:
    case = TYPING["formats"][index]
    differences = list(_compare(case["answers"], _steps(typing_answers["formats"][index]), TYPING["format_reads"]))
    assert not differences


@pytest.mark.parametrize("index", range(len(TYPING["presets"])),
                         ids=[f"{case['writer']} {case['code']} {case['text']}" for case in TYPING["presets"]])
def test_typing_into_a_formatted_cell(typing_answers: dict[str, list[str]], index: int) -> None:
    case = TYPING["presets"][index]
    assert not list(_compare([case["answers"]], [typing_answers["presets"][index]], TYPING["preset_reads"]))


@pytest.mark.parametrize("index", range(len(TYPING["sequences"])),
                         ids=[" then ".join(case["steps"]) for case in TYPING["sequences"]])
def test_writes_one_after_another(typing_answers: dict[str, list[str]], index: int) -> None:
    case = TYPING["sequences"][index]
    model = _steps(typing_answers["sequences"][index])
    assert not list(_compare(case["answers"], model, TYPING["sequence_reads"]))


@pytest.mark.parametrize("index", range(len(TYPING["arrays"])), ids=[case["write"] for case in TYPING["arrays"]])
def test_an_array_is_typed_item_by_item(typing_answers: dict[str, list[str]], index: int) -> None:
    case = TYPING["arrays"][index]
    assert _steps(typing_answers["arrays"][index]) == case["answers"]


@pytest.mark.parametrize("index", range(len(FORMATS["matrix"])), ids=[row["code"] for row in FORMATS["matrix"]])
def test_typing_keeps_or_replaces_a_format(formats_answers: dict[str, list[str]], index: int) -> None:
    row = FORMATS["matrix"][index]
    got = _steps(formats_answers["matrix"][index])
    wrong = [(typed, want, found) for typed, want, found in zip(FORMATS["typed"], row["answers"], got, strict=True)
             if want != found]
    assert not wrong


@pytest.mark.parametrize("index", range(len(FORMATS["chains"])), ids=[row["first"] for row in FORMATS["chains"]])
def test_a_format_typing_gave_is_kept_or_replaced_the_same_way(formats_answers: dict[str, list[str]],
                                                              index: int) -> None:
    row = FORMATS["chains"][index]
    got = _steps(formats_answers["chains"][index])
    wrong = [(typed, want, found) for typed, want, found in zip(FORMATS["typed"], row["answers"], got, strict=True)
             if want != found]
    assert not wrong


@pytest.mark.parametrize("index", range(len(FORMATS["sequences"])),
                         ids=[" then ".join(case["steps"]) for case in FORMATS["sequences"]])
def test_prefixes_text_cells_and_limits(formats_answers: dict[str, list[str]], index: int) -> None:
    case = FORMATS["sequences"][index]
    model = _steps(formats_answers["sequences"][index])
    assert not list(_compare(case["answers"], model, FORMATS["reads"]))


#: Codes whose rewriting the model does not follow: Excel moves the thousands separator when a quoted
#: literal splits the digits, and one case is not enough to say where it goes.
_UNFOLLOWED = {'#,##0"."00'}


@pytest.mark.parametrize("index", range(len(FORMATS["normalized"])),
                         ids=[entry["code"] for entry in FORMATS["normalized"]])
def test_a_number_format_is_kept_as_excel_rewrites_it(formats_answers: dict[str, list[str]], index: int) -> None:
    entry = FORMATS["normalized"][index]
    if entry["code"] in _UNFOLLOWED:
        pytest.xfail("Excel re-lays the thousands separator of a code split by a quoted literal")
    assert formats_answers["normalized"][index] == entry["stored"]


# --- the file side ----------------------------------------------------------------------------
#
# tests/fixtures/typing/ is what scripts/measure_typing_file.py saw: cases typed into cells of their own and
# the workbook Excel saved. The model types the same and saves; each cell is compared by what its xf says --
# its number format's id and spelling, font, fill, border, flags and prefix -- and by the kind and value it
# holds. How a number is spelt in <v>, and whether text is shared or inline, are the writer's own questions.


def _xf_descriptions(styles: str) -> list[str]:
    def children(section: str, tag: str) -> list[str]:
        block = re.search(rf"<{section}\b[^>]*>(.*?)</{section}>", styles, re.DOTALL)
        return re.findall(rf"<{tag}\b[^>]*/>|<{tag}\b[^>]*>.*?</{tag}>", block.group(1), re.DOTALL) if block else []

    spelled = dict(re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles))
    fonts, fills, borders = children("fonts", "font"), children("fills", "fill"), children("borders", "border")
    out: list[str] = []
    for xf in children("cellXfs", "xf"):
        head = dict(re.findall(r'(\w+)="([^"]*)"', xf[: xf.index(">")]))
        number = head.pop("numFmtId", "0")
        parts = [f"numFmt={number}:{spelled.get(number, '')}", fonts[int(head.pop("fontId", "0"))],
                 fills[int(head.pop("fillId", "0"))], borders[int(head.pop("borderId", "0"))],
                 *(f"{key}={value}" for key, value in sorted(head.items()))]
        out.append(";".join([*parts, xf[xf.index(">") + 1: xf.rindex("<")] if not xf.endswith("/>") else ""]))
    return out


def _saved_cells(path: Path) -> dict[int, tuple[str, str, object]]:
    """Each row's cell in column A: its xf described, what kind of value it holds, and the value."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        shared = package.read("xl/sharedStrings.xml").decode("utf-8") if "xl/sharedStrings.xml" in \
            package.namelist() else ""
    strings = ["".join(html.unescape(piece) for piece in re.findall(r"<t[^>]*>(.*?)</t>", item, re.DOTALL))
               for item in re.findall(r"<si>(.*?)</si>", shared, re.DOTALL)]
    xfs = _xf_descriptions(styles)
    out: dict[int, tuple[str, str, object]] = {}
    for cell in re.findall(r"<c\b[^>]*?(?:/>|>.*?</c>)", sheet, re.DOTALL):
        head = dict(re.findall(r'(\w+)="([^"]*)"', cell[: cell.index(">")]))
        row = int(head["r"][1:])
        raw = re.search(r"<v>(.*?)</v>", cell, re.DOTALL)
        inline = re.search(r"<is>(.*?)</is>", cell, re.DOTALL)
        kind = head.get("t", "n")
        value: object = None
        if kind == "s" and raw is not None:
            kind, value = "text", strings[int(raw.group(1))]
        elif kind == "inlineStr" and inline is not None:
            kind, value = "text", "".join(html.unescape(piece) for piece in re.findall(r"<t[^>]*>(.*?)</t>",
                                                                                       inline.group(1), re.DOTALL))
        elif raw is not None:
            value = float(raw.group(1)) if kind == "n" else html.unescape(raw.group(1))
        out[row] = (xfs[int(head.get("s", "0"))], kind if value is not None else "none", value)
    return out


@pytest.fixture(scope="module")
def model_saved(tmp_path_factory: pytest.TempPathFactory) -> dict[int, tuple[str, str, object]]:
    lines = ["Public Sub Build()", "Dim c As Object"]
    for row, steps in enumerate(SAVED["cases"], start=1):
        lines += [f"Set c = Cells({row}, 1)", *(_step(step) for step in steps)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*lines, "End Sub"]) + "\n", name="Builder")
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(_typing, "_this_year", lambda: SAVED["year"])
        app.run("Build")
    path = tmp_path_factory.mktemp("typing") / "typing.xlsx"
    app.save(path)
    return _saved_cells(path)


@pytest.fixture(scope="module")
def excel_saved() -> dict[int, tuple[str, str, object]]:
    return _saved_cells(FIXTURES / "typing" / "typing.xlsx")


@pytest.mark.parametrize("row", range(1, len(SAVED["cases"]) + 1),
                         ids=[" then ".join(steps) for steps in SAVED["cases"]])
def test_a_typed_cell_is_saved_as_excel_saves_it(model_saved: dict[int, tuple[str, str, object]],
                                                 excel_saved: dict[int, tuple[str, str, object]], row: int) -> None:
    assert model_saved.get(row) == excel_saved.get(row)
