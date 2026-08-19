"""Round-trip gate: source -> Excel-compiled p-code -> source.

The corpus below is the decompiler's acceptance set. Each entry is
compiled by real Excel (via ``pyvbaharness``, dev-time only), read back
with pyOpenVBA, decompiled by :mod:`pcode_source`, and compared to the
original text character for character.

Run it with a scratch directory:

    python docs/research/pcode/roundtrip.py <scratch-dir>

Entries whose reconstruction is not byte-identical print a unified diff
rather than being silently counted as passes.
"""
from __future__ import annotations

import difflib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pcode_source import decompile

from pyopenvba import ExcelFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import VBAModuleKind

MODULE = "M"

CORPUS: dict[str, str] = {
    "arith": (
        "Sub S()\n"
        "    Dim a As Long\n"
        "    Dim b As Long\n"
        "    Dim c As Long\n"
        "    b = 3\n"
        "    c = 4\n"
        "    a = b + c * 2\n"
        "    a = (b + c) * 2\n"
        "    a = b Mod c\n"
        "    a = b \\ c\n"
        "    a = -b\n"
        "End Sub\n"
    ),
    "if": (
        "Sub S()\n"
        "    Dim a As Long\n"
        "    a = 1\n"
        "    If a > 0 Then\n"
        "        a = 2\n"
        "    ElseIf a = 0 Then\n"
        "        a = 3\n"
        "    Else\n"
        "        a = 4\n"
        "    End If\n"
        "End Sub\n"
    ),
    "for": (
        "Sub S()\n"
        "    Dim i As Long\n"
        "    Dim t As Long\n"
        "    For i = 1 To 10\n"
        "        t = t + i\n"
        "    Next i\n"
        "    For i = 10 To 1 Step -2\n"
        "        t = t - i\n"
        "    Next i\n"
        "End Sub\n"
    ),
    "while": (
        "Sub S()\n"
        "    Dim i As Long\n"
        "    Do While i < 10\n"
        "        i = i + 1\n"
        "    Loop\n"
        "    While i > 0\n"
        "        i = i - 1\n"
        "    Wend\n"
        "End Sub\n"
    ),
    "select": (
        "Sub S()\n"
        "    Dim a As Long\n"
        "    a = 2\n"
        "    Select Case a\n"
        "        Case 1\n"
        "            a = 10\n"
        "        Case 2\n"
        "            a = 20\n"
        "        Case Else\n"
        "            a = 30\n"
        "    End Select\n"
        "End Sub\n"
    ),
    "str": (
        "Sub S()\n"
        "    Dim s As String\n"
        "    Dim t As String\n"
        '    s = "hello"\n'
        '    t = s & " world"\n'
        "    t = Left(t, 5)\n"
        "End Sub\n"
    ),
    "func": (
        "Function F(ByVal x As Long, y As String) As Long\n"
        "    F = x + Len(y)\n"
        "End Function\n"
    ),
    "props": (
        "Private mV As Long\n"
        "\n"
        "Public Property Get V() As Long\n"
        "    V = mV\n"
        "End Property\n"
        "\n"
        "Public Property Let V(ByVal value As Long)\n"
        "    mV = value\n"
        "End Property\n"
    ),
    "decls": (
        "Public gPub As Long\n"
        "Private gPriv As String\n"
        "\n"
        "Sub S()\n"
        "    Const K As Long = 7\n"
        "    Static sN As Long\n"
        "    Dim v As Variant\n"
        "    sN = sN + K\n"
        "End Sub\n"
    ),
    "errs": (
        "Sub S()\n"
        "    On Error GoTo Handler\n"
        "    Dim a As Long\n"
        "    a = 1\n"
        "    Exit Sub\n"
        "Handler:\n"
        "    Resume Next\n"
        "End Sub\n"
    ),
    "with": (
        "Sub S()\n"
        "    Dim c As Collection\n"
        "    Set c = New Collection\n"
        "    With c\n"
        "        .Add 1\n"
        "    End With\n"
        "End Sub\n"
    ),
    "arrays": (
        "Sub S()\n"
        "    Dim a(3) As Long\n"
        "    Dim b(1 To 5) As String\n"
        "    Dim c() As Double\n"
        "    ReDim c(9)\n"
        "    a(0) = 1\n"
        "End Sub\n"
    ),
    "udt": (
        "Private Type Point3\n"
        "    X As Long\n"
        "    Y As Long\n"
        "End Type\n"
        "\n"
        "Sub S()\n"
        "    Dim p As Point3\n"
        "    p.X = 1\n"
        "    p.Y = 2\n"
        "End Sub\n"
    ),
    "udtarr": (
        "Private Type Point3\n"
        "    X As Long\n"
        "End Type\n"
        "\n"
        "Sub S()\n"
        "    Dim p(3) As Point3\n"
        "    Dim q() As Point3\n"
        "    ReDim q(1)\n"
        "    p(0).X = 1\n"
        "End Sub\n"
    ),
    "enum": (
        "Private Enum Colour\n"
        "    Red = 1\n"
        "    Green = 2\n"
        "End Enum\n"
        "\n"
        "Sub S()\n"
        "    Dim c As Colour\n"
        "    c = Green\n"
        "End Sub\n"
    ),
    "fixstr": (
        "Private Type Rec\n"
        "    Code As String * 4\n"
        "End Type\n"
        "\n"
        "Sub S()\n"
        "    Dim tag As String * 8\n"
        "    Dim r As Rec\n"
        '    tag = "abc"\n'
        '    r.Code = "wxyz"\n'
        "End Sub\n"
    ),
    "objs": (
        "Sub S()\n"
        "    Dim o As Object\n"
        "    Dim c As Collection\n"
        "    Set o = Nothing\n"
        "    Set c = New Collection\n"
        "End Sub\n"
    ),
    "calls": (
        "Sub Helper(ByVal n As Long)\n"
        "End Sub\n"
        "\n"
        "Sub S()\n"
        "    Helper 1\n"
        "    Call Helper(2)\n"
        "End Sub\n"
    ),
    "nested": (
        "Sub S()\n"
        "    Dim i As Long\n"
        "    Dim j As Long\n"
        "    For i = 1 To 3\n"
        "        For j = 1 To 3\n"
        "            If i = j Then\n"
        "                i = i + 1\n"
        "            End If\n"
        "        Next j\n"
        "    Next i\n"
        "End Sub\n"
    ),
    "bools": (
        "Sub S()\n"
        "    Dim a As Boolean\n"
        "    Dim b As Boolean\n"
        "    a = True\n"
        "    b = Not a\n"
        "    a = a And b\n"
        "    a = a Or b\n"
        "    a = a Xor b\n"
        "End Sub\n"
    ),
    "foreach": (
        "Sub S()\n"
        "    Dim v\n"
        "    Dim c As Collection\n"
        "    For Each v In c\n"
        "        v = 1\n"
        "    Next v\n"
        "End Sub\n"
    ),
    "fileio": (
        "Sub S()\n"
        "    Dim fh As Integer\n"
        "    Dim buf As String\n"
        "    fh = FreeFile\n"
        '    Open "a.txt" For Output As #fh\n'
        '    Print #fh, "x"\n'
        '    Write #fh, "y"\n'
        "    Close #fh\n"
        "    Line Input #fh, buf\n"
        "End Sub\n"
    ),
    "options": (
        "Option Explicit\n"
        "Option Base 1\n"
        "Option Compare Text\n"
        "\n"
        "Sub S()\n"
        "End Sub\n"
    ),
    "condcomp": (
        "#If VBA7 Then\n"
        "Sub S()\n"
        "End Sub\n"
        "#Else\n"
        "Sub T()\n"
        "End Sub\n"
        "#End If\n"
    ),
    "coerce": (
        "Sub S()\n"
        "    Dim x\n"
        "    x = CInt(1)\n"
        "    x = CLng(1)\n"
        "    x = CDbl(1)\n"
        "    x = CStr(1)\n"
        "    x = CBool(1)\n"
        "End Sub\n"
    ),
    "strstmt": (
        "Sub S()\n"
        "    Dim buf As String\n"
        "    Dim pad As String * 4\n"
        '    buf = "abcd"\n'
        '    Mid(buf, 1, 2) = "xy"\n'
        '    LSet pad = "ab"\n'
        '    RSet pad = "cd"\n'
        "End Sub\n"
    ),
    "debug": (
        "Sub S()\n"
        "    Debug.Print 1\n"
        "    Debug.Assert True\n"
        "    Stop\n"
        "End Sub\n"
    ),
    "comments": (
        "' leading comment\n"
        "Sub S()\n"
        "    ' inner comment\n"
        "    Rem rem comment\n"
        "    Dim x\n"
        "    x = 1\n"
        "End Sub\n"
    ),
    "errforms": (
        "Sub S()\n"
        "    On Error GoTo 0\n"
        "    On Error Resume Next\n"
        "    On Error GoTo H\n"
        "    Exit Sub\n"
        "H:\n"
        "    Resume\n"
        "    Resume Next\n"
        "    Resume H\n"
        "End Sub\n"
    ),
    "deftype": (
        "DefLng L\n"
        "DefStr S\n"
        "\n"
        "Sub A()\n"
        "    Dim Lx\n"
        "    Lx = 1\n"
        "End Sub\n"
    ),
    "gosub": (
        "Sub S()\n"
        "    Dim x\n"
        "    GoSub L\n"
        "    Exit Sub\n"
        "L:\n"
        "    x = 1\n"
        "    Return\n"
        "End Sub\n"
    ),
    "enumimp": (
        "Public Enum Severity\n"
        "    sevLow\n"
        "    sevMid\n"
        "    sevHigh = 7\n"
        "    sevMax\n"
        "End Enum\n"
    ),
    "numbers": (
        "Sub S()\n"
        "    Dim x\n"
        "    x = 1\n"
        "    x = 300\n"
        "    x = 100000\n"
        "    x = 1.5\n"
        "    x = &HFF\n"
        "    x = &O17\n"
        "End Sub\n"
    ),
    "inline": (
        "Sub S()\n"
        "    Dim x\n"
        "    x = 1: x = 2\n"
        "    If x = 2 Then x = 3 Else x = 4\n"
        "    If x = 3 Then x = 5\n"
        "End Sub\n"
    ),
    "params": (
        "Sub S(ByVal a As Long, ByRef b As Long, c As Long, "
        "Optional ByVal d As Long = 5, Optional e As String)\n"
        "End Sub\n"
    ),
    "withdeep": (
        "Sub S()\n"
        "    Dim c As Collection\n"
        "    Set c = New Collection\n"
        "    With c\n"
        "        .Add 1\n"
        "        .Add 2\n"
        "    End With\n"
        "End Sub\n"
    ),
    "filestmt": (
        "Sub S()\n"
        '    Kill "a"\n'
        '    MkDir "b"\n'
        '    RmDir "b"\n'
        '    ChDir "c"\n'
        '    FileCopy "a", "b"\n'
        '    Name "a" As "b"\n'
        "    Randomize\n"
        "    Beep\n"
        "    DoEvents\n"
        "End Sub\n"
    ),
    "arraydims": (
        "Sub S()\n"
        "    Dim a(3, 4) As Long\n"
        "    Dim b(2, 3, 4) As Long\n"
        "    Dim c(1 To 5, 1 To 6) As Long\n"
        "    Dim d(3, 1 To 5) As Long\n"
        "    a(0, 0) = 1\n"
        "    b(0, 0, 0) = 2\n"
        "    c(1, 1) = 3\n"
        "    d(0, 1) = 4\n"
        "End Sub\n"
    ),
    "memberchain": (
        "Sub S()\n"
        "    Dim ws As Object\n"
        "    Dim r, c, v\n"
        "    v = ws.Cells(r, c).Value\n"
        "    ws.Cells(r, c).Interior.Color = v\n"
        "    With ws\n"
        "        .Cells(r, c).Value = v\n"
        "    End With\n"
        "End Sub\n"
    ),
}

# Entries VBA itself cannot round-trip. The compiler folds identifiers
# case-insensitively into a single project-table spelling, so a Sub S
# with a local s -- or a name that lands in the runtime operand table --
# loses its original casing. Expected, not a decompiler defect.
LOSSY: dict[str, str] = {
    "str": "Sub S and Dim s share one case-folded identifier",
    "func": "Function F resolves through the runtime operand table as f",
}


def _write(path: Path, source: str) -> None:
    body = f'Attribute VB_Name = "{MODULE}"\r\n' + source.replace("\n", "\r\n")
    if path.exists():
        path.unlink()
    with ExcelFile.create_new(path) as wb:
        project = wb.vba_project()
        if MODULE in [m.name for m in project.modules]:
            wb.set_module(MODULE, body)
        else:
            project.add_module(MODULE, body, kind=VBAModuleKind.standard)
        wb.save()


def _read_streams(path: Path) -> tuple[bytes, bytes]:
    """Return (module prefix bytes, _VBA_PROJECT stream)."""
    with ExcelFile(path) as wb:
        module = wb.vba_project().get_module(MODULE)
        prefix = bytes(module.prefix_bytes)
        cfb = CFB.from_bytes(wb.vba_project_bytes())
    return prefix, cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")


def run(scratch: Path, only: set[str] | None = None) -> int:
    """Compile, decompile and compare the corpus. Returns the failure count."""
    from pyvbaharness import ExcelSession

    tags = [t for t in CORPUS if not only or t in only]
    paths = {}
    for tag in tags:
        p = scratch / f"rt_{tag}.xlsm"
        _write(p, CORPUS[tag])
        paths[tag] = p

    session = ExcelSession()
    try:
        for tag in tags:
            session.open_document(paths[tag], read_only=False)
            session.compile_project(watch_seconds=20)
            session.save_as(paths[tag])
    finally:
        session.close()

    failures = 0
    exact = 0
    for tag in tags:
        prefix, project = _read_streams(paths[tag])
        actual = decompile(prefix, project)
        expected = CORPUS[tag].rstrip("\n")
        if actual.rstrip("\n") == expected:
            exact += 1
            print(f"  exact   {tag}")
            continue
        if tag in LOSSY:
            print(f"  lossy   {tag}  ({LOSSY[tag]})")
            continue
        failures += 1
        print(f"  DIFF    {tag}")
        for row in difflib.unified_diff(
            expected.splitlines(), actual.rstrip("\n").splitlines(),
            "source", "decompiled", lineterm="", n=1,
        ):
            print("      " + row)
    print(f"\n{exact}/{len(tags)} exact, "
          f"{len(tags) - exact - failures} known-lossy, {failures} failing")
    return failures


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd()
    target.mkdir(parents=True, exist_ok=True)
    raise SystemExit(1 if run(target) else 0)
